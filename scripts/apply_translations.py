"""Apply translations to PPTX slides, force Arial font, and repack.
Usage: python apply_translations.py <work_dir> <translations.json> <output.pptx> --original <original.pptx>

translations.json format: {"Japanese text": "English translation", ...}
"""
import sys, os, re, json, pathlib, subprocess, argparse

JP_RE = re.compile(r'[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uffef]')
ARIAL = 'typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"'

TXBODY_RE = re.compile(r'(<p:txBody>)(.*?)(</p:txBody>)', re.DOTALL)
PARA_RE   = re.compile(r'(<a:p[^>]*>)(.*?)(</a:p>)', re.DOTALL)
AT_RE     = re.compile(r'<a:t[^>]*>([^<]*)</a:t>')
RUN_BR_RE = re.compile(r'(<a:r>.*?</a:r>|<a:br[^>]*/\s*>)', re.DOTALL)
RPR_RE    = re.compile(r'(<a:rPr.*?</a:rPr>|<a:rPr[^>]*/\s*>)', re.DOTALL)

def find_pptx_scripts():
    sd = pathlib.Path(__file__).resolve().parent.parent
    pd = sd.parent / "pptx" / "scripts"
    return pd / "clean.py", pd / "office" / "pack.py"

def escape_xml(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')

def update_rpr_to_arial_en(rpr):
    """Update rPr: set lang=en-US, force Arial, remove Japanese-only attrs.
    Preserves bold, italic, color, size, underline, and other formatting attrs."""
    # Ensure lang="en-US"
    if 'lang=' in rpr:
        rpr = re.sub(r'lang="[^"]*"', 'lang="en-US"', rpr)
    else:
        rpr = rpr.replace('<a:rPr', '<a:rPr lang="en-US"')

    # Remove japanese-specific kumimoji if present
    rpr = re.sub(r'\s*kumimoji="1"', '', rpr)

    # Strip any existing typeface/font elements
    rpr = re.sub(r'<a:latin[^>]*/>', '', rpr)
    rpr = re.sub(r'<a:ea[^>]*/>', '', rpr)
    rpr = re.sub(r'<a:cs[^>]*/>', '', rpr)
    rpr = re.sub(r'<a:sym[^>]*/>', '', rpr)

    # Insert Arial latin/ea fonts
    if rpr.endswith('/>'):
        rpr = rpr[:-2] + f'>\n                <a:latin {ARIAL}/>\n                <a:ea {ARIAL}/>\n              </a:rPr>'
    else:
        rpr = re.sub(r'(</a:rPr>)', f'<a:latin {ARIAL}/>\n                <a:ea {ARIAL}/>\n              \\1', rpr)

    return rpr


def extract_run_elements(p_inner):
    """Extract list of run/br elements with their rPr from paragraph inner XML.
    Returns list of dicts: {'type': 'text'|'br', 'text': str, 'rpr': str}
    """
    elements = RUN_BR_RE.findall(p_inner)
    result = []
    last_rpr = None
    for el in elements:
        if '<a:br' in el:
            result.append({'type': 'br', 'text': '\n', 'rpr': last_rpr})
        else:
            rpr_m = RPR_RE.search(el)
            rpr = rpr_m.group(1) if rpr_m else None
            if rpr:
                last_rpr = rpr
            texts = AT_RE.findall(el)
            for t in texts:
                result.append({'type': 'text', 'text': t, 'rpr': rpr or last_rpr or '<a:rPr lang="en-US" dirty="0"/>'})
    return result


def apply_paragraph_translations(content, translations):
    """Replace entire paragraphs if their extracted text matches translations.json.
    Only processes paragraphs inside <p:txBody> blocks to avoid corrupting other XML.
    Preserves per-run formatting (bold, italic, color, size, etc.) by using the
    source run rPr proportionally mapped to the translated text position."""

    def replace_para(match):
        p_open  = match.group(1)
        p_inner = match.group(2)
        p_close = match.group(3)

        runs = extract_run_elements(p_inner)
        full_text = ''.join(r['text'] for r in runs).strip()

        if full_text not in translations:
            return match.group(0)

        translated = translations[full_text]

        # Extract pPr (paragraph properties)
        pPr_match = re.match(r'^(\s*<a:pPr.*?</a:pPr>\s*|\s*<a:pPr[^>]*/?>\s*)', p_inner, re.DOTALL)
        pPr_str = pPr_match.group(1) if pPr_match else ''

        # Extract endParaRPr
        endPr_match = re.search(r'(\s*<a:endParaRPr.*?>\s*|\s*<a:endParaRPr[^>]*/?>\s*)$', p_inner, re.DOTALL)
        endPr_str = endPr_match.group(1) if endPr_match else ''
        if endPr_str:
            endPr_str = endPr_str.replace('lang="ja-JP"', 'lang="en-US"')

        # Build character-level rPr map from source runs (proportional mapping).
        # Each source character maps to its run's rPr so we can preserve per-run colors.
        text_runs = [r for r in runs if r['type'] == 'text']
        default_rpr = text_runs[0]['rpr'] if text_runs else '<a:rPr lang="en-US" dirty="0"/>'

        # Build aligned rPr list matching each char of full_text (stripped)
        raw_text = ''.join(r['text'] for r in runs)
        strip_offset = len(raw_text) - len(raw_text.lstrip())
        strip_end    = len(raw_text.rstrip())
        rpr_aligned = []
        pos = 0
        for r in runs:
            rpr = r['rpr'] or default_rpr
            for _ in r['text']:
                if strip_offset <= pos < strip_end:
                    rpr_aligned.append(rpr)
                pos += 1

        src_chars_count = len(rpr_aligned)
        
        # Build translated runs, splitting on '\n' → <a:br/>
        new_parts = []
        
        if isinstance(translated, list):
            if len(translated) > 0 and isinstance(translated[0], dict):
                # Explicit mapping mode: [{'en': '...', 'src': '...'}, ...]
                for chunk in translated:
                    en_text = chunk.get("en", "")
                    src_text = chunk.get("src", "")
                    
                    rpr = default_rpr
                    if src_text and src_text in raw_text:
                        idx = raw_text.find(src_text)
                        align_idx = idx - strip_offset
                        center = int(align_idx + len(src_text) / 2.0)
                        if rpr_aligned and 0 <= center < len(rpr_aligned):
                            rpr = rpr_aligned[center]
                        elif rpr_aligned and center >= len(rpr_aligned):
                            rpr = rpr_aligned[-1]
                        elif rpr_aligned and center < 0:
                            rpr = rpr_aligned[0]
                    
                    lines = en_text.split('\n')
                    for i, line in enumerate(lines):
                        if i > 0:
                            new_parts.append('<a:br/>')
                        if line:
                            new_rPr = update_rpr_to_arial_en(rpr)
                            new_parts.append(f'<a:r>\n              {new_rPr}\n              <a:t>{escape_xml(line)}</a:t>\n            </a:r>')
            else:
                # 1-to-1 mapping mode: `translated` is a list of strings matching `text_runs`
                # If counts don't match, fallback safely or just zip them.
                run_idx = 0
                for r in runs:
                    if r['type'] == 'br':
                        new_parts.append('<a:br/>')
                    else:
                        if run_idx < len(translated):
                            chunk = translated[run_idx]
                        else:
                            chunk = ""
                        run_idx += 1
                        
                        if chunk:
                            lines = chunk.split('\n')
                            for i, line in enumerate(lines):
                                if i > 0:
                                    new_parts.append('<a:br/>')
                                if line:
                                    new_rPr = update_rpr_to_arial_en(r['rpr'] or default_rpr)
                                    new_parts.append(f'<a:r>\n              {new_rPr}\n              <a:t>{escape_xml(line)}</a:t>\n            </a:r>')
        else:
            # Proportional mapping mode with word-boundary snapping
            translated_flat = translated.replace('\n', '')
            tgt_len = max(len(translated_flat), 1)

            def get_rpr_for_word(word_start, word_end):
                """Map a target word to a source rPr by its center position."""
                if not rpr_aligned:
                    return default_rpr
                center_idx = (word_start + word_end) / 2.0
                src_idx = min(int(center_idx * src_chars_count / tgt_len), src_chars_count - 1)
                return rpr_aligned[src_idx]

            lines = translated.split('\n')
            tgt_global = 0  # position in translated_flat
            for i, line in enumerate(lines):
                if i > 0:
                    new_parts.append('<a:br/>')
                if not line:
                    continue
                
                # Tokenize line by words/spaces to keep whole words in same color
                # We use regex to yield words and non-words
                tokens = [m.group(0) for m in re.finditer(r'\S+|\s+', line)]
                groups = [] # list of (rpr_str, text_chunk)
                cur_rpr = None
                cur_text = []
                
                for token in tokens:
                    tok_len = len(token)
                    t_rpr = get_rpr_for_word(tgt_global, tgt_global + tok_len)
                    if t_rpr != cur_rpr and cur_rpr is not None:
                        groups.append((cur_rpr, ''.join(cur_text)))
                        cur_text = [token]
                        cur_rpr = t_rpr
                    else:
                        cur_text.append(token)
                        cur_rpr = t_rpr
                    tgt_global += tok_len
                    
                if cur_text:
                    groups.append((cur_rpr, ''.join(cur_text)))
                    
                for rpr_str, chunk in groups:
                    new_rPr = update_rpr_to_arial_en(rpr_str)
                    new_parts.append(f'<a:r>\n              {new_rPr}\n              <a:t>{escape_xml(chunk)}</a:t>\n            </a:r>')

        runs_str = '\n            '.join(new_parts)

        return p_open + '\n            ' + pPr_str + runs_str + '\n            ' + endPr_str + p_close

    def replace_txbody(tb_match):
        tb_open  = tb_match.group(1)
        tb_inner = tb_match.group(2)
        tb_close = tb_match.group(3)
        return tb_open + PARA_RE.sub(replace_para, tb_inner) + tb_close

    return TXBODY_RE.sub(replace_txbody, content)


def fix_broken_rels(work_dir):
    """Remove broken Relationship entries (Target='NULL' or empty) from .rels files.
    These exist in some source PPTXs and cause PowerPoint repair dialogs."""
    import glob
    rels_pattern = os.path.join(work_dir, '**', '*.rels')
    fixed = 0
    for rels_path in glob.glob(rels_pattern, recursive=True):
        with open(rels_path, 'r', encoding='utf-8') as f:
            content = f.read()
        original = content
        # Remove Relationship elements with Target="NULL" or Target=""
        content = re.sub(r'<Relationship[^>]*Target="NULL"[^>]*/>\s*', '', content)
        content = re.sub(r'<Relationship[^>]*Target=""[^>]*/>\s*', '', content)
        if content != original:
            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(content)
            fixed += 1
            print(f'Fixed broken rels: {os.path.relpath(rels_path, work_dir)}')
    if fixed:
        print(f'Fixed {fixed} .rels file(s) with broken relationships')


def main():
    p = argparse.ArgumentParser()
    p.add_argument("work_dir")
    p.add_argument("translations_json")
    p.add_argument("output_pptx")
    p.add_argument("--original", required=True)
    a = p.parse_args()

    with open(a.translations_json, 'r', encoding='utf-8') as f:
        translations = json.load(f)

    slides_dir = os.path.join(a.work_dir, 'ppt', 'slides')
    modified = []

    for f in sorted(os.listdir(slides_dir)):
        if not (f.endswith('.xml') and f.startswith('slide')):
            continue
        path = os.path.join(slides_dir, f)
        with open(path, 'r', encoding='utf-8') as fh:
            content = fh.read()

        orig = content
        content = apply_paragraph_translations(content, translations)

        if content != orig:
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write(content)
            modified.append(f)
            print(f'Updated: {f}')

    print(f"\nModified {len(modified)} slides")

    # Fix broken relationships (Target="NULL" or Target="") in .rels files
    # These exist in some source PPTXs and cause PowerPoint repair dialogs
    fix_broken_rels(a.work_dir)

    clean_py, pack_py = find_pptx_scripts()
    print("Cleaning...")
    subprocess.run([sys.executable, str(clean_py), a.work_dir], check=True)
    print("Packing...")
    subprocess.run([sys.executable, str(pack_py), a.work_dir, a.output_pptx,
                    "--original", a.original, "--validate", "false"], check=True)
    print(f"\nDone! Output: {a.output_pptx}")

if __name__ == "__main__":
    main()
