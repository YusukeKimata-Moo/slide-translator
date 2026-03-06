"""Apply translations to PPTX slides, force Arial font, and repack.
Usage: python apply_translations.py <work_dir> <translations.json> <output.pptx> --original <original.pptx>

translations.json format: {"日本語テキスト": "English translation", ...}
"""
import sys, os, re, json, pathlib, subprocess, argparse

JP_RE = re.compile(r'[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uffef]')
ARIAL = 'typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"'

PARA_RE = re.compile(r'(<a:p[^>]*>)(.*?)(</a:p>)', re.DOTALL)
AT_RE = re.compile(r'<a:t[^>]*>([^<]*)</a:t>')
RUN_BR_RE = re.compile(r'(<a:r>.*?</a:r>|<a:br[^>]*/>)', re.DOTALL)

def find_pptx_scripts():
    sd = pathlib.Path(__file__).resolve().parent.parent
    pd = sd.parent / "pptx" / "scripts"
    return pd / "clean.py", pd / "office" / "pack.py"

def escape_xml(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')

def update_rpr_to_arial_en(rpr):
    # Ensure lang="en-US"
    rpr = re.sub(r'lang="[^"]*"', 'lang="en-US"', rpr)
    # Remove japanese-specific kumimoji if present
    rpr = re.sub(r'\s*kumimoji="1"', '', rpr)
    
    # Strip any existing typeface/fonts
    rpr = re.sub(r'<a:latin[^>]*/>', '', rpr)
    rpr = re.sub(r'<a:ea[^>]*/>', '', rpr)
    rpr = re.sub(r'<a:cs[^>]*/>', '', rpr)
    
    if rpr.endswith('/>'):
        rpr = rpr[:-2] + f'>\n                <a:latin {ARIAL}/>\n                <a:ea {ARIAL}/>\n              </a:rPr>'
    else:
        rpr = re.sub(r'(</a:rPr>)', f'<a:latin {ARIAL}/>\n                <a:ea {ARIAL}/>\n              \\1', rpr)
        
    if 'lang=' not in rpr:
        rpr = rpr.replace('<a:rPr', '<a:rPr lang="en-US"')
        
    return rpr

def apply_paragraph_translations(content, translations):
    """Replace entire paragraphs if their extracted text matches translations.json"""
    
    def replace_para(match):
        p_open = match.group(1)
        p_inner = match.group(2)
        p_close = match.group(3)
        
        elements = RUN_BR_RE.findall(p_inner)
        text_parts = []
        for el in elements:
            if '<a:br' in el:
                text_parts.append('\n')
            else:
                for t in AT_RE.findall(el):
                    text_parts.append(t)
                    
        full_text = "".join(text_parts).strip()
        
        if full_text in translations:
            translated = translations[full_text]
            
            # Extract pPr
            pPr_match = re.match(r'^(\s*<a:pPr.*?</a:pPr>\s*|\s*<a:pPr[^>]*/>\s*)', p_inner, re.DOTALL)
            pPr_str = pPr_match.group(1) if pPr_match else ''
            
            # Extract endParaRPr
            endPr_match = re.search(r'(\s*<a:endParaRPr.*?>\s*|\s*<a:endParaRPr[^>]*/>\s*)$', p_inner, re.DOTALL)
            endPr_str = endPr_match.group(1) if endPr_match else ''
            if endPr_str:
                endPr_str = endPr_str.replace('lang="ja-JP"', 'lang="en-US"')
            
            # Extract rPr template from the first run
            rPr_match = re.search(r'(<a:rPr.*?</a:rPr>|<a:rPr[^>]*/>)', p_inner, re.DOTALL)
            rPr_str = rPr_match.group(1) if rPr_match else '<a:rPr lang="en-US" dirty="0"/>'
            
            new_rPr = update_rpr_to_arial_en(rPr_str)
            
            # Build new runs
            new_runs = []
            lines = translated.split('\n')
            for line in lines:
                run = f'<a:r>\n              {new_rPr}\n              <a:t>{escape_xml(line)}</a:t>\n            </a:r>'
                new_runs.append(run)
                
            runs_str = '\n            <a:br/>\n            '.join(new_runs)
            
            # Reconstruct paragraph
            return p_open + '\n            ' + pPr_str + runs_str + '\n            ' + endPr_str + p_close
            
        return match.group(0)

    return PARA_RE.sub(replace_para, content)

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

    clean_py, pack_py = find_pptx_scripts()
    print("Cleaning...")
    subprocess.run([sys.executable, str(clean_py), a.work_dir], check=True)
    print("Packing...")
    subprocess.run([sys.executable, str(pack_py), a.work_dir, a.output_pptx,
                    "--original", a.original, "--validate", "false"], check=True)
    print(f"\nDone! Output: {a.output_pptx}")

if __name__ == "__main__":
    main()
