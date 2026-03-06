"""Apply translations to PPTX slides, force Arial font, and repack.
Usage: python apply_translations.py <work_dir> <translations.json> <output.pptx> --original <original.pptx>

translations.json format: {"日本語テキスト": "English translation", ...}
"""
import sys, os, re, json, pathlib, subprocess, argparse

JP_RE = re.compile(r'[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uffef]')
ARIAL = 'typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"'

# Common Japanese font names to replace
JP_FONTS = [
    '\u30e1\u30a4\u30ea\u30aa','\u6e38\u30b4\u30b7\u30c3\u30af',
    '\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af','\uff2d\uff33 \u30b4\u30b7\u30c3\u30af',
    'HGP\u5275\u82f1\u89d2\uff7a\uff9e\uff7c\uff6f\uff78UB',
    'HG\u5275\u82f1\u89d2\uff7a\uff9e\uff7c\uff6f\uff78UB',
    '\uff2d\uff33 \u660e\u671d','\uff2d\uff33 \uff30\u660e\u671d',
    '\u6e38\u660e\u671d',
]

def find_pptx_scripts():
    sd = pathlib.Path(__file__).resolve().parent.parent
    pd = sd.parent / "pptx" / "scripts"
    return pd / "clean.py", pd / "office" / "pack.py"

def replace_jp_fonts(content):
    """Replace all Japanese font names with Arial."""
    for fn in JP_FONTS:
        pat = re.compile(
            re.escape(f'typeface="{fn}"')
            + r'(\s+panose="[^"]*")?(\s+pitchFamily="[^"]*")?(\s+charset="[^"]*")?'
        )
        content = pat.sub(ARIAL, content)
    # Catch any remaining CJK-named fonts
    content = re.sub(
        r'typeface="([^"]*[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uffef][^"]*)"'
        r'(\s+panose="[^"]*")?(\s+pitchFamily="[^"]*")?(\s+charset="[^"]*")?',
        ARIAL, content
    )
    return content

def update_lang(content):
    """Change lang=ja-JP to en-US for runs whose text is no longer Japanese."""
    def fix(m):
        if JP_RE.search(m.group(2)):
            return m.group(0)
        r = m.group(1).replace('lang="ja-JP"', 'lang="en-US"')
        r = re.sub(r'\s*kumimoji="1"', '', r)
        return r + m.group(2) + m.group(3)
    content = re.sub(
        r'(<a:rPr[^>]*lang="ja-JP"[^>]*/>\s*<a:t[^>]*>)([^<]*)(</a:t>)',
        fix, content
    )
    def fix2(m):
        if JP_RE.search(m.group(4)):
            return m.group(0)
        r = m.group(1).replace('lang="ja-JP"', 'lang="en-US"')
        r = re.sub(r'\s*kumimoji="1"', '', r)
        return r + m.group(2) + m.group(3) + m.group(4) + m.group(5)
    content = re.sub(
        r'(<a:rPr[^>]*lang="ja-JP"[^>]*>)(.*?)(</a:rPr>\s*<a:t[^>]*>)([^<]*)(</a:t>)',
        fix2, content, flags=re.DOTALL
    )
    return content

def add_arial_to_rpr(content):
    """Add Arial font spec to self-closing en-US rPr that lack <a:latin>."""
    def add_font(m):
        rpr = m.group(0)
        # Check if the next sibling already has latin font (look ahead in content)
        pos = m.end()
        next_chunk = content[pos:pos+200]
        if '<a:latin' in next_chunk.split('</a:rPr>')[0] if '</a:rPr>' in next_chunk else False:
            return rpr
        if rpr.endswith('/>'):
            return (rpr[:-2] + '>\n'
                    '                <a:latin ' + ARIAL + '/>\n'
                    '                <a:ea ' + ARIAL + '/>\n'
                    '              </a:rPr>')
        return rpr
    # Only target self-closing rPr with en-US that were formerly ja-JP (have dirty="0" etc)
    content = re.sub(r'<a:rPr[^>]*lang="en-US"[^>]*/>', add_font, content)
    return content

def main():
    p = argparse.ArgumentParser()
    p.add_argument("work_dir")
    p.add_argument("translations_json")
    p.add_argument("output_pptx")
    p.add_argument("--original", required=True)
    a = p.parse_args()

    with open(a.translations_json, 'r', encoding='utf-8') as f:
        translations = json.load(f)

    # Sort longest first to avoid partial matches
    sorted_t = sorted(translations.items(), key=lambda x: len(x[0]), reverse=True)
    slides_dir = os.path.join(a.work_dir, 'ppt', 'slides')
    modified = []

    for f in sorted(os.listdir(slides_dir)):
        if not (f.endswith('.xml') and f.startswith('slide')):
            continue
        path = os.path.join(slides_dir, f)
        with open(path, 'r', encoding='utf-8') as fh:
            content = fh.read()
        orig = content
        for jp, en in sorted_t:
            content = content.replace(f'<a:t>{jp}</a:t>', f'<a:t>{en}</a:t>')
        if content != orig:
            content = update_lang(content)
            content = replace_jp_fonts(content)
            content = add_arial_to_rpr(content)
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
