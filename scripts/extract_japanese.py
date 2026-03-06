"""Extract Japanese text from PPTX slides for translation.
Usage: python extract_japanese.py <input.pptx> [work_dir] [--exclude-slides 1 2 ...]
"""
import sys, os, re, json, pathlib, subprocess, argparse

JP_RE = re.compile(r'[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\uff00-\uffef]')
AT_RE = re.compile(r'<a:t[^>]*>([^<]*)</a:t>')
RUN_BR_RE = re.compile(r'(<a:r>.*?</a:r>|<a:br[^>]*/>)', re.DOTALL)
PARA_RE = re.compile(r'<a:p[^>]*>(.*?)</a:p>', re.DOTALL)

def find_unpack():
    skill_dir = pathlib.Path(__file__).resolve().parent.parent
    return skill_dir.parent / "pptx" / "scripts" / "office" / "unpack.py"

def extract(slides_dir, exclude):
    exclude = set(exclude or [])
    result = {}
    files = [f for f in os.listdir(slides_dir) if f.endswith('.xml') and f.startswith('slide')]
    files.sort(key=lambda x: int(re.search(r'\d+', x).group()))
    
    for f in files:
        num = int(re.search(r'\d+', f).group())
        if num in exclude:
            continue
        with open(os.path.join(slides_dir, f), 'r', encoding='utf-8') as fh:
            content = fh.read()
        if not JP_RE.search(content):
            continue
            
        entries, seen = [], set()
        
        for pm in PARA_RE.finditer(content):
            para_content = pm.group(1)
            elements = RUN_BR_RE.findall(para_content)
            
            # Combine text within a paragraph, replacing <a:br/> with \n
            combined_texts = []
            for el in elements:
                if '<a:br' in el:
                    combined_texts.append("\n")
                else:
                    for t in AT_RE.findall(el):
                        combined_texts.append(t)
                        
            # Re-split by \n to keep lines separate IF they don't form a single sentence, 
            # BUT for now, let's treat the whole paragraph combined with \n as a single translation unit 
            # if it contains Japanese.
            full_text = "".join(combined_texts).strip()
            
            if JP_RE.search(full_text) and full_text not in seen:
                seen.add(full_text)
                entries.append({"text": full_text})
                
        if entries:
            result[f"slide{num}"] = entries
            
    return result

def main():
    p = argparse.ArgumentParser()
    p.add_argument("input_pptx")
    p.add_argument("work_dir", nargs="?", default=None)
    p.add_argument("--exclude-slides", nargs="*", type=int, default=[])
    a = p.parse_args()
    
    if a.work_dir is None:
        a.work_dir = os.path.splitext(a.input_pptx)[0] + '_unpacked'
        
    unpack = find_unpack()
    if not unpack.exists():
        print(f"ERROR: pptx skill not found at {unpack}", file=sys.stderr); sys.exit(1)
        
    print(f"Unpacking {a.input_pptx}...")
    subprocess.run([sys.executable, str(unpack), a.input_pptx, a.work_dir], check=True)
    
    slides_dir = os.path.join(a.work_dir, 'ppt', 'slides')
    result = extract(slides_dir, a.exclude_slides)
    
    jp_path = os.path.join(a.work_dir, 'japanese_texts.json')
    with open(jp_path, 'w', encoding='utf-8') as f:
        # Use indent=2 and ensure_ascii=False for readable Japanese, but we need to handle \n properly in JSON
        json.dump(result, f, ensure_ascii=False, indent=2)
        
    total = sum(len(v) for v in result.values())
    print(f"\nFound {total} Japanese text items in {len(result)} slides\n")
    
    for slide, entries in result.items():
        print(f"--- {slide} ---")
        for e in entries:
            # Print with explicit \n for visibility
            text_disp = e['text'].replace('\n', '\\n')
            print(f"  {text_disp}")
            
    print(f"\nJSON: {jp_path}\nWork dir: {a.work_dir}")

if __name__ == "__main__":
    main()
