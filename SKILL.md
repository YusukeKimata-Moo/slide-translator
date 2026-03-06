---
name: slide-translator
description: "Translate Japanese text in molecular biology research presentation slides (.pptx) to English. Triggers on: 'translate slides', 'translate presentation', 'スライド英訳', 'プレゼン翻訳', 'スライドの翻訳', or any request to translate a PPTX from Japanese to English for scientific presentations."
---

# Slide Translator Skill

Translates Japanese text in PPTX slides to English suitable for molecular biology research presentations. Forces Arial font on all translated text. Designed for **2 commands only** to minimize token consumption.

## Prerequisites

- **pptx skill** must be installed (uses its unpack/clean/pack scripts)

## Workflow

### Step 1: Extract Japanese text (1 command)

```bash
python scripts/extract_japanese.py <input.pptx> <work_dir> [--exclude-slides 1 2 ...]
```

- Unpacks the PPTX and scans all slides for Japanese text
- Saves structured data to `<work_dir>/japanese_texts.json`
- Prints a human-readable summary with context (preceding/following English text)
- Use `--exclude-slides` to skip title slides or other non-translatable slides

### Step 2: Create translations JSON

Read the extracted summary and create `translations.json` (flat mapping):

```json
{
  "日本語テキスト": "English translation",
  "細胞分裂の過程": "Process of cell division"
}
```

#### Translation guidelines (molecular biology)

- **Do NOT use literal word-for-word translations.** Instead, use natural, concise, and academic English expressions appropriate for molecular biology research presentations.
- **Line breaks (`\n`):** The extracted text may contain `\n` representing manual line breaks (`<a:br/>`) in the text box.
  - If the `\n` is just for visual wrapping of a single continuous sentence (e.g., `"Vacuoles\nexpanded"`), **you may remove the `\n` or change its position** to fit the flow of the English translation cleanly.
  - If the `\n` represents a genuine structural break, such as a bullet point or a distinct new line of thought, **you should keep the `\n`** in the English translation to preserve the layout.
- Use standard nomenclature for proteins, genes, and organelles
- Keep gene/protein names (e.g., Gene A, Protein B) as-is — they are already English
- Use appropriate abbreviations: WT (wild type), GFP (Green Fluorescent Protein), etc.
- Species names should not be italicized in the JSON — PowerPoint handles formatting

> [!IMPORTANT]
> **Context-dependent translation**: When the extract output shows `(after: Gene A)` for text like `は高発現していた`, **do NOT repeat the preceding English text** in the translation. Translate only the Japanese portion: `was highly expressed`.

### Step 3: Apply translations and repack (1 command)

```bash
python scripts/apply_translations.py <work_dir> translations.json <output.pptx> --original <input.pptx>
```

This single command:

1. Replaces Japanese text with English translations
2. Changes `lang="ja-JP"` → `lang="en-US"` on translated runs
3. **Forces Arial font** on all translated slides (replaces all Japanese fonts)
4. Cleans orphaned files and repacks to PPTX

### Step 4: User Review

Ask the user to review the generated output file to ensure the translation is acceptable and the layout is well-maintained.

### Step 5: Cleanup (1 command)

Once the user approves the result, delete the temporary unpacked directory to save storage space:

```powershell
Remove-Item -Recurse -Force <work_dir>
```

## Notes

- Speaker notes are **not** translated by default (extract script only reads slide text, not notes)
- The original PPTX is never modified — output goes to a new file
- `--validate false` is used during pack to avoid false positives from external video references
