---
name: slide-translator
description: "Translate Japanese molecular biology research slides in a local .pptx file into concise academic English, preserving scientific names, meaningful line breaks, and text-color mappings. Use for Japanese-to-English slide translation, not manuscript translation, English proofreading, or creating a new deck."
---

# Slide Translator Skill

Translates Japanese text in PPTX slides to English suitable for molecular biology research presentations. Forces Arial font on translated runs while preserving their other run formatting.

## Prerequisites

- **pptx skill** must be installed alongside this directory: scripts resolve `../pptx/scripts/` from the resolved skill path and use its unpack/clean/pack scripts.
- Use the host-configured Python executable; run examples from this skill root or resolve the script paths explicitly.
- Use a dedicated disposable `<work_dir>` for this run. Keep the original, output, and `translations.json` outside it: the apply script deletes non-package entries at its root before packing.

## Workflow

### Step 1: Extract Japanese text

```bash
python scripts/extract_japanese.py <input.pptx> <work_dir> [--exclude-slides 1 2 ...]
```

- Unpacks the PPTX and scans all slides for Japanese text
- Saves structured data to `<work_dir>/japanese_texts.json`
- Prints whole-paragraph text, including adjacent English runs and manual line breaks
- `--exclude-slides` skips extraction for the numbered `slideN.xml` files, not necessarily presentation-order slide numbers. Check the slide mapping first.
- Apply has no exclusion filter: a translation key is replaced wherever it matches across slides. Before applying, check for keys also present on excluded slides or needing different translations in different contexts. For such collisions, use the `pptx` skill's targeted XML editing workflow on only the intended slides instead.

### Step 2: Draft and Check Translations

Do NOT create the JSON immediately. First, draft the English translations based on the extracted text and **evaluate** if each translation sounds like a natural, concise expression appropriate for an English molecular biology research presentation.

- Does it sound like a literal word-for-word translation? (If yes, rephrase it to be more natural)
- Is it too wordy for a presentation slide? (If yes, make it concise)
- Are the structural line breaks preserved properly?

After confirming the quality of the drafted translations, create UTF-8 `translations.json` outside `<work_dir>` (flat mapping). Copy keys exactly from `japanese_texts.json`, including embedded English and `\n`; translate the complete keyed paragraph:

```json
{
  "日本語テキスト": "English translation",
  "細胞分裂の過程": "Process of cell division"
}
```

#### Preserving text color for specific words

If a Japanese sentence contains words with specific colors and their order changes in English, you can explicitly map the original Japanese word (`src`) to the translated English word (`en`) using an array of objects. This ensures the English word inherits the correct color from the original Japanese word.

```json
{
  "微小管重合阻害時の色素体": [
    { "en": "Plastids under ", "src": "色素体" },
    { "en": "microtubule", "src": "微小管" },
    { "en": " polymerization inhibition", "src": "重合阻害時" }
  ]
}
```

_Note: Only use this explicit array format when preserving specific colors is necessary. For normal text, use the simple string format._

#### Translation guidelines (molecular biology)

- **Do NOT use literal word-for-word translations.** Instead, use natural, concise, and academic English expressions appropriate for molecular biology research presentations.
- **Line breaks (`\n`):** The extracted text may contain `\n` representing manual line breaks (`<a:br/>`) in the text box. Since the source text is Japanese, judge whether a line break is necessary based on the meaning and context of the **Japanese original**.
  - If the `\n` is just for visual wrapping within a single continuous phrase in Japanese (e.g., `"分裂前には\n核よりも上に移動"` — one continuous thought split for box width), **you may remove the `\n` or reposition it** to fit the English translation naturally (e.g., `"Moved above the nucleus\nbefore division"` or `"Moved above the nucleus before division"`).
  - If the `\n` separates structurally distinct lines in Japanese (e.g., `"微小管\n核"` — two separate items listed), **keep the `\n`** in the English translation to preserve the layout (e.g., `"Microtubules\nNucleus"`).
- **Capitalization after removed line breaks:** When you remove a `\n` from the Japanese source, the word that followed the line break should **NOT** be capitalized unless it is a proper noun or the start of a sentence. The apply script preserves text exactly as written in translations.json — do NOT capitalize mid-sentence words.
- Use standard nomenclature for proteins, genes, and organelles
- Keep gene/protein names (e.g., Gene A, Protein B) as-is — they are already English
- Use appropriate abbreviations: WT (wild type), GFP (Green Fluorescent Protein), etc.
- Species names should not be italicized in the JSON — PowerPoint handles formatting

> [!IMPORTANT]
> **Translation unit**: The current extractor includes English within the same paragraph in the key. For `Gene Aは高発現していた`, output `Gene A was highly expressed`; omitting `Gene A` would delete it. Do not add context from a separate paragraph or text box that is not part of the key.

### Step 3: Apply translations and repack

```bash
python scripts/apply_translations.py <work_dir> translations.json <output.pptx> --original <input.pptx>
```

This single command:

1. Replaces Japanese text with English translations
2. Changes `lang="ja-JP"` → `lang="en-US"` on translated runs
3. **Forces Arial font** on translated runs (does not replace every font on the slide)
4. Cleans orphaned files and repacks to PPTX

### Step 4: Verify and Request User Review

Compare output text with the mapping: check missing replacements, retained English, scientific names, meaningful line breaks, and excluded slides. The apply script handles `<p:txBody>` paragraphs; text extracted from other structures, such as table cells, can remain unchanged. Use targeted PPTX editing for those cases rather than reporting complete translation.

Render the output with the `pptx` skill's QA workflow and inspect translated slides for overflow, color mapping, and layout changes. Packing uses `--validate false` for external-video false positives, so successful packing is not proof of schema or visual validity. Report any unavailable checks and ask the user to review the output.

### Step 5: Cleanup

Only after the user approves the result and cleanup, verify that `<work_dir>` resolves to this run's disposable unpacked directory, then delete that exact directory:

```powershell
Remove-Item -LiteralPath <work_dir> -Recurse -Force
```

## Notes

- Speaker notes are **not** translated by default (extract script only reads slide text, not notes)
- The original PPTX is never modified — output goes to a new file
- `--validate false` is used during pack to avoid false positives from external video references
