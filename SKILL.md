---
name: NVivo-transcript
description: Turns anonymised Zoom transcripts (the *_anon.vtt files from anonymise-transcripts) into NVivo-ready Word tables (Timestamp | Speaker | Content) with naturalised text, plus corpus .txt for AntConc/LancsBox and an identifiable twin with identical timestamps. Step 3 of the anonymise-first pipeline. Use when the user says "NVivo transcript", "convert transcript for NVivo", "prepare transcript for NVivo", "transform Zoom transcript", "naturalise", or mentions importing interview transcripts into NVivo.
---

# NVivo transcript (v2)

Step 3 of a three-step pipeline that takes Zoom interview transcripts to NVivo **without an AI ever reading identifiable text**. It starts from the anonymised `.vtt` that step 2 (`anonymise-transcripts`) writes. Claude naturalises the text; a script checks Claude's work and writes the Word transcript, so every timestamp comes from the transcript itself and cannot drift.

## Rules that do not bend

1. **Anonymised input only.** Work on `*_anon.vtt` files from step 2. If the user gives a raw Zoom `.vtt` (no `_anon` in the name, real names in the speaker labels), stop: say that the pipeline anonymises first and point to the `anonymise-transcripts` skill. Do not read the raw file to check.
2. **Never read identifiable files.** Not the raw `.vtt`, not anything in `_anonymisation/`, not the `identifiable/` twins that `twin` writes. `twin` prints counts only; do not inspect its output.
3. **Keep the anonymisation intact.** Every bracketed marker (`[Researcher 1]`, `[independent secondary school A]`, `[phone number redacted]`) stays exactly as it is, and pseudonyms stay as they are. Never "correct" a word into a proper noun that was not spoken: a misheard word that might be a name is left alone or marked `[UNCLEAR]`. The script refuses rows that drop a marker and warns on new capitalised words.
4. **Complete and in order.** Every cue is covered, nothing is summarised, no word order is changed.

## The pipeline end to end

```
raw/*.vtt  (Zoom transcripts: IDENTIFIABLE)
   |
   |  STEP 1  anon_table.py scan  -> _anonymisation/Anonymisation_Review.xlsx
   |          researcher fills it in (Excel, by hand)
   |          anon_table.py build -> _anonymisation/Anonymisation_Table.xlsx        NO AI
   |
   |  STEP 2  anonymise_vtt.py   -> anonymised/*_anon.vtt (same cues, same times)
   |                               _anonymisation/audit/*_audit.md, manifest.json   Claude runs the script,
   |  STEP 2b Claude reads anonymised/*_anon.vtt + _review/candidates.md             reads anonymised text only
   |          -> suggested rows -> researcher approves -> anon_table.py add -> re-run step 2
   |
   |  STEP 3  nvivo_build.py prepare -> Claude naturalises chunk by chunk -> rows-NN.json     (THIS SKILL)
   |          nvivo_build.py build   -> nvivo/*_anon_NVivo.docx (Timestamp | Speaker | Content)
   |                                    nvivo/txt/*_anon.txt (corpus text, researchers stripped)
   |          nvivo_build.py twin    -> identifiable/*_NVivo.docx (same rows, same times, real text)
   v
downstream: qdpx-augment builds the anonymised project from nvivo/ and the identifiable one from identifiable/;
qdpx-auto-code codes the anonymised project and `transfer`s codings by timestamp; qdpx-media-sync plays the
recordings against the identifiable site; AntConc / LancsBox read nvivo/txt/.
```

| Component | Where | Who runs it | Reads | Writes |
|---|---|---|---|---|
| `anon_table.py` (step 1) | `anonymise-transcripts/scripts/` | researcher, or Claude (prints counts only) | raw `.vtt`, table | review workbook, table |
| `anonymise_vtt.py` (step 2) | `anonymise-transcripts/scripts/` | Claude | raw `.vtt`, table | `anonymised/*_anon.vtt`, audit logs, manifest |
| review (step 2b) | `anonymise-transcripts` SKILL.md | Claude | anonymised text only | suggested table rows |
| `nvivo_build.py` (step 3) | `nvivo-transcript/scripts/` | Claude | `*_anon.vtt`, Claude's rows; `twin` also raw `.vtt` via the manifest | NVivo `.docx`, corpus `.txt`, identifiable twins |

How the parts connect:

- **Cues are the backbone.** Step 2 keeps every Zoom cue, one for one, with its time. Step 3 numbers the cues; Claude's rows say only which cue each row starts at. The script takes the time from that cue, so timestamps are exact by construction.
- **The twin replays the same rows over the raw text.** `build` saves the row structure (`_nvivo_work/<name>/rows.final.json`: cue ranges and times, no text). `twin` finds the raw file through step 2's `manifest.json` and writes the identifiable transcript with the same rows and times and the raw words (mechanically cleaned, not naturalised). The anonymised and identifiable projects built from `nvivo/` and `identifiable/` therefore match turn for turn, which is what `qdpx-auto-code transfer` and `qdpx-media-sync` need.
- **A gap in anonymisation is fixed upstream.** If you notice a name while naturalising, do not fix it in the rows: report it, have it added to the table (step 2b), re-run step 2, and run `prepare` again.

Folder layout: `raw/` and `_anonymisation/` and `identifiable/` are IDENTIFIABLE; `anonymised/`, `_nvivo_work/` and `nvivo/` are anonymised. See the `anonymise-transcripts` skill for the full layout and steps 1 and 2.

Needs Python 3.8+ with `python-docx` (`pip install python-docx`). On Windows type `python` where this file says `python3`. Run with `-B` so no `__pycache__` is left in the user's folders.

## Steps

### 1. Prepare

```bash
python3 -B <skill>/scripts/nvivo_build.py prepare anonymised --work _nvivo_work
```

For each `*_anon.vtt` it writes `_nvivo_work/<name>/chunk-01.md`, `chunk-02.md`... (about 2,000 words each, split at a change of speaker). Each line is one cue: `[12] 00:01:02.340 Hannah: text`.

### 2. Naturalise, chunk by chunk

Read one chunk in full and write `rows-NN.json` beside it (same number):

```json
[
  {"cue": 1, "content": "Hi Hannah, thanks for joining. I'm [Researcher 1] and this is Dr. [Researcher 2]. Can you tell us about your work at [independent secondary school A]?"},
  {"cue": 3, "content": "Sure. So I've been at [independent secondary school A] for six years; I started in [early 2020]."}
]
```

- A row **starts at a cue** and runs to the cue before the next row. The first row of a chunk starts at the chunk's first cue. A row never crosses a change of speaker. The speaker and time come from the start cue; do not write them.
- **One row per speaker turn**, or per natural paragraph in a long turn: break long turns at about 100-150 words, at a topic shift or completed thought, starting the new row at the cue where that paragraph begins.
- **A sentence that runs across cues goes whole into the row where it starts.** That is the only time words move between cues.

What naturalising means:

- **Do:** join fragments into flowing sentences; correct punctuation; remove stutters and doubled words ("the the" -> "the"); correct a misheard or misspelled word when context makes the intended word obvious ("definately compleated" -> "definitely completed"), unless the correction would be a name.
- **Do not:** change word order; skip, omit or summarise anything; add words that were not spoken; touch bracketed markers or pseudonyms.
- **Unclear words:** mark with `[UNCLEAR]` or a capitalised bracket only when a word is truly indecipherable, an acronym or abbreviation that cannot be identified, or garbled. Decision: is the meaning clear from context? Yes: correct it. No: is it an acronym, abbreviation or garble? Yes: bracket it. No: leave it as it is and mention it in your report.

Finish every chunk; do not stop to ask whether to continue. For many or long transcripts, give each transcript (or each chunk) to a subagent with this section as its brief, so the main context stays clear.

### 3. Check, then build

```bash
python3 -B <skill>/scripts/nvivo_build.py build anonymised --work _nvivo_work --out nvivo --check
```

**Errors** (nothing is written until they are fixed): a missing or out-of-order cue, a row across a change of speaker, an empty row, a dropped bracketed marker, a time that cannot be made unique. **Warnings** to read and judge: a row missing more than a quarter of its spoken words (omission?) or with more than a quarter new words (addition?), a new capitalised word not in the transcript (an introduced name?), and a time nudged forward (under a second) to stay 0.1 s after the previous row. Fix the rows and check again; then run without `--check`:

```bash
python3 -B <skill>/scripts/nvivo_build.py build anonymised --work _nvivo_work --out nvivo
```

It writes `nvivo/<name>_NVivo.docx` (header row Timestamp | Speaker | Content, one table, author fields cleared) and `nvivo/txt/<name>.txt`: the corpus text, with the researchers' turns left out (`[Researcher N]`, Interviewer, Moderator, Facilitator; add others with `--strip "A; B"`, or keep all with `--keep-all`), timestamps dropped, one paragraph per row, Unicode normalised (straight quotes, ligatures split, zero-width characters removed, accents kept), UTF-8, LF. If no rows were stripped it says so: the researchers may have other labels.

**No AI at all:** `build --mechanical` skips naturalising and uses Zoom's own turns, split near 150 words, with doubled words removed. Useful for a quick first look, or where the project does not allow even anonymised text to go to an AI.

### 4. Identifiable twin (when the project needs one)

Needed for `qdpx-media-sync` (recordings played against the transcript) and for `qdpx-auto-code transfer` (coding carried to the identifiable project).

```bash
python3 -B <skill>/scripts/nvivo_build.py twin --manifest _anonymisation/audit/manifest.json --out identifiable --work _nvivo_work
```

It prints only counts: twins written, and any skipped (not yet built in step 3, raw file moved, cue counts or times differing). Twin files are named after the raw transcript (`<raw name>_NVivo.docx`) so `qdpx-media-sync` can match them to recordings by name. Never open them.

### 5. Report

Files written, rows per file, any warnings you accepted and why, words you left or marked unclear, and anything that looked like a missed identifier (to go back to step 2b). Remind the user that `identifiable/` is identifiable data.

## Validation checklist

- [ ] Input was `*_anon.vtt` from step 2.
- [ ] Every chunk naturalised completely; no summarising, no reordering.
- [ ] `build --check` passed with no errors; every warning looked at.
- [ ] `nvivo/*_NVivo.docx` and `nvivo/txt/*.txt` written; stripped-row counts plausible.
- [ ] Twin built if the project needs one; its output not opened.
- [ ] Possible missed identifiers reported for step 2b, not fixed in the rows.

## Changelog

### v2.0.0 (November 2026)
- Works on anonymised `.vtt` from `anonymise-transcripts` v2 (anonymise first, then naturalise), so the AI never reads identifiable text.
- New `scripts/nvivo_build.py` (`prepare`, `build`, `twin`): Claude writes rows by cue number; the script takes times from the cues, checks completeness, speakers, markers, omissions and additions, and writes the `.docx` and corpus `.txt`. `--mechanical` builds without AI.
- `twin` writes an identifiable transcript with identical rows and times from the raw `.vtt`, for `qdpx-media-sync` and `qdpx-auto-code transfer`.
- The corpus `.txt` (researchers stripped) moves here from the anonymiser's old Mode B.

### v1
- Claude converted a Zoom `.vtt` directly to an NVivo `.docx`, writing the timestamps itself, from the reference script in this file.
