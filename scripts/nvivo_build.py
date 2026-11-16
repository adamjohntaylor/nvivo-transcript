#!/usr/bin/env python3
"""nvivo_build.py - STEP 3 of the transcript pipeline: turn anonymised Zoom
.vtt transcripts into NVivo-ready Word transcripts and corpus text.

    nvivo_build.py prepare <anon.vtt|folder> [--work DIR] [--chunk-words 2000]
    nvivo_build.py build   <anon.vtt|folder> [--work DIR] [--out DIR] [--strip "A; B"] [--keep-all]
                           [--mechanical] [--check]
    nvivo_build.py twin    --manifest <_anonymisation/audit/manifest.json> --out <DIR> [--work DIR]

prepare  splits each transcript into numbered cues and writes chunk files
         (work/<name>/chunk-NN.md) for Claude to naturalise. Claude writes
         one rows-NN.json per chunk: [{"cue": 12, "content": "..."}, ...].
         A row starts at a cue and runs to the cue before the next row, so
         every timestamp comes from the transcript itself.
build    checks the rows (every cue covered, one speaker per row, times
         ascending and unique, bracketed anonymisation markers all kept,
         nothing large dropped or added, no new capitalised words) and, if
         they pass, writes <out>/<name>_NVivo.docx (Timestamp | Speaker |
         Content) and <out>/txt/<name>.txt (the corpus text: researchers'
         turns stripped, timestamps dropped, Unicode normalised). With
         --mechanical no AI is involved: rows are the speaker turns as Zoom
         gave them, split near 150 words, with doubled words removed.
         --check validates without writing.
twin     builds the IDENTIFIABLE twin of each built transcript from its raw
         .vtt (found through the anonymiser's manifest): the same rows and
         the same timestamps, raw text, mechanical clean-up only. This is
         what qdpx-media-sync (recordings) and qdpx-auto-code `transfer`
         need: two versions whose turns carry the same times. It prints
         counts only; never open its output with an AI.

Input must be the anonymised .vtt from anonymise_vtt.py (step 2). Needs
python-docx (pip install python-docx).
"""
import argparse
import glob
import json
import os
import re
import sys
import unicodedata
from collections import Counter

sys.dont_write_bytecode = True
VERSION = "2.0.0"

# --------------------------------------------------------------------------
# .vtt
# --------------------------------------------------------------------------
TIME_RE = re.compile(r"^\s*((?:\d+:)?\d{1,2}:\d{2}[.,]\d{1,3})\s*-->\s*((?:\d+:)?\d{1,2}:\d{2}[.,]\d{1,3})")
SPEAKER_RE = re.compile(r"^([^:\n]{1,80}?):\s+(.*)$", re.S)


def norm_time(t):
    t = t.replace(",", ".")
    hms, frac = t.split(".")
    parts = [int(p) for p in hms.split(":")]
    while len(parts) < 3:
        parts.insert(0, 0)
    return "%02d:%02d:%02d.%s" % (parts[0], parts[1], parts[2], (frac + "000")[:3])


def time_ms(t):
    h, m, s = t.split(":")
    sec, ms = s.split(".")
    return ((int(h) * 60 + int(m)) * 60 + int(sec)) * 1000 + int(ms)


def ms_time(ms):
    s, ms = divmod(int(ms), 1000)
    m, s = divmod(s, 60)
    h, m = divmod(m, 60)
    return "%02d:%02d:%02d.%03d" % (h, m, s, ms)


def looks_like_name(s):
    words = s.split()
    if not 1 <= len(words) <= 5 or re.search(r"[?!,;\"]", s):
        return False
    return all(w[:1].isupper() or w[:1] == "[" or re.match(r"^[\w]+[._][\w.]+$", w) or w.endswith("]")
               or w.lower() in ("van", "de", "von", "da", "-") for w in words)


def read_vtt(path):
    """Cues: {n, start, speaker, text}. Same rules as anon_table.read_vtt."""
    with open(path, encoding="utf-8-sig") as f:
        raw = f.read().replace("\r\n", "\n").replace("\r", "\n")
    cues = []
    for b in re.split(r"\n\s*\n", raw):
        lines = [l for l in b.split("\n") if l.strip()]
        ti = next((i for i, l in enumerate(lines) if TIME_RE.match(l)), None)
        if ti is None:
            continue
        cues.append({"start": norm_time(TIME_RE.match(lines[ti]).group(1)),
                     "text": " ".join(l.strip() for l in lines[ti + 1:])})
    prefixes = Counter(m.group(1).strip() for c in cues for m in [SPEAKER_RE.match(c["text"])] if m)
    prev = ""
    for i, c in enumerate(cues):
        c["n"] = i + 1
        m = SPEAKER_RE.match(c["text"])
        if m and (prefixes[m.group(1).strip()] >= 2 or looks_like_name(m.group(1).strip())):
            c["speaker"], c["text"] = m.group(1).strip(), m.group(2).strip()
        else:
            c["speaker"] = prev
        prev = c["speaker"]
    return cues


def inputs(target):
    if os.path.isdir(target):
        return sorted(glob.glob(os.path.join(target, "*_anon.vtt")))
    return [target]


def stem_of(path):
    return os.path.splitext(os.path.basename(path))[0]


# --------------------------------------------------------------------------
# text helpers
# --------------------------------------------------------------------------
BRACKET_RE = re.compile(r"\[[^\]\n]{1,80}\]")
TOKEN_RE = re.compile(r"[a-z0-9]+(?:'[a-z]+)?")


def tokens(s):
    return TOKEN_RE.findall(s.lower().replace("\u2019", "'"))


def destutter(s):
    s = re.sub(r"\b(\w+)(\s+\1\b)+", r"\1", s, flags=re.I)
    return " ".join(s.split())


def mechanical_text(texts):
    s = destutter(" ".join(t.strip() for t in texts if t.strip()))
    if s and s[-1] not in ".!?\"'":
        s += "."
    return s[:1].upper() + s[1:] if s else s


def corpus_normalise(s):
    s = unicodedata.normalize("NFKC", s)  # splits ligatures, keeps accents
    for a, b in (("\u2018", "'"), ("\u2019", "'"), ("\u201a", "'"), ("\u201b", "'"), ("\u201c", '"'),
                 ("\u201d", '"'), ("\u201e", '"'), ("\u00a0", " ")):
        s = s.replace(a, b)
    s = re.sub("[\u200b\u200c\u200d\u2060\ufeff]", "", s)
    return s


DEFAULT_STRIP = re.compile(r"^\[?(researcher\s*\d*|interviewer|moderator|facilitator)\b", re.I)


# --------------------------------------------------------------------------
# prepare
# --------------------------------------------------------------------------
def cmd_prepare(a):
    files = inputs(a.target)
    if not files:
        sys.exit("No *_anon.vtt files at %s (step 3 works on step 2's anonymised output)." % a.target)
    for path in files:
        cues = read_vtt(path)
        st = stem_of(path)
        wd = os.path.join(a.work, st)
        os.makedirs(wd, exist_ok=True)
        with open(os.path.join(wd, "cues.json"), "w", encoding="utf-8") as f:
            json.dump({"source": os.path.abspath(path), "cues": cues}, f, indent=0)
        # chunks end at a speaker change once they pass the word budget
        chunks, cur, words = [], [], 0
        for i, c in enumerate(cues):
            if cur and words >= a.chunk_words and c["speaker"] != cues[i - 1]["speaker"]:
                chunks.append(cur)
                cur, words = [], 0
            cur.append(c)
            words += len(c["text"].split())
        if cur:
            chunks.append(cur)
        for old in glob.glob(os.path.join(wd, "chunk-*.md")):
            os.remove(old)
        for k, ch in enumerate(chunks, 1):
            with open(os.path.join(wd, "chunk-%02d.md" % k), "w", encoding="utf-8") as f:
                f.write("# %s, chunk %d of %d: cues %d-%d\n\n" % (st, k, len(chunks), ch[0]["n"], ch[-1]["n"]))
                f.write("Write rows-%02d.json beside this file: a JSON list of {\"cue\": N, \"content\": \"...\"}. "
                        "The first row must start at cue %d. Each row runs from its cue to the cue before the next "
                        "row, and may not cross a change of speaker.\n\n" % (k, ch[0]["n"]))
                for c in ch:
                    f.write("[%d] %s %s: %s\n" % (c["n"], c["start"], c["speaker"], c["text"]))
        print("%s: %d cues, %d words, %d chunk(s) in %s" % (st, len(cues), sum(len(c["text"].split()) for c in cues),
                                                          len(chunks), wd))


# --------------------------------------------------------------------------
# build
# --------------------------------------------------------------------------
def mechanical_rows(cues, max_words=150):
    rows, cur, words = [], None, 0
    for c in cues:
        if cur is None or c["speaker"] != cues[cur[-1] - 1]["speaker"] or \
                (words >= max_words and re.search(r"[.?!]$", cues[cur[-1] - 1]["text"].strip())) or words >= 2 * max_words:
            cur = [c["n"]]
            rows.append(cur)
            words = 0
        else:
            cur.append(c["n"])
        words += len(c["text"].split())
    return [{"cue": r[0], "content": mechanical_text([cues[n - 1]["text"] for n in r])} for r in rows]


def load_rows(wd):
    files = sorted(glob.glob(os.path.join(wd, "rows-*.json")))
    rows = []
    for p in files:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            data = data.get("rows", [])
        rows.extend(data)
    return rows, files


def validate(cues, rows):
    """Return (rows with ranges and times, errors, warnings)."""
    errors, warnings = [], []
    n = len(cues)
    starts = []
    for i, r in enumerate(rows):
        try:
            starts.append(int(r["cue"]))
        except (KeyError, TypeError, ValueError):
            errors.append("row %d has no valid 'cue'" % (i + 1))
            return [], errors, warnings
    if not rows:
        return [], ["no rows"], warnings
    if starts[0] != 1:
        errors.append("the first row starts at cue %d, not cue 1: cues 1-%d are missing" % (starts[0], starts[0] - 1))
    for i in range(1, len(starts)):
        if starts[i] <= starts[i - 1]:
            errors.append("row %d starts at cue %d, not after cue %d" % (i + 1, starts[i], starts[i - 1]))
    if starts[-1] > n:
        errors.append("a row starts at cue %d; there are only %d cues" % (starts[-1], n))
    if errors:
        return [], errors, warnings
    out, prev_ms = [], -1
    for i, r in enumerate(rows):
        a, b = starts[i], (starts[i + 1] - 1 if i + 1 < len(starts) else n)
        span = cues[a - 1:b]
        speakers = {c["speaker"] for c in span}
        if len(speakers) > 1:
            errors.append("row at cue %d (%s) runs across a change of speaker (cues %d-%d)"
                          % (a, cues[a - 1]["start"], a, b))
        content = " ".join(str(r.get("content", "")).split())
        if not content:
            errors.append("row at cue %d is empty" % a)
        src = " ".join(c["text"] for c in span)
        need = Counter(BRACKET_RE.findall(src))
        have = Counter(BRACKET_RE.findall(content))
        lost = need - have
        if lost:
            errors.append("row at cue %d (%s) dropped anonymisation marker(s): %s"
                          % (a, cues[a - 1]["start"], ", ".join(sorted(lost))))
        st, ct = Counter(tokens(BRACKET_RE.sub(" ", src))), Counter(tokens(BRACKET_RE.sub(" ", content)))
        if sum(st.values()) >= 8:
            miss = sum((st - ct).values()) / sum(st.values())
            if miss > 0.25:
                warnings.append("row at cue %d (%s): %d%% of the spoken words are not in the row (omission?)"
                                % (a, cues[a - 1]["start"], round(100 * miss)))
        if sum(ct.values()) >= 8:
            add = sum((ct - st).values()) / sum(ct.values())
            if add > 0.25:
                warnings.append("row at cue %d (%s): %d%% of the row's words were not spoken (addition?)"
                                % (a, cues[a - 1]["start"], round(100 * add)))
        src_caps = {w.lower() for w in re.findall(r"\b[A-Za-z][\w'-]*", BRACKET_RE.sub(" ", src))}
        new_caps = sorted({w for w in re.findall(r"\b[A-Z][\w'-]+", BRACKET_RE.sub(" ", content))
                           if w.lower() not in src_caps and w not in ("I", "I'm", "I've", "I'll", "I'd", "UNCLEAR")})
        if new_caps:
            warnings.append("row at cue %d (%s): new capitalised word(s) not in the transcript: %s "
                            "(a correction must not introduce a name)" % (a, cues[a - 1]["start"], ", ".join(new_caps)))
        ms = time_ms(cues[a - 1]["start"])
        if ms < prev_ms + 100:
            new = prev_ms + 100
            if new - ms >= 1000:
                errors.append("row at cue %d: time %s cannot be made unique within 1 second" % (a, cues[a - 1]["start"]))
            else:
                warnings.append("row at cue %d: time moved %s -> %s to stay 0.1 s after the previous row"
                                % (a, cues[a - 1]["start"], ms_time(new)))
                ms = new
        prev_ms = ms
        out.append({"cue": a, "last": b, "time": ms_time(ms), "speaker": cues[a - 1]["speaker"], "content": content})
    return out, errors, warnings


def write_docx(path, rows):
    try:
        from docx import Document
        from docx.shared import Inches
    except ImportError:
        sys.exit("python-docx is not installed: pip install python-docx")
    doc = Document()
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for cell, h in zip(table.rows[0].cells, ("Timestamp", "Speaker", "Content")):
        cell.text = h
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
    for r in rows:
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = r["time"], r["speaker"], r["content"]
    for row in table.rows:
        for cell, w in zip(row.cells, (1.3, 1.5, 4.7)):
            cell.width = Inches(w)
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.save(path)


def write_corpus(path, rows, strip):
    keep = [corpus_normalise(r["content"]) for r in rows if not strip(r["speaker"])]
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n\n".join(keep) + "\n")
    return len(rows) - len(keep)


def make_strip(a):
    extra = {s.strip().lower() for s in (a.strip or "").replace(",", ";").split(";") if s.strip()}
    if a.keep_all:
        return lambda sp: False
    return lambda sp: bool(DEFAULT_STRIP.match(sp)) or sp.lower() in extra


def cmd_build(a):
    files = inputs(a.target)
    if not files:
        sys.exit("No *_anon.vtt files at %s." % a.target)
    strip = make_strip(a)
    bad = 0
    for path in files:
        st = stem_of(path)
        cues = read_vtt(path)
        wd = os.path.join(a.work, st)
        if a.mechanical:
            rows, src = mechanical_rows(cues), ["mechanical"]
        else:
            rows, src = load_rows(wd)
            if not src:
                print("%s: no rows-*.json in %s (run prepare, naturalise, then build; or use --mechanical)" % (st, wd))
                bad += 1
                continue
        built, errors, warnings = validate(cues, rows)
        print("%s: %d cues -> %d rows; %d error(s), %d warning(s)" % (st, len(cues), len(built) or len(rows),
                                                                      len(errors), len(warnings)))
        for e in errors:
            print("  ERROR   " + e)
        for w in warnings:
            print("  warning " + w)
        if errors:
            bad += 1
            continue
        if a.check:
            continue
        os.makedirs(a.out, exist_ok=True)
        os.makedirs(wd, exist_ok=True)
        with open(os.path.join(wd, "rows.final.json"), "w", encoding="utf-8") as f:
            json.dump({"source": os.path.abspath(path), "mechanical": a.mechanical,
                       "rows": [{"cue": r["cue"], "last": r["last"], "time": r["time"]} for r in built]}, f, indent=0)
        docx_path = os.path.join(a.out, st + "_NVivo.docx")
        write_docx(docx_path, built)
        stripped = write_corpus(os.path.join(a.out, "txt", st + ".txt"), built, strip)
        print("  wrote %s and txt/%s.txt (%d of %d rows stripped from the corpus text)"
              % (os.path.basename(docx_path), st, stripped, len(built)))
        if stripped == 0:
            print("  note: no researcher rows were stripped; pass --strip if researchers have other labels")
    if bad:
        sys.exit(1)


# --------------------------------------------------------------------------
# twin
# --------------------------------------------------------------------------
def cmd_twin(a):
    with open(a.manifest, encoding="utf-8") as f:
        manifest = json.load(f)
    os.makedirs(a.out, exist_ok=True)
    done = skipped = 0
    problems = Counter()
    for anon_name, info in sorted(manifest.items()):
        st = stem_of(anon_name)
        final = os.path.join(a.work, st, "rows.final.json")
        if not os.path.exists(final):
            problems["not built yet in step 3"] += 1
            skipped += 1
            continue
        with open(final, encoding="utf-8") as f:
            rows = json.load(f)["rows"]
        raw = info["raw"]
        if not os.path.exists(raw):
            problems["raw transcript not found"] += 1
            skipped += 1
            continue
        cues = read_vtt(raw)
        if len(cues) < rows[-1]["last"]:
            problems["raw and anonymised cue counts differ"] += 1
            skipped += 1
            continue
        out_rows, prev = [], -1
        mismatch = False
        for r in rows:
            c = cues[r["cue"] - 1]
            if abs(time_ms(c["start"]) - time_ms(r["time"])) >= 1000:
                mismatch = True
                break
            out_rows.append({"time": r["time"], "speaker": c["speaker"],
                             "content": mechanical_text([x["text"] for x in cues[r["cue"] - 1:r["last"]]])})
        if mismatch:
            problems["times differ between raw and anonymised"] += 1
            skipped += 1
            continue
        write_docx(os.path.join(a.out, stem_of(raw) + "_NVivo.docx"), out_rows)
        done += 1
    print("Identifiable twins written: %d; skipped: %d" % (done, skipped))
    for k, v in problems.items():
        print("  %s: %d" % (k, v))
    print("Folder (IDENTIFIABLE, never open with an AI): %s" % os.path.abspath(a.out))


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = p.add_subparsers(dest="cmd", required=True)
    pr = sub.add_parser("prepare")
    pr.add_argument("target", help="an *_anon.vtt file, or a folder of them")
    pr.add_argument("--work", default="_nvivo_work")
    pr.add_argument("--chunk-words", type=int, default=2000)
    b = sub.add_parser("build")
    b.add_argument("target")
    b.add_argument("--work", default="_nvivo_work")
    b.add_argument("--out", default="nvivo")
    b.add_argument("--strip", default="", help="extra speaker labels to leave out of the corpus text, ';'-separated")
    b.add_argument("--keep-all", action="store_true", help="strip no speakers from the corpus text")
    b.add_argument("--mechanical", action="store_true", help="no naturalisation: Zoom's turns, cleaned")
    b.add_argument("--check", action="store_true", help="validate only")
    t = sub.add_parser("twin")
    t.add_argument("--manifest", required=True)
    t.add_argument("--out", required=True)
    t.add_argument("--work", default="_nvivo_work")
    a = p.parse_args(argv)
    {"prepare": cmd_prepare, "build": cmd_build, "twin": cmd_twin}[a.cmd](a)


if __name__ == "__main__":
    main()
