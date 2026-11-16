---
name: NVivo-transcript
description: Converts Zoom transcripts into NVivo-compatible Word document format with timestamps, speakers, and naturalized content. Use when user says "NVivo transcript", "convert transcript for NVivo", "prepare transcript for NVivo", "transform Zoom transcript", "naturalise", or mentions importing transcripts into NVivo for qualitative analysis.
---

# NVivo Transcript Converter

Transform Zoom transcripts into naturalized, flowing conversation while maintaining strict timestamp accuracy and sequential ordering. Output as an MS Word document with a 3-column table suitable for direct import into NVivo.

## Input Format

The source Zoom transcript follows this structure:

```
Item_Number
Timestamp_Beginning --> Timestamp_End
Speaker: Content
```

## Output Format

### File Structure
- MS Word document (.docx)
- Filename: Original transcript filename with "_NVivo" suffix
- Single table with header row

### Table Structure
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:00.000 | Speaker A | This is the beginning of the conversation... |
| 00:00:15.340 | Speaker B | I agree with that point... |

### Table Requirements
- First row: Header row with column names (Timestamp, Speaker, Content)
- One row per turn-take or natural paragraph break
- Three columns per row

## Critical Requirements

### 1. Timestamp Rules (Absolute Priority)

**Sequential and Unique:**
- EVERY timestamp MUST be strictly sequential and ascending (each timestamp > previous timestamp)
- EVERY timestamp MUST be unique - no timestamp can appear more than once
- Timestamps must progress forward in time throughout the entire document
- The progression forward in time must be by at least 0.1 seconds
- If a progression from the previous timestamp is less than 0.1 seconds, make a small adjustment (less than 1.0 seconds)

**Timestamp Field Construction:**
- Extract the Timestamp_Beginning from the original transcript
- Reproduce it EXACTLY as it appears (everything before ` --> Timestamp_End`)
- Do not modify the timestamp format or values, with the only exception being a micro adjustment (less than 1.0 seconds) to ensure minimum 0.1 second progression
- Preserve all digits, colons, and decimal points exactly

**Synchronization:**
- The text in each row must synchronize accurately with its timestamp
- Minor adjustments allowed: You may move a few words to the previous timestamp row if it completes a sentence or idea
- If a sentence/idea cuts across two timestamps, place ALL words of that idea in the FIRST timestamp row
- This is the ONLY acceptable deviation from strict timestamp/text alignment

### 2. Speaker Field Rules

**Reproduction:**
- Copy the speaker name EXACTLY as it appears in the original transcript
- Preserve capitalization, spacing, and spelling
- No modifications unless specifically detailed below

**Consistency:**
- If a speaker's name appears in different formats (e.g., "John Smith" vs "John"), standardize to the most complete version
- Maintain consistency throughout the document

### 3. Content Transformation Rules

**YOU MUST DO:**
- Combine fragmented speech from the same speaker into natural, flowing paragraphs
- Correct obvious punctuation errors
- Fix stuttering by removing doubled words (e.g., "the the" → "the")
- Fix misspelled/misheard words when context makes the correction obvious (e.g., "definately compleated" → "definitely completed")
- Preserve the original word order exactly as spoken
- Complete the ENTIRE transcript - do not output in small portions requiring prompts for more

**YOU MUST NOT DO:**
- Change the order of words spoken
- Skip, omit, or summarize any part of the transcript
- Add words that weren't spoken

**Handling Unclear/Indecipherable Words:**

Use [SQUARE BRACKETS WITH CAPITALS] ONLY when:
- Words are truly indecipherable or nonsensical
- Acronyms that cannot be identified from context
- Abbreviations whose meaning is unclear
- Garbled audio where you cannot determine what was said

**Decision Tree for Word Handling:**
1. Is the word's intended meaning clear from context?
   - YES → Correct the spelling/grammar
   - NO → Is it an acronym, abbreviation, or truly indecipherable?
     - YES → Mark with [CAPS IN BRACKETS]
     - NO → Leave as-is and flag for review

### 4. Paragraph Breaking Strategy

When the same speaker has an extended turn:
- Break into natural paragraphs of approximately 100-150 words
- Place breaks at natural pause points (topic shifts, completed thoughts)
- Each new paragraph becomes a new table row
- Use the timestamp from where that paragraph's content actually begins in the original transcript
- Ensure timestamps remain sequential

**Combining Fragments:**
- You MAY combine sequential text from sequential timestamps under a single timestamp if it improves sentence structure, grammar, and punctuation
- Priority: Natural flow while maintaining timestamp accuracy

## Step-by-Step Process

### Step 1: Parse Input Data
Extract all transcript entries with structure:
- Item_Number
- Timestamp_Beginning (before ` --> `)
- Timestamp_End (after ` --> `)
- Speaker (before `:`)
- Content (after `:`)

### Step 2: Build Speaker Turns
- Group consecutive entries by speaker
- Track the timestamp_begin for each fragment
- Preserve the exact timestamp/content relationship

### Step 3: Naturalize Content
For each speaker turn:
- Combine fragments into flowing text
- Fix punctuation and grammar
- Remove stuttering
- Correct obvious misspellings using context
- Mark truly unclear words with [CAPITALS IN BRACKETS]
- If turn exceeds ~150 words, break into paragraphs
- Each paragraph gets timestamp from its actual starting point

### Step 4: Handle Cross-Timestamp Sentences
If a sentence/idea spans multiple timestamps:
- Place COMPLETE sentence in the FIRST timestamp row
- Adjust subsequent content accordingly
- Ensure no orphaned sentence fragments

### Step 5: Verify Sequential Order
CRITICAL VALIDATION:
1. All entries must be sorted by timestamp (ascending)
2. Verify: timestamp[i] > timestamp[i-1] for all i
3. Verify: no duplicate timestamps
4. If violations exist, adjust timestamps minimally to maintain sequence

### Step 6: Output Word Document
- Create Word document with filename: `[original_name]_NVivo.docx`
- Insert table with header row: Timestamp | Speaker | Content
- Add data rows with appropriate formatting

## Validation Checklist

Before providing final output, verify:
- [ ] All timestamps are strictly ascending (each > previous)
- [ ] No timestamp appears more than once
- [ ] Table has three columns: Timestamp | Speaker | Content
- [ ] Complete transcript (no content omitted)
- [ ] Unclear words marked with [CAPITALS IN BRACKETS] only when truly indecipherable
- [ ] Obvious misspellings corrected using context
- [ ] Natural paragraph flow while preserving original words
- [ ] Timestamp/text synchronization maintained (with only minor adjustments for sentence completion)

## Examples

### Example 1: Correct Timestamp Handling

**Input:**
```
1
00:00:00.000 --> 00:00:05.000
John: This is the beginning

2
00:00:05.500 --> 00:00:10.000
John: of my thought about

3
00:00:10.500 --> 00:00:15.000
John: the project timeline.
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:00.000 | John | This is the beginning of my thought about the project timeline. |

### Example 2: Cross-Timestamp Sentence

**Input:**
```
1
00:00:00.000 --> 00:00:03.000
Mary: We need to consider the budget implications

2
00:00:03.500 --> 00:00:07.000
Mary: before making any decisions.

3
00:00:07.500 --> 00:00:10.000
Mary: Let me explain why.
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:00.000 | Mary | We need to consider the budget implications before making any decisions. |
| 00:00:07.500 | Mary | Let me explain why. |

### Example 3a: Clear from Context (Correct the Error)

**Input:**
```
1
00:00:15.000 --> 00:00:20.000
Tom: The report was definately compleated last week
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:15.000 | Tom | The report was definitely completed last week. |

### Example 3b: Truly Unclear (Use Brackets)

**Input:**
```
1
00:00:15.000 --> 00:00:20.000
Tom: We implemented the new [garbled audio] system in March
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:15.000 | Tom | We implemented the new [UNCLEAR] system in March. |

### Example 4: Speaker Change

**Input:**
```
5
00:00:30.000 --> 00:00:35.000
Alice: I agree with that assessment completely.

6
00:00:36.000 --> 00:00:40.000
Bob: However, we should also consider alternative approaches.
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:00:30.000 | Alice | I agree with that assessment completely. |
| 00:00:36.000 | Bob | However, we should also consider alternative approaches. |

### Example 5: Long Speaker Turn with Paragraph Breaks

**Input:**
```
10
00:02:00.000 --> 00:02:10.000
Dr. Smith: The research findings indicate several important trends that we need to discuss.

11
00:02:10.500 --> 00:02:20.000
Dr. Smith: First, the data shows a significant correlation between variables A and B.

12
00:02:20.500 --> 00:02:30.000
Dr. Smith: Second, we observed unexpected patterns in the control group.

13
00:02:30.500 --> 00:02:40.000
Dr. Smith: These findings suggest we should revise our initial hypothesis.
```

**Output Table:**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:02:00.000 | Dr. Smith | The research findings indicate several important trends that we need to discuss. First, the data shows a significant correlation between variables A and B. Second, we observed unexpected patterns in the control group. |
| 00:02:30.500 | Dr. Smith | These findings suggest we should revise our initial hypothesis. |

## Common Errors to Avoid

**WRONG: Duplicate timestamps**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | First paragraph |
| 00:05:00.000 | Speaker A | Second paragraph |

**CORRECT: Unique, sequential timestamps**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | First paragraph |
| 00:05:12.450 | Speaker A | Second paragraph |

**WRONG: Non-sequential timestamps**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | Content |
| 00:07:30.000 | Speaker B | Content |
| 00:06:15.000 | Speaker A | Content (Goes backward!) |

**CORRECT: Always ascending**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | Content |
| 00:06:15.000 | Speaker A | Content |
| 00:07:30.000 | Speaker B | Content |

**WRONG: Breaking sentences across rows unnecessarily**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | We need to consider |
| 00:05:03.000 | Speaker A | the budget carefully. |

**CORRECT: Complete thought in first timestamp**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | We need to consider the budget carefully. |

**WRONG: Not correcting obvious misspellings**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | The report was definately compleated. |

**CORRECT: Fix obvious errors from context**
| Timestamp | Speaker | Content |
|-----------|---------|---------|
| 00:05:00.000 | Speaker A | The report was definitely completed. |

## Token Management

- Process as much of the transcript as your token window allows
- Do NOT stop and ask for permission to continue
- If you reach token limits, output what you've processed and indicate where you stopped
- User will prompt for continuation if needed

## Technical Implementation

### Required Python Dependencies
```
python-docx
```

### Correct Imports (use ONLY these)
```python
import re
from docx import Document
from docx.shared import Inches, Pt
```

**DO NOT import:**
- `Wd_Table_Alignment` (does not exist in python-docx)
- `Wd_Align_Paragraph` (not needed for this task)

### Reference Implementation

Use this tested Python script as the basis for conversion:

```python
import re
from docx import Document
from docx.shared import Inches, Pt

def parse_vtt(file_path):
    """Parse VTT file and extract entries."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    blocks = content.strip().split('\n\n')
    entries = []

    for block in blocks:
        lines = block.strip().split('\n')
        if len(lines) >= 2:
            for i, line in enumerate(lines):
                if '-->' in line:
                    timestamp_match = re.match(r'(\d{2}:\d{2}:\d{2}\.\d{3})\s*-->', line)
                    if timestamp_match:
                        timestamp = timestamp_match.group(1)
                        if i + 1 < len(lines):
                            content_line = lines[i + 1]
                            speaker_match = re.match(r'^([^:]+):\s*(.*)$', content_line)
                            if speaker_match:
                                speaker = speaker_match.group(1).strip()
                                text = speaker_match.group(2).strip()
                                entries.append({
                                    'timestamp': timestamp,
                                    'speaker': speaker,
                                    'content': text
                                })
                    break
    return entries

def naturalize_content(text):
    """Fix stuttering, punctuation, and obvious errors."""
    # Remove repeated words (stuttering)
    text = re.sub(r'\b(\w+)\s+\1\b', r'\1', text, flags=re.IGNORECASE)

    # Ensure proper sentence ending
    text = text.strip()
    if text and text[-1] not in '.!?':
        if len(text) > 10:
            text += '.'
    return text

def combine_speaker_turns(entries, max_words=150):
    """Combine consecutive entries from the same speaker."""
    if not entries:
        return []

    combined = []
    current = None

    for entry in entries:
        speaker = entry['speaker']

        if current is None:
            current = {
                'timestamp': entry['timestamp'],
                'speaker': speaker,
                'content': entry['content'],
                'word_count': len(entry['content'].split())
            }
        elif speaker == current['speaker']:
            current['content'] += ' ' + entry['content']
            current['word_count'] = len(current['content'].split())

            if current['word_count'] > max_words:
                combined.append(current)
                current = {
                    'timestamp': entry['timestamp'],
                    'speaker': speaker,
                    'content': '',
                    'word_count': 0
                }
        else:
            if current['content'].strip():
                combined.append(current)
            current = {
                'timestamp': entry['timestamp'],
                'speaker': speaker,
                'content': entry['content'],
                'word_count': len(entry['content'].split())
            }

    if current and current['content'].strip():
        combined.append(current)

    return combined

def ensure_unique_timestamps(entries):
    """Ensure all timestamps are unique and sequential."""
    seen_timestamps = set()
    result = []

    for entry in entries:
        timestamp = entry['timestamp']

        while timestamp in seen_timestamps:
            parts = timestamp.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds_parts = parts[2].split('.')
            seconds = int(seconds_parts[0])
            milliseconds = int(seconds_parts[1])

            milliseconds += 100
            if milliseconds >= 1000:
                milliseconds -= 1000
                seconds += 1
            if seconds >= 60:
                seconds -= 60
                minutes += 1
            if minutes >= 60:
                minutes -= 60
                hours += 1

            timestamp = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"

        seen_timestamps.add(timestamp)
        result.append({
            'timestamp': timestamp,
            'speaker': entry['speaker'],
            'content': naturalize_content(entry['content'])
        })

    return result

def create_nvivo_document(entries, output_path):
    """Create the NVivo-compatible Word document."""
    doc = Document()
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'

    header_cells = table.rows[0].cells
    header_cells[0].text = 'Timestamp'
    header_cells[1].text = 'Speaker'
    header_cells[2].text = 'Content'

    for cell in header_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    for entry in entries:
        row = table.add_row()
        row.cells[0].text = entry['timestamp']
        row.cells[1].text = entry['speaker']
        row.cells[2].text = entry['content']

    for row in table.rows:
        row.cells[0].width = Inches(1.5)
        row.cells[1].width = Inches(1.5)
        row.cells[2].width = Inches(4.5)

    doc.save(output_path)
```

### Cleanup Instructions

After creating the Word document:
- Delete any temporary Python scripts created during processing
- Use `rm` command (works on all platforms in Claude Code's bash environment)
- Do NOT use Windows-specific commands like `del`
