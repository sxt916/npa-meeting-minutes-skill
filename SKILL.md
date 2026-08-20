---
name: npa-meeting-minutes
description: Use when the user provides Chinese meeting raw materials, transcript text, audio-to-text records, or rough notes and asks to整理、生成、输出新力通/公司标准格式会议纪要或会议记录 Word 文件.
---

# NPA Meeting Minutes

## Purpose

Turn raw Chinese meeting transcripts or rough notes into a formal company-standard meeting minutes `.docx`.

Use this skill with `officecli` when the deliverable is a Word file.

## Inputs

Expected input is one or more of:

- raw meeting transcript text, usually with speakers and timestamps
- meeting notes or bullet points
- a reference/template meeting-minutes PDF or DOCX
- optional metadata: meeting name, date, place, host, attendees, copied departments

If key metadata is missing, use `待补充` rather than inventing names, places, or attendees.

## Output Standard

Create a formal `.docx` matching the company style:

- top header with 新力通 logo/contact strip
- centered title: `{会议主题}会议纪要`
- one fixed-width master table containing the meeting-info rows, the bordered body row, and the sign-off rows
- meeting info zone: `会议名称 / 地点 / 主持人 / 会议时间 / 参会人员`
- body inside one full-width bordered cell of the same master table
- numbered items using Chinese section headings
- final sign-off area: `编制： 审核： 批准：` and `抄送：相关部门`
- footer with company values strip: `诚信 责任 和谐 改善`

## Workflow

1. Read all raw materials completely. If there is a template file, inspect it visually or extract text to confirm header, table, body, and footer details.
2. Normalize transcript content: remove repeated filler, speaker hesitation, partial duplicates, and obvious ASR noise.
3. Extract facts conservatively:
   - confirmed meeting topic
   - core conclusions
   - important requirements or decisions
   - action items, owners, deadlines, and pending confirmations
   - risks, constraints, and phased implementation suggestions
4. Draft the minutes in a concise formal tone. Do not quote the transcript verbatim unless a short exact term is important.
5. Prepare `minutes.json` using this schema:

```json
{
  "title": "AI写标书工作台需求沟通会议纪要",
  "meeting_name": "AI写标书工作台需求沟通会",
  "place": "待补充",
  "host": "待补充",
  "meeting_time": "2026.5.28",
  "attendees": "需求方代表、项目实施方代表、系统原型演示及相关业务人员（具体名单待补充）",
  "cc": "相关部门",
  "sections": [
    {
      "heading": "一、会议背景与总体结论",
      "items": ["本次会议围绕……", "会议认为……"]
    }
  ]
}
```

6. Run the bundled builder:

```powershell
& $PYTHON "C:/Users/13971/.codex/skills/npa-meeting-minutes/scripts/build_minutes_docx.py" `
  --input minutes.json `
  --output output.docx
```

Use the bundled Python path from `load_workspace_dependencies` as `$PYTHON`.

7. Verify with OfficeCLI:
   - `officecli validate output.docx`
   - `officecli view output.docx issues`
   - PowerShell: `$env:OFFICECLI_NO_AUTO_RESIDENT='1'; officecli view output.docx screenshot --render html --grid auto -o preview.png`
   - run `scripts/check_border_continuity.py output.docx`

## Border Integrity

- The document must contain exactly one top-level master table after the title. Do not create separate info, body, or sign-off tables.
- Use a shared 12-column grid: merge columns in groups of 3 for the four-column meeting-info rows, merge all 12 columns for the body row, and merge columns in groups of 4 for the three-column sign-off rows.
- The master table must use `fixed` layout, an explicit `dxa` preferred width, no table indent, and one shared grid total for every row.
- Normalize every cell `tcW` to the exact sum of the grid columns it spans (`gridSpan` included). Re-run width normalization after all merges because merging can rewrite `tcW`.
- The body box must be one full-width merged cell in the master table with explicit top/left/bottom/right borders. Keep every section heading and numbered item inside that cell.
- The body row may split across pages; do not add `cantSplit`.
- The sign-off rows must remain inside the master table but override all cell borders to `nil`, so the body box ends cleanly above `编制 / 审核 / 批准`.
- Never simulate the body box by putting left/right borders on separate paragraphs. Paragraph spacing creates visible border gaps in Word/WPS.
- Do not convert the body table to paragraph borders merely because the OfficeCLI HTML renderer moves a long table to the next page. Treat that pagination as a renderer limitation; border continuity in the DOCX structure takes priority.
- Do not rely on `auto` or percentage widths. Do not use several tables with merely matching widths: WPS can independently reflow them and create a visible few-pixel edge offset.
- The continuity check must reject multiple top-level tables, require the body as row 4 of the master table, require the 12-column `fixed`/`dxa` grid, verify every cell width against its span, and verify borderless sign-off rows.

### WPS Alignment Red Lines

- Separate meeting-info, body, and sign-off tables are a failure even when their numeric widths match.
- A body made from paragraph borders is a failure because paragraph spacing creates border gaps.
- A body cell that does not span the complete master grid is a failure.
- Passing OfficeCLI structural validation alone is insufficient; also run the bundled continuity checker and inspect the OfficeCLI HTML screenshot.

## Writing Rules

- Keep numbered items clear and executable.
- Use `待补充` for unknown host, place, exact attendees, owners, or deadlines.
- Do not manufacture names, responsibilities, or dates not present in the materials.
- Prefer headings like:
  - `一、会议背景与总体结论`
  - `二、主要讨论内容`
  - `三、重点需求及决议`
  - `四、后续工作安排`
- For requirement meetings, use:
  - `一、会议背景与总体结论`
  - `二、核心业务痛点`
  - `三、主要功能需求`
  - `四、建设路径建议`
  - `五、后续工作安排`

## Assets

The builder uses:

- `assets/header_strip.jpg`
- `assets/footer_strip.jpg`

Do not replace these unless the company template changes.
