from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── PAGE SETUP ────────────────────────────────────────────────────────────────
sec = doc.sections[0]
sec.page_width  = Inches(6.5)
sec.page_height = Inches(9.5)
sec.left_margin   = Inches(1.1)
sec.right_margin  = Inches(1.0)
sec.top_margin    = Inches(1.1)
sec.bottom_margin = Inches(1.0)

# ── COLOUR PALETTE ────────────────────────────────────────────────────────────
C_BLACK   = RGBColor(0x10, 0x0D, 0x08)
C_BODY    = RGBColor(0x22, 0x1C, 0x12)
C_GOLD    = RGBColor(0x8A, 0x6E, 0x28)
C_GOLD_H  = RGBColor(0xB8, 0x98, 0x3C)
C_RED     = RGBColor(0x8B, 0x1A, 0x1A)
C_GREEN   = RGBColor(0x28, 0x6E, 0x28)
C_BLUE    = RGBColor(0x2C, 0x3E, 0x7A)
C_GREY    = RGBColor(0x55, 0x4E, 0x3A)
C_CREAM   = RGBColor(0xF5, 0xF0, 0xE4)

BODY_FONT   = 'Garamond'
MONO_FONT   = 'Courier New'
TITLE_FONT  = 'Garamond'

# ── HELPERS ───────────────────────────────────────────────────────────────────

def set_para_spacing(para, before=0, after=0, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing       = Pt(line)

def add_border_to_para(para, color="8A6E28", size=4, space=6, val="single"):
    """Add a bottom border to a paragraph."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'),   val)
    bottom.set(qn('w:sz'),    str(size))
    bottom.set(qn('w:space'), str(space))
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_full_border(para, color="8A6E28", size=4, space=4):
    """Box border around a paragraph."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for side in ('top','left','bottom','right'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    str(size))
        el.set(qn('w:space'), str(space))
        el.set(qn('w:color'), color)
        pBdr.append(el)
    pPr.append(pBdr)

def add_shade(para, fill="F5F0E4"):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    pPr.append(shd)

def para_indent(para, left=0.35, first=0):
    para.paragraph_format.left_indent  = Inches(left)
    para.paragraph_format.first_line_indent = Inches(first)

def run(para, text, bold=False, italic=False, size=12, color=None, font=None, small_caps=False):
    r = para.add_run(text)
    r.bold       = bold
    r.italic     = italic
    r.font.size  = Pt(size)
    r.font.name  = font or BODY_FONT
    if color:
        r.font.color.rgb = color
    if small_caps:
        r.font.small_caps = True
    return r

def heading_para(text, level=1, page_break_before=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if page_break_before:
        p.runs[0].add_break() if p.runs else None
        p.paragraph_format.page_break_before = True
    if level == 1:
        run(p, text, bold=True, size=22, color=C_GOLD_H, font=TITLE_FONT, small_caps=True)
        set_para_spacing(p, before=36, after=6)
        add_border_to_para(p, color="B8983C", size=6, space=4)
    elif level == 2:
        run(p, text, bold=True, size=13, color=C_RED, font=BODY_FONT, small_caps=True)
        set_para_spacing(p, before=28, after=4)
    return p

def chapter_opener(num, title, subtitle):
    doc.add_page_break()
    # Chapter number
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, f'— {num} —', bold=False, size=11, color=C_GREY, font=BODY_FONT)
    set_para_spacing(p, before=36, after=4)

    # Chapter title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, title.upper(), bold=True, size=20, color=C_GOLD_H, font=TITLE_FONT)
    set_para_spacing(p, before=4, after=6)
    add_border_to_para(p, color="B8983C", size=8, space=6)

    # Subtitle
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, subtitle, italic=True, size=11, color=C_GREY, font=BODY_FONT)
    set_para_spacing(p, before=10, after=30)

def body(text, indent=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p, text, size=12, color=C_BODY, font=BODY_FONT)
    set_para_spacing(p, before=0, after=8)
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.3)
    return p

def spacer(n=1):
    for _ in range(n):
        p = doc.add_paragraph()
        set_para_spacing(p, before=0, after=0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(4)

def divider(label=''):
    spacer()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sym = f'— {label} —' if label else '✦   ✦   ✦'
    run(p, sym, size=10, color=C_GREY, font=BODY_FONT, small_caps=True)
    set_para_spacing(p, before=10, after=10)
    spacer()

def document_block(title, lines, stamp='CLASSIFIED', shade='EFEFEA'):
    spacer()
    # stamp line
    p_stamp = doc.add_paragraph()
    p_stamp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run(p_stamp, f'[ {stamp} ]', bold=True, size=8, color=C_RED, font=MONO_FONT, small_caps=True)
    set_para_spacing(p_stamp, before=6, after=0)
    para_indent(p_stamp, left=0.3)

    # title row
    p_title = doc.add_paragraph()
    run(p_title, title, bold=True, size=9.5, color=C_GOLD, font=MONO_FONT)
    set_para_spacing(p_title, before=0, after=4)
    para_indent(p_title, left=0.3)
    add_border_to_para(p_title, color="8A6E28", size=4, space=2)

    # body lines
    for line in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run(p, line, italic=True, size=10, color=RGBColor(0x44,0x3A,0x22), font=MONO_FONT)
        set_para_spacing(p, before=0, after=2)
        para_indent(p, left=0.3)
        add_shade(p, fill=shade)

    spacer()

def tape_block(label, text, corrupted=False):
    spacer()
    p_label = doc.add_paragraph()
    col = C_RED if corrupted else C_GREEN
    run(p_label, f'▶  {label}', bold=True, size=9, color=col, font=MONO_FONT)
    set_para_spacing(p_label, before=4, after=2)
    para_indent(p_label, left=0.25)

    p_body = doc.add_paragraph()
    p_body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    shade = 'F5EEE4' if not corrupted else 'F5E4E4'
    run(p_body, f'"{text}"', italic=True, size=10.5,
        color=RGBColor(0x33,0x4A,0x22) if not corrupted else RGBColor(0x6A,0x22,0x22),
        font=MONO_FONT)
    set_para_spacing(p_body, before=0, after=6)
    para_indent(p_body, left=0.25)
    add_shade(p_body, fill=shade)
    spacer()

def note_block(header, lines, sig=''):
    spacer()
    p_hdr = doc.add_paragraph()
    run(p_hdr, header, bold=True, size=9, color=C_BLUE, font=MONO_FONT, small_caps=True)
    set_para_spacing(p_hdr, before=6, after=4)
    para_indent(p_hdr, left=0.25)
    add_border_to_para(p_hdr, color="2C3E7A", size=4, space=2)

    for line in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run(p, line, italic=True, size=11, color=RGBColor(0x2A,0x30,0x55), font=BODY_FONT)
        set_para_spacing(p, before=0, after=4)
        para_indent(p, left=0.25)
        add_shade(p, fill='EEF0F8')

    if sig:
        p_sig = doc.add_paragraph()
        p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(p_sig, f'— {sig}', italic=True, size=10, color=C_BLUE, font=BODY_FONT)
        set_para_spacing(p_sig, before=4, after=4)
        para_indent(p_sig, left=0.25)

    spacer()

def timeline_entry(date, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run(p, f'{date:<22}', bold=True, size=10, color=C_GOLD, font=MONO_FONT)
    run(p, text, size=11, color=C_BODY, font=BODY_FONT)
    set_para_spacing(p, before=2, after=2)
    para_indent(p, left=0.1)

def pull_quote(text, attribution=''):
    spacer()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, f'\u201c{text}\u201d', italic=True, size=14, color=C_GOLD, font=TITLE_FONT)
    set_para_spacing(p, before=14, after=6)
    para_indent(p, left=0.4)
    add_full_border(p, color="B8983C", size=6, space=8)

    if attribution:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run(p2, f'\u2014 {attribution}', italic=True, size=10, color=C_GREY, font=BODY_FONT)
        set_para_spacing(p2, before=0, after=14)
        para_indent(p2, left=0.4)
    spacer()

# ══════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════════════════════════════════════

# Classification banner
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'CLASSIFIED  —  DEPT. OF ANOMALOUS RESEARCH  —  EYES ONLY',
    bold=True, size=9, color=C_RED, font=MONO_FONT, small_caps=True)
set_para_spacing(p, before=30, after=30)

# Decorative top rule
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, '─' * 52, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=0, after=20)

spacer()
spacer()
spacer()

# Main title
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'CASE FILE 1994', bold=True, size=36, color=C_GOLD_H, font=TITLE_FONT)
set_para_spacing(p, before=0, after=10)

# Decorative rule
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, '─' * 40, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=0, after=14)

# Subtitle
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'THE ALDRIC BUILDING INCIDENT', bold=False, size=14, color=C_GREY, font=TITLE_FONT, small_caps=True)
set_para_spacing(p, before=0, after=6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'A Full Incident Record', italic=True, size=12, color=C_GREY, font=BODY_FONT)
set_para_spacing(p, before=0, after=60)

spacer()
spacer()
spacer()
spacer()

# Cover pull quote
pull_quote(
    'The building did not contain the signal.\nThe signal contained the building.',
    'Core Log — Final Entry — Aldric Building, 1987'
)

spacer()
spacer()
spacer()

# File reference
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'FILE REF: DAR-1994-CF-0047', size=9, color=C_GREY, font=MONO_FONT)
set_para_spacing(p, before=0, after=4)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'COMPILED: NOVEMBER 1994  |  STATUS: SEALED INDEFINITELY', size=9, color=C_GREY, font=MONO_FONT)
set_para_spacing(p, before=0, after=4)

spacer()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, '─' * 52, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=10, after=6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'Created and Developed by\n', italic=False, size=9, color=C_GREY, font=MONO_FONT)
run(p, 'Ale Espinoza  &  Vanessa Baldueza', bold=True, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=0, after=20)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER I
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('I', 'Project Veil', '"Something is still inside."')

body(
    'In 1985, a division of the federal government known only as the Department '
    'of Anomalous Research made contact with something it could not explain.'
)
body(
    'It was not a creature. It was not a signal in any conventional sense. It was '
    'a frequency — a pattern of sound that existed at the edge of human perception, '
    'appearing spontaneously in an abandoned industrial district outside the city. '
    'Workers in the surrounding blocks reported hearing it in their sleep. Three '
    'reported that it had begun to respond to them. Two could no longer leave the '
    'area voluntarily.'
)
body(
    'The Department mobilised within 48 hours. The entity was contained, documented, '
    'and assigned a project designation: VEIL.'
)
body(
    'A research facility was requisitioned — the Aldric Building, Block 7. Staff '
    'were assigned. Clearances were issued. And the frequency was brought inside.'
)
body('None of the senior researchers believed it could escape.')

document_block(
    'INTERNAL MEMO — ARCHIVE DIVISION  |  JAN 14, 1987',
    [
        'All employees on Floor 3 have been instructed to avoid Sub-corridor',
        '7B until further notice.',
        '',
        'Reason: unspecified auditory anomaly.',
        'Reported by 9 separate workers.',
        'Acknowledged by management: NO.',
    ],
    stamp='ARCHIVE DIV.',
    shade='F5F0E4'
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER II
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('II', 'The Researchers', '"We should have stopped in September."')

body(
    'The early research logs describe a frequency that was initially passive. '
    'It did not respond to stimuli. It did not move. It simply existed in the '
    'lower sublevels of the Aldric Building, broadcasting at 40Hz — a range '
    'below normal speech, felt more than heard.'
)
body('That changed on Day 4.')

document_block(
    'FIELD NOTE — DR. VALE  |  SUBLEVEL B1',
    [
        'The signal does not repeat. It responds.',
        '',
        'We changed the frequency on day 4.',
        'It changed back before we could log it.',
        'Kellner refuses to enter the room anymore.',
        'I understand why.',
        'We should have stopped in September.',
    ],
    stamp='FIELD NOTE',
    shade='F4F0E8'
)

body(
    'Dr. Vale was one of the lead researchers assigned to the Sublevel B1 '
    'containment zone. Her notes — recovered during the 1994 investigation — '
    'paint a portrait of a team that understood what was happening to them '
    'and could not stop it.'
)
body(
    'The frequency was not simply broadcasting. It was learning. It absorbed '
    'identity from proximity. The researchers closest to the entity began to '
    'change — not physically, not immediately — but at the edges. Small things first.'
)

document_block(
    'PERSONAL LOG — DR. VALE  |  SUBLEVEL B1  |  NOV 1, 1987',
    [
        'The frequency is changing us.',
        '',
        'Kellner started humming it in his sleep.',
        'None of us can remember what it sounded like before.',
        'I tried to leave today. My feet brought me back.',
        'I think it wants us here.',
    ],
    stamp='PERSONAL LOG',
    shade='F4F0E8'
)

body(
    'Dr. Kellner — never formally identified beyond his surname in Vale\'s '
    'notes — reportedly refused to re-enter the research wing in early November '
    '1987. He was not recorded among the 47 missing. What became of him is unknown.'
)
body(
    'Dr. Neel was assigned to Sublevel C2. A recording was found near his last '
    'known position. No researcher who recovered it ever listened to it. It was '
    'sealed and filed as Exhibit 7-N. The file has since been destroyed.'
)

tape_block(
    'TAPE 003 — CORRUPTED  |  TIMESTAMP: 1991-03-22',
    'This recording predates your investigation by three years. The voice on the tape is yours.',
    corrupted=True
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER III
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('III', 'The Containment Breach', '"The frequency became them."')

body(
    'The official cause of death listed in the sealed government record is a '
    'gas leak. The building, it states, was decommissioned following an industrial '
    'accident. The files were sealed. The record was archived. The case was closed.'
)

pull_quote('That is a lie.')

body(
    'In March 1987, the containment around Project Veil failed. Forty-seven city '
    'workers — maintenance staff, researchers, and administrative personnel — were '
    'present in the building at the time of the breach. None reached the surface.'
)
body('The frequency did not kill them. It absorbed them.')

tape_block('TAPE 004', 'PROJECT VEIL was containment. They caught it in 1985. It does not die. It LEARNS. It absorbs identity from proximity.')
tape_block('TAPE 005', 'The containment failed March 1987. 47 workers. None made it to the surface. The frequency became them.')
tape_block('TAPE 006', 'It mimics human posture to move. It is not what you think you see. Do not look at it directly.')

heading_para('Timeline of Events', level=2)

timeline_entry('NOV 2, 1987', 'Worker 41 unaccounted for. Search ongoing.')
timeline_entry('NOV 3, 1987', 'Workers 42–44. Doors found open from inside.')
timeline_entry('NOV 4, 1987', 'Transmitter signal strength: 400% baseline.')
timeline_entry('NOV 5  —  2:44 AM', 'Emergency evacuation order issued. Building sealed by Dept. order 0312.')
timeline_entry('NOV 5, 1987', '47 unaccounted for. File sealed. Official record: gas leak.')

spacer()

document_block(
    'EMERGENCY MEMO — SUBLEVEL B3  |  NOV 5, 1987  —  2:44 AM  |  EYES ONLY',
    [
        'Containment breach confirmed. All 47 subjects: status unknown.',
        'Building sealed by Dept. order 0312.',
        'This file has been classified retroactively.',
        '',
        'If you are reading this inside the building —',
        'you were never meant to leave either.',
    ],
    stamp='EYES ONLY',
    shade='F5E8E8'
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER IV
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('IV', 'The Nature of the Entity', '"Do not look at it directly."')

body(
    'Two distinct manifestations of the absorbed frequency were encountered '
    'during the 1994 investigation. They are designated here by the names used '
    'in the recovered field recordings.'
)

heading_para('The Figure', level=2)

body(
    'The primary manifestation. It moves. It hunts. It learns from each encounter, '
    'adapting patrol patterns and response speed over time. It does not communicate. '
    'It does not warn.'
)
body(
    'Researchers speculate it is a composite of multiple absorbed identities — '
    'the 47 workers, compressed into a single mobile form that retains only one '
    'imperative: prevent exit.'
)

heading_para('The Listener', level=2)

body(
    'Encountered in Sublevels C2 and the Containment Core. Blind. Responds '
    'exclusively to acoustic stimuli — footsteps, breathing, vibration. '
    'Threat level exceeds that of The Figure in confined environments.'
)
body(
    'It does not sleep. It locates sound sources faster with each detection. '
    'It is not just detecting. It is learning.'
)

document_block(
    'DEPT. MEMO — SUBJECT DESIGNATION: THE LISTENER',
    [
        'Blind. Responds exclusively to acoustic stimuli.',
        'Threat level: CRITICAL.',
        '',
        'Do not run. Do not sprint.',
        'It hears heartbeats at close range.',
        'Silence is your only protection.',
    ],
    stamp='THREAT: CRITICAL',
    shade='F5E8E8'
)

tape_block('TAPE C2-04', 'The sound triggers it. It cannot see. It only listens.')
tape_block('TAPE C2-06', 'Move slow. Stay quiet. You have a chance. —Vale')

body(
    'Subject 47 — the last of the original workers — was confirmed as a distinct '
    'contained entity in the lowest level of the building: the Containment Core. '
    'It had been in that room for seven years. It was still waiting.'
)

document_block(
    'CORE LOG — FINAL ENTRY  |  DATE UNKNOWN',
    [
        'Subject 47. Experiment 7. Day unknown.',
        '',
        'The signal is no longer external.',
        'We confirmed it yesterday. It is inside the subjects.',
        'They are broadcasting it.',
        '',
        'There is no transmitter. There never was.',
        '',
        'The building did not contain the signal.',
        'The signal contained the building.',
    ],
    stamp='CORE LOG',
    shade='F0EEE4'
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER V
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('V', 'Detective Harris', '"He was not sent there. He was called there."')

body(
    'Seven years after the Aldric Building was sealed, Detective S. Harris — '
    'a private investigator working three unrelated open cases — received an '
    'anonymous email.'
)

note_block(
    'EMAIL  |  RECEIVED: NOV 3, 1994 — 11:47 PM  |  FROM: [REDACTED]@anomalousresearch.gov',
    [
        'Detective Harris,',
        '',
        'Forty-seven city workers disappeared from this building in 1987.',
        'The official cause of death: gas leak.',
        'The official record: sealed by government order.',
        'That is a lie.',
        '',
        'There are recordings inside. Documents. Evidence of what really',
        'happened. What they were studying. What they found.',
        '',
        'You are the only person I trust to go in.',
        'Go alone. Tell no one. Do not delay.',
        '',
        '— SENDER ID CORRUPTED —',
    ],
    sig=''
)

note_block(
    'HARRIS — PRIVATE CASE NOTES  |  NOVEMBER 4, 1994 — 11:58 PM',
    [
        'Received an anonymous email tonight.',
        'The sender address no longer exists. IP trace: dead end.',
        '',
        'Official record on the Aldric Building: gas leak, decommissioned.',
        'Too clean. Sealed too fast.',
        '',
        'I have three open cases that trace back to that block.',
        'No connection I can prove. But I feel it.',
        '',
        'This could be nothing.',
        'Or it could be everything.',
        '',
        'Tomorrow. 3 AM. No backup.',
    ],
    sig='S. Harris'
)

body(
    'Harris arrived at the Aldric Building on the morning of November 5, 1994, '
    'at approximately 3:06 AM. It was raining. The guardhouse was unoccupied. '
    'The gate was unlocked.'
)
body(
    'He noted something else — a sound difficult to place. Not the rain. '
    'Something beneath it.'
)
body('He went in alone.')

note_block(
    'ALDRIC BUILDING — BLOCK 7  |  NOVEMBER 5, 1994 — 3:06 AM',
    [
        'Still raining.',
        '',
        'The guardhouse is unoccupied. The gate is unlocked.',
        'There should be a security key inside.',
        '',
        'The building entrance is sealed.',
        'Whatever is in there — it has been waiting seven years.',
        '',
        'I can hear the rain on the roof.',
        'Something else, too. Hard to place.',
        '',
        'Somewhere in that building is the truth.',
        'I am going in.',
    ],
    sig='S. Harris'
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER VI
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('VI', 'Inside the Aldric Building', '"Trust the clues. Run."')

body(
    'What Harris found inside defied every assumption in his case notes. The '
    'building had not decayed in the normal sense. It was occupied — not by '
    'squatters, not by wildlife — by something that moved with intention. '
    'That tracked. That learned. That blocked exit points.'
)
body(
    'Scattered across five floors were eleven evidence tapes and a series of '
    'sealed documents, each encoding a fragment of the truth. The researchers '
    'who had placed them there understood what was happening and had left behind '
    'the only thing they could: a cipher. A way out for anyone who came after them.'
)
body(
    'Each tape held a clue. Each clue was a letter pair. Together they decoded '
    'the word that unlocked each floor\'s sealed exit.'
)

tape_block('TAPE 010', 'The researchers left a cipher. The exit is locked by frequency-code. It was designed so ONLY the signal could decode it.')
tape_block('TAPE 011  [FINAL]', 'You were never meant to leave. But the researchers left one way out. Find the cipher. Trust the clues. Run.')

body(
    'On the third floor, Harris recovered a tape predating his investigation by '
    'three years — a recording from 1991, in a voice that matched his own. He '
    'had no memory of making it. He had no memory of ever entering the Aldric '
    'Building before.'
)
body('He kept moving.')

tape_block(
    'TAPE 009 — CORRUPTED',
    'You — Detective — this is a message from yourself. Three months from now. Do not look at it. Do not—',
    corrupted=True
)
tape_block('TAPE CORE-06', 'If you find this — I made it out. I think. —S.H.')


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER VII
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('VII', 'The Truth About the Signal', '"There was never a transmitter."')

body(
    'The last floor — the Containment Core — held the final answer.'
)
body(
    'There was no transmitter. There had never been a transmitter. The '
    'Department of Anomalous Research had spent years studying a device they '
    'believed was broadcasting the frequency. They had the cause and effect inverted.'
)
body(
    'The frequency was not coming from a machine in the building. The frequency '
    'was coming from the workers. All forty-seven of them. Absorbed into the '
    'signal, they had become the source. They were broadcasting it. They were '
    'still broadcasting it, from inside the walls, from inside the floors, from '
    'the darkness on every level of the Aldric Building.'
)

pull_quote(
    'The building did not contain the signal.\nThe signal contained the building.',
    'Core Log — Final Entry'
)

body(
    'And for seven years, it had been waiting for someone to walk back in.'
)

document_block(
    'INCIDENT LOG — CONTAINMENT CORE',
    [
        'All 47 subjects confirmed contained — Level 5.',
        'Acoustic sensitivity: extreme. Movement triggers response.',
        'Evacuation order: Nov 5, 1987 — too late.',
        '',
        'This floor was never meant to be opened again.',
        '',
        'If you are reading this — run. Slowly.',
    ],
    stamp='SEALED',
    shade='F0EEE4'
)


# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER VIII
# ══════════════════════════════════════════════════════════════════════════════
chapter_opener('VIII', 'November 7, 1994', '"I pray that is enough."')

body(
    'Harris exited the Aldric Building on November 7, 1994 — two days after '
    'entry. His final recorded statement, recovered from a personal audio device '
    'found near the building\'s south exit, reads as follows:'
)

note_block(
    'HARRIS — AUDIO LOG  |  NOVEMBER 7, 1994',
    [
        '"The transmitter. I destroyed it. The frequency stopped."',
        '',
        '"I pray that is enough."',
    ],
    sig='Det. Harris   Nov. 7, 1994'
)

document_block(
    'DEPT. OF ANOMALOUS RESEARCH — INTERNAL NOTE  [RESTRICTED]',
    [
        'Harris was not assigned to this investigation.',
        'Harris received an anonymous tip. Caller: unidentified.',
        'Voice pattern on tip recording: matches no known individual.',
        'Harris has no memory of receiving the tip.',
        'No record of who convinced him to enter the building.',
        '',
        'The Aldric Building was demolished. May 1995.',
        'No transmitter was found.',
        '',
        '[ CASE FILE — SEALED INDEFINITELY ]',
    ],
    stamp='RESTRICTED',
    shade='F5E8E8'
)

body('Harris never filed a report. No one knew he had entered the building.')
body('The anonymous tip came three weeks before. The voice on the recording was his own.')
body('He was not sent there.')
body('He was called there.', indent=False)


# ══════════════════════════════════════════════════════════════════════════════
# EPILOGUE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'EPILOGUE', bold=True, size=16, color=C_GOLD_H, font=TITLE_FONT, small_caps=True)
set_para_spacing(p, before=60, after=30)
add_border_to_para(p, color="B8983C", size=6, space=6)

pull_quote(
    'The signal does not repeat.\nIt responds.\n\n'
    'We changed the frequency on day 4.\n'
    'It changed back before we could log it.',
    'Dr. Vale  —  Field Note  —  Sublevel B1  —  1987'
)

spacer()

for line in [
    'The Aldric Building no longer stands.',
    '',
    'The frequency has not been detected since November 1994.',
    '',
    'Detective Harris closed his remaining open cases in December 1994.',
    'He retired from private investigation in 1995.',
    '',
    'He does not discuss the Aldric Building.',
    'He does not remember the tape dated 1991.',
    '',
    'He still hears something at night.',
    'Just below the threshold of speech.',
    'Something that responds when he listens back.',
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run(p, line, italic=True, size=12, color=C_GREY, font=BODY_FONT)
    set_para_spacing(p, before=0, after=3)

spacer()
spacer()

divider()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, '─' * 40, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=20, after=10)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'CASE FILE 1994', bold=True, size=14, color=C_GOLD_H, font=TITLE_FONT)
set_para_spacing(p, before=0, after=6)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'Created and Developed by', size=10, color=C_GREY, font=BODY_FONT)
set_para_spacing(p, before=0, after=4)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, 'Ale Espinoza  &  Vanessa Baldueza', bold=True, size=12, color=C_GOLD, font=BODY_FONT)
set_para_spacing(p, before=0, after=20)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, '─' * 40, size=10, color=C_GOLD, font=MONO_FONT)
set_para_spacing(p, before=0, after=0)

# ── SAVE ──────────────────────────────────────────────────────────────────────
out = r'c:\Users\ale.espinoza\Documents\testtt\Case File 1994 — Story.docx'
doc.save(out)
print(f'Saved: {out}')
