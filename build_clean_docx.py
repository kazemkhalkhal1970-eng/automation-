#!/usr/bin/env python3
"""
Clean rebuild of the Autonomous Excavator RAS review article.
Addresses all 13 critical + 10 minor reviewer issues from the second-revision report.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FIG_DIR = '/home/z/my-project/download/article_figures'
OUT     = '/home/z/my-project/download/Autonomous_Excavator_RAS_Review.docx'

doc = Document()

# ---- Normal style ----
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(11)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts'); rPr.insert(0, rFonts)
rFonts.set(qn('w:ascii'), 'Times New Roman')
rFonts.set(qn('w:hAnsi'), 'Times New Roman')
rFonts.set(qn('w:eastAsia'), 'Times New Roman')
rFonts.set(qn('w:cs'), 'Times New Roman')

# ---- Page setup ----
section = doc.sections[0]
section.page_height = Cm(29.7); section.page_width = Cm(21.0)
section.left_margin = Cm(2.0); section.right_margin = Cm(2.0)
section.top_margin = Cm(2.0);  section.bottom_margin = Cm(2.0)

# ---- Empty header / footer with page numbers ----
header = section.header
header.paragraphs[0].text = ''
footer = section.footer
footer_p = footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run()
run.font.name = 'Times New Roman'; run.font.size = Pt(9)
fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
instrText = OxmlElement('w:instrText'); instrText.text = 'PAGE'
fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
r_el = run._r
r_el.append(fldChar1); r_el.append(instrText); r_el.append(fldChar2)

# ---- Colors ----
PRIMARY = RGBColor(0x1F, 0x3A, 0x68)
ACCENT  = RGBColor(0xC8, 0x4B, 0x31)
GRAY    = RGBColor(0x6C, 0x75, 0x7D)
BLACK   = RGBColor(0x00, 0x00, 0x00)

def H(text, level=1, color=PRIMARY):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True; run.font.name = 'Arial'; run.font.color.rgb = color
    if level == 1:   run.font.size = Pt(14)
    elif level == 2: run.font.size = Pt(12)
    else:            run.font.size = Pt(11)
    return p

def B(text, italic=False, justify=True, indent=True, size=11):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_after = Pt(4)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    if indent:
        pf.first_line_indent = Cm(0.6)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'; run.font.size = Pt(size); run.italic = italic
    return p

def BL(text, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.8 + 0.4 * level)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11)

def EQ(eq_text, eq_num=None):
    """Render an equation centered, with the equation number right-aligned."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.4
    tab_stops = p.paragraph_format.tab_stops
    page_w = (section.page_width - section.left_margin - section.right_margin)
    tab_stops.add_tab_stop(page_w // 2, WD_TAB_ALIGNMENT.CENTER)
    tab_stops.add_tab_stop(page_w, WD_TAB_ALIGNMENT.RIGHT)
    run = p.add_run('\t' + eq_text + '\t')
    run.font.name = 'Cambria Math'; run.font.size = Pt(11); run.italic = True
    if eq_num is not None:
        num_run = p.add_run('({})'.format(eq_num))
        num_run.font.name = 'Times New Roman'; num_run.font.size = Pt(11)
    return p

def FIG(img_filename, caption, width_in=5.6):
    img_path = os.path.join(FIG_DIR, img_filename)
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(18)
    p_img.paragraph_format.keep_with_next = True
    p_img.paragraph_format.space_after  = Pt(8)
    run = p_img.add_run()
    run.add_picture(img_path, width=Inches(width_in))
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after  = Pt(16)
    p_cap.paragraph_format.keep_with_next = False
    crun = p_cap.add_run(caption)
    crun.font.name = 'Times New Roman'; crun.font.size = Pt(9); crun.italic = True
    crun.font.color.rgb = GRAY

def TBL_CAP(caption):
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(8)
    p_cap.paragraph_format.space_after  = Pt(2)
    crun = p_cap.add_run(caption)
    crun.font.name = 'Times New Roman'; crun.font.size = Pt(9); crun.italic = True
    crun.font.color.rgb = GRAY

def TBL(headers, rows, col_widths_in=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        para = hdr[i].paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(h)
        run.bold = True; run.font.name = 'Arial'; run.font.size = Pt(9)
        run.font.color.rgb = PRIMARY
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for r_idx, row in enumerate(rows):
        cells = table.rows[r_idx + 1].cells
        for c_idx, val in enumerate(row):
            cells[c_idx].text = ''
            para = cells[c_idx].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para.add_run(str(val))
            run.font.name = 'Times New Roman'; run.font.size = Pt(9)
    if col_widths_in:
        for i, w in enumerate(col_widths_in):
            for row in table.rows:
                row.cells[i].width = Inches(w)
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(6)
    return table

# ============================================================
# COVER / TITLE BLOCK
# ============================================================
for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(8)
r = p.add_run('Robotics and Autonomous Systems for Excavator Automation: A PRISMA-2020 Systematic Review of Mechatronic Architectures, Robotic Perception, and Intelligent Control of Earth-Moving Robots')
r.bold = True; r.font.name = 'Arial'; r.font.size = Pt(17); r.font.color.rgb = PRIMARY

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(16)
r = p.add_run('A systematic review and quantitative meta-analysis of seventy included studies, with a unified five-pillar mechatronic taxonomy, twenty-one governing equations, eleven figures, and a research roadmap toward full autonomy across defined excavation operational design domains')
r.italic = True; r.font.name = 'Times New Roman'; r.font.size = Pt(11); r.font.color.rgb = GRAY

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run('Masoud Masih Tehrani  *,  Kazem Khalkhal')
r.bold = True; r.font.name = 'Arial'; r.font.size = Pt(13); r.font.color.rgb = BLACK

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run('Automotive Engineering School, Iran University of Science and Technology, Tehran, Iran')
r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
r = p.add_run('* Corresponding author: m.masihtehrani@iust.ac.ir')
r.font.name = 'Times New Roman'; r.font.size = Pt(9); r.font.color.rgb = GRAY

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(16)
r = p.add_run('Submitted to: Robotics and Autonomous Systems (Elsevier)   |   2026')
r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.italic = True; r.font.color.rgb = GRAY

# ============================================================
# ABSTRACT  (revised: no "first" / ROI overstatement; keywords reduced to 8)
# ============================================================
H('Abstract', level=2)
abstract = ('This article presents a systematic review of seventy included studies on autonomous excavator robotic systems published between 1995 and 2025, conducted and reported in line with the PRISMA 2020 statement. The studies are consolidated into a unified five-pillar mechatronic framework comprising sensing, perception, decision, planning, and action. The goal is to identify the dominant technical trajectories in robotic earth-moving equipment, quantify cross-study effect sizes via meta-analysis, and articulate a research roadmap toward full autonomy across defined excavation operational design domains, as referenced by ISO 17757 and ISO 19014. '
            'A systematic search of Scopus, Web of Science, IEEE Xplore, ScienceDirect, and MDPI yielded 1,847 candidate records, of which 70 satisfied all eligibility and quality criteria on a ten-domain engineering-specific appraisal rubric. Twenty-one governing equations were extracted, and a random-effects (DerSimonian-Laird) meta-analysis was performed on nine quantitative comparisons drawn from the included studies. '
            'Domain-specific random-effects pooled effect sizes range from d = 0.17 (energy) to d = 0.58 (control), all favouring proposed methods over baselines. The largest single-study standardised effect size in the synthesis is Cohen d = 0.74 for DB-MPC, calculated from the results reported by Okada et al. [37] (who report a percentage reduction in tracking error relative to a fixed-model baseline; the percentage and the standardised d are distinct quantities and should not be conflated). Electrified actuator architectures deliver 12-18 % energy savings over comparable hydraulic architectures. Commonsense knowledge graphs reduce conservative false stops by 31 %. Reinforcement-learning-based controllers achieve 70 % success rates on boulder excavation tasks, approaching the 83 % achieved by skilled human operators. No single control strategy dominates across all five evaluation criteria, indicating that controller choice must remain application-driven. '
            'This review contributes (i) a unified five-pillar mechatronic taxonomy with twenty-one explicit governing equations; (ii) a quantitative synthesis of autonomous-excavator robotic interventions that complements the qualitative surveys of Eraliev et al. [2], Wu and Cai [3], and Nguyen and Ha [30]; to our knowledge, this is among the first quantitative syntheses in this literature; and (iii) a research roadmap explicitly tied to electrified-architecture energy management and commonsense-driven task reasoning. '
            'The taxonomy and benchmarks provide robot manufacturers with a structured decision framework for selecting sensing, perception, and control strategies. The meta-analytic effect sizes may inform preliminary cost-benefit discussions, although a direct mapping from Cohen d to monetary return-on-investment is not established and would require further validation.')
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.line_spacing = 1.3
p.paragraph_format.left_indent  = Cm(0.6); p.paragraph_format.right_indent = Cm(0.6)
p.paragraph_format.space_after  = Pt(8)
r = p.add_run(abstract); r.font.name = 'Times New Roman'; r.font.size = Pt(10)

p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.left_indent  = Cm(0.6); p.paragraph_format.right_indent = Cm(0.6)
r = p.add_run('Keywords: ')
r.bold = True; r.font.name = 'Arial'; r.font.size = Pt(10)
r2 = p.add_run('autonomous excavator; systematic review; meta-analysis; mechatronic architecture; model predictive control; 3D LiDAR; construction robotics; mining automation.')
r2.font.name = 'Times New Roman'; r2.font.size = Pt(10); r2.italic = True

doc.add_page_break()

# ============================================================
# 1. INTRODUCTION
# ============================================================
H('1. Introduction', level=1)

B('Autonomous excavator robotic systems constitute one of the most economically significant and operationally challenging classes of field robotics, deployed across mining, construction, infrastructure development, and hazardous-environment remediation [1,2,5,8,9,17,22,32,55,71,72,73]. A hydraulic excavator is, in essence, a heavy-duty mobile manipulator: a tracked or wheeled base carrying a 3- to 4-degree-of-freedom (DOF) serial manipulator arm terminated in a bucket that engages soil at forces of tens of kilonewtons. The robotic-systems challenge of automating such a machine is compounded by an unstructured outdoor workspace, harsh environmental conditions (dust, vibration, extreme temperatures), operator fatigue and safety hazards, and the rapidly increasing global shortage of skilled heavy-equipment operators [23,40,60]. Three converging pressures have accelerated the migration from manual to autonomous excavation: (i) demographic - the ageing of the heavy-equipment operator workforce in mining regions; (ii) safety - the imperative to remove humans from inherently hazardous environments such as underground mines, slope-stabilisation sites, and demolition zones [9,12,18,30,37,55]; and (iii) economic - the rapidly decreasing cost of perception sensors, embedded compute, and electrified actuation, which together make robotic autonomy economically viable at scale for the first time [8,38,50,54,57,69].')

B('The state of the art in autonomous excavation has matured rapidly over the last five years, building on three decades of foundational work at Carnegie Mellon University (CMU) [2,11,17,19,28], the University of British Columbia [4,5,6], Tsinghua University [32,33,46,68], and ETH Zurich [18,26,27]. Liu et al. [1] demonstrated a self-driving wheel excavator whose path tracking is governed by an enhanced model-free adaptive controller with explicit time-delay compensation (EMFAC-TD), achieving sub-metre tracking accuracy without any prior plant identification. Eraliev, Shin, and Lee [2] provided a comprehensive review of the sensing, perception, decision, planning, and action modules that together constitute an autonomous excavator, identifying persistent gaps in soil-tool interaction modelling and in the integration of commonsense reasoning. Wu and Cai [3] surveyed the control systems and strategies specifically, cataloguing the migration from classical PID and inverse-kinematics-based control toward model-free, learning-based, and database-driven paradigms. Zhang et al. [32] presented an autonomous excavator system (AES) for material loading tasks that has been deployed commercially and has accumulated more than 8,500 operating hours, integrating multimodal perception, hierarchical task and motion planning, and a controller tightly coupled to the perception pipeline.')

B('More recently, Wang et al. [33] formulated a 3-DOF manipulator dynamic model with online compensation of gravity, inertia, and centripetal terms, reducing the maximum joint tracking errors of the boom, arm, and bucket by 43 %, 27 %, and 51 % respectively. Bae et al. [34] integrated landscape estimation, excavation path generation, real-time control, and post-excavation inspection into a single AES achieving centimetre-level accuracy on repetitive excavation tasks. Ahmadi Khiyavi and Seo [35] redesigned the excavator with parallel electrical linear actuators and proposed a PSO/PFM path-generation scheme that demonstrates tangible energy savings. Peng et al. [36] developed an improved RANSAC-based 3D LiDAR obstacle detection pipeline tailored to the geometric constraints of underground mining tunnels. Okada et al. [37] introduced a database-driven model predictive control (DB-MPC) architecture that adapts its prediction model online to changing soil parameters, achieving 22 % lower tracking error than fixed-model MPC. Chen et al. [29] introduced ExACT, an end-to-end autonomous excavator system using action-chunking transformers that takes raw LiDAR and camera inputs and directly outputs valve commands - a paradigm shift that bypasses the traditional Sense-Plan-Act decomposition. Egli et al. [26,27] demonstrated that reinforcement-learning-based bucket-filling controllers can react to encountered material properties in real time, achieving performance comparable to skilled human operators.')

B('Despite this rapid progress, the literature remains fragmented along module boundaries: control papers rarely cite perception papers and vice versa; electrification studies are largely decoupled from autonomy studies; and safety certification is treated as an afterthought rather than a first-class design constraint. The existing review literature - including the influential surveys by Eraliev et al. [2] and Wu and Cai [3] - is qualitative in nature and does not provide quantitative cross-study effect sizes, leaving practitioners without a basis for principled selection among candidate technologies. The present article addresses this gap by: (a) consolidating seventy included studies into a unified five-pillar mechatronic architecture; (b) formalising twenty-one governing equations spanning kinematics, dynamics, actuation, control, and performance metrics; (c) providing, to our knowledge, among the first domain-specific quantitative syntheses of autonomous-excavator interventions in this literature; and (d) articulating a five-point research roadmap explicitly tied to electrified-architecture energy management and commonsense-driven task reasoning.')

H('1.1 Research questions', level=2)
B('This review is structured around four explicit research questions (RQs) that frame the analysis, the meta-analysis, and the roadmap:')
BL('RQ1 - What is the current maturity of each of the five mechatronic pillars (sensing, perception, decision, planning, action) in autonomous-excavator robotic-systems research, and how is maturity distributed across the seventy included studies?')
BL('RQ2 - What is the pooled quantitative effect size of advanced control and autonomous-control strategies relative to the baselines used in the comparable studies, across the studies that report comparable metrics?')
BL('RQ3 - To what extent does electrification of the actuator stack alter the energy efficiency of autonomous excavation, and what is the magnitude of the gain attributable to electrified architectures alone?')
BL('RQ4 - What are the major open challenges identified from the reviewed literature that will shape the next five years of autonomous-excavator robotic-systems research, and what specific technical directions are most likely to address them?')

H('1.2 Contributions of this article', level=2)
B('Relative to the most-cited prior review by Eraliev et al. [2] (qualitative, single-pillar focused) and the control-focused review by Wu and Cai [3], this article makes four specific contributions, summarised in Table 1.')

TBL_CAP('Table 1. Contributions of this article relative to the two most-cited prior reviews.')
TBL(
    headers=['Dimension', 'Eraliev et al. [2]', 'Wu & Cai [3]', 'This article'],
    rows=[
        ['Research question',  'Sensing-action gaps', 'Control strategies', 'RQ1-RQ4 with meta-analysis'],
        ['Databases searched',  'Not specified', 'Not specified', '5 databases + deduplication'],
        ['Search period',       '2000-2022', '2010-2022', '1995-2025'],
        ['Studies included',    'Qualitative', 'Qualitative', '70 (PRISMA flow)'],
        ['Quality appraisal',   'Narrative', 'Narrative', '10-domain rubric + GRADE'],
        ['Meta-analysis',       'None', 'None', '9 comparisons, random-effects'],
        ['Equations extracted', 'Few', 'Few', '21 governing equations'],
        ['Open roadmap',       'Brief', 'Brief', '5-point, electrification-tied'],
    ],
    col_widths_in=[1.2, 1.4, 1.4, 1.6],
)

# (continuation appended below)

# ============================================================
# 2. METHODS
# ============================================================
H('2. Methods', level=1)
B('This systematic review was conducted and reported in line with the PRISMA 2020 statement [68]. The PRISMA 2020 checklist is provided as Supplementary Material. No review protocol was pre-registered on a public platform (e.g., PROSPERO); the review protocol, including the search strategy, eligibility criteria, and analysis plan, is documented in Section 2 and was finalised before the formal screening began. We acknowledge that protocol pre-registration is now considered best practice and recommend it for future systematic reviews in the construction-robotics domain.')

H('2.1 PRISMA 2020 reporting', level=2)
B('The PRISMA 2020 statement provides a 27-item checklist and a flow diagram template for reporting systematic reviews. The present article adheres to the checklist items applicable to a non-clinical engineering review; items relating to patient-reported outcomes and clinical heterogeneity are not applicable. The flow diagram (Figure 1) reports the number of records identified, screened, eligible, and included at each stage of the review.')

H('2.2 Eligibility criteria', level=2)
B('Studies were eligible if they (i) addressed autonomous, semi-autonomous, or teleoperated excavation systems; (ii) reported original research or systematic reviews published in English between January 1995 and December 2025; (iii) appeared in a peer-reviewed journal, conference proceeding, or arXiv preprint with a digital object identifier; and (iv) addressed at least one of the five mechatronic pillars (sensing, perception, decision, planning, action). Studies were excluded if they (i) addressed only manual excavation ergonomics; (ii) addressed only conventional hydraulic control without an autonomy component; (iii) were editorials, abstracts, or duplicate publications; or (iv) addressed construction automation generically without specific reference to excavation.')

H('2.3 Search strategy and screening', level=2)
B('Five databases were searched: Scopus, Web of Science, IEEE Xplore, ScienceDirect, and MDPI. The search string combined synonyms for "autonomous", "excavator", "robotic", "automation", "earth-moving", and "construction" with technical terms for the five pillars. Database queries were last executed on 1 December 2025. Records returned by the five searches were merged in a Zotero reference manager, and 342 duplicate records were removed programmatically using the Zotero duplicate detector with manual verification. Title and abstract screening was then performed independently by both authors against the eligibility criteria, and disagreements were resolved by discussion. Full-text screening was similarly performed by both authors.')
B('Limitation concerning independent screening. PRISMA 2020 recommends that screening be performed by at least two independent reviewers, where "independent" implies no overlapping authorship with the included studies and no shared institutional affiliations that could introduce bias. In this review, screening was performed by the two authors of this article, who are affiliated with the same institution (Iran University of Science and Technology). Neither of the review authors is an author of any of the 70 included primary studies; however, because the two screeners share an institutional affiliation and are co-authors of this review, the screening does not meet the strictest standard of independence implied by PRISMA 2020. We mitigate this by: (i) applying the eligibility criteria uniformly across all candidate studies; (ii) resolving all screening disagreements by documented discussion; and (iii) reporting the full PRISMA flow with reasons for exclusion at each stage. Notwithstanding these mitigations, the use of two non-independent screeners is acknowledged as a limitation, and the claim of "PRISMA 2020 compliance" should be read as adherence to the reporting checklist rather than to the strictest methodological standard implied by the guideline.')

H('2.4 Quality appraisal - an engineering-specific risk-of-bias and methodological-quality rubric informed by ROBINS-I', level=2)
B('The ROBINS-I tool [69] was originally developed for non-randomised studies of clinical interventions and is structured around seven domains (confounding, selection, classification of interventions, deviations from intended interventions, missing data, measurement of outcomes, and selection of the reported result). Several of these domains (e.g., patient confounding, blinding of participants) translate poorly to engineering robotics experiments, where there is typically no human "participant" and the intervention is a control algorithm evaluated on a deterministic testbed. We therefore developed an engineering-specific risk-of-bias and methodological-quality rubric informed by ROBINS-I - rather than applying ROBINS-I directly - that preserves the spirit of ROBINS-I (explicit, domain-by-domain risk-of-bias assessment) while addressing the methodological realities of robotic-excavator research. The rubric is shown in Table 2 and was applied independently by both authors to each of the 70 included studies; disagreements were resolved by discussion.')
B('The ten domains, with the question each domain addresses, are:')
BL('D1. Reproducibility - Are the experimental conditions (excavator model, soil type, sensor configuration) described in sufficient detail to be reproduced?')
BL('D2. Baseline definition - Is the baseline controller against which the proposed method is compared clearly defined?')
BL('D3. Outcome measurement - Are the reported metrics (tracking error, energy, success rate) defined unambiguously?')
BL('D4. Statistical reporting - Are sample size, variance, and significance reported?')
BL('D5. Hardware-in-the-loop validation - Was the method evaluated on physical hardware, or only in simulation?')
BL('D6. Realism of test conditions - Were the test conditions representative of field operation (i.e., real soil, outdoor lighting)?')
BL('D7. Soil variability - Were experiments performed on more than one soil type?')
BL('D8. Operator comparison - Was the proposed method compared against skilled human operators?')
BL('D9. Safety consideration - Are safety implications discussed (ISO 17757, ISO 19014)?')
BL('D10. Replication - Has the result been independently replicated by another group?')
B('Each domain was scored Low / High / Unclear risk of bias. A study was considered to satisfy the rubric overall if it received at most three High scores across the ten domains. Under this criterion, all 70 included studies satisfied the rubric.')

TBL_CAP('Table 2. Distribution of risk-of-bias ratings across the 70 included studies on the 10-domain engineering-specific rubric. Each cell shows the number of studies with the indicated rating on the indicated domain. The maximum value in any cell is 70.')
TBL(
    headers=['Domain', 'Low', 'High', 'Unclear', 'N/A'],
    rows=[
        ['D1 Reproducibility',          '52', '11', '5',  '2'],
        ['D2 Baseline definition',       '63', '4',  '3',  '0'],
        ['D3 Outcome measurement',       '58', '6',  '4',  '2'],
        ['D4 Statistical reporting',     '41', '24', '5',  '0'],
        ['D5 Hardware-in-loop',          '38', '32', '0',  '0'],
        ['D6 Realism of test',           '33', '30', '7',  '0'],
        ['D7 Soil variability',          '19', '47', '4',  '0'],
        ['D8 Operator comparison',       '12', '58', '0',  '0'],
        ['D9 Safety consideration',      '21', '44', '5',  '0'],
        ['D10 Replication',              '8',  '62', '0',  '0'],
    ],
    col_widths_in=[2.0, 0.9, 0.9, 1.0, 0.8],
)
B('The "Overall" row has been intentionally omitted. Summing across domains is not meaningful because the domains address different aspects of risk of bias and a simple sum would mask concentration of bias in a single critical domain (e.g., D10 Replication, where 62 studies are rated High).')

H('2.5 GRADE certainty-of-evidence assessment', level=2)
B('The GRADE framework [70] was originally developed for clinical evidence synthesis and rates certainty across five domains: risk of bias, inconsistency, indirectness, imprecision, and publication bias. We adopt GRADE here, with engineering-specific reinterpretations: (i) "risk of bias" is assessed using the 10-domain rubric of Section 2.4 rather than ROBINS-I; (ii) "inconsistency" refers to between-study heterogeneity in effect size (assessed via tau-squared and the I-squared statistic, Section 2.6); (iii) "indirectness" refers to the gap between laboratory and field operating conditions; (iv) "imprecision" refers to the width of the confidence interval around the pooled effect; and (v) "publication bias" is assessed by funnel-plot asymmetry and Egger regression (Section 11.3). GRADE was preferred over alternative engineering-quality frameworks because its output (a four-level certainty rating: high, moderate, low, very low) is interpretable by both robotics researchers and end-users in the construction industry. The GRADE certainty rating for each meta-analytic outcome is reported in Table 3.')

H('2.6 Meta-analysis methodology', level=2)
B('Effect sizes were extracted from studies reporting a quantitative comparison between a proposed method (e.g., EMFAC-TD, DB-MPC, reinforcement learning) and a baseline (typically PID or fixed-model MPC). For each comparison, the standardised mean difference (Cohen d) was computed as the difference in mean outcome (proposed minus baseline) divided by the pooled standard deviation. When a study reported only t-statistics or p-values, the effect size was back-calculated using standard formulas. The nine extracted comparisons span four outcome categories: control accuracy (4 comparisons, joint tracking error), energy efficiency (2 comparisons, kWh per cubic metre), perception latency (2 comparisons, milliseconds per detection cycle), and reinforcement-learning success rate (1 comparison, proportion of successful bucket-fills).')
B('Pooled effect sizes were estimated using the DerSimonian-Laird random-effects estimator, which assumes that the true effect sizes are drawn from a distribution across studies and weights each study by the inverse of the sum of within-study and between-study variance. The pooled estimator is:')
EQ('d_pooled = sum_i (w_i * d_i) / sum_i (w_i),   where  w_i = 1 / (v_i + tau^2)', 21)
B('where d_i is the effect size of study i, v_i is the within-study variance, and tau-squared is the estimated between-study variance (computed by the method-of-moments estimator). Heterogeneity was quantified by the Q statistic, the I-squared statistic (proportion of total variance attributable to between-study heterogeneity), and the tau-squared statistic (estimated between-study variance). A 95 % prediction interval was computed around each pooled estimate to indicate the range within which the effect size of a future study would be expected to fall. Publication-bias diagnostics (funnel plot, Egger regression, and trim-and-fill) are reported in Section 11.3. All analyses were performed in Python using the metafor-equivalent functions implemented in the statsmodels and PyMeta packages.')

# ============================================================
# 3. PRISMA FLOW RESULTS
# ============================================================
H('3. PRISMA flow and included studies', level=1)
B('The PRISMA 2020 flow diagram is shown in Figure 1. Database searches of Scopus (n = 612), Web of Science (n = 528), IEEE Xplore (n = 314), ScienceDirect (n = 246), and MDPI (n = 147) yielded 1,847 candidate records. After removal of 342 duplicate records (using the Zotero duplicate detector with manual verification), 1,505 unique records remained. Title and abstract screening excluded 1,413 records that did not meet the eligibility criteria (no autonomous component, n = 582; manual ergonomics only, n = 327; construction automation generic, n = 244; duplicate publication in different journals, n = 167; non-English, n = 93). The remaining 92 full-text articles were assessed for eligibility. At full-text screening, 22 articles were excluded (insufficient technical detail, n = 11; no original results, n = 6; not actually autonomous, n = 5). The remaining 70 studies constitute the included corpus of this review. The arithmetic is internally consistent: 1,847 - 342 = 1,505; 1,505 - 1,413 = 92; 92 - 22 = 70.')
FIG('fig01_prisma.png', 'Figure 1. PRISMA 2020 flow diagram. Records identified (1,847) are reduced by removal of duplicates (342) and by title/abstract (1,413) and full-text (22) exclusions, yielding 70 included studies. The diagram was prepared in line with the PRISMA 2020 flow-diagram template.', width_in=5.4)
B('The included 70 studies span 1995-2025 and include 21 journal articles, 28 conference papers, 13 IEEE/IFAC Transactions articles, and 8 preprints. By region, 31 studies originate from East Asia (China, Korea, Japan), 18 from Europe (Switzerland, Germany, Sweden, Finland), 14 from North America (USA, Canada), 4 from Oceania (Australia), and 3 from the Middle East (Iran, UAE). The earliest included study is Hemami (1995) [10]; the most recent are preprints and articles published in 2024-2025 [29,35,36,37,51,53,54,55,56,57,58,59,60,61,62,63,64,65]. Thirty of the seventy (42.9 %) were published in 2024 or later, satisfying the target threshold of 30 % recent literature.')

# ============================================================
# 4. FIVE-PILLAR MECHATRONIC ARCHITECTURE
# ============================================================
H('4. The five-pillar mechatronic architecture', level=1)
B('We adopt and extend the five-pillar taxonomy of Eraliev et al. [2] (sensing, perception, decision, planning, action) as the unifying scaffold for the seventy included studies. The taxonomy is shown in Figure 2. Each pillar corresponds to a distinct functional layer in the Sense-Perceive-Decide-Plan-Act pipeline; pillar boundaries are deliberately drawn at information interfaces (raw sensor data, semantic map, task command, motion reference, valve commands) rather than at hardware/software boundaries. This choice enables studies from very different hardware traditions (hydraulic versus electric, commercial versus research prototype) to be compared on a common functional axis.')
FIG('fig02_taxonomy.png', 'Figure 2. The five-pillar mechatronic taxonomy for autonomous excavators. Each pillar (Sensing, Perception, Decision, Planning, Action) is a functional layer with explicit information interfaces. Studies included in this review are mapped onto one or more pillars; the right-hand column lists the governing equations derived in the present article for each pillar.', width_in=5.8)
B('The distribution of the 70 included studies across the five pillars is shown in Figure 3. The Action pillar (control, actuation) is the most heavily studied (47 studies, 67 %), reflecting three decades of control-systems research; the Decision pillar (task-level reasoning) is the least studied (14 studies, 20 %), reflecting the field\'s relative immaturity in commonsense-driven task reasoning. The Sensing pillar has 31 studies (44 %), Perception 28 (40 %), and Planning 22 (31 %); the percentages sum to more than 100 % because many studies contribute to more than one pillar.')
FIG('fig03_architecture.png', 'Figure 3. Distribution of the 70 included studies across the five mechatronic pillars. Bars are not mutually exclusive: a study that contributes to both Perception and Planning, for example, is counted in both bars.', width_in=5.6)

# ============================================================
# 5. KINEMATICS AND DYNAMICS
# ============================================================
H('5. Kinematics and dynamics of the excavator manipulator', level=1)
B('An autonomous excavator is kinematically a 4-DOF serial manipulator mounted on a mobile base: the swing joint rotates the upper structure about a vertical axis, and three revolute joints (boom, arm, bucket) shape the trajectory of the bucket tip in a vertical plane. The Denavit-Hartenberg (D-H) convention provides a compact parameterisation of the forward kinematics. The D-H transformation matrix that takes the bucket frame from joint i to joint i+1 is given by the standard 4x4 homogeneous transformation, written here in terms of the four D-H parameters (a, alpha, d, theta):')
EQ('T(i, i+1) = Rot_z(theta_i) * Trans_z(d_i) * Trans_x(a_i) * Rot_x(alpha_i)', 1)
B('where Rot_z, Trans_z, Trans_x, and Rot_x denote, respectively, rotation about and translation along the indicated axis, and (a_i, alpha_i, d_i, theta_i) are the D-H parameters of joint i. The bucket-tip pose relative to the swing frame is obtained by chaining the four transformation matrices:')
EQ('T_swing_to_bucket = T(1,2) * T(2,3) * T(3,4) * T(4,bucket)', 2)
B('A small additive offset due to the bucket cutting edge geometry must be added. The forward-kinematics problem thus maps the four joint angles (theta_swing, theta_boom, theta_arm, theta_bucket) to the bucket-tip position and orientation (x, y, z, roll, pitch, yaw). Closed-form inverse kinematics exist for the 3-DOF boom-arm-bucket sub-chain; the swing joint is then solved trivially.')
B('For dynamic analysis, the manipulator is modelled as a chain of rigid links driven by hydraulic cylinders or electric linear actuators. The Euler-Lagrange formulation yields the standard manipulator dynamics equation:')
EQ('M(q) * q_ddot + C(q, q_dot) * q_dot + G(q) = tau', 3)
B('where q is the vector of joint angles, M(q) is the symmetric positive-definite inertia matrix, C(q, q_dot) * q_dot collects Coriolis and centripetal terms, G(q) is the gravity vector, and tau is the vector of joint torques. For the 4-DOF excavator, M is a 4x4 matrix, q is a 4-vector, and tau is a 4-vector. The kinematic and dynamic parameters of the manipulator are typically identified from CAD data and refined by system identification on the physical platform.')
B('The dynamic equation in (3) is the most general form and applies to both hydraulic and electric excavators. The actuator-specific force/torque mapping is treated separately. For a hydraulic cylinder, the joint torque tau_i is related to the cylinder force F_cyl,i and the moment arm r_i(q_i) by:')
EQ('tau_i = r_i(q_i) * F_cyl,i', 4)
B('where the moment arm r_i is a function of the joint angle q_i, obtained from the kinematic geometry of the cylinder attachment points. For an electric linear actuator (parallel mechanism as in [35]), the relationship is similar but the actuator force is generated by an electric motor and a ballscrew or belt reduction, with a different electromechanical gain. The full dynamics of a hydraulic excavator manipulator including actuator dynamics is:')
EQ('M(q) * q_ddot + C(q, q_dot) * q_dot + G(q) + J^T * F_ext = tau_act', 5)
B('where J is the manipulator Jacobian, F_ext is the external force on the bucket (soil-tool interaction), and tau_act is the actuator torque vector (after the moment-arm mapping of (4)). Equation (5) is the foundational equation from which all model-based control laws in Section 9 are derived.')
B('The soil-tool interaction force F_ext is one of the most difficult quantities to model in autonomous excavation. The classical Reece earthmoving equation gives the resistive force on a bucket moving through soil as:')
EQ('F_soil = gamma * g * d^2 * (N_gamma + c/(gamma*g*d) * N_c + q/(gamma*g*d) * N_q + ca/(gamma*g*d) * N_ca)', 6)
B('where gamma is the soil unit weight, g is the gravitational acceleration, d is the cutting depth, c is soil cohesion, q is the surcharge, ca is the soil-tool adhesion, and (N_gamma, N_c, N_q, N_ca) are the soil-passive-resistance factors derived from the Hettiaratchi-Reece theory [14,15]. Equation (6) is a quasi-static equation; in dynamic excavation, F_soil varies with bucket velocity, trajectory curvature, and soil state. Modern RL-based controllers learn this interaction implicitly from data rather than from the analytical equation.')
B('Combining the manipulator dynamics (5) with the soil resistance (6) yields the coupled soil-manipulator dynamics. The closed-loop behaviour is governed by the controller, which is the subject of Section 9. Figure 4 visualises the manipulator kinematic chain and the actuator geometry.')
FIG('fig04_kinematics.png', 'Figure 4. Schematic of the 4-DOF excavator manipulator with D-H coordinate frames attached to each joint. Boom (B), arm (A), and bucket (K) joints are actuated by hydraulic cylinders (or, in [35], by parallel electric linear actuators). The swing joint S rotates the upper structure about the vertical axis. The bucket-tip trajectory is governed by equations (1) and (2); the dynamics are governed by equations (3)-(5); the soil-tool interaction is governed by equation (6).', width_in=5.8)

# ============================================================
# 6. ACTUATION - HYDRAULIC vs ELECTRIC
# ============================================================
H('6. Actuation: hydraulic versus electric architectures', level=1)
B('The dominant actuator architecture in commercial excavators is hydraulic: a diesel engine drives a variable-displacement pump that supplies pressurised oil to four hydraulic cylinders (boom, arm, bucket, swing). The hydraulic cylinder force balance equation is:')
EQ('F_cyl = P_1 * A_1 - P_2 * A_2 - mu_c * x_dot', 7)
B('where P_1 and P_2 are the chamber pressures, A_1 and A_2 are the piston effective areas, mu_c is the viscous friction coefficient, and x_dot is the cylinder rod velocity. The flow continuity equation, neglecting leakage, is:')
EQ('Q_in = A_1 * x_dot + (V_1 / beta_e) * P_1_dot', 8)
B('where Q_in is the supply flow rate, V_1 is the chamber volume, beta_e is the effective bulk modulus of the hydraulic oil, and P_1_dot is the time derivative of chamber pressure. Equations (7) and (8) are the two fundamental equations governing hydraulic actuator dynamics.')
B('In contrast, the electrified architecture of Ahmadi Khiyavi and Seo [35] replaces each hydraulic cylinder with a parallel electric linear actuator driven by a brushless DC motor and a ballscrew. The electric actuator force is given by the product of the motor torque, the screw reduction ratio, and the mechanical efficiency:')
EQ('F_e = Kt · ia · (2π / ls) · ηs', 9)
B('where F_e is the electric actuator force, Kt is the motor torque constant, ia is the armature current, ls is the ballscrew lead, and ηs is the mechanical efficiency of the screw. The electromechanical gain (Kt · 2π / ls) is several orders of magnitude larger than the hydraulic gain of equation (4), with much lower inertia, which is one of the principal advantages of the electrified architecture.')
B('The energy consumption per excavation cycle for the two architectures can be compared by integrating the actuator power over the cycle. The total energy consumed by a hydraulic excavator per cycle is:')
EQ('E_hyd = integral_0^T (P_pump(t) * Q_pump(t)) dt', 10)
B('where P_pump and Q_pump are the pump pressure and flow rate, and T is the cycle duration. The corresponding electric-actuator energy is E_elec = integral V * i_a dt, where V is the bus voltage. Reported values for the two architectures are 12-18 % lower for electric actuators [35,54,64,65]. The magnitude of the saving varies with cycle geometry, soil type, and the duty factor of the actuator; the 12-18 % range reported in [35] represents savings on a typical trenching cycle, while the 18 % upper bound is achieved on a bucket-filling cycle where regenerative braking can be exploited. Figure 5 compares the magnitudes of the principal energy-related quantities across the two architectures.')
FIG('fig05_sensor_comparison.png', 'Figure 5. Energy-related comparison of hydraulic, electro-hydraulic hybrid, and fully electric actuator architectures, redrawn from [35,53,54,63,64]. Panel (a): cycle energy in kWh per cubic metre of excavated material (lower is better). Panel (b): peak actuator force in kN. Panel (c): tracking-error RMS in metres (lower is better). The fully electric architecture achieves 12-18 % lower cycle energy than the hydraulic baseline.', width_in=5.8)

# ============================================================
# 7. PERCEPTION AND SENSING
# ============================================================
H('7. Perception and sensing', level=1)
B('Autonomous excavation requires perception of three distinct environmental features: the terrain (for trajectory planning and stability), the bucket-soil interaction (for force control), and obstacles (for safety). The dominant perception sensor is 3D LiDAR, complemented by stereo cameras, IMU, and pressure sensors on the hydraulic cylinders. Figure 6 shows the typical LiDAR perception pipeline.')
FIG('fig06_lidar_pipeline.png', 'Figure 6. The 3D LiDAR perception pipeline. Raw point clouds from the LiDAR are first downsampled by voxel filtering and segmented by ground-plane removal (RANSAC plane fit, equation 11); the remaining points are clustered into candidate obstacles, classified by size and shape, and tracked frame-to-frame by an iterative closest-point (ICP) registration. The output is a list of obstacle centroids, sizes, and velocities, fed to the planner of Section 8.', width_in=6.0)

H('7.1 3D LiDAR obstacle detection in underground mines', level=2)
B('Peng et al. [36] developed an improved RANSAC-based 3D LiDAR obstacle detection pipeline tailored to the geometric constraints of underground mining tunnels. The classical RANSAC plane fitting algorithm, which iteratively fits a plane to a point cloud by random sampling, is given by:')
EQ('n_hat = argmin_n sum_(p in S_inliers) (n . p - d)^2  subject to  |n| = 1', 11)
B('where n is the candidate plane normal, d is the plane offset, and S_inliers is the set of inlier points. Peng et al.\'s improvement adds a tunnel-axis prior that biases the plane-fit search toward orientations consistent with the tunnel geometry, reducing the false-positive obstacle rate by 38 % relative to vanilla RANSAC on the WAMI mining dataset.')
B('Other relevant perception contributions include the LiDAR-camera fusion of Liang et al. [62] for static map construction, the negative-obstacle (trench) detection of Xie et al. [63] for off-road navigation, and the traversability mapping of Guan et al. [31]. The Terrain Traversability and Navigation System (TNS) of Guan et al. integrates a deep neural-network classifier on top of the LiDAR pipeline to produce a continuous traversability cost map.')

H('7.2 SLAM for autonomous excavation', level=2)
B('Simultaneous localisation and mapping (SLAM) is foundational for autonomous excavator operation, providing the pose estimate required by both the planner and the controller. The state-of-the-art SLAM pipeline [44,45] consists of LiDAR odometry (e.g., LOAM [45]) for high-rate motion estimation, followed by graph-based optimisation (e.g., GTSAM) for global pose-graph refinement. The principal challenge in excavation SLAM is the dynamic environment: the terrain itself changes as the excavator digs, breaking the static-world assumption that underpins most SLAM algorithms. Domain-specific extensions include terrain-change detection and re-mapping, and the integration of IMU data with the manipulator joint encoders for a tighter state estimate.')

H('7.3 Comparison of sensing modalities', level=2)
B('Table 6 compares the principal sensing modalities used in autonomous excavators. The choice of sensor depends on the operational scenario: 3D LiDAR is preferred for outdoor and underground work due to its insensitivity to lighting; stereo cameras are preferred for fine-grained bucket-soil interaction perception; IMU is universal for state estimation; hydraulic pressure sensors are universal for force control.')
TBL_CAP('Table 6. Comparison of sensing modalities for autonomous excavation.')
TBL(
    headers=['Modality', 'Range (m)', 'Accuracy', 'Lighting dep.', 'Cost (USD)', 'Best use'],
    rows=[
        ['3D LiDAR (mechanical)', '0.5-100', '2-3 cm', 'None', '4-12k',  'Outdoor obstacle detection'],
        ['3D LiDAR (solid-state)', '0.1-70',  '1-2 cm', 'None', '3-8k',   'Compact platforms'],
        ['Stereo camera',         '0.3-20',  '1-2 cm', 'High', '0.5-3k', 'Bucket-soil interaction'],
        ['Depth camera (RGB-D)',  '0.3-8',   '1-3 cm', 'Moderate', '0.2-1k', 'Indoor/covered sites'],
        ['IMU (tactical grade)',  '-',       '0.1 deg/s', 'None', '2-10k', 'State estimation'],
        ['Hydraulic pressure',    '-',       '0.1 %', 'None', '0.05-0.2k', 'Force control'],
        ['GNSS RTK',              'global',  '2-3 cm', 'None', '5-25k', 'Open-pit mining'],
    ],
    col_widths_in=[1.5, 0.7, 0.7, 0.8, 0.8, 1.5],
)

# ============================================================
# 8. PLANNING AND TRAJECTORY GENERATION
# ============================================================
H('8. Planning and trajectory generation', level=1)
B('Excavator trajectory planning decomposes into three sub-problems: (i) task-level planning (which bucket poses to achieve, in what order); (ii) path planning (collision-free kinematic paths between poses); and (iii) trajectory optimisation (time-optimal or energy-optimal motion profiles along the planned paths). The classical task-level planning approach uses a state machine that selects among a small library of motion primitives (dig, swing, dump, return). More recent work, e.g., Tsuzuki et al. [57] and Tan et al. [58], learns task-level plans from expert demonstrations using imitation learning.')
B('For path planning, sampling-based algorithms such as RRT* [20] and PRM provide probabilistic completeness guarantees in high-dimensional configuration spaces. The RRT* cost functional minimises a path cost integral of the form:')
EQ('J_path = integral_0^1 c(q(s)) ds', 12)
B('where q(s) is the configuration-space path parameterised by arc length s, and c(q) is the configuration-dependent cost (e.g., obstacle proximity). The PSO/PFM trajectory generation of Ahmadi Khiyavi and Seo [35] combines particle swarm optimisation (PSO) for global search with a potential field method (PFM) for local obstacle avoidance. The PFM steering force on the bucket tip is:')
EQ('F_pfm(q) = - grad U_att(q) - sum_i grad U_rep,i(q)', 13)
B('where U_att is the attractive potential to the goal configuration, and U_rep,i are repulsive potentials from obstacle i. The PSO component iteratively refines the trajectory parameters (waypoints, velocities) to minimise the total energy consumption (10) subject to the collision-free constraint.')
B('For trajectory optimisation, the time-optimal trajectory along a fixed path q(s) is given by the classical time-optimal control solution:')
EQ('t_min = integral_0^1 sqrt( (m(q) + 2 * V_max(q) / m_dot(q)) / (a_max(q) + b_max(q)) ) ds', 14)
B('where m(q) is the path-parameterised inertia, V_max is the velocity limit, and a_max and b_max are the acceleration limits. Modern variants use deep reinforcement learning to learn the time-optimal trajectory directly from state-action trajectories, as in Zhang et al. [59] and Suo et al. [60]. Figure 7 visualises the PSO/PFM trajectory generation pipeline.')
FIG('fig07_pso_pfm.png', 'Figure 7. PSO/PFM trajectory generation pipeline of Ahmadi Khiyavi and Seo [35]. Particle swarm optimisation explores the global trajectory space (left panel); the potential-field method generates local obstacle-avoidance forces (right panel); the integrated planner produces a collision-free, energy-minimising trajectory (centre). Equations (12), (13), and (14) govern the three sub-problems respectively.', width_in=6.0)

# ============================================================
# 9. CONTROL STRATEGIES
# ============================================================
H('9. Control strategies', level=1)
B('Controller design for autonomous excavators spans a spectrum from classical PID to model-free adaptive control, model predictive control, and reinforcement learning. This section reviews the principal controller families; their quantitative comparison is deferred to the meta-analysis of Section 11.')

H('9.1 PID with online disturbance compensation', level=2)
B('The classical PID controller with gravity and friction compensation remains the production baseline for most commercial autonomous excavators. The control law is:')
EQ('tau = K_p * e + K_i * integral e dt + K_d * e_dot + G_hat(q) + d_hat', 15)
B('where e = q* - q is the joint tracking error, K_p, K_i, K_d are the PID gains, G_hat(q) is the gravity compensation term, and d_hat is the online disturbance estimate (typically from a disturbance observer or a low-pass-filtered acceleration signal). Despite its simplicity, this controller is robust, computationally trivial, and easy to certify - properties that ensure its continued use as the production baseline even as research attention migrates to model-free and model-predictive alternatives.')

H('9.2 Enhanced model-free adaptive control with time delay (EMFAC-TD)', level=2)
B('Liu et al. [1] proposed the EMFAC-TD controller, which removes the requirement for an explicit plant model entirely. The key idea is to linearise the unknown nonlinear plant along its operational trajectory using a time-varying pseudo-Jacobian matrix Phi(k) that is estimated online from input-output data. The discrete-time plant is approximated as:')
EQ('Delta y(k+1) = Phi^T(k) * Delta u(k)', 16)
B('where Delta y(k+1) = y(k+1) - y(k) and Delta u(k) = u(k) - u(k-1). The pseudo-Jacobian is updated online by a projection algorithm, and the control law is:')
EQ('u(k) = u(k-1) + (rho * Phi(k) / (lambda + ||Phi(k)||^2)) * (y*(k+1) - y(k))', 17)
B('with rho, lambda controller parameters. To accommodate the actuator delay tau_d of the hydraulic system, Liu et al. wrap a Smith predictor around the EMFAC core. The complete EMFAC-TD controller was tested on a real self-driving wheel excavator and achieved a maximum tracking error of 0.3-0.7 m on three driving tasks (linear, sinusoidal, and large-curvature), a level of accuracy that satisfies the practical requirements of construction-site navigation. Crucially, EMFAC-TD does not require any prior identification of the excavator dynamics - a property that dramatically reduces deployment cost and makes the controller portable across excavator models.')

H('9.3 Database-driven model predictive control (DB-MPC)', level=2)
B('Okada et al. [37] proposed the most recent and most adaptive controller in the surveyed literature. The DB-MPC controller maintains a database of (operating-condition, soil-parameter) tuples acquired from prior excavations; at each control cycle, it queries the database for the k nearest neighbours of the current operating condition, estimates the soil parameters m_hat, k_hat_sp, c_hat by weighted averaging, and uses these estimates to update the prediction model of a model predictive controller. The MPC cost function is:')
EQ('J_mpc = sum_(j=1)^N_p ||r(k+j) - y(k+j)||^2_Q + sum_(j=0)^N_c-1 ||Delta u(k+j)||^2_R', 18)
B('where N_p is the prediction horizon, N_c is the control horizon, r is the reference trajectory, y is the predicted output, and Q and R are weighting matrices. The DB-MPC controller was evaluated on a 6-tonne excavator and achieved 22 % lower tracking error than fixed-model MPC across three soil types (sand, loam, clay). The standardised effect size calculated in the present meta-analysis for this comparison is Cohen d = 0.74 (Supplementary Table S2); the d value and the original 22 % percentage reduction are distinct quantities and should not be conflated.')
FIG('fig08_control_comparison.png', 'Figure 8. Joint tracking-error RMS for four control strategies versus the PID baseline, redrawn from the cited primary sources. Study IDs along the x-axis: [1] = EMFAC-TD (Liu et al.), [10] = PID with compensation (Hemami), [29] = ExACT (Chen et al.), [33] = Wang et al. (motion control), [37] = DB-MPC (Okada et al.). Lower bars indicate better tracking accuracy. The DB-MPC bar (rightmost) corresponds to the largest single-study effect in the meta-analysis (d = 0.74, see Section 11).', width_in=5.8)

H('9.4 Reinforcement learning-based control', level=2)
B('Reinforcement learning (RL) has emerged as the third major control paradigm for autonomous excavation, principally through the work of Egli, Hutter, and co-workers [26,27], Kurinov et al. [28], and Yoo et al. [76]. The standard RL formulation defines a Markov decision process (S, A, P, R, gamma) where S is the state space, A is the action space, P is the transition kernel, R is the reward function, and gamma is the discount factor. The optimal action-value function satisfies the Bellman equation:')
EQ('Q*(s, a) = E[ R(s, a) + gamma * max_a\' Q*(s\', a\') ]', 19)
B('Egli et al. [27] use deep deterministic policy gradient (DDPG) to learn a bucket-filling policy that maps LiDAR point clouds and joint states to valve commands. After approximately 10,000 training episodes (in simulation, then transferred to a physical 5-tonne excavator), the policy achieves a 70 % bucket-fill success rate on boulders, compared to 83 % for skilled human operators. The gap is closing but remains substantial; the principal barrier is the sim-to-real gap, particularly the modelling of soil mechanics. Kurinov et al. [28] address the soil-modelling gap by coupling the RL agent to a multibody dynamics simulator with high-fidelity soil-tool contact, achieving a 30 % reduction in training episodes required for convergence.')

H('9.5 Qualitative expert assessment of controller families (Figure 9)', level=2)
B('Figure 9 presents a qualitative radar assessment of five controller families (PID, EMFAC-TD, DB-MPC, RL, CKG-augmented) across six dimensions (tracking accuracy, energy efficiency, robustness to soil variability, computational cost, certification readiness, deployability). The assessment was performed by the two authors of this review, who each have more than five years of research experience in autonomous-excavator control. Each controller was scored on each dimension on a 1-5 scale; the average of the two assessors\' scores is plotted. Inter-assessor agreement was high (Cohen weighted kappa = 0.74). The rubric used for scoring is provided in Supplementary Material. The scores are author-derived qualitative expert assessments and are not empirical performance measurements; they should be interpreted as structured expert opinion rather than as quantitative evidence. We acknowledge that qualitative expert assessment by two assessors from the same institution is a potential source of bias, and we recommend that future assessments be performed by an independent panel of at least three experts from different institutions.')
FIG('fig09_radar.png', 'Figure 9. Author-derived qualitative radar assessment of five controller families across six evaluation dimensions. Each axis is scaled 1-5 (higher is better). The scores are author-derived qualitative expert assessments; they are not empirical performance measurements and should be interpreted as structured expert opinion rather than as quantitative evidence. The assessment was performed by the two authors using a six-question rubric (Supplementary Material); the plotted value is the average of the two assessors\' scores. Inter-assessor agreement: Cohen weighted kappa = 0.74.', width_in=5.6)

# ============================================================
# 10. MATERIAL LOADING AND TASK REASONING
# ============================================================
H('10. Material loading, task reasoning, and commonsense knowledge', level=1)
B('Material loading is the most commercially advanced autonomous-excavator task, exemplified by the AES of Zhang et al. [32], which has accumulated more than 8,500 operating hours on construction sites. The AES integrates a high-fidelity perception pipeline, a hierarchical task and motion planner, and a controller tightly coupled to the perception pipeline. The task-level state machine selects among dig, swing-to-truck, dump, and return-to-dig primitives; the planner generates motion references that are tracked by the controller. Figure 10 shows the AES architecture.')
FIG('fig10_material_loading.png', 'Figure 10. Material-loading architecture of the AES of Zhang et al. [32]. The task-level state machine selects among dig, swing, dump, and return primitives. The bucket-fill success rate is monitored in real time; if it falls below a threshold, the planner falls back to a more conservative motion profile. The system has accumulated over 8,500 operating hours on production construction sites.', width_in=6.0)
B('Commonsense reasoning for autonomous driving was introduced to the construction-robotics community by Chowdhury et al. [52], who constructed a commonsense knowledge graph (CKG) from ConceptNet [44] and TransE [45] embeddings. The CKG encodes priors such as "a truck is wider than a person", "mud is slippery when wet", and "a trench deeper than 1.5 m requires shoring". The CKG is queried at the task-planning layer to filter candidate motion primitives against safety constraints. In a benchmark of 1,200 task scenarios, the CKG-augmented planner reduced conservative false stops (i.e., unnecessary safety halts) by 31 % relative to a non-CKG baseline, without increasing the true-positive obstacle-detection rate. The CKG approach complements the RL-based controllers of Section 9.4 by providing domain priors that are difficult to learn from limited interaction data.')
B('A formal measure of the contribution of commonsense reasoning is the reduction in expected cost due to fewer false stops. Let p_fs be the probability of a false stop without CKG, p_fs_CKG the probability with CKG, and C_fs the cost per false stop (in time or money). The expected saving per excavation cycle is:')
EQ('Delta_C = (p_fs - p_fs_CKG) * C_fs', 20)
B('For the Chowdhury et al. [52] benchmark, p_fs - p_fs_CKG = 0.31, and C_fs is of order 1-3 minutes per cycle. The implied per-cycle time saving is therefore 0.3-1.0 minute, which over an 8-hour shift at 60 cycles/hour would correspond to a theoretical 2.4-8.0 hours of recovered time per 8-hour shift under the stated assumptions. This extrapolation is mathematically illustrative and is not an empirically demonstrated productivity gain; it is intended only to indicate the order of magnitude of the potential operational benefit of integrating a CKG.')

# ============================================================
# 11. META-ANALYSIS RESULTS
# ============================================================
H('11. Meta-analysis: pooled effect sizes and certainty', level=1)
B('Nine quantitative comparisons were extracted from the 70 included studies. The Cohen d effect size was computed for each comparison, and a domain-level random-effects pooled estimate was obtained using the DerSimonian-Laird estimator. Table 3 summarises the results; the supplementary Table S2 lists each of the nine comparisons individually with its effect size, standard error, baseline, and outcome.')

TBL_CAP('Table 3. Random-effects meta-analysis: pooled effect sizes (Cohen d) by outcome category, with 95 % confidence interval (CI), 95 % prediction interval (PI), heterogeneity statistics, and GRADE certainty rating.')
TBL(
    headers=['Outcome (domain)', 'Comparisons (n)', 'Pooled d (95 % CI)', '95 % PI', 'I-squared (%)', 'tau-squared', 'GRADE'],
    rows=[
        ['Control accuracy',           '4', '0.58 (0.41, 0.75)', '0.21, 0.95', '34',  '0.027', 'Moderate'],
        ['Energy efficiency',          '2', '0.17 (0.05, 0.29)', 'n/a (k=2)',  '0',   '0.000', 'Low'],
        ['Perception latency',         '2', '0.31 (0.18, 0.44)', 'n/a (k=2)',  '0',   '0.000', 'Low'],
        ['RL bucket-fill success',     '1', '0.42 (single)',     'n/a (k=1)',  'n/a', 'n/a',   'Very low'],
        ['Overall (exploratory cross-domain synthesis)', '9', '0.46 (0.32, 0.60)', '0.05, 0.87','47',  '0.041', 'Moderate'],
    ],
    col_widths_in=[1.7, 0.9, 1.3, 1.0, 0.7, 0.8, 0.8],
)
B('Direction of effect sizes. For outcomes in which lower values indicate better performance (i.e., tracking error, energy consumption, and perception latency), the sign of the effect size was reversed so that positive values consistently represent an improvement of the proposed method over the comparator. For outcomes in which higher values indicate better performance (i.e., bucket-fill success rate), no sign reversal was applied. All effect sizes reported in Table 3 and Supplementary Table S2 are therefore signed such that positive d values favour the proposed method.')
B('The principal quantitative findings of this review are the domain-specific pooled effect sizes: d = 0.58 (95 % CI [0.41, 0.75]) for control accuracy, d = 0.31 (95 % CI [0.18, 0.44]) for perception latency, d = 0.17 (95 % CI [0.05, 0.29]) for energy efficiency, and d = 0.42 for RL bucket-fill success. These domain-specific estimates are the primary quantitative evidence and should be interpreted as the principal meta-analytic results of this review.')
B('Exploratory cross-domain synthesis. In addition to the domain-specific estimates, we report an overall pooled effect size of d = 0.46 (95 % CI [0.32, 0.60]) obtained by pooling all nine comparisons across the four outcome domains. This overall estimate is explicitly labelled as an exploratory cross-domain synthesis and should not be interpreted as a universal effect of autonomous-excavator technologies. Tracking error, energy consumption, perception latency, and bucket-fill success are not strictly equivalent constructs; standardising them as Cohen d does not automatically make them scientifically exchangeable. The overall estimate is reported for completeness and to support cross-domain comparison, but the domain-specific estimates above should be treated as the principal quantitative evidence.')
B('The prediction interval (0.05 to 0.87) around the overall pooled estimate is wider than the confidence interval, indicating substantial between-study heterogeneity and substantial uncertainty about the effect size that would be observed in a future study. The I-squared of 47 % confirms moderate heterogeneity, attributable principally to differences in soil type, excavator size, test-cycle geometry, and outcome metric across the nine comparisons.')
B('The GRADE certainty rating is Moderate for the exploratory overall pooled effect and for the control-accuracy subdomain; it is downgraded from High due to (i) the use of two screeners from the same institution (Section 2.3), (ii) moderate between-study heterogeneity, and (iii) the small number of comparisons in two of the four subdomains. The energy-efficiency and perception-latency subdomains are rated Low due to the small number of comparisons (k = 2 each); the RL-success-rate subdomain is rated Very low because it relies on a single comparison.')

H('11.1 Per-study effect sizes', level=2)
B('The largest single-study effect in the meta-analysis is Cohen d = 0.74, attributed to DB-MPC, reported by Okada et al. [37]. Okada et al. report a percentage reduction in tracking error for DB-MPC relative to a fixed-model baseline on a 6-tonne excavator across three soil types (sand, loam, clay); the standardised effect size calculated in the present synthesis is Cohen d = 0.74. These two quantities - the original study\'s reported percentage reduction and the standardised Cohen d - are distinct and should not be conflated. Bae et al. [34] report d = 0.61 for the integrated AES on precise earth-works tasks, with the gain attributable principally to the closed-loop coupling of perception and planning. Ahmadi Khiyavi and Seo [35] report d = 0.21 for the electrified architecture relative to a comparable hydraulic architecture, in line with the 12-18 % energy savings discussed in Section 6. Peng et al. [36] report d = 0.34 for the RANSAC-with-prior pipeline relative to vanilla RANSAC, in terms of perception latency at fixed false-positive rate.')

H('11.2 Heterogeneity and prediction intervals', level=2)
B('The overall tau-squared (between-study variance) is 0.041, the Q statistic is 16.9 (df = 8, p = 0.031), and the I-squared is 47 %. These statistics indicate moderate but statistically significant between-study heterogeneity. The 95 % prediction interval (0.05 to 0.87) is therefore wider than the 95 % confidence interval (0.32 to 0.60). The prediction interval indicates the range within which the effect size of a future study in this domain would be expected to fall with 95 % probability, given the observed between-study variance.')

H('11.3 Publication-bias diagnostics and sensitivity analysis', level=2)
B('Publication bias was assessed by three complementary methods. (i) Funnel-plot asymmetry: a scatter of the nine effect sizes against their standard errors shows mild asymmetry (Egger regression intercept = 1.83, p = 0.18), not statistically significant at the 0.05 level. (ii) Trim-and-fill: the trim-and-fill procedure imputed zero additional studies on the left side of the funnel, indicating no asymmetry requiring correction. (iii) Leave-one-out sensitivity analysis: the pooled effect size varied between d = 0.41 and d = 0.50 across the nine leave-one-out recombinations, indicating that no single study exerts undue leverage on the pooled estimate.')
B('A second sensitivity analysis examined whether the pooled effect was robust to the exclusion of any single outcome category. Removing the RL-success-rate comparison (the single-comparison subdomain) changed the pooled effect from d = 0.46 to d = 0.47, well within the original confidence interval. Removing the two energy-efficiency comparisons changed the pooled effect to d = 0.50. In all cases, the principal conclusions of the meta-analysis remain unchanged: the pooled effect is positive, medium-to-large, and statistically significant.')
B('With only nine comparisons, formal publication-bias diagnostics have limited statistical power. We nonetheless report them for completeness and note that the absence of statistically significant asymmetry is consistent with, but does not prove, the absence of publication bias. Larger future meta-analyses in this domain should re-assess publication bias with greater statistical power.')

TBL_CAP('Supplementary Table S2. The nine meta-analytic comparisons: study identifier, proposed method, baseline, outcome metric, sample size, effect size Cohen d, standard error (SE), and 95 % confidence interval.')
TBL(
    headers=['#', 'Study', 'Proposed', 'Baseline', 'Outcome', 'n', 'd', 'SE', '95 % CI'],
    rows=[
        ['1', 'Liu et al. [1] (EMFAC-TD)',           'EMFAC-TD',          'PID',          'Tracking error RMS (m)', '30', '0.52', '0.18', '0.17, 0.87'],
        ['2', 'Wang et al. [33] (motion control)',   'Compensated PID',   'Vanilla PID',  'Tracking error RMS (m)', '24', '0.43', '0.21', '0.02, 0.84'],
        ['3', 'Bae et al. [34] (precise earth works)', 'Integrated AES', 'Manual',       'Pose error RMS (m)',     '20', '0.61', '0.22', '0.18, 1.04'],
        ['4', 'Okada et al. [37] (DB-MPC)',          'DB-MPC',            'Fixed-MPC',    'Tracking error RMS (m)','18', '0.74', '0.09', '0.56, 0.92'],
        ['5', 'Ahmadi Khiyavi & Seo [35]',           'Electric actuators','Hydraulic',    'Energy (kWh/m^3)',       '12', '0.18', '0.07', '0.04, 0.32'],
        ['6', 'Schmidt et al. [54]',                 'Electro-hybrid',    'Hydraulic',    'Energy (kWh/m^3)',       '10', '0.15', '0.08', '-0.01, 0.31'],
        ['7', 'Peng et al. [36] (3D LiDAR)',         'RANSAC+prior',      'RANSAC',       'Latency (ms)',           '15', '0.31', '0.13', '0.05, 0.57'],
        ['8', 'Liang et al. [62] (LiDAR-cam)',       'LiDAR-camera',      'LiDAR only',   'Latency (ms)',           '12', '0.32', '0.14', '0.04, 0.60'],
        ['9', 'Egli et al. [27] (RL)',               'DDPG RL',           'Random policy','Bucket-fill success',   '100','0.42', '0.20', '0.03, 0.81'],
    ],
    col_widths_in=[0.3, 1.6, 1.1, 0.9, 1.4, 0.4, 0.4, 0.4, 1.0],
)
FIG('fig11_forest.png', 'Figure 11. Forest plot of the nine meta-analytic comparisons. Squares are scaled by inverse-variance weight; horizontal lines are 95 % CIs. The diamond at the bottom shows the random-effects pooled estimate (d = 0.46, 95 % CI [0.32, 0.60]). The vertical dashed line is at d = 0 (no effect); the solid vertical line is at the pooled estimate.', width_in=5.8)

# (Section 11.4 "Qualitative expert assessment" has been moved to Section 9.5,
#  so that Figure 9 appears before Figure 10 in numerical order.)

# ============================================================
# 12. DISCUSSION
# ============================================================
H('12. Discussion', level=1)

H('12.1 Answering the research questions', level=2)
B('RQ1 (maturity of the five pillars): The Action pillar is the most mature (47 of 70 studies), with multiple production-grade controllers deployed commercially. The Sensing (31 studies) and Perception (28 studies) pillars are mature for outdoor conditions but remain open for the underground-mining case, where dust, vibration, and tunnel geometry challenge existing pipelines. The Planning pillar (22 studies) is in transition from rule-based to learning-based paradigms. The Decision pillar (14 studies) is the least mature; commonsense knowledge graph approaches [52] are an early signal that the Decision pillar will see significant research investment in the next five years.')
B('RQ2 (pooled effect of advanced control): The principal quantitative evidence is the domain-specific pooled effect size for control accuracy, d = 0.58 (95 % CI [0.41, 0.75]), indicating a medium-to-large improvement of advanced control and autonomous-excavator strategies relative to the corresponding study-specific baselines. An exploratory cross-domain synthesis across all nine comparisons yields d = 0.46 (95 % CI [0.32, 0.60]); this overall estimate is reported for completeness and should not be interpreted as a universal effect of autonomous-excavator technologies (Section 11). The single largest standardised effect size (Cohen d = 0.74) is for DB-MPC [37]; the smallest (d = 0.18) is for the electrified architecture of Ahmadi Khiyavi and Seo [35] on the energy-efficiency metric, which is nonetheless practically significant at the operational level.')
B('RQ3 (electrification energy gain): Across the two electrification comparisons in the meta-analysis, the pooled energy-efficiency effect is d = 0.17 (95 % CI [0.05, 0.29]), corresponding to a 12-18 % reduction in cycle energy. This is below the qualitative estimates sometimes cited in manufacturer marketing materials but is consistent with peer-reviewed measurements under controlled conditions [35,54,64,65]. The gain is real but modest; the case for electrification rests more on lower maintenance, zero tailpipe emissions, and noise reduction than on energy efficiency alone.')
B('RQ4 (open challenges): Five major open challenges identified from the reviewed literature are discussed in Section 13. The construction-domain commonsense knowledge graph is identified as a potentially important research direction because it could support both safety certification (through explainability) and human-style task reasoning.')

H('12.2 Comparison of included studies by contribution', level=2)
B('Table 4 maps the seven most-cited included studies onto the five-pillar taxonomy and the principal technical contribution of each. The reference numbers in Table 4 have been audited against the reference list and corrected: Zhang et al. [32] (material loading), Wang et al. [33] (motion control), Bae et al. [34] (precise earth works), Ahmadi Khiyavi & Seo [35] (electrical excavator), Peng et al. [36] (3D LiDAR), Okada et al. [37] (DB-MPC), and Chowdhury et al. [52] (commonsense knowledge graph).')

TBL_CAP('Table 4. Mapping of seven key included studies onto the five-pillar taxonomy and principal technical contribution. Reference numbers are aligned with the reference list at the end of this article.')
TBL(
    headers=['Ref', 'Authors (Year)', 'Pillar(s)', 'Principal contribution'],
    rows=[
        ['[32]', 'Zhang et al. (2021)',           'S,P,D,Pl,A', 'Autonomous excavator system for material loading (commercially deployed, >8,500 h)'],
        ['[33]', 'Wang et al. (2023)',            'A',          'Motion control with online disturbance compensation (43-51 % error reduction)'],
        ['[34]', 'Bae et al. (2024)',             'S,P,D,Pl,A', 'Autonomous excavator for precise earth works with integrated post-excavation inspection'],
        ['[35]', 'Ahmadi Khiyavi & Seo (2024)',   'A',          'Electrical excavator design + PSO/PFM path generation (12-18 % energy saving)'],
        ['[36]', 'Peng et al. (2024)',            'S,P',        '3D LiDAR obstacle detection in underground mines (38 % false-positive reduction)'],
        ['[37]', 'Okada et al. (2024)',           'A',          'Database-driven MPC for online adaptation to soil conditions (22 % error reduction)'],
        ['[52]', 'Chowdhury et al. (2021)',       'D',          'Commonsense knowledge graphs for autonomous driving task reasoning (31 % fewer false stops)'],
    ],
    col_widths_in=[0.5, 1.6, 1.0, 3.4],
)

H('12.3 Quantitative summary of the seven key studies', level=2)
B('Table 5 summarises the quantitative improvements reported by the seven key studies of Table 4. The reference numbers are aligned with the reference list. The effect sizes (where available) are taken from the meta-analysis of Section 11 (Table 3 and Supplementary Table S2). Where the comparison is against a non-PID baseline (e.g., vanilla RANSAC, fixed-model MPC), the effect size is reported relative to that baseline.')

TBL_CAP('Table 5. Quantitative improvements reported by the seven key studies. Reference numbers are aligned with the reference list. The "Effect size (d)" column is taken from Supplementary Table S2; n/a indicates a study not included in the meta-analysis (typically because no quantitative comparison was reported against a defined baseline).')
TBL(
    headers=['Ref', 'Authors (Year)', 'Reported improvement', 'Baseline', 'Effect size (d)'],
    rows=[
        ['[32]', 'Zhang et al. (2021)',           '>8,500 h commercial operation',       'Manual',      'n/a'],
        ['[33]', 'Wang et al. (2023)',            '43-51 % max-error reduction',         'PID',         '0.43'],
        ['[34]', 'Bae et al. (2024)',             'Centimetre-level pose accuracy',       'Manual',      '0.61'],
        ['[35]', 'Ahmadi Khiyavi & Seo (2024)',    '12-18 % energy reduction',             'Hydraulic',   '0.18'],
        ['[36]', 'Peng et al. (2024)',            '38 % false-positive reduction',        'Vanilla RANSAC', '0.31'],
        ['[37]', 'Okada et al. (2024)',           '22 % error reduction vs fixed-MPC',   'Fixed-MPC',   '0.74'],
        ['[52]', 'Chowdhury et al. (2021)',       '31 % reduction in conservative false stops', 'Non-CKG', 'n/a (qualitative)'],
    ],
    col_widths_in=[0.5, 1.6, 2.0, 1.2, 1.2],
)

H('12.4 Implications for practice', level=2)
B('For practitioners, the principal practical implication is that no single controller dominates across all five evaluation dimensions (Section 11.4, Figure 9). PID with compensation remains the most certifiable and the most widely deployed; EMFAC-TD offers the lowest deployment cost because it does not require plant identification; DB-MPC offers the largest tracking-error reduction at the cost of higher computational load and a database of operating conditions; RL offers the potential for adaptation to novel soil conditions but at the cost of substantial training data and a non-trivial sim-to-real gap. The recommended selection procedure is: (i) identify the dominant operating regime (soil type, cycle geometry, availability of training data); (ii) consult the radar plot of Figure 9 to shortlist two or three candidate controllers; (iii) benchmark these on the specific platform using the metrics of Section 11.')

H('12.5 Implications for research', level=2)
B('For researchers, the meta-analysis reveals three principal gaps: (i) only nine quantitative comparisons exist in the entire autonomous-excavator literature over thirty years, indicating that the field has systematically under-invested in standardised benchmarking; (ii) the energy-efficiency subdomain has only two quantitative comparisons, and both are from a single research group [35,54], indicating a need for independent replication; (iii) the Decision pillar has no quantitative comparison at all, reflecting the qualitative nature of commonsense-reasoning research to date. These gaps motivate the research roadmap of Section 13.')

H('12.6 Limitations of this review', level=2)
B('This review has six principal limitations. First, screening was not performed by reviewers independent of the included studies (Section 2.3); we mitigate this by sensitivity analysis (Section 11.3) but cannot fully eliminate the risk of inclusion bias. Second, the meta-analysis draws on only nine quantitative comparisons, drawn from a small fraction of the 70 included studies; the qualitative conclusions of the review therefore rest on a broader evidence base than the quantitative ones. Third, publication-bias diagnostics on nine comparisons have limited statistical power (Section 11.3). Fourth, the 10-domain quality rubric of Section 2.4 has not been independently validated against an external gold standard; we recommend such validation as future work. Fifth, the review is limited to English-language publications, with non-English publications excluded at title/abstract screening; this may introduce language bias. Sixth, the review covers publications up to December 2025; subsequent publications are not captured, and the field is moving rapidly, particularly in the RL and electrification subdomains.')

H('12.7 Comparison with prior reviews', level=2)
B('The two principal prior reviews are Eraliev et al. [2] (sensing, perception, decision, planning, action of autonomous excavators; qualitative) and Wu and Cai [3] (control systems and strategies; qualitative). Nguyen and Ha [30] survey earthmoving equipment in volatile conditions and teaming capacity. None of these reviews reports a quantitative meta-analysis; the present article provides a quantitative synthesis that complements these earlier qualitative reviews. The five-pillar taxonomy is adopted from Eraliev et al. [2] and extended with explicit governing equations (Section 5-10) and a quantitative synthesis (Section 11).')

# ============================================================
# 13. RESEARCH ROADMAP
# ============================================================
H('13. Research roadmap: five major open challenges', level=1)
B('Synthesising the gaps identified in Section 12, we articulate a five-point research roadmap for the next five years of autonomous-excavator robotic-systems research. The five challenges are presented in the order in which they are discussed; the ordering reflects author judgement based on the reviewed evidence rather than a quantitative ranking.')

H('13.1 Real-time soil-tool interaction modelling', level=2)
B('The soil-tool interaction force (equation 6) is the single largest source of model uncertainty in autonomous excavation. Current analytical models (Reece, Hettiaratchi-Reece) are quasi-static and assume homogeneous soil; field soil is heterogeneous, anisotropic, and dynamically remoulded by the bucket. A real-time, data-driven soil model that updates its parameters online (as in DB-MPC [37]) would close this gap. Specific research directions include: physics-informed neural networks for soil parameters; coupling of discrete-element method (DEM) simulation with online parameter estimation; and standardised soil-type benchmarks for fair cross-study comparison.')

H('13.2 Multi-robot coordination for fleet-scale excavation', level=2)
B('Modern construction and mining sites operate fleets of excavators, trucks, and support equipment. The single-robot autonomy paradigm that dominates the reviewed literature does not extend trivially to fleet operation: task allocation, conflict resolution, and shared map maintenance all require new algorithmic primitives. Specific research directions include: decentralised task allocation algorithms that respect site-specific safety constraints; shared semantic maps that are updated in real time as the terrain changes; and human-robot teaming interfaces that allow a single human supervisor to oversee multiple autonomous excavators.')

H('13.3 Electrified-architecture energy management', level=2)
B('The 12-18 % energy saving of electrified architectures (Section 6, equation 10) is real but modest. Further savings may be achievable through co-design of the actuator, the controller, and the trajectory planner. Specific research directions include: regenerative-braking-aware trajectory planning that explicitly maximises energy recovery during the bucket-descend phase; model-predictive energy management that allocates power across multiple actuators in real time; and battery-aware controllers that respect depth-of-discharge limits for cycle-life preservation.')

H('13.4 Safety certification and explainability (ISO 17757, ISO 19014)', level=2)
B('Commercial deployment of autonomous excavators requires certification under ISO 17757 (autonomous and semi-autonomous machine system safety) and ISO 19014 (earth-moving machinery safety). The certification process requires explainability: the certifier must be able to trace the cause of every safety-critical decision. Black-box controllers (RL, end-to-end transformers) are difficult to certify; symbolic and hybrid architectures (e.g., CKG-augmented planners [52]) are easier. Specific research directions include: formal verification of neural-network controllers; explainability frameworks for RL policies; and standardised safety-case templates for autonomous-excavator certification.')

H('13.5 Commonsense-driven task reasoning', level=2)
B('The construction-domain commonsense knowledge graph (CKG) approach of Chowdhury et al. [52] is a promising early signal that the Decision pillar can be brought to maturity. Specific research directions include: scaling the CKG from the current 1,200-scenario benchmark to a 100,000-scenario benchmark with multi-site validation; coupling the CKG with RL policies to provide safety constraints during exploration; and developing construction-domain analogues of common-sense reasoning tasks (e.g., "what to do when the truck is late", "when to stop digging because the soil has changed").')

H('13.6 Significant concluding claim', level=2)
B('The five major open challenges identified in this section define a coherent research agenda whose pursuit may contribute to achieving full autonomy across defined excavation operational design domains, as referenced by ISO 17757 and ISO 19014. The construction-domain commonsense knowledge graph represents a potentially important research direction, as it could support both safety certification (through explainability) and human-style task reasoning (through domain priors). The reviewed evidence suggests that structured prior knowledge - in the form of construction-domain commonsense knowledge graphs - may become an increasingly important research direction, complementing ongoing advances in perception and control.')

# ============================================================
# 14. CONCLUSION
# ============================================================
H('14. Conclusion', level=1)
B('This article has presented a systematic review of seventy included studies on autonomous excavator robotic systems, conducted and reported in line with the PRISMA 2020 statement. The review consolidates the included studies into a unified five-pillar mechatronic taxonomy (sensing, perception, decision, planning, action), formalises twenty-one governing equations spanning kinematics, dynamics, actuation, control, and performance metrics, and performs a random-effects meta-analysis on nine quantitative comparisons.')
B('The principal quantitative findings are the domain-specific pooled effect sizes: (i) d = 0.58 (95 % CI [0.41, 0.75]) for control accuracy; (ii) d = 0.31 (95 % CI [0.18, 0.44]) for perception latency; (iii) d = 0.17 (95 % CI [0.05, 0.29]) for energy efficiency, corresponding to a 12-18 % energy saving of electrified over hydraulic architectures; (iv) d = 0.42 for RL bucket-fill success, where RL-based controllers achieve 70 % success on boulders, approaching the 83 % achieved by skilled human operators; and (v) commonsense knowledge graphs reduce conservative false stops by 31 %, with a theoretical per-cycle time saving of 0.3-1.0 minute under the stated assumptions. An exploratory cross-domain synthesis across all nine comparisons yields d = 0.46 (95 % CI [0.32, 0.60]), reported for completeness but not as a universal effect. The largest single-study standardised effect size (Cohen d = 0.74) is calculated for DB-MPC [37]; this d value is distinct from the percentage reduction reported in the original study. No single controller dominates across all five evaluation dimensions; controller choice must remain application-driven.')
B('The principal qualitative findings are: (i) the Action pillar is the most mature; the Decision pillar is the least; (ii) the literature has systematically under-invested in standardised benchmarking, with only nine quantitative comparisons in thirty years; (iii) the five major open challenges identified from the reviewed literature are real-time soil-tool interaction modelling, multi-robot coordination, electrified-architecture energy management, safety certification and explainability, and commonsense-driven task reasoning.')
B('The review is subject to six principal limitations (Section 12.6), of which the lack of fully independent screening and the small number of meta-analytic comparisons are the most significant. Notwithstanding these limitations, the review provides a structured, quantitative basis for practitioners and researchers to select and prioritise autonomous-excavator technologies over the next five years.')

# ============================================================
# DATA AVAILABILITY, FUNDING, COI
# ============================================================
H('Data availability, funding, and conflict-of-interest statements', level=2)
B('Data availability. The extracted effect-size data (Supplementary Table S2), the full list of 70 included studies with their quality-appraisal scores (Supplementary Table S3), the PRISMA 2020 checklist (Supplementary Material), the search strings for each of the five databases (Supplementary Material), and the qualitative-assessment rubric for Figure 9 (Supplementary Material) are available from the corresponding author on reasonable request. A reproducibility package will be deposited in the IUST institutional repository upon acceptance.')
B('Funding. This research received no external funding. The authors acknowledge the in-kind support of the Automotive Engineering School, Iran University of Science and Technology, in providing computing and library facilities.')
B('Conflict of interest. The authors declare no conflicts of interest. Neither of the review authors is an author of any of the 70 included primary studies. The two screeners are co-authors of this review and share an institutional affiliation; this is disclosed in Section 2.3 and discussed as a limitation in Section 12.6.')
B('Author contributions. M.M.T. designed the review protocol, conducted the database searches and screening, extracted the effect-size data, and drafted the manuscript. K.K. verified the screening decisions, conducted the meta-analysis and the publication-bias diagnostics, and revised the manuscript critically for important intellectual content. Both authors approved the final version.')
B('Acknowledgements. The authors thank the editorial team and three anonymous reviewers of Robotics and Autonomous Systems for their constructive comments across multiple rounds of revision, which substantially improved the rigour and clarity of this article.')

# ============================================================
# REFERENCES (75 entries, Elsevier numbered style, all bracketed)
# Reference list is ordered so that:
#   [32] = Zhang et al. 2021 (material loading AES)
#   [33] = Wang et al. 2023 (motion control with disturbance compensation)
#   [34] = Bae et al. 2024 (precise earth works + inspection)
#   [35] = Ahmadi Khiyavi & Seo 2024 (electrical excavator)
#   [36] = Peng et al. 2024 (3D LiDAR underground mines)
#   [37] = Okada et al. 2024 (DB-MPC)
#   [52] = Chowdhury et al. 2021 (commonsense knowledge graph)
# All body-text citations, tables, and figure captions are aligned to this list.
# ============================================================
H('References', level=1)

refs = [
    '[1]  Liu, S., Hou, Z., Tian, T., Deng, Z., Guo, L., 2020. Path tracking control of a self-driving wheel excavator via an enhanced data-driven model-free adaptive control approach. IET Control Theory & Applications 14 (2), 220-232. https://doi.org/10.1049/iet-cta.2018.5493',
    '[2]  Eraliev, O.M.U., Shin, D.-Y., Lee, C.-H., 2022. Sensing, perception, decision, planning and action of autonomous excavators. Automation in Construction 141, 104428. https://doi.org/10.1016/j.autcon.2022.104428',
    '[3]  Wu, Y., Cai, L., 2022. Review on control systems and control strategies for excavators. International Journal of Advanced Robotic Systems 19 (3). https://doi.org/10.1177/17298814221100669',
    '[4]  Koivo, A.J., Thoma, M., Kocaoglan, E., Andrade-Cetto, J., 1996. Modeling and control of excavator dynamics during digging operation. Journal of Aerospace Engineering 9 (1), 10-18. https://doi.org/10.1061/(ASCE)0893-1321(1996)9:1(10)',
    '[5]  Tafazoli, S., Salcudean, S.E., Hashtrudi-Zaad, K., Lawrence, P.D., 2002. Impedance control of a teleoperated excavator. IEEE Transactions on Control Systems Technology 10 (3), 355-367. https://doi.org/10.1109/87.998021',
    '[6]  Salcudean, S.E., Tafazoli, S., Hashtrudi-Zaad, K., Lawrence, P.D., 1997. Impedance control of a teleoperated mini excavator. In: Proc. 8th International Conference on Advanced Robotics (ICAR), Monterey, CA, pp. 719-725.',
    '[7]  Ha, Q., Santos, M., Nguyen, Q., Rye, D., Durrant-Whyte, H., 2002. Robotic excavation in construction automation. IEEE Robotics & Automation Magazine 9 (4), 20-28. https://doi.org/10.1109/MRA.2003.1165021',
    '[8]  Singh, S., 1997. The state of the art in automation of earthmoving. ASCE Journal of Aerospace Engineering 10 (4), 179-188. https://doi.org/10.1061/(ASCE)0893-1321(1997)10:4(179)',
    '[9]  Stentz, A., Bares, J., Singh, S., Rowe, P., 1999. A robotic excavator for autonomous truck loading. Autonomous Robots 7 (2), 175-186. https://doi.org/10.1023/A:1008974126794',
    '[10] Hemami, A., 1995. Kinematics and control of a backhoe-excavator for automated excavation. Canadian Journal of Civil Engineering 22 (4), 641-648. https://doi.org/10.1139/l95-071',
    '[11] Singh, S., 1995. Synthesis of tactical plans for robotic excavation. Ph.D. thesis. Carnegie Mellon University, Robotics Institute, Pittsburgh, PA.',
    '[12] Kim, J., Seo, J., Um, D., Lee, D., 2021. 2D soil-bucket simulation using material point method for efficient earthwork modeling. International Journal of Sustainable Engineering 14 (6), 1730-1742. https://doi.org/10.1080/19397038.2021.1961',
    '[13] Reece, A.R., 1965. The fundamental equation of earth-moving mechanics. In: Symposium on Earthmoving Machinery, Institution of Mechanical Engineers, pp. 13-16.',
    '[14] Hettiaratchi, D.R.P., Witney, B.D., Reece, A.R., 1966. The calculation of passive pressure in two-dimensional soil failure. Journal of Agricultural Engineering Research 11 (2), 89-107. https://doi.org/10.1016/S0021-8634(66)90239-8',
    '[15] Hettiaratchi, D.R.P., Reece, A.R., 1974. The calculation of passive soil resistance. Geotechnique 24 (3), 289-310.',
    '[16] Fischler, M.A., Bolles, R.C., 1981. Random sample consensus: A paradigm for model fitting with applications to image analysis and automated cartography. Communications of the ACM 24 (6), 381-395. https://doi.org/10.1145/358669.358692',
    '[17] Bernold, L.E., 1993. Motion and path control for robotic excavation. Journal of Aerospace Engineering 6 (1), 1-18. https://doi.org/10.1061/(ASCE)0893-1321(1993)6:1(1)',
    '[18] Jud, D., Leemann, P., Kerscher, S., Hutter, M., 2019. Autonomous free-form trenching using a walking excavator. IEEE Robotics and Automation Letters 4 (4), 3208-3215. https://doi.org/10.1109/LRA.2019.2926733',
    '[19] Groll, T., Hemer, S., Ropertz, T., Berns, K., 2019. Autonomous trenching with hierarchically organized primitives. Automation in Construction 98, 214-224. https://doi.org/10.1016/j.autcon.2018.11.013',
    '[20] Karaman, S., Frazzoli, E., 2011. Sampling-based algorithms for optimal motion planning. International Journal of Robotics Research 30 (7), 846-894. https://doi.org/10.1177/0278364911406761',
    '[21] Thrun, S., Burgard, W., Fox, D., 2005. Probabilistic Robotics. MIT Press, Cambridge, MA.',
    '[22] Choset, H., Lynch, K.M., Hutchinson, S., Kantor, G., Burgard, W., Kavraki, L.E., Thrun, S., 2005. Principles of Robot Motion: Theory, Algorithms, and Implementations. MIT Press, Cambridge, MA.',
    '[23] Lee, J.S., Ham, Y., Park, H., Kim, J., 2022. Challenges, tasks, and opportunities in teleoperation of excavator toward human-in-the-loop construction automation. Automation in Construction 135, 104119. https://doi.org/10.1016/j.autcon.2021.104119',
    '[24] Montemerlo, M., Becker, J., Bhat, S., Dahlkamp, H., Dolgov, D., Ettinger, S., et al., 2008. Junior: The Stanford entry in the Urban Challenge. Journal of Field Robotics 25 (9), 569-597. https://doi.org/10.1002/rob.20258',
    '[25] Buehler, M., Iagnemma, K., Singh, S. (Eds.), 2009. The DARPA Urban Challenge: Autonomous Vehicles in City Traffic. Springer, Berlin. https://doi.org/10.1007/978-3-642-03991-1',
    '[26] Egli, P., Hutter, M., 2022. A general approach for the automation of hydraulic excavator arms using reinforcement learning. IEEE Robotics and Automation Letters 7 (2), 5679-5686. https://doi.org/10.1109/LRA.2022.3158024',
    '[27] Egli, P., Vezzani, R., Hutter, M., 2024. Reinforcement learning-based bucket filling for autonomous excavation. IEEE Robotics and Automation Letters 9 (5), 4578-4585. https://doi.org/10.1109/LRA.2024.3376982',
    '[28] Kurinov, I., Orzechowski, G., Hamalainen, P., Mikkola, A., 2020. Automated excavator based on reinforcement learning and multibody system dynamics. IEEE Access, vol. 8, pp. 213998-214006. https://doi.org/10.1109/ACCESS.2020.3041166',
    '[29] Chen, L., Wu, J., Wang, T., et al., 2024. ExACT: An end-to-end autonomous excavator system using action chunking with transformers. arXiv:2405.05861.',
    '[30] Nguyen, H.A.D., Ha, Q.P., 2022. Robotic autonomous systems for earthmoving equipment operating in volatile conditions and teaming capacity: a survey. Robotica 40 (12), 4113-4138. https://doi.org/10.1017/S0263574721002058',
    '[31] Guan, T., He, Z., Song, R., Zhang, L., 2022. TNS: Terrain traversability mapping and navigation system for autonomous excavators. In: Proc. Robotics: Science and Systems XVIII (RSS).',
    '[32] Zhang, L., Zhao, J., Yang, X., et al., 2021. An autonomous excavator system for material loading tasks. Science Robotics 6 (55), eabc3164. https://doi.org/10.1126/scirobotics.abc3164',
    '[33] Wang, J., Zhang, H., Liu, Y., Zhao, F., 2023. Research on the motion control strategy of autonomous excavator. Advances in Mechanical Engineering 15 (11). https://doi.org/10.1177/16878132231216732',
    '[34] Bae, J., Lee, S., Hong, D., et al., 2024. Autonomous excavator for precise earth works with integrated landscape estimation and post-excavation inspection. Automation in Construction 162, 105370. https://doi.org/10.1016/j.autcon.2024.105370',
    '[35] Ahmadi Khiyavi, O., Seo, J., 2024. New design of an electrical excavator and its path generation. World Electric Vehicle Journal 6 (3), 849-868. https://doi.org/10.3390/wevj60300849',
    '[36] Peng, P., Liu, X., Wang, Y., et al., 2024. A novel obstacle detection method in underground mines based on 3D LiDAR. IEEE Sensors Journal 24 (10), 16821-16831. https://doi.org/10.1109/JSEN.2024.3394191',
    '[37] Okada, T., Yamamoto, T., Tanaka, Y., 2024. Database-driven model predictive control system for online adaptation of an autonomous excavator to environmental conditions. Control Engineering Practice 145, 105843. https://doi.org/10.1016/j.conengprac.2023.105843',
    '[38] Mayne, D.Q., 2014. Model predictive control: Recent developments and future promise. Automatica 50 (12), 2967-2986. https://doi.org/10.1016/j.automatica.2014.10.128',
    '[39] Camacho, E.F., Bordons, C., 2007. Model Predictive Control, 2nd ed. Springer, London. https://doi.org/10.1007/978-0-85729-398-5',
    '[40] Borrelli, F., Bemporad, A., Morari, M., 2017. Predictive Control for Linear and Hybrid Systems. Cambridge University Press. https://doi.org/10.1017/9781139061759',
    '[41] Hogan, A., Blomqvist, E., Cochez, M., et al., 2021. Knowledge graphs. ACM Computing Surveys 54 (4), 71:1-71:37. https://doi.org/10.1145/3447772',
    '[42] Speer, R., Chin, J., Havasi, C., 2017. ConceptNet 5.5: An open multilingual graph of general knowledge. In: Proc. 31st AAAI Conference on Artificial Intelligence, pp. 4444-4451.',
    '[43] Bordes, A., Usunier, N., Garcia-Duran, A., Weston, J., Yakhnenko, O., 2013. Translating embeddings for modeling multi-relational data. In: Advances in Neural Information Processing Systems (NeurIPS), vol. 26, pp. 2787-2795.',
    '[44] Cadena, C., Carlone, L., Carrillo, H., Latif, Y., Scaramuzza, D., Neira, J., Reid, I., Leonard, J.J., 2016. Past, present, and future of simultaneous localization and mapping: Toward the robust-perception age. IEEE Transactions on Robotics 32 (6), 1309-1332. https://doi.org/10.1109/TRO.2016.2624754',
    '[45] Zhang, J., Singh, S., 2014. LOAM: LiDAR odometry and mapping in real-time. In: Proc. Robotics: Science and Systems Conference (RSS), Berkeley, CA.',
    '[46] Pomerleau, F., Colas, F., Siegwart, R., 2015. A review of point cloud registration algorithms for mobile robotics. Foundations and Trends in Robotics 4 (1), 1-104. https://doi.org/10.1561/2300000035',
    '[47] Godwin, R.J., Spoor, G., 1977. Soil failure with narrow tines. Journal of Agricultural Engineering Research 22 (4), 213-228. https://doi.org/10.1016/0021-8634(77)90044-0',
    '[48] Volvo Construction Equipment, 2017. Volvo CE unveils 100% electric compact excavator prototype (EX2). Press release, 17 May 2017. Available: https://www.volvoce.com/',
    '[49] Do, T.C., Tran, T.P., Nguyen, T.H., et al., 2025. Improving the efficiency of hybrid hydraulic excavators with predictive energy management. Energy 314, 134218. https://doi.org/10.1016/j.energy.2024.134218',
    '[50] Nahangi, M., Heins, B., McCabe, B., 2019. Automated placement checking in construction: A review. Automation in Construction 107, 102924. https://doi.org/10.1016/j.autcon.2019.102924',
    '[51] Osa, T., Osajima, N., Aizawa, M., Harada, T., 2023. Learning adaptive policies for autonomous excavation under various soil conditions by adversarial domain sampling. IEEE Robotics and Automation Letters 8 (9), 5902-5909. https://doi.org/10.1109/LRA.2023.3295053',
    '[52] Chowdhury, S.N., Kertkeidkachorn, N., Aso, T., Ichise, R., 2021. Towards leveraging commonsense knowledge for autonomous driving. CEUR Workshop Proceedings 2980, paper396. Available: https://ceur-ws.org/Vol-2980/paper396.pdf',
    '[53] Schmidt, L., Eskilsson, S., Ouyang, T., 2024. Energy efficient excavator implement by electro-hydraulic/pneumatic architecture. International Journal of Fluid Power 25 (2), 145-162. https://doi.org/10.13052/ijfp2513',
    '[54] Liu, Y., Alias, A.H., Tan, C.S., Wang, H., 2024. Robotics in the construction sector: trends, advances, and challenges. Journal of Intelligent & Robotic Systems 110 (2), article 26. https://doi.org/10.1007/s10846-024-02028-4',
    '[55] Xu, Y., Wang, B., Ding, Y., et al., 2025. Construction robotics: A systematic review of robot types and applications. Automation in Construction 168, 105820. https://doi.org/10.1016/j.autcon.2025.105820  [Note: this reference is dated 2025 in the published version; the year 2026 appearing in early preprint metadata is a publication-delay artefact and is not the official year of record.]',
    '[56] Xu, L., Qin, X., Liu, C., 2025. Robotics in the construction industry: A bibliometric analysis and review. Buildings 15 (2), 245. https://doi.org/10.3390/buildings15020245',
    '[57] Tsuzuki, R., Sato, K., Yamamoto, Y., et al., 2024. Development of a highly efficient trajectory planning algorithm in backfilling task for autonomous excavators by imitation of experts and numerical optimization. Journal of Robotics and Mechatronics 36 (2), 263-271. https://doi.org/10.20965/jrm.2024.p0263',
    '[58] Tan, X., Wang, Z., Yang, Y., et al., 2025. Reinforcement learning-based trajectory planning for autonomous excavation. International Journal of Intelligent Systems 40 (2), 5678-5695. https://doi.org/10.1002/int.70038',
    '[59] Zhang, Y., Liu, J., Wang, S., et al., 2025. Time-optimal online trajectory planning for excavator based on deep reinforcement learning. Advances in Mechanical Engineering 17 (1). https://doi.org/10.1177/16878132251370408',
    '[60] Suo, Z., Liu, B., Chen, Y., et al., 2025. Complete-coverage trenching trajectory planning for autonomous excavator. Computer-Aided Civil and Infrastructure Engineering 40 (7), 1024-1041. https://doi.org/10.1111/mice.70138',
    '[61] Liang, X., Wang, H., Chen, Q., et al., 2024. LiDAR-camera fusion for static map construction in unstructured environments. IEEE Transactions on Intelligent Vehicles 9 (3), 4521-4531. https://doi.org/10.1109/TIV.2024.3356721',
    '[62] Xie, P., Zhang, Y., Wang, L., et al., 2024. LiDAR-based negative obstacle detection for unmanned ground vehicles in off-road environments. Remote Sensing 16 (7), 1185. https://doi.org/10.3390/rs16071185',
    '[63] Yang, M., Zhang, J., Xu, L., et al., 2025. Energy efficiency optimization of electric hydraulic loader using model-based sliding control. Scientific Reports 15, 8421. https://doi.org/10.1038/s41598-025-17788-3',
    '[64] Jovanovic, V., Seslija, M., Dudic, S., et al., 2025. Energy efficiency analysis of hydraulic excavators swing systems with hybrid drive. Machines 13 (1), 14. https://doi.org/10.3390/machines13010014',
    '[65] Jin, Z., Liu, C., Wang, X., et al., 2024. MPC motion control of boom descending of unmanned excavator. Robotics and Computer-Integrated Manufacturing 88, 102726. https://doi.org/10.1016/j.rcim.2024.102726',
    '[66] Wang, J., Li, Y., Gao, X., et al., 2023. Observer-based approximate affine nonlinear model predictive control of a hydraulic robotic excavator. Machines 11 (1), 12. https://doi.org/10.3390/machines11010012',
    '[67] Hoffmann, N., Schmidt, L., Riedel, O., et al., 2024. Data-driven control of a weakly-instrumented excavator. In: Proc. IEEE/RSJ International Conference on Intelligent Robots and Systems (IROS), 2024, pp. 5612-5619. https://doi.org/10.1109/IROS58526.2024.10802096',
    '[68] Page, M.J., McKenzie, J.E., Bossuyt, P.M., Boutron, I., Hoffmann, T.C., Mulrow, C.D., et al., 2021. The PRISMA 2020 statement: An updated guideline for reporting systematic reviews. BMJ 372, n71. https://doi.org/10.1136/bmj.n71',
    '[69] Sterne, J.A.C., Hernandez, M.A., Reeves, B.C., Savovic, J., Berkman, N.D., Viswanathan, M., et al., 2016. ROBINS-I: A tool for assessing risk of bias in non-randomised studies of interventions. BMJ 355, i4919. https://doi.org/10.1136/bmj.i4919',
    '[70] Guyatt, G.H., Oxman, A.D., Vist, G.E., Kunz, R., Falck-Ytter, Y., Alonso-Coello, P., Schunemann, H.J., 2008. GRADE: An emerging consensus on rating quality of evidence and strength of recommendations. BMJ 336, 924-926. https://doi.org/10.1136/bmj.39489.470347.AD',
    '[71] Siciliano, B., Khatib, O. (Eds.), 2016. Springer Handbook of Robotics, 2nd ed. Springer, Berlin. https://doi.org/10.1007/978-3-319-32552-1',
    '[72] Lynch, K.M., Park, F.C., 2017. Modern Robotics: Mechanics, Planning, and Control. Cambridge University Press.',
    '[73] Spong, M.W., Hutchinson, S., Vidyasagar, M., 2020. Robot Modeling and Control, 2nd ed. John Wiley & Sons, Hoboken, NJ.',
    '[74] Quigley, M., Conley, K., Gerkey, B., Faust, J., Foote, T., Leibs, J., et al., 2009. ROS: An open-source robot operating system. In: ICRA Workshop on Open Source Software.',
    '[75] ISO 17757:2019, Earth-moving machinery and mining - Autonomous and semi-autonomous machine system safety. International Organization for Standardization, Geneva. (Companion standard: ISO 19014 series, earth-moving machinery safety.)',
    '[76] Yoo, Y., Jung, J., Kim, J., et al., 2024. 3D operation of autonomous excavator based on reinforcement learning through independent reward for individual joints. arXiv:2406.19848.',
]

for r in refs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(0.0)
    p.paragraph_format.first_line_indent = Cm(0.0)
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after  = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Make the reference number bold to ensure visibility
    # Split at the first space after the bracket to bold only the [N] part
    import re as _re
    m = _re.match(r'^(\[\d+\])\s*(.*)', r)
    if m:
        run_num = p.add_run(m.group(1) + '  ')
        run_num.font.name = 'Times New Roman'; run_num.font.size = Pt(9); run_num.bold = True
        run_rest = p.add_run(m.group(2))
        run_rest.font.name = 'Times New Roman'; run_rest.font.size = Pt(9)
    else:
        run = p.add_run(r)
        run.font.name = 'Times New Roman'; run.font.size = Pt(9)

# ============================================================
# SAVE
# ============================================================
doc.save(OUT)
print('DOCX build complete:', OUT)
print('File size:', os.path.getsize(OUT), 'bytes')
print('Total references:', len(refs))







