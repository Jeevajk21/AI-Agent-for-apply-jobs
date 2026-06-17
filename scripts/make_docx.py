"""
Generate Word .docx files for Deloitte Internal Audit & Assurance application.
Produces a CV and a Cover Letter matching the formatting spec in prompts/.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

BLACK = RGBColor(0, 0, 0)
FONT = "Calibri"


def set_font(run, size, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = BLACK


def para_space(para, before=0, after=0, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        from docx.shared import Pt as _Pt
        pf.line_spacing = _Pt(line)


def add_hline(doc):
    """Insert a horizontal rule by adding a bottom border to a blank paragraph."""
    p = doc.add_paragraph()
    para_space(p, before=2, after=2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def set_narrow_margins(doc):
    for section in doc.sections:
        section.top_margin = Cm(1.27)
        section.bottom_margin = Cm(1.27)
        section.left_margin = Cm(1.27)
        section.right_margin = Cm(1.27)


# ── CV ────────────────────────────────────────────────────────────────────────

def build_cv():
    doc = Document()
    set_narrow_margins(doc)

    # ── Name header ──
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_space(name_p, after=2)
    r = name_p.add_run("JEEVAKUMAR JAYACHANDRAN")
    set_font(r, 15, bold=True)

    # ── Contact line ──
    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_space(contact_p, after=4)
    r = contact_p.add_run(
        "Lancaster, UK  |  jeevajk2112@gmail.com  |  +44 7466 480480  |  linkedin.com/in/jeevakumar-j"
    )
    set_font(r, 10)

    add_hline(doc)

    def section_heading(text):
        p = doc.add_paragraph()
        para_space(p, before=6, after=3)
        r = p.add_run(text)
        set_font(r, 11, bold=True)
        return p

    def body(text, indent=False):
        p = doc.add_paragraph()
        para_space(p, after=2)
        if indent:
            p.paragraph_format.left_indent = Cm(0.4)
        r = p.add_run(text)
        set_font(r, 10)
        return p, r

    def bullet(text):
        p = doc.add_paragraph(style="List Bullet")
        para_space(p, after=2)
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        r = p.add_run(text)
        set_font(r, 10)
        return p

    def role_line(title, org_loc, dates):
        p = doc.add_paragraph()
        para_space(p, before=4, after=1)
        r1 = p.add_run(title)
        set_font(r1, 10, bold=True)
        r2 = p.add_run(f"  |  {org_loc}  |  ")
        set_font(r2, 10)
        r3 = p.add_run(dates)
        set_font(r3, 10, bold=False)
        r3.font.italic = True
        return p

    def proj_heading(text):
        p = doc.add_paragraph()
        para_space(p, before=4, after=1)
        r = p.add_run(text)
        set_font(r, 10, bold=True)
        return p

    # ── EDUCATION ──
    section_heading("EDUCATION")

    role_line(
        "MSc Finance",
        "Lancaster University Management School, Lancaster, UK",
        "Oct 2025 – Aug 2026",
    )
    p, _ = body(
        "First Class expected. Modules: Advanced Investment Management, Financial Modelling & Valuation, "
        "Python for Data Analysis, Financial Databases (Bloomberg Terminal). Dissertation: ESG Debt Premium "
        "in the Syndicated Loan Market — panel regression on 135,000+ loan tranches (DealScan, Compustat, RepRisk) in R.",
        indent=True,
    )

    role_line(
        "Master of Commerce (M.Com), 82%",
        "KCS College of Arts and Science, Chennai, India",
        "Jun 2019 – Apr 2021",
    )
    role_line(
        "Bachelor of Commerce (B.Com), 72%",
        "KCS College of Arts and Science, Chennai, India",
        "Jun 2016 – Apr 2019",
    )

    # ── EXPERIENCE ──
    section_heading("PROFESSIONAL EXPERIENCE")

    role_line("Pricing Associate (Financial Analyst)", "Accenture, Bengaluru, India", "Jun 2023 – Sep 2025")
    for b in [
        "Assessed the design and operating effectiveness of pricing controls for 600+ multimillion-dollar outsourcing contracts, ensuring commercial terms were governed by clearly defined risk parameters and approval thresholds.",
        "Identified and quantified margin risks by running multi-variable sensitivity analysis across deal scenarios, isolating key cost drivers and providing structured recommendations to senior finance and deal-approval stakeholders.",
        "Reduced manual reporting time by 30% by designing automated performance dashboards in Excel VBA, improving the accuracy and timeliness of financial control reporting across the pricing team.",
        "Conducted variance and trend analysis on large pricing datasets, distilling findings into actionable recommendations for non-financial stakeholders and senior leadership.",
        "Collaborated cross-functionally with sales, legal, and delivery teams to validate deal assumptions and ensure commercial commitments were aligned with internal risk and profitability controls.",
    ]:
        bullet(b)

    role_line("Finance Assistant", "SMS LLP (Mtandt Group), Chennai, India", "Feb 2022 – May 2023")
    for b in [
        "Produced monthly management accounts and financial reports that supported executive decision-making, consolidating financial statements into concise reporting packages for senior leaders.",
        "Reduced potential financial losses by 12% by conducting inventory and asset-performance analysis, identifying control weaknesses and recommending targeted provisions to management.",
        "Supported the annual budgeting cycle and delivered monthly variance analysis, tracking financial performance against targets and flagging material deviations for management review.",
    ]:
        bullet(b)

    # ── PROJECTS ──
    section_heading("INVESTMENT & FINANCIAL ANALYSIS PROJECTS")

    proj_heading("Unilever PLC: Intrinsic Valuation via FCFF DCF Model")
    for b in [
        "Evaluated the robustness of Unilever's financial controls and business model by constructing a full FCFF discounted cash flow model, stress-testing outputs through scenario and sensitivity analysis across WACC, terminal growth, and operating-margin assumptions.",
        "Produced an equity-research-style recommendation assessing Unilever's risk profile and identifying macro and discount-rate risks as key drivers of valuation uncertainty.",
    ]:
        bullet(b)

    proj_heading("FTSE 350 Portfolio: Equity Valuation & Risk Analysis")
    for b in [
        "Assessed systematic risk and return-risk trade-offs for 10 FTSE 350 companies applying CAPM, producing a structured investment rationale with buy/sell/hold recommendations grounded in quantitative evidence.",
        "Formulated analytical findings into clear, evidence-based recommendations — mirroring the assurance reporting process of translating analysis into actionable client insight.",
    ]:
        bullet(b)

    # ── SKILLS ──
    section_heading("SKILLS & INTERESTS")

    def skill_line(label, content):
        p = doc.add_paragraph()
        para_space(p, after=2)
        r1 = p.add_run(f"{label}: ")
        set_font(r1, 10, bold=True)
        r2 = p.add_run(content)
        set_font(r2, 10)

    skill_line(
        "Technical Tools",
        "Advanced Excel (VBA, Power Query), Bloomberg Terminal, Python (pandas, data visualisation), R (panel regression), SQL, PowerPoint, Power BI",
    )
    skill_line(
        "Financial Competencies",
        "Internal Control Assessment, Risk Identification & Quantification, Variance Analysis, Financial Reporting, Management Accounting, Process Auditing, Scenario & Sensitivity Analysis",
    )
    skill_line("Certifications", "CMA Intermediate, Institute of Cost Accountants of India (ICAI)")
    skill_line("Interests", "Chess, Painting")

    add_hline(doc)

    out = "/home/user/AI-Agent-for-apply-jobs/applications/cvs/Deloitte_Internal_Audit_Assurance_Graduate_CV.docx"
    doc.save(out)
    print(f"Saved CV → {out}")


# ── COVER LETTER ──────────────────────────────────────────────────────────────

def build_cover_letter():
    doc = Document()
    set_narrow_margins(doc)

    def para(text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size=10.5, before=0, after=6, italic=False):
        p = doc.add_paragraph()
        p.alignment = align
        para_space(p, before=before, after=after)
        if text:
            r = p.add_run(text)
            set_font(r, size, bold=bold)
            r.font.italic = italic
        return p

    # Header
    para("Jeevakumar Jayachandran", align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, size=12, after=2)
    para("Lancaster, UK", after=2)
    para("jeevajk2112@gmail.com  |  +44 7466 480480  |  linkedin.com/in/jeevakumar-j", after=10)

    para("17 June 2026", after=10)

    para("Hiring Manager", after=2)
    para("Deloitte", after=2)
    para("London, UK", after=10)

    # Subject line
    p = doc.add_paragraph()
    para_space(p, after=10)
    r = p.add_run(
        "Re: Application for Internal Audit & Assurance – Full Time Graduate Programme "
        "(Audit & Assurance, London, September 2026)"
    )
    set_font(r, 10.5, bold=True)

    para("Dear Hiring Manager,", after=8)

    body_paras = [
        (
            "Deloitte's Internal Audit & Change team occupies a uniquely privileged position — working with "
            "senior leaders across financial services and corporate sectors to understand the real inner workings "
            "of businesses and make their control environments stronger. That combination of cross-sector breadth, "
            "genuine commercial exposure, and the ACA qualification pathway is precisely why this programme stands "
            "out to me."
        ),
        (
            "At Accenture, I spent over two years assessing the design and operating effectiveness of pricing "
            "controls across 600+ multimillion-dollar outsourcing contracts. My work involved identifying margin "
            "risks through sensitivity analysis, validating deal assumptions against internal risk parameters, and "
            "communicating findings clearly to senior finance and deal-approval stakeholders — a process that "
            "mirrors the core assurance workflow of evaluating controls, identifying weaknesses, and presenting "
            "structured recommendations. I also built automated Excel VBA dashboards that reduced manual reporting "
            "time by 30%, improving the reliability of financial control reporting across the team. Earlier, at "
            "SMS LLP, I conducted inventory and asset-performance analysis that reduced financial losses by 12% "
            "by surfacing control weaknesses and recommending targeted provisions to management."
        ),
        (
            "My MSc Finance at Lancaster University Management School — where I am on track for a First Class — "
            "has deepened my quantitative skills through financial modelling, Bloomberg Terminal, Python, and an "
            "econometric dissertation analysing ESG risk premia across 135,000+ loan tranches. My CMA Intermediate "
            "qualification provides a strong accounting foundation ahead of the ACA pathway available within "
            "the programme."
        ),
        (
            "I would welcome the opportunity to bring my risk assessment and financial analysis experience to "
            "Deloitte's Internal Audit & Change team. I am available for interview at your convenience and can "
            "be reached at jeevajk2112@gmail.com or +44 7466 480480."
        ),
    ]

    for text in body_paras:
        para(text, after=8)

    para("Yours sincerely,", after=20)

    p = doc.add_paragraph()
    para_space(p, after=2)
    r = p.add_run("Jeevakumar Jayachandran")
    set_font(r, 10.5, bold=True)

    out = "/home/user/AI-Agent-for-apply-jobs/applications/coverletters/Deloitte_Internal_Audit_Assurance_Graduate_CoverLetter.docx"
    doc.save(out)
    print(f"Saved Cover Letter → {out}")


if __name__ == "__main__":
    build_cv()
    build_cover_letter()
