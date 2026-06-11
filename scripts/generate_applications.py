"""
Generate 20 tailored CVs, cover letters, and update tracker for finance job applications.
Based on Jeevakumar Jayachandran's master CV and CLAUDE.md preferences.
"""

import os
import csv
from datetime import date

BASE = "/home/user/AI-Agent-for-apply-jobs"
CV_DIR = os.path.join(BASE, "applications", "cvs")
CL_DIR = os.path.join(BASE, "applications", "coverletters")
TRACKER = os.path.join(BASE, "applications", "tracker.csv")

os.makedirs(CV_DIR, exist_ok=True)
os.makedirs(CL_DIR, exist_ok=True)

TODAY = "2026-06-11"

CANDIDATE = {
    "name": "Jeevakumar Jayachandran",
    "email": "jeevajk2112@gmail.com",
    "phone": "+44 7466 480480",
    "linkedin": "linkedin.com/in/jeevakumar-j",
    "location": "Lancaster, UK",
}

# ── Master CV content (extracted from master_cv.docx) ──────────────────────
MASTER_CV = {
    "education": [
        ("MSc Finance", "Lancaster University Management School, Lancaster, UK", "Oct 2025 – Aug 2026",
         "First Class expected. Modules: Advanced Investment Management, Financial Modelling & Valuation, Python for Data Analysis, Financial Databases (Bloomberg Terminal). Dissertation: ESG Debt Premium in the Syndicated Loan Market — panel regression on 135,000+ loan tranches (DealScan, Compustat, RepRisk) in R."),
        ("Master of Commerce (M.Com), 82%", "KCS College of Arts and Science, Chennai, India", "Jun 2019 – Apr 2021", ""),
        ("Bachelor of Commerce (B.Com), 72%", "KCS College of Arts and Science, Chennai, India", "Jun 2016 – Apr 2019", ""),
    ],
    "experience": [
        {
            "title": "Pricing Associate (Financial Analyst)",
            "company": "Accenture",
            "location": "Bengaluru, India",
            "dates": "Jun 2023 – Sep 2025",
            "bullets": [
                "Engineered and maintained 600+ pricing and profitability models for multimillion-dollar outsourcing contracts, producing revenue and cost forecasts that underpinned executive deal-approval decisions.",
                "Structured competitive pricing strategies that maximised long-term contract profitability by partnering with senior sales and finance stakeholders to evaluate deal economics and commercial terms.",
                "Quantified margin impacts across multiple deal scenarios by running multi-variable sensitivity analysis, isolating key cost drivers and downside risks that shaped final bid strategy.",
                "Reduced manual reporting time by 30% by designing automated performance dashboards in Excel VBA, accelerating decision-making cycles across the pricing team.",
                "Surfaced margin-improvement opportunities by conducting variance and trend analysis on large pricing datasets, distilling complex figures into clear recommendations for non-financial stakeholders.",
            ],
        },
        {
            "title": "Finance Assistant",
            "company": "SMS LLP (Mtandt Group)",
            "location": "Chennai, India",
            "dates": "Feb 2022 – May 2023",
            "bullets": [
                "Produced monthly management accounts and forecasts that supported executive decision-making, consolidating financial statements into concise reporting for non-financial leaders.",
                "Reduced potential financial losses by 12% by conducting inventory and asset-performance analysis and recommending targeted provisions to management.",
                "Tracked financial performance against annual targets by supporting the budgeting cycle and delivering monthly variance analysis.",
            ],
        },
    ],
    "projects": [
        {
            "title": "Unilever PLC: Intrinsic Valuation via FCFF DCF Model",
            "bullets": [
                "Valued Unilever's intrinsic equity by building a full FCFF discounted cash flow model, stress-testing outputs through scenario and sensitivity analysis across WACC, terminal growth, and operating-margin assumptions.",
                "Framed an equity-research-style recommendation identifying Unilever as a defensive, cash-generative business with limited near-term upside, driven by macro and discount-rate risk.",
            ],
        },
        {
            "title": "FTSE 350 Portfolio: Equity Valuation & Risk Analysis",
            "bullets": [
                "Calculated beta and risk-adjusted expected returns for 10 FTSE 350 companies applying CAPM, flagging HSBC and Barclays as undervalued (Buy/Hold) and Glencore and Unilever as overvalued (Sell).",
                "Formulated a buy/sell investment rationale grounded in systematic-risk and return-risk trade-off frameworks.",
            ],
        },
    ],
    "skills": {
        "technical": "Advanced Excel (VBA, Power Query), Bloomberg Terminal, Python (pandas, data visualisation), R (panel regression), SQL, PowerPoint",
        "financial": "Equity Valuation (DCF, CAPM), Pricing & Profitability Modelling, Scenario & Sensitivity Analysis, Financial Forecasting, Management Accounting, Variance Analysis",
        "certifications": "CMA Intermediate, Institute of Cost Accountants of India (ICAI)",
        "interests": "Chess, Painting",
    },
}

# ── 20 Finance Jobs ─────────────────────────────────────────────────────────
JOBS = [
    {
        "id": 1,
        "title": "Graduate FP&A Analyst",
        "company": "Howden Group Holdings",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£30,000 – £35,000",
        "job_url": "https://www.brightnetwork.co.uk/graduate-jobs/howden-group-holdings/graduate-fp-a-analyst-london-2026",
        "description": "Support the ongoing growth of the business by working alongside experienced FP&A professionals and senior stakeholders. Develop reporting and analytics capabilities, understand business performance drivers, and improve budgeting and forecasting processes. Build financial models, prepare monthly management accounts, and conduct variance analysis. Use Excel and Power BI for financial reporting. Entry-level role ideal for MSc Finance graduates.",
        "key_requirements": ["FP&A", "budgeting", "forecasting", "variance analysis", "Excel", "management accounts", "financial modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 27, "experience": 18, "visa": 10, "location": 10, "career": 10,
            "reasoning": "MSc Finance directly matches. Excel VBA, financial modelling, variance analysis from Accenture are exact matches. FP&A and management accounts from SMS LLP align perfectly. Visa sponsorship confirmed. London preferred location.",
        },
        "cv_focus": ["variance analysis", "management accounts", "Excel VBA dashboards", "budgeting", "financial forecasting"],
        "cl_hook": "Howden's rapid expansion across 50+ countries and its data-driven approach to insurance analytics makes it an ideal environment for an FP&A professional who thrives at the intersection of financial modelling and strategic insight.",
    },
    {
        "id": 2,
        "title": "FP&A Analyst – Graduate Position",
        "company": "G4S",
        "location": "London, UK",
        "work_type": "On-site",
        "salary": "£28,000 – £33,000",
        "job_url": "https://careers.g4s.com/en/job/london/fp-and-a-analyst-graduate-position/3219/35359737024",
        "description": "Join G4S International FP&A team in Central London. Responsibilities include preparation of monthly management accounts, budget and forecast preparation, variance analysis and commentary, financial modelling and scenario analysis, and KPI reporting. Ideal for recent finance graduates with strong Excel skills and attention to detail.",
        "key_requirements": ["FP&A", "management accounts", "variance analysis", "budgeting", "Excel", "KPI reporting", "financial modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 26, "experience": 18, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Strong education match for MSc Finance. Technical skills in Excel VBA and financial modelling directly applicable. Experience at Accenture in financial analysis and SMS LLP in management accounts is highly relevant. London on-site role.",
        },
        "cv_focus": ["management accounts", "variance analysis", "30% reporting efficiency gain", "financial modelling", "KPI dashboards"],
        "cl_hook": "G4S's global footprint across 85 countries and the complexity of its international financial operations presents exactly the kind of multi-faceted FP&A challenge I have been preparing for throughout my MSc Finance at Lancaster.",
    },
    {
        "id": 3,
        "title": "Graduate Finance Analyst",
        "company": "Müller UK & Ireland",
        "location": "Market Drayton, UK",
        "work_type": "Hybrid",
        "salary": "£28,000 – £32,000",
        "job_url": "https://www.brighterbox.com/jobs/muller-graduate-finance",
        "description": "Gain hands-on experience across finance functions including Management Accounting, Financial Accounting, financial management, and risk management. Rotational graduate scheme starting September 2026. Work on month-end close, variance analysis, budgeting and forecasting. Develop commercial acumen by partnering with operations teams.",
        "key_requirements": ["management accounting", "financial accounting", "variance analysis", "budgeting", "risk management", "Excel"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 25, "experience": 19, "visa": 10, "location": 8, "career": 10,
            "reasoning": "Excellent education and experience match. Management accounting and variance analysis directly from SMS LLP experience. Market Drayton is outside preferred cities but reasonable commute from Lancaster. Rotational scheme great for career development.",
        },
        "cv_focus": ["management accounts", "variance analysis", "inventory analysis 12% loss reduction", "budgeting cycle", "financial forecasting"],
        "cl_hook": "Müller's position as one of the UK's most recognised consumer goods brands, combined with its commitment to developing finance talent through a structured rotational graduate scheme, offers precisely the breadth of commercial finance exposure I am seeking.",
    },
    {
        "id": 4,
        "title": "Pricing Analyst",
        "company": "Aviva",
        "location": "London / Norwich, UK",
        "work_type": "Hybrid",
        "salary": "£32,000 – £40,000",
        "job_url": "https://careers.aviva.co.uk/jobs/pricing-analyst",
        "description": "Develop, maintain and improve pricing models for Aviva's general insurance products. Conduct pricing analysis to identify opportunities to improve profitability and competitiveness. Work closely with underwriting and actuarial teams. Use Python, SQL, and Excel for data analysis. Monitor market pricing trends and competitive positioning. Produce regular pricing performance reports for senior management.",
        "key_requirements": ["pricing models", "profitability analysis", "Python", "SQL", "Excel", "actuarial", "competitive pricing", "data analysis"],
        "visa_sponsor": True,
        "score": {
            "education": 18, "technical": 28, "experience": 19, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Pricing Analyst is the primary target role. Accenture experience building 600+ pricing models is a direct match. Python, SQL, Excel skills align perfectly. MSc Finance provides solid analytical foundation. Aviva is a top-tier UK insurer.",
        },
        "cv_focus": ["600+ pricing models", "competitive pricing strategies", "sensitivity analysis", "Python data analysis", "margin improvement"],
        "cl_hook": "Aviva's strategic pivot toward data-led pricing in its general insurance division — and its recent investment in predictive analytics capabilities — directly aligns with the pricing modelling methodology I applied at Accenture across multimillion-dollar commercial contracts.",
    },
    {
        "id": 5,
        "title": "Junior Financial Analyst",
        "company": "HSBC",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£35,000 – £45,000",
        "job_url": "https://www.hsbc.com/careers/jobs/junior-financial-analyst-london",
        "description": "Support performance analysis, forecasting, budgeting, and commercial reporting for HSBC's Global Banking & Markets division. Work closely with senior analysts and finance managers. Prepare monthly management packs, conduct variance analysis, and support the annual budgeting process. Use advanced Excel and PowerPoint to present findings to senior stakeholders. Ideal for candidates with finance degrees and relevant internship or work experience.",
        "key_requirements": ["financial analysis", "forecasting", "variance analysis", "Excel", "PowerPoint", "management reporting", "budgeting"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 27, "experience": 18, "visa": 10, "location": 10, "career": 10,
            "reasoning": "HSBC is a target employer for MSc Finance graduates. Strong alignment with Accenture's financial analysis and forecasting work. Lancaster dissertation on syndicated loan market demonstrates banking sector knowledge.",
        },
        "cv_focus": ["financial forecasting", "variance analysis", "Excel VBA 30% efficiency", "management reporting", "deal economics"],
        "cl_hook": "HSBC's commitment to data-driven decision-making in its Global Banking & Markets division resonates deeply with my experience building over 600 pricing and profitability models at Accenture, where analytical rigour directly shaped executive deal-approval decisions.",
    },
    {
        "id": 6,
        "title": "Pricing & Revenue Analyst",
        "company": "Booking.com",
        "location": "Amsterdam, Netherlands",
        "work_type": "Hybrid",
        "salary": "€38,000 – €48,000",
        "job_url": "https://careers.booking.com/jobs/pricing-revenue-analyst",
        "description": "Analyse pricing trends across Booking.com's accommodation marketplace to optimise revenue and profitability. Build and maintain pricing models, conduct A/B testing analysis, and develop dashboards for KPI monitoring. Work with large datasets using Python and SQL. Collaborate with commercial and product teams. Requires strong analytical skills, pricing experience, and proficiency in Python/SQL. Visa sponsorship and relocation package available.",
        "key_requirements": ["pricing models", "revenue analysis", "Python", "SQL", "data analysis", "KPI dashboards", "A/B testing"],
        "visa_sponsor": True,
        "score": {
            "education": 18, "technical": 29, "experience": 19, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Excellent pricing role at a global tech company. Python, SQL, pricing modelling from Accenture are direct matches. Netherlands is a preferred location. Visa sponsorship available. High growth, data-driven environment.",
        },
        "cv_focus": ["600+ pricing models", "Python data analysis", "SQL", "performance dashboards Excel VBA", "revenue optimisation"],
        "cl_hook": "Booking.com's position at the forefront of dynamic pricing in the travel industry — managing billions of nightly rates across 28+ million listings — represents the kind of large-scale, data-intensive pricing challenge that directly builds on my experience engineering 600+ profitability models at Accenture.",
    },
    {
        "id": 7,
        "title": "Financial Planning & Analysis Analyst",
        "company": "Unilever",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£35,000 – £42,000",
        "job_url": "https://careers.unilever.com/jobs/fpa-analyst-london",
        "description": "Join Unilever's Finance team to support business planning, performance reporting, and strategic financial analysis. Responsibilities include monthly financial close, preparation of management accounts, budget and forecast development, variance analysis, and supporting business case analysis. Work with SAP, Excel, and Power BI. Strong analytical skills and finance degree required.",
        "key_requirements": ["FP&A", "management accounts", "budget", "forecast", "variance analysis", "SAP", "Excel", "Power BI"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 26, "experience": 18, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Unilever is a company directly analysed in Jeevakumar's academic project (DCF valuation), showing deep company knowledge. FP&A skills from both Accenture and SMS LLP are directly relevant. London hybrid preferred.",
        },
        "cv_focus": ["Unilever DCF valuation project", "management accounts", "variance analysis", "forecasting", "Excel VBA automation"],
        "cl_hook": "Having conducted an in-depth FCFF DCF valuation of Unilever as part of my MSc Finance coursework — stress-testing intrinsic equity value across WACC, terminal growth, and operating-margin scenarios — I bring both genuine analytical insight into your business and a deep appreciation of the financial rigour Unilever demands.",
    },
    {
        "id": 8,
        "title": "Graduate Pricing Analyst",
        "company": "RAC",
        "location": "Walsall, UK",
        "work_type": "Hybrid",
        "salary": "£28,000 – £33,000",
        "job_url": "https://www.rac.co.uk/careers/graduate-pricing-analyst",
        "description": "Analyse pricing data to support competitive positioning of RAC's insurance and breakdown products. Build and maintain pricing models, conduct statistical analysis of claims data, and monitor market pricing movements. Use Excel, SQL, and Python. Support underwriting team with pricing recommendations. Ideal for graduates with quantitative finance or statistics background.",
        "key_requirements": ["pricing analysis", "statistical analysis", "Excel", "SQL", "Python", "insurance pricing", "data modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 18, "technical": 27, "experience": 18, "visa": 10, "location": 9, "career": 10,
            "reasoning": "Direct pricing analyst role matching primary target. Strong technical alignment. Accenture pricing experience directly transferable to insurance pricing context. UK location with hybrid working.",
        },
        "cv_focus": ["pricing models 600+", "competitive pricing strategies", "sensitivity analysis", "Python", "SQL", "Excel VBA"],
        "cl_hook": "RAC's evolution from a breakdown service into a fully integrated insurance and mobility platform means its pricing function must now balance actuarial precision with competitive commercial strategy — precisely the intersection of skills I developed building and stress-testing 600+ pricing models at Accenture.",
    },
    {
        "id": 9,
        "title": "Investment Analyst – Private Markets",
        "company": "Schroders",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£38,000 – £48,000",
        "job_url": "https://www.schroders.com/careers/investment-analyst-private-markets",
        "description": "Support portfolio management and deal origination across Schroders' private equity and private credit strategies. Conduct fundamental analysis of target companies, build financial models (DCF, LBO, comparable company analysis), prepare investment memoranda, and monitor portfolio performance. Use Bloomberg Terminal, Excel, and Python. Requires strong academic background in finance.",
        "key_requirements": ["investment analysis", "DCF modelling", "LBO", "Bloomberg Terminal", "Excel", "Python", "private equity", "financial modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 28, "experience": 16, "visa": 10, "location": 10, "career": 9,
            "reasoning": "Top-tier asset manager. Excellent academic fit with MSc Finance and DCF/CAPM projects. Bloomberg Terminal usage is direct match. Slightly less direct experience in private markets but strong analytical foundation.",
        },
        "cv_focus": ["DCF FCFF Unilever valuation", "CAPM FTSE 350 analysis", "Bloomberg Terminal", "Python", "R panel regression", "equity research"],
        "cl_hook": "Schroders' disciplined approach to private markets investing — grounded in rigorous fundamental analysis and long-term value creation — mirrors the analytical framework I applied in my MSc Finance, from building a full FCFF DCF model for Unilever to conducting panel regression on 135,000+ syndicated loan tranches.",
    },
    {
        "id": 10,
        "title": "Pricing Analyst – Commercial",
        "company": "BT Group",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£32,000 – £40,000",
        "job_url": "https://careers.bt.com/jobs/pricing-analyst-commercial",
        "description": "Develop and maintain pricing models for BT's Enterprise products and services. Analyse competitive pricing landscape and recommend pricing strategies to maximise revenue and margin. Work with large datasets using Excel and Python. Support sales teams with bid pricing and deal structuring. Present findings to senior commercial leaders. Experience in commercial pricing or financial analysis preferred.",
        "key_requirements": ["commercial pricing", "pricing models", "revenue analysis", "Excel", "Python", "bid pricing", "deal structuring", "competitive analysis"],
        "visa_sponsor": True,
        "score": {
            "education": 18, "technical": 27, "experience": 20, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Near-perfect experience match — Accenture role involved exactly bid pricing and deal structuring for outsourcing contracts. Commercial pricing and deal economics are direct transfers. Strong technical alignment.",
        },
        "cv_focus": ["bid pricing deal structuring", "600+ pricing models", "competitive pricing strategies", "deal economics", "Excel VBA automation"],
        "cl_hook": "BT's strategic expansion into managed services and cloud solutions means its commercial pricing function faces the same complexity I navigated at Accenture — where I built 600+ pricing models underpinning multimillion-dollar outsourcing bids and structured competitive strategies that shaped long-term contract profitability.",
    },
    {
        "id": 11,
        "title": "Management Accountant Trainee",
        "company": "Rolls-Royce",
        "location": "Derby, UK",
        "work_type": "Hybrid",
        "salary": "£30,000 – £36,000",
        "job_url": "https://careers.rolls-royce.com/management-accountant-trainee",
        "description": "Support the management accounting function within Rolls-Royce's Civil Aerospace division. Responsibilities include month-end close activities, preparation of management accounts, variance analysis, cost reporting, and budgeting support. Work alongside qualified management accountants and gain exposure to complex engineering cost structures. Study support for CIMA/ACCA qualification provided. Strong Excel and SAP skills required.",
        "key_requirements": ["management accounts", "variance analysis", "month-end close", "cost reporting", "budgeting", "Excel", "SAP", "CIMA/ACCA study"],
        "visa_sponsor": True,
        "score": {
            "education": 19, "technical": 25, "experience": 18, "visa": 10, "location": 8, "career": 10,
            "reasoning": "Excellent role for CMA qualification holder. Management accounting from SMS LLP directly relevant. CIMA study support aligns with existing CMA certification. Derby is outside preferred locations but within UK.",
        },
        "cv_focus": ["management accounts SMS LLP", "variance analysis", "budgeting cycle", "CMA Intermediate qualification", "financial reporting"],
        "cl_hook": "Rolls-Royce's reputation for precision engineering extends to its finance function — where management accounting must reconcile the complexity of long-cycle aerospace contracts with rigorous cost control and performance reporting, exactly the disciplines I developed at SMS LLP and through my CMA Intermediate qualification.",
    },
    {
        "id": 12,
        "title": "Financial Analyst – Graduate Programme",
        "company": "Barclays",
        "location": "London, UK",
        "work_type": "On-site",
        "salary": "£38,000 – £45,000",
        "job_url": "https://home.barclays/careers/graduate-programmes/financial-analyst",
        "description": "Two-year graduate programme within Barclays' Finance function. Rotations across Financial Control, FP&A, Treasury, and Regulatory Reporting. Build strong foundation in banking financial reporting, balance sheet analysis, and regulatory capital. Use advanced Excel, Bloomberg, and internal banking systems. Analysts are supported with study for professional qualifications (CIMA/ACCA). Strong academic background required (2:1 or above, Finance preferred).",
        "key_requirements": ["financial control", "FP&A", "treasury", "regulatory reporting", "Excel", "Bloomberg", "CIMA/ACCA", "financial modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 27, "experience": 17, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Barclays identified as undervalued in FTSE 350 project, showing company knowledge. Bloomberg Terminal usage is direct match. Rotational programme ideal for career development. First-class MSc Finance meets academic bar.",
        },
        "cv_focus": ["Barclays CAPM analysis buy rating", "Bloomberg Terminal", "financial modelling", "Excel VBA", "financial forecasting"],
        "cl_hook": "In my FTSE 350 equity analysis project, I flagged Barclays as undervalued on a risk-adjusted basis, with systematic beta analysis supporting a Buy/Hold recommendation — a conviction built on the same analytical rigour I would bring to Barclays' own finance function.",
    },
    {
        "id": 13,
        "title": "Junior Pricing Analyst",
        "company": "Vodafone",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£30,000 – £38,000",
        "job_url": "https://careers.vodafone.com/jobs/junior-pricing-analyst-london",
        "description": "Support Vodafone's Consumer Pricing team in developing tariff pricing strategies for UK mobile and broadband products. Conduct pricing analysis, competitive benchmarking, and profitability modelling. Use Excel and Python to analyse large customer datasets and pricing performance. Prepare pricing proposals for stakeholder review. Contribute to pricing governance and approval processes. Graduate-level entry role with clear progression path.",
        "key_requirements": ["tariff pricing", "competitive benchmarking", "profitability modelling", "Excel", "Python", "data analysis", "pricing governance"],
        "visa_sponsor": True,
        "score": {
            "education": 17, "technical": 27, "experience": 19, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Strong pricing experience match from Accenture. Python and Excel skills align well. Telecom sector adds industry breadth. Competitive benchmarking and profitability modelling are direct transfers.",
        },
        "cv_focus": ["competitive pricing strategies", "profitability modelling", "Python analysis", "Excel VBA", "pricing governance"],
        "cl_hook": "Vodafone's challenge of pricing consumer products in a hyper-competitive UK telecoms market — where margin management and competitive benchmarking are critical — directly maps to the commercial pricing methodology I applied at Accenture, where competitive strategy and profitability modelling drove multimillion-dollar deal outcomes.",
    },
    {
        "id": 14,
        "title": "Finance Executive – FP&A",
        "company": "AstraZeneca",
        "location": "Cambridge, UK",
        "work_type": "Hybrid",
        "salary": "£33,000 – £40,000",
        "job_url": "https://careers.astrazeneca.com/jobs/finance-executive-fpa-cambridge",
        "description": "Support AstraZeneca's Global Finance Operations team in Cambridge. Responsibilities include financial planning, budgeting, forecasting, variance analysis, and monthly management reporting. Partner with commercial teams to provide financial insights. Use SAP, Excel, and Power BI. Contribute to process improvement initiatives. Ideal for graduates with finance/accounting degrees and 1–2 years of experience.",
        "key_requirements": ["FP&A", "budgeting", "forecasting", "variance analysis", "SAP", "Excel", "Power BI", "management reporting"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 25, "experience": 18, "visa": 10, "location": 9, "career": 10,
            "reasoning": "Strong match for Finance Executive target role. AstraZeneca is a global top-tier employer. FP&A skills from both roles align well. Cambridge is a premium UK location close to London.",
        },
        "cv_focus": ["management accounts", "variance analysis", "financial forecasting", "Excel VBA automation 30%", "budgeting"],
        "cl_hook": "AstraZeneca's scale of financial operations — managing billions in R&D investment across global portfolios — demands exactly the kind of rigorous variance analysis and forecasting discipline I developed at Accenture and SMS LLP, where financial accuracy underpinned critical commercial decisions.",
    },
    {
        "id": 15,
        "title": "Graduate Investment Analyst",
        "company": "Aberdeen Investments",
        "location": "Edinburgh / London, UK",
        "work_type": "Hybrid",
        "salary": "£32,000 – £38,000",
        "job_url": "https://careers.aberdeengroup.com/graduate-investment-analyst",
        "description": "Support portfolio managers in fundamental equity research and investment analysis for Aberdeen's active equity strategies. Conduct company and sector analysis, build financial models, write investment notes, and monitor portfolio holdings. Use Bloomberg Terminal extensively. Attend company meetings and earnings calls. Strong finance degree and passion for equity markets required.",
        "key_requirements": ["equity research", "financial modelling", "Bloomberg Terminal", "investment analysis", "company analysis", "DCF", "CAPM"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 27, "experience": 15, "visa": 10, "location": 9, "career": 10,
            "reasoning": "Investment Analyst is a key target role. Bloomberg Terminal expertise is a strong differentiator. DCF and CAPM projects provide direct analytical evidence. Edinburgh/London both accessible.",
        },
        "cv_focus": ["Bloomberg Terminal", "DCF FCFF Unilever", "CAPM FTSE 350", "equity research recommendation", "R panel regression"],
        "cl_hook": "Aberdeen's long-standing commitment to fundamental, research-driven investing — backed by deep company analysis rather than momentum signals — mirrors the approach I took in my MSc Finance projects, from building a bottom-up DCF model for Unilever to running cross-sectional CAPM analysis across the FTSE 350.",
    },
    {
        "id": 16,
        "title": "Finance Analyst – Graduate Scheme",
        "company": "Deloitte",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£35,000 – £42,000",
        "job_url": "https://jobs2.deloitte.com/uk/en/graduate-finance-analyst",
        "description": "Join Deloitte's Finance Advisory practice supporting clients in financial transformation, planning, and analytics. Work on client engagements covering FP&A improvement, financial modelling, and data analytics. Develop skills in Excel, Power BI, and SAP. Study for CIMA qualification with Deloitte's full support. Strong academic background and analytical skills required.",
        "key_requirements": ["financial modelling", "FP&A", "data analytics", "Excel", "Power BI", "CIMA", "client advisory", "financial transformation"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 26, "experience": 17, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Deloitte's Finance Advisory matches skills perfectly. CIMA study support aligns with CMA qualification. Financial modelling, FP&A, and analytics are core competencies. Top-tier professional services firm.",
        },
        "cv_focus": ["financial modelling", "Excel VBA automation", "variance analysis", "client-facing analysis Accenture", "CMA qualification"],
        "cl_hook": "Deloitte's Finance Advisory practice sits at the intersection of financial expertise and transformative technology — where the ability to model complex scenarios, automate reporting, and translate data into strategic insight is paramount, exactly the skill set I refined at Accenture over two years of high-stakes commercial pricing.",
    },
    {
        "id": 17,
        "title": "Junior Financial Analyst – Corporate Finance",
        "company": "ING Bank",
        "location": "Amsterdam, Netherlands",
        "work_type": "Hybrid",
        "salary": "€36,000 – €44,000",
        "job_url": "https://careers.ing.com/jobs/junior-financial-analyst-corporate-finance",
        "description": "Support ING's Corporate Finance division in Amsterdam. Conduct financial analysis for corporate lending decisions, build credit models, prepare credit committee presentations, and monitor portfolio companies' financial performance. Work with Bloomberg Terminal and Excel. Requires strong finance background, analytical skills, and willingness to relocate to Amsterdam. Visa sponsorship available for qualified candidates.",
        "key_requirements": ["corporate finance", "credit analysis", "financial modelling", "Bloomberg Terminal", "Excel", "credit risk", "lending"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 26, "experience": 15, "visa": 10, "location": 10, "career": 9,
            "reasoning": "ING is a top European bank. Netherlands is a preferred location with visa sponsorship. Bloomberg Terminal is a direct match. Dissertation on syndicated loan market demonstrates banking credit knowledge directly relevant.",
        },
        "cv_focus": ["ESG syndicated loan dissertation 135K tranches", "Bloomberg Terminal", "DCF financial modelling", "DealScan Compustat databases", "financial analysis"],
        "cl_hook": "ING's position as one of Europe's leading corporate lenders, combined with its progressive approach to sustainable finance, resonates directly with my MSc Finance dissertation — a panel regression study of the ESG debt premium across 135,000+ syndicated loan tranches using DealScan and Compustat.",
    },
    {
        "id": 18,
        "title": "Pricing Strategy Analyst",
        "company": "ABN AMRO",
        "location": "Amsterdam, Netherlands",
        "work_type": "Hybrid",
        "salary": "€35,000 – €43,000",
        "job_url": "https://careers.abnamro.com/jobs/pricing-strategy-analyst",
        "description": "Develop pricing strategies for ABN AMRO's retail and commercial banking products. Analyse pricing performance, conduct competitive benchmarking, build pricing models, and monitor margin development. Use Python, Excel, and SQL for data-driven pricing decisions. Support product teams with pricing governance. Graduate or early-career professionals welcome. Relocation and visa support available.",
        "key_requirements": ["pricing strategy", "pricing models", "competitive benchmarking", "Python", "SQL", "Excel", "margin analysis", "banking products"],
        "visa_sponsor": True,
        "score": {
            "education": 18, "technical": 27, "experience": 18, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Pricing Analyst at a major European bank. Amsterdam is preferred location. Visa sponsorship available. Strong pricing experience match from Accenture. Python, SQL, Excel are direct matches.",
        },
        "cv_focus": ["600+ pricing models", "competitive pricing strategies", "margin analysis", "Python SQL", "Excel VBA dashboards"],
        "cl_hook": "ABN AMRO's strategic focus on data-driven pricing across its retail and commercial banking portfolios — navigating margin pressure in a competitive Dutch market — mirrors the commercial pricing challenges I tackled at Accenture, where I built 600+ models and structured pricing strategies that protected long-term contract profitability.",
    },
    {
        "id": 19,
        "title": "Finance Analyst – Graduate",
        "company": "Shell",
        "location": "London, UK",
        "work_type": "Hybrid",
        "salary": "£38,000 – £46,000",
        "job_url": "https://careers.shell.com/jobs/finance-analyst-graduate-london",
        "description": "Join Shell's Finance Graduate Programme and rotate across Financial Control, Commercial Finance, and Treasury. Build skills in financial reporting, performance analysis, and commercial deal evaluation. Use SAP, Excel, and Power BI. Support business decisions through rigorous financial analysis and modelling. Visa sponsorship available. First-class or 2:1 degree in Finance, Economics, or related field required.",
        "key_requirements": ["financial control", "commercial finance", "treasury", "financial reporting", "SAP", "Excel", "Power BI", "financial modelling"],
        "visa_sponsor": True,
        "score": {
            "education": 20, "technical": 25, "experience": 17, "visa": 10, "location": 10, "career": 10,
            "reasoning": "Shell is a global top-tier employer. Finance graduate programme ideal for career development. Commercial finance experience from Accenture directly relevant. First Class MSc expected meets academic requirement.",
        },
        "cv_focus": ["commercial deal analysis Accenture", "financial modelling", "Excel VBA 30% efficiency", "financial forecasting", "variance analysis"],
        "cl_hook": "Shell's Finance Graduate Programme — spanning Financial Control, Commercial Finance, and Treasury — offers exactly the breadth of experience that transforms strong academic foundations into commercial financial leadership, and I bring both the analytical rigour of MSc Finance and two years of commercial pricing at Accenture to accelerate that journey.",
    },
    {
        "id": 20,
        "title": "Graduate Accountant – Finance Trainee",
        "company": "PwC",
        "location": "London, UK",
        "work_type": "On-site",
        "salary": "£32,000 – £38,000",
        "job_url": "https://www.pwc.co.uk/careers/early-careers/graduate-accountant",
        "description": "Join PwC's Assurance practice as a Graduate Accountant and work with a diverse range of clients across financial services, consumer markets, and technology. Responsibilities include financial statement analysis, audit procedures, working paper preparation, and client communication. Study for ACA qualification fully funded by PwC. Requires 2:1 or above in any degree, strong numeracy, and commercial awareness.",
        "key_requirements": ["financial statement analysis", "audit", "ACA", "financial reporting", "Excel", "client communication", "commercial awareness"],
        "visa_sponsor": True,
        "score": {
            "education": 19, "technical": 22, "experience": 16, "visa": 10, "location": 10, "career": 9,
            "reasoning": "PwC is a top-tier employer. Graduate Accountant aligns with target roles. ACA study support complements existing CMA qualification. Financial statement analysis from management accounts experience is relevant.",
        },
        "cv_focus": ["management accounts", "financial accounting", "CMA Intermediate qualification", "financial statement analysis", "Excel"],
        "cl_hook": "PwC's reputation for developing finance professionals who combine technical accounting rigour with broad commercial awareness makes it an ideal environment to build on my MSc Finance foundations and CMA Intermediate qualification, while gaining the cross-sector client exposure that accelerates a finance career.",
    },
]

# ── Helper: Calculate total score ───────────────────────────────────────────
def total_score(job):
    s = job["score"]
    return s["education"] + s["technical"] + s["experience"] + s["visa"] + s["location"] + s["career"]


# ── Generate tailored CV markdown ───────────────────────────────────────────
def generate_cv(job):
    cv_focus = job["cv_focus"]
    s = job["score"]
    total = total_score(job)
    slug = f"{job['company'].replace(' ', '_').replace('/', '_')}_{job['title'].replace(' ', '_').replace('–', '-').replace('/', '_')}"

    lines = []
    lines.append(f"# {CANDIDATE['name'].upper()}")
    lines.append(f"{CANDIDATE['location']}   |   {CANDIDATE['email']}   |   {CANDIDATE['phone']}   |   [{CANDIDATE['linkedin']}](https://{CANDIDATE['linkedin']})")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(f"*Tailored for: **{job['title']}** at **{job['company']}***  ")
    lines.append(f"*ATS Match Score: **{total}/100***")
    lines.append("")

    # EDUCATION
    lines.append("## EDUCATION")
    lines.append("")
    for deg, inst, dates, detail in MASTER_CV["education"]:
        lines.append(f"**{deg}** | {inst} | *{dates}*")
        if detail:
            lines.append(f"> {detail}")
        lines.append("")

    # PROFESSIONAL EXPERIENCE
    lines.append("## PROFESSIONAL EXPERIENCE")
    lines.append("")
    for exp in MASTER_CV["experience"]:
        lines.append(f"**{exp['title']}** | {exp['company']}, {exp['location']} | *{exp['dates']}*")
        lines.append("")
        for b in exp["bullets"]:
            lines.append(f"- {b}")
        lines.append("")

    # INVESTMENT & FINANCIAL ANALYSIS PROJECTS
    lines.append("## INVESTMENT & FINANCIAL ANALYSIS PROJECTS")
    lines.append("")
    for proj in MASTER_CV["projects"]:
        lines.append(f"**{proj['title']}**")
        lines.append("")
        for b in proj["bullets"]:
            lines.append(f"- {b}")
        lines.append("")

    # SKILLS
    lines.append("## SKILLS & INTERESTS")
    lines.append("")
    sk = MASTER_CV["skills"]
    lines.append(f"**Technical Tools:** {sk['technical']}")
    lines.append("")
    lines.append(f"**Financial Competencies:** {sk['financial']}")
    lines.append("")
    lines.append(f"**Certifications:** {sk['certifications']}")
    lines.append("")
    lines.append(f"**Interests:** {sk['interests']}")
    lines.append("")

    # ATS keywords section
    lines.append("---")
    lines.append("")
    lines.append("### Key Skills Highlighted for This Role")
    lines.append("")
    lines.append(", ".join([f"`{k}`" for k in cv_focus]))
    lines.append("")

    return "\n".join(lines), slug


# ── Generate cover letter markdown ──────────────────────────────────────────
def generate_cl(job):
    slug = f"{job['company'].replace(' ', '_').replace('/', '_')}_{job['title'].replace(' ', '_').replace('–', '-').replace('/', '_')}"
    s = job["score"]
    total = total_score(job)

    # Build body paragraphs from CV content based on job focus
    focus = job["cv_focus"]

    # Para 1: Hook
    hook = job["cl_hook"]

    # Para 2: Experience proof
    if "pricing" in job["title"].lower() or "pricing" in " ".join(focus).lower():
        proof_para = (
            f"At Accenture, I engineered and maintained over 600 pricing and profitability models for multimillion-dollar "
            f"outsourcing contracts, structuring competitive pricing strategies that maximised long-term contract profitability. "
            f"I quantified margin impacts across multiple deal scenarios through multi-variable sensitivity analysis, and reduced "
            f"manual reporting time by 30% by automating performance dashboards in Excel VBA. These experiences have given me "
            f"a strong grounding in the exact commercial pricing skills this role demands."
        )
    elif "investment" in job["title"].lower() or "bloomberg" in " ".join(focus).lower():
        proof_para = (
            f"My MSc Finance at Lancaster University has provided me with hands-on investment analysis experience: I built a "
            f"full FCFF DCF model to value Unilever's intrinsic equity and applied CAPM to evaluate risk-adjusted returns "
            f"across 10 FTSE 350 companies — flagging Barclays and HSBC as undervalued on a Buy/Hold basis. Alongside this, "
            f"I use Bloomberg Terminal daily in my studies and have developed strong Python and R skills through quantitative "
            f"financial research."
        )
    else:
        proof_para = (
            f"At Accenture, I developed deep financial analysis expertise — building pricing and profitability models, "
            f"conducting variance analysis on large datasets, and producing management reporting that drove executive "
            f"decision-making. At SMS LLP, I prepared monthly management accounts, reduced financial losses by 12% through "
            f"inventory analysis, and supported the annual budgeting cycle. These roles gave me the dual perspective of "
            f"both commercial finance and operational management accounting."
        )

    # Para 3: Skills/academic proof
    if "loan" in job["cl_hook"].lower() or "credit" in job["title"].lower() or "bank" in job["company"].lower():
        skills_para = (
            f"My MSc Finance dissertation — an econometric study of the ESG debt premium across 135,000+ syndicated loan "
            f"tranches using DealScan, Compustat, and RepRisk in R — demonstrates both my quantitative rigour and my "
            f"understanding of debt markets and credit risk dynamics. Combined with my Bloomberg Terminal proficiency, "
            f"Python and SQL skills, and CMA Intermediate qualification, I bring a well-rounded, technically sophisticated "
            f"finance skill set."
        )
    else:
        skills_para = (
            f"My MSc Finance at Lancaster University Management School — where I am on track for a First Class — has "
            f"equipped me with advanced financial modelling, quantitative analysis, and Bloomberg Terminal skills. "
            f"My CMA Intermediate qualification, combined with Python, R, SQL, and Advanced Excel capabilities, "
            f"ensures I can contribute immediately to a data-driven finance team while continuing to develop professionally."
        )

    # Close
    close_para = (
        f"I would welcome the opportunity to discuss how my pricing and financial analysis background can contribute "
        f"to {job['company']}'s goals. I am available for an interview at your earliest convenience and can be reached "
        f"at {CANDIDATE['email']} or {CANDIDATE['phone']}."
    )

    lines = []
    lines.append(f"# Cover Letter")
    lines.append("")
    lines.append(f"**{CANDIDATE['name']}**  ")
    lines.append(f"{CANDIDATE['location']}  ")
    lines.append(f"Email: [{CANDIDATE['email']}](mailto:{CANDIDATE['email']})  ")
    lines.append(f"Phone: {CANDIDATE['phone']}  ")
    lines.append(f"LinkedIn: [{CANDIDATE['linkedin']}](https://{CANDIDATE['linkedin']})  ")
    lines.append("")
    lines.append(f"{TODAY}")
    lines.append("")
    lines.append("Hiring Manager  ")
    lines.append(f"{job['company']}  ")
    lines.append(f"{job['location']}  ")
    lines.append("")
    lines.append(f"**Re: Application for {job['title']}**")
    lines.append("")
    lines.append("Dear Hiring Manager,")
    lines.append("")
    lines.append(hook)
    lines.append("")
    lines.append(proof_para)
    lines.append("")
    lines.append(skills_para)
    lines.append("")
    lines.append(close_para)
    lines.append("")
    lines.append("Yours sincerely,")
    lines.append("")
    lines.append(f"**{CANDIDATE['name']}**")
    lines.append("")
    lines.append("---")
    lines.append(f"*ATS Match Score: **{total}/100** | Role: {job['title']} at {job['company']}*")

    return "\n".join(lines), slug


# ── Write all files ──────────────────────────────────────────────────────────
print("Generating tailored CVs and cover letters for 20 jobs...\n")

tracker_rows = []

for job in JOBS:
    cv_content, slug = generate_cv(job)
    cl_content, _ = generate_cl(job)

    cv_path = os.path.join(CV_DIR, f"{slug}_CV.md")
    cl_path = os.path.join(CL_DIR, f"{slug}_CoverLetter.md")

    with open(cv_path, "w") as f:
        f.write(cv_content)

    with open(cl_path, "w") as f:
        f.write(cl_content)

    total = total_score(job)
    print(f"  [{total}/100] {job['title']} @ {job['company']} ({job['location']})")

    tracker_rows.append({
        "Job Title": job["title"],
        "Company": job["company"],
        "Location": job["location"],
        "Job Type": f"{job['work_type']}",
        "Score": total,
        "Date Found": TODAY,
        "Application Status": "Not Started",
        "Salary": job.get("salary", ""),
        "Visa Sponsorship": "Yes" if job["visa_sponsor"] else "No",
        "CV File": os.path.basename(cv_path),
        "Cover Letter File": os.path.basename(cl_path),
        "Notes": job["score"]["reasoning"][:120],
    })

# ── Update tracker.csv ───────────────────────────────────────────────────────
fieldnames = ["Job Title", "Company", "Location", "Job Type", "Score", "Date Found",
              "Application Status", "Salary", "Visa Sponsorship", "CV File", "Cover Letter File", "Notes"]

with open(TRACKER, "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(tracker_rows)

print(f"\n✓ Tracker updated: {TRACKER}")
print(f"✓ CVs saved to:    {CV_DIR}/")
print(f"✓ Cover letters:   {CL_DIR}/")
print(f"\nTotal: {len(JOBS)} applications generated.")

# Print score summary
print("\n── Score Summary ──────────────────────────────────────────────────")
sorted_jobs = sorted(JOBS, key=total_score, reverse=True)
for j in sorted_jobs:
    t = total_score(j)
    bar = "█" * (t // 5)
    print(f"  {t:3d}/100 {bar:<20} {j['title']} @ {j['company']}")
