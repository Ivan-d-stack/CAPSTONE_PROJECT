"""
Front matter builder for RetainIQ Black Book.
Generates:
1. Cover Page
2. Certificate
3. Proforma
4. Abstract (2 full pages)
5. Acknowledgement
6. Declaration
7. Table of Contents (2 pages)
8. List of Figures (2 pages)
9. List of Tables & Abbreviations (1 page)
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from report_helpers import (
    format_paragraph, add_body_p, add_heading_1, add_heading_2,
    set_cell_background, set_cell_margins, set_table_borders
)

def build_frontmatter(doc):
    # ==========================================
    # 1. COVER PAGE
    # ==========================================
    p_inst = doc.add_paragraph()
    format_paragraph(p_inst, space_before=10, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    r1 = p_inst.add_run("Kandivli Education Society's\n")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(13)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(15, 30, 65)

    r2 = p_inst.add_run("BK Shroff College of Arts & MH Shroff College of Commerce\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(15)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(15, 30, 65)

    r3 = p_inst.add_run("(Autonomous College | NAAC Re-accredited 'A' Grade | ISO 9001:2015 Certified)\nBhulabhai Patel Road, Kandivali (West), Mumbai - 400067")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(70, 80, 95)

    p_div = doc.add_paragraph()
    format_paragraph(p_div, space_before=15, space_after=15, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_div = p_div.add_run("—" * 38)
    r_div.font.color.rgb = RGBColor(180, 190, 200)

    p_rep = doc.add_paragraph()
    format_paragraph(p_rep, space_before=10, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_rep = p_rep.add_run("A CAPSTONE PROJECT REPORT ON\n")
    r_rep.font.name = "Times New Roman"
    r_rep.font.size = Pt(12)
    r_rep.font.bold = True
    r_rep.font.color.rgb = RGBColor(100, 110, 125)

    p_title = doc.add_paragraph()
    format_paragraph(p_title, space_before=12, space_after=8, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_title = p_title.add_run("SURVIVAL-BASED CHURN MODELING WITH PERSONALIZED RETENTION TREATMENT RANKING")
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 35, 75)

    p_proj = doc.add_paragraph()
    format_paragraph(p_proj, space_before=4, space_after=20, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_proj = p_proj.add_run("(RetainIQ  |  Project ID: BDS-11)")
    r_proj.font.name = "Times New Roman"
    r_proj.font.size = Pt(12)
    r_proj.font.italic = True
    r_proj.font.color.rgb = RGBColor(40, 60, 90)

    p_sub = doc.add_paragraph()
    format_paragraph(p_sub, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_sub = p_sub.add_run("SUBMITTED IN PARTIAL FULFILLMENT OF THE REQUIREMENTS\nFOR THE DEGREE OF\n")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(10.5)
    r_sub.font.color.rgb = RGBColor(80, 90, 100)

    r_deg = p_sub.add_run("BACHELOR OF SCIENCE (DATA SCIENCE)")
    r_deg.font.name = "Times New Roman"
    r_deg.font.size = Pt(13)
    r_deg.font.bold = True
    r_deg.font.color.rgb = RGBColor(20, 30, 50)

    p_auth = doc.add_paragraph()
    format_paragraph(p_auth, space_before=25, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_sb = p_auth.add_run("SUBMITTED BY\n")
    r_sb.font.name = "Times New Roman"
    r_sb.font.size = Pt(10.5)
    r_sb.font.bold = True
    r_sb.font.color.rgb = RGBColor(90, 100, 115)

    r_name = p_auth.add_run("MR. IVAN DSOUZA\n")
    r_name.font.name = "Times New Roman"
    r_name.font.size = Pt(14)
    r_name.font.bold = True
    r_name.font.color.rgb = RGBColor(15, 30, 65)

    r_rn = p_auth.add_run("Roll No.: TDDS013A  |  PRN: 2023016401824103\nT.Y. B.Sc. Data Science  |  Semester VI")
    r_rn.font.name = "Times New Roman"
    r_rn.font.size = Pt(11)
    r_rn.font.color.rgb = RGBColor(50, 60, 75)

    p_gui = doc.add_paragraph()
    format_paragraph(p_gui, space_before=22, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_ug = p_gui.add_run("UNDER THE ESTEEMED GUIDANCE OF\n")
    r_ug.font.name = "Times New Roman"
    r_ug.font.size = Pt(10.5)
    r_ug.font.bold = True
    r_ug.font.color.rgb = RGBColor(90, 100, 115)

    r_gname = p_gui.add_run("Mr./Ms. __________________________\nDepartment of Data Science & Information Technology")
    r_gname.font.name = "Times New Roman"
    r_gname.font.size = Pt(12)
    r_gname.font.bold = True
    r_gname.font.color.rgb = RGBColor(20, 30, 50)

    p_yr = doc.add_paragraph()
    format_paragraph(p_yr, space_before=30, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_yr = p_yr.add_run("ACADEMIC YEAR 2025 – 2026")
    r_yr.font.name = "Times New Roman"
    r_yr.font.size = Pt(12)
    r_yr.font.bold = True
    r_yr.font.color.rgb = RGBColor(15, 30, 65)

    doc.add_page_break()

    # ==========================================
    # 2. CERTIFICATE
    # ==========================================
    p_c1 = doc.add_paragraph()
    format_paragraph(p_c1, space_before=15, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    rc1 = p_c1.add_run("Kandivli Education Society's\nBK Shroff College of Arts & MH Shroff College of Commerce\n")
    rc1.font.name = "Times New Roman"
    rc1.font.size = Pt(13)
    rc1.font.bold = True
    rc1.font.color.rgb = RGBColor(15, 30, 65)

    rc2 = p_c1.add_run("Kandivali (West), Mumbai - 400067\nDEPARTMENT OF DATA SCIENCE")
    rc2.font.name = "Times New Roman"
    rc2.font.size = Pt(11)
    rc2.font.bold = True
    rc2.font.color.rgb = RGBColor(40, 55, 80)

    p_ct = doc.add_paragraph()
    format_paragraph(p_ct, space_before=20, space_after=18, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_ct = p_ct.add_run("CERTIFICATE")
    r_ct.font.name = "Times New Roman"
    r_ct.font.size = Pt(17)
    r_ct.font.bold = True
    r_ct.font.underline = True
    r_ct.font.color.rgb = RGBColor(15, 30, 65)

    cert_text = (
        "This is to certify that Mr. IVAN DSOUZA, a bonafide student of THIRD YEAR Bachelor of Science "
        "in DATA SCIENCE (Semester VI), Division: A, Roll No.: TDDS013A, has successfully completed the "
        "Capstone Project entitled 'SURVIVAL-BASED CHURN MODELING WITH PERSONALIZED RETENTION TREATMENT RANKING' "
        "(RetainIQ, Project ID: BDS-11) during the Academic Year 2025 – 2026, in partial fulfillment of the requirements "
        "for the award of the Degree of Bachelor of Science (Data Science) affiliated with the University of Mumbai, "
        "under the curriculum guidelines of KES' Shroff College of Arts & Commerce (Autonomous)."
    )
    add_body_p(doc, cert_text, space_before=8, space_after=12, line_spacing=1.2)

    cert_text2 = (
        "The project work incorporates original research, empirical survival analysis, heterogeneous treatment effect "
        "estimation, prescriptive decision optimization, cryptographic audit logging, automated unit/integration testing, "
        "and production web microservice deployment carried out under institutional supervision."
    )
    add_body_p(doc, cert_text2, space_before=4, space_after=40, line_spacing=1.2)

    # Signatures Table
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    
    col_widths = [Inches(2.1), Inches(2.1), Inches(2.1)]
    for row in sig_table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    # Row 0: Guide, Teacher In-Charge, Principal
    c0 = sig_table.cell(0, 0)
    c1 = sig_table.cell(0, 1)
    c2 = sig_table.cell(0, 2)
    
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p0.add_run("_____________________\nMr./Ms. ____________________\nProject Guide\nDept. of Data Science")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(20, 30, 45)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p1.add_run("_____________________\nMs. Sweta Suman\nTeacher In-charge\nCoordinator, B.Sc. DS")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(20, 30, 45)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run("_____________________\nDr. Lily Bhushan\nPrincipal\nKES' Shroff College")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(20, 30, 45)

    # Row 1: External Examiner & College Seal
    c_ex0 = sig_table.cell(1, 0)
    c_ex1 = sig_table.cell(1, 1)
    c_ex2 = sig_table.cell(1, 2)
    
    p_ex0 = c_ex0.paragraphs[0]
    format_paragraph(p_ex0, space_before=30, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_ex0.add_run("_____________________\nExternal Examiner\nDate: ____ / ____ / 2026")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)

    p_ex1 = c_ex1.paragraphs[0]
    format_paragraph(p_ex1, space_before=30, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_ex1.add_run("COLLEGE SEAL / STAMP")
    r.font.name = "Times New Roman"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(120, 130, 140)

    p_ex2 = c_ex2.paragraphs[0]
    format_paragraph(p_ex2, space_before=30, space_after=0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = p_ex2.add_run("_____________________\nInternal Examiner\nDate: ____ / ____ / 2026")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)

    doc.add_page_break()

    # ==========================================
    # 3. PROFORMA FOR APPROVAL OF PROJECT PROPOSAL
    # ==========================================
    p_phead = doc.add_paragraph()
    format_paragraph(p_phead, space_before=10, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_ph = p_phead.add_run("PROFORMA FOR THE APPROVAL OF PROJECT PROPOSAL\nACADEMIC YEAR 2025 – 2026")
    r_ph.font.name = "Times New Roman"
    r_ph.font.size = Pt(14)
    r_ph.font.bold = True
    r_ph.font.color.rgb = RGBColor(15, 30, 65)

    prof_tbl = doc.add_table(rows=7, cols=2)
    prof_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prof_tbl, color="90A0B0", sz="6", val="single")

    prof_data = [
        ("Roll No. & Division:", "TDDS013A  |  Division A"),
        ("Student Name:", "IVAN DSOUZA"),
        ("Programme / Class:", "Bachelor of Science (Data Science)  |  T.Y. B.Sc. Semester VI"),
        ("Project Title:", "SURVIVAL-BASED CHURN MODELING WITH PERSONALIZED RETENTION TREATMENT RANKING (RetainIQ)"),
        ("Project Identification ID:", "BDS-11"),
        ("Name of Project Guide:", "Mr./Ms. __________________________"),
        ("Specialization Track:", "Advanced Machine Learning, Causal Inference & Time-to-Event Analytics")
    ]

    for idx, (label, val) in enumerate(prof_data):
        row = prof_tbl.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.2)
        set_cell_background(c0, "F0F4F8")
        set_cell_margins(c0, top=90, bottom=90, left=120, right=100)
        set_cell_margins(c1, top=90, bottom=90, left=120, right=100)
        
        p_l = c0.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_l = p_l.add_run(label)
        r_l.font.name = "Times New Roman"
        r_l.font.size = Pt(10.5)
        r_l.font.bold = True
        r_l.font.color.rgb = RGBColor(20, 35, 60)

        p_v = c1.paragraphs[0]
        p_v.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_v = p_v.add_run(val)
        r_v.font.name = "Times New Roman"
        r_v.font.size = Pt(10.5)
        r_v.font.color.rgb = RGBColor(30, 40, 50)

    p_spc = doc.add_paragraph()
    format_paragraph(p_spc, space_before=25, space_after=10)

    # Proforma signatures
    p_psig = doc.add_paragraph()
    format_paragraph(p_psig, space_before=20, space_after=10)
    
    prof_sig_tbl = doc.add_table(rows=1, cols=3)
    prof_sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, width in enumerate([Inches(2.1), Inches(2.1), Inches(2.1)]):
        prof_sig_tbl.rows[0].cells[idx].width = width

    pc0, pc1, pc2 = prof_sig_tbl.rows[0].cells[0], prof_sig_tbl.rows[0].cells[1], prof_sig_tbl.rows[0].cells[2]
    
    p0 = pc0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run("_____________________\nSignature of the Student\nDate: ____ / ____ / 2026")
    r0.font.name = "Times New Roman"
    r0.font.size = Pt(10)

    p1 = pc1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run("_____________________\nSignature of the Guide\nDate: ____ / ____ / 2026")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(10)

    p2 = pc2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("_____________________\nSignature of Coordinator\nDate: ____ / ____ / 2026")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10)

    doc.add_page_break()

    # ==========================================
    # 4. ABSTRACT (2 FULL PAGES)
    # ==========================================
    p_absh = doc.add_paragraph()
    format_paragraph(p_absh, space_before=10, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_absh = p_absh.add_run("ABSTRACT")
    r_absh.font.name = "Times New Roman"
    r_absh.font.size = Pt(16)
    r_absh.font.bold = True
    r_absh.font.color.rgb = RGBColor(15, 30, 65)

    add_body_p(
        doc,
        "Customer churn represents one of the most critical operational risks confronting modern subscription-based and "
        "recurring-revenue enterprises. In competitive markets such as telecommunications, Software-as-a-Service (SaaS), "
        "and digital media, customer acquisition costs frequently exceed retention investments by a factor of five to seven. "
        "Consequently, preserving existing customer relationships constitutes a primary driver of enterprise valuation and "
        "long-term profitability. However, conventional commercial churn prediction pipelines remain fundamentally limited by "
        "an over-reliance on static binary classification formulations (e.g., Logistic Regression, standard XGBoost, and "
        "Random Forests). These legacy systems frame attrition as an arbitrary point-in-time classification problem: predicting "
        "whether an account will churn within an ad-hoc window such as 30 or 90 days. This paradigm suffers from two fatal "
        "theoretical and economic flaws: survival blindness and right-censoring distortion.",
        bold_prefix="Background & Problem Formulation: "
    )

    add_body_p(
        doc,
        "First, binary classifiers treat time as a discarded nuisance variable. A customer departing at Day 3 post-onboarding "
        "is treated identically to a tenured subscriber departing at Day 1,000. Second, binary models cannot handle "
        "right-censoring—the ubiquitous statistical reality that active subscribers have not yet experienced the churn event "
        "at the snapshot date. By forcibly coercing right-censored observations into zero-labels or dropping them altogether, "
        "conventional models introduce severe downward survival bias and miss the underlying temporal hazard trajectories. "
        "Even more destructively, commercial systems routinely mistake prediction for prescription: marketing teams "
        "systematically target the highest-risk churn cohorts with aggressive financial discounts. This practice squanders "
        "retention budgets on 'Lost Causes' (customers determined to churn regardless of intervention) and subsidizes "
        "'Sure Things' (subscribers who would have renewed organically), while risking friction among 'Sleeping Dogs' who "
        "are prompted to churn upon contact.",
        bold_prefix="Theoretical Deficiencies of Legacy Pipelines: "
    )

    add_body_p(
        doc,
        "To resolve these structural deficiencies, this capstone project introduces RetainIQ (Project ID: BDS-11), an end-to-end, "
        "production-grade decision-support ecosystem that unifies statistical survival analysis with causal uplift modeling and "
        "prescriptive economic optimization. RetainIQ replaces point-in-time churn probabilities with continuous survival "
        "curves S(t) = P(T > t) and hazard functions h(t). The system benchmarks four survival paradigms on an authentic "
        "telecommunications cohort of 7,043 subscribers: non-parametric Kaplan-Meier estimation, semi-parametric Cox Proportional "
        "Hazards (M1), parametric Weibull Accelerated Failure Time (M2), and non-linear Random Survival Forests (M3). Rigorous "
        "time-dependent validation demonstrates that Weibull AFT and Cox Proportional Hazards achieve peak discriminatory power "
        "with test Concordance Indices (C-Index) of 0.8697 and 0.8694 respectively, vastly outperforming a naive binary baseline (0.8283). "
        "At the operational 90-day horizon, the Weibull AFT model yields an exceptional Integrated Brier Score (IBS) of 0.0818, "
        "demonstrating near-perfect probabilistic calibration across deciles.",
        bold_prefix="The RetainIQ Architectural Framework: "
    )

    # Abstract Page 2
    add_body_p(
        doc,
        "Beyond predicting survival durations, RetainIQ formulates personalized retention treatment allocation through the lens of "
        "causal uplift modeling. The system models heterogeneous treatment effects across four distinct operational interventions: "
        "Basic Retention Check-in ($5.00), Proactive Tech Support ($12.00), 15% Bill Discount ($18.50), and Long-Term Contract "
        "Incentive ($35.00). By deploying Meta-Learner architectures (S-Learner, T-Learner, and X-Learner), RetainIQ estimates the "
        "Conditional Average Treatment Effect (CATE) tau(x) = Delta S(t|x) representing the absolute counterfactual gain in survival "
        "probability attributable to each specific intervention. Empirical uplift evaluation reveals that the T-Learner achieves "
        "the highest ranking effectiveness with a mean Qini score of 0.6920, establishing that treatment responsiveness is "
        "orthogonal to raw baseline risk.",
        bold_prefix="Heterogeneous Treatment Effects & Uplift Optimization: "
    )

    add_body_p(
        doc,
        "The estimated individual treatment effects are fed into RetainIQ's Prescriptive Ranking Engine, which solves a constrained "
        "knapsack optimization problem grounded in Expected Retention Value (ERV). ERV is defined mathematically as "
        "ERV_ik = Delta S_ik(t) * LTV_i - C_k, where LTV_i denotes customer lifetime value and C_k is the unit intervention cost. "
        "RetainIQ enforces a strict economic fallback policy: if no candidate treatment yields a strictly positive ERV, the customer "
        "is prescribed 'No Intervention' (Cost: $0.00). Across a simulated campaign portfolio of 1,873 high-risk accounts, RetainIQ "
        "optimizes treatment allocation under a $15,000 budget cap, allocating 1,248 accounts to generate $48,920 in Net Expected "
        "Retention Value, delivering a remarkable 329.2% campaign return on investment (ROI).",
        bold_prefix="Prescriptive Economics & Expected Retention Value (ERV): "
    )

    add_body_p(
        doc,
        "To ensure institutional trust and auditability, RetainIQ incorporates a tamper-evident cryptographic audit ledger using "
        "SHA-256 hash chaining, guaranteeing immutable logging of every algorithmic prescription and manual analyst override. "
        "Production monitoring implements Population Stability Index (PSI) tracking (recording a healthy max PSI of 0.0120), "
        "while adversarial stress testing confirms model resilience under 30% missing data and Gaussian noise perturbations. "
        "The system is delivered through a low-latency FastAPI microservice (<15ms response latency) and an 11-page interactive "
        "Streamlit decision dashboard. RetainIQ bridges the gap between theoretical survival analysis and commercial prescriptive "
        "economics, providing enterprise retention leaders with a scientifically grounded, economically optimal retention framework.",
        bold_prefix="Enterprise Governance, Auditability & System Deployment: "
    )

    p_kw = doc.add_paragraph()
    format_paragraph(p_kw, space_before=12, space_after=6)
    r_kw = p_kw.add_run("Keywords: ")
    r_kw.font.name = "Times New Roman"
    r_kw.font.size = Pt(10.5)
    r_kw.font.bold = True
    r_kwt = p_kw.add_run(
        "Survival Analysis, Right Censoring, Kaplan-Meier, Cox Proportional Hazards, Weibull AFT, "
        "Random Survival Forests, Causal Uplift Modeling, T-Learner, Expected Retention Value (ERV), "
        "Prescriptive Optimization, Cryptographic Audit Trail, FastAPI, Streamlit."
    )
    r_kwt.font.name = "Times New Roman"
    r_kwt.font.size = Pt(10)
    r_kwt.font.italic = True

    doc.add_page_break()

    # ==========================================
    # 5. ACKNOWLEDGEMENT
    # ==========================================
    p_ackh = doc.add_paragraph()
    format_paragraph(p_ackh, space_before=10, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_ah = p_ackh.add_run("ACKNOWLEDGEMENT")
    r_ah.font.name = "Times New Roman"
    r_ah.font.size = Pt(16)
    r_ah.font.bold = True
    r_ah.font.color.rgb = RGBColor(15, 30, 65)

    add_body_p(
        doc,
        "The completion of this Capstone Project, entitled 'SURVIVAL-BASED CHURN MODELING WITH PERSONALIZED "
        "RETENTION TREATMENT RANKING' (RetainIQ, Project ID: BDS-11), has been an enriching academic and technical "
        "journey. I take this opportunity to express my profound gratitude to everyone who contributed their invaluable "
        "guidance, constructive criticism, and encouragement throughout the ideation, implementation, and reporting phases."
    )

    add_body_p(
        doc,
        "First and foremost, I wish to express my deepest sense of gratitude and indebtedness to my esteemed Project Guide, "
        "Mr./Ms. __________________________, Department of Data Science, for their constant guidance, intellectual mentorship, "
        "and meticulous scrutiny of this work. Their profound insights into survival statistics, causal machine learning, "
        "and software architecture provided the technical clarity necessary to execute this rigorous undertaking."
    )

    add_body_p(
        doc,
        "I extend my sincere thanks to Ms. Sweta Suman, Teacher In-charge and Course Coordinator of the B.Sc. Data Science "
        "programme, for providing an academically stimulating environment, essential departmental infrastructure, and "
        "continuous administrative support throughout the academic year."
    )

    add_body_p(
        doc,
        "I also convey my deep respect and gratitude to our respected Principal, Dr. Lily Bhushan, whose visionary leadership "
        "and commitment to fostering advanced applied research at KES' Shroff College of Arts & Commerce has inspired us to "
        "strive for academic and technical excellence."
    )

    add_body_p(
        doc,
        "I am immensely grateful to all the esteemed faculty members of the Department of Data Science and Information "
        "Technology for imparting foundational mathematical and computational principles that formed the bedrock of this project. "
        "Special thanks are due to the laboratory assistants and administrative staff for their seamless technical assistance."
    )

    add_body_p(
        doc,
        "Finally, I owe an immeasurable debt of gratitude to my parents and family members for their unconditional love, "
        "patience, and moral support throughout my academic tenure. I also thank my peers and fellow classmates for their "
        "spirited discussions, collaborative feedback, and camaraderie.",
        space_after=35
    )

    p_asig = doc.add_paragraph()
    format_paragraph(p_asig, space_before=15, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    ras = p_asig.add_run("MR. IVAN DSOUZA\nRoll No.: TDDS013A\nT.Y. B.Sc. Data Science, Semester VI\nKES' Shroff College, Mumbai")
    ras.font.name = "Times New Roman"
    ras.font.size = Pt(11)
    ras.font.bold = True
    ras.font.color.rgb = RGBColor(15, 30, 65)

    doc.add_page_break()

    # ==========================================
    # 6. DECLARATION
    # ==========================================
    p_dech = doc.add_paragraph()
    format_paragraph(p_dech, space_before=10, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_dh = p_dech.add_run("DECLARATION")
    r_dh.font.name = "Times New Roman"
    r_dh.font.size = Pt(16)
    r_dh.font.bold = True
    r_dh.font.color.rgb = RGBColor(15, 30, 65)

    add_body_p(
        doc,
        "I, MR. IVAN DSOUZA, student of Third Year Bachelor of Science in Data Science (Semester VI), Division: A, "
        "Roll No.: TDDS013A, studying at Kandivli Education Society's BK Shroff College of Arts & MH Shroff College of "
        "Commerce (Autonomous), Kandivali (West), Mumbai - 400067, hereby declare that the Capstone Project report entitled:",
        space_before=10, space_after=10
    )

    p_dtitle = doc.add_paragraph()
    format_paragraph(p_dtitle, space_before=10, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    rdt = p_dtitle.add_run("“SURVIVAL-BASED CHURN MODELING WITH PERSONALIZED RETENTION TREATMENT RANKING”\n(RetainIQ  |  Project ID: BDS-11)")
    rdt.font.name = "Times New Roman"
    rdt.font.size = Pt(13)
    rdt.font.bold = True
    rdt.font.color.rgb = RGBColor(15, 35, 75)

    add_body_p(
        doc,
        "represents my own authentic, original work carried out under the supervision of my project guide, "
        "Mr./Ms. __________________________, Department of Data Science, KES' Shroff College, during the Academic Year 2025 – 2026."
    )

    add_body_p(
        doc,
        "I further affirm that this project report has been compiled by me in partial fulfillment of the academic requirements "
        "for the degree of Bachelor of Science (Data Science). To the best of my knowledge and belief, this work contains no "
        "material previously published or written by another person, nor has any part of this work been submitted previously "
        "to this or any other university, institute, or examining body for the award of any degree, diploma, or qualification."
    )

    add_body_p(
        doc,
        "All data sources, analytical formulas, algorithms, software libraries, and literature consulted have been duly acknowledged "
        "and cited in accordance with academic standards and ethical reporting guidelines. I bear full and sole responsibility "
        "for any inaccuracies, misstatements, or omissions present in this report.",
        space_after=45
    )

    p_dsig = doc.add_paragraph()
    format_paragraph(p_dsig, space_before=20, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    rds = p_dsig.add_run("_______________________________\nMR. IVAN DSOUZA\nRoll No.: TDDS013A\nDate: ____ / ____ / 2026\nPlace: Mumbai")
    rds.font.name = "Times New Roman"
    rds.font.size = Pt(11)
    rds.font.bold = True
    rds.font.color.rgb = RGBColor(15, 30, 65)

    doc.add_page_break()

    # ==========================================
    # 7. TABLE OF CONTENTS (2 FULL PAGES)
    # ==========================================
    p_toch = doc.add_paragraph()
    format_paragraph(p_toch, space_before=10, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_th = p_toch.add_run("TABLE OF CONTENTS")
    r_th.font.name = "Times New Roman"
    r_th.font.size = Pt(16)
    r_th.font.bold = True
    r_th.font.color.rgb = RGBColor(15, 30, 65)

    toc_table = doc.add_table(rows=1, cols=3)
    toc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(toc_table, color="A0B0C0", sz="4", val="single")
    
    # Headers
    h0, h1, h2 = toc_table.rows[0].cells[0], toc_table.rows[0].cells[1], toc_table.rows[0].cells[2]
    h0.width, h1.width, h2.width = Inches(1.1), Inches(4.5), Inches(0.9)
    for h, title in [(h0, "SR. NO."), (h1, "TOPIC / CHAPTER TITLE"), (h2, "PAGE")]:
        set_cell_background(h, "152035")
        set_cell_margins(h, top=80, bottom=80, left=100, right=80)
        p = h.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if h != h1 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    toc_rows = [
        # Preliminary
        ("", "PRELIMINARY PAGES", "i–xii"),
        ("", "Certificate", "ii"),
        ("", "Proforma for Approval of Project Proposal", "iii"),
        ("", "Abstract", "iv–v"),
        ("", "Acknowledgement", "vi"),
        ("", "Declaration", "vii"),
        ("", "Table of Contents", "viii–ix"),
        ("", "List of Figures", "x–xi"),
        ("", "List of Tables & List of Abbreviations", "xii"),
        # Chapter 1
        ("1", "CHAPTER 1: INTRODUCTION", "13–17"),
        ("1.1", "Significance of the Project", "14"),
        ("1.2", "Project Objectives", "15"),
        ("1.3", "Purpose and Scope", "15"),
        ("1.3.1", "Purpose", "15"),
        ("1.3.2", "Scope and Boundary Conditions", "16"),
        ("1.4", "Enterprise Applicability", "16"),
        ("1.5", "Technical and Empirical Achievements", "17"),
        # Chapter 2
        ("2", "CHAPTER 2: SYSTEM ANALYSIS", "18–23"),
        ("2.1", "Existing System and Theoretical Limitations", "18"),
        ("2.2", "Proposed System: RetainIQ Architecture", "19"),
        ("2.3", "Requirements Analysis (Functional, ML, Dashboard, API, Governance)", "19"),
        ("2.4", "Hardware Requirements", "21"),
        ("2.5", "Software Requirements", "22"),
        ("2.6", "Survey of Technology and Algorithmic Foundations", "22"),
        # Chapter 3
        ("3", "CHAPTER 3: SYSTEM DESIGN", "24–32"),
        ("3.1", "Module Division and Functional Decomposition", "24"),
        ("3.2", "System Architecture (Figure 3.1)", "25"),
        ("3.3", "Project Development Gantt Chart (Figure 3.2)", "26"),
        ("3.4", "Data Flow Representation (DFD Level 0 & Level 1) (Figure 3.3)", "27"),
        ("3.5", "UML Modeling", "28"),
        ("3.5.1", "UML Class Diagram (Figure 3.4)", "28"),
        ("3.5.2", "UML Sequence Diagram: Scoring & Prescription (Figure 3.5)", "29"),
        ("3.5.3", "UML State Chart: Churn State Machine (Figure 3.6a)", "30"),
        ("3.5.4", "UML Use-Case Model (Figure 3.6b)", "30"),
        ("3.6", "Data Model & Treatment Configuration Architecture", "31"),
        # Chapter 4
        ("4", "CHAPTER 4: IMPLEMENTATION AND TESTING", "33–39"),
        ("4.1", "Data Ingestion and Cleaning Implementation", "33"),
        ("4.2", "Leakage Prevention Pipeline", "33"),
        ("4.3", "Survival Modeling Engines (KM, Cox PH, Weibull AFT, RSF)", "34"),
        ("4.4", "Causal Uplift & Meta-Learner Implementation", "34"),
        ("4.5", "Personalized Ranking & ERV Optimization Engine", "35"),
        ("4.6", "Representative Source Code Listings (Listings 4.1 – 4.6)", "35"),
        ("4.7", "Testing Methodology & Quality Assurance", "38"),
        ("4.8", "Automated Test Results Matrix (27 Test Cases Verified)", "39"),
        # Chapter 5
        ("5", "CHAPTER 5: RESULTS AND DISCUSSIONS", "40–51"),
        ("5.1", "Dataset Composition and Exploratory Demographics", "40"),
        ("5.2", "Data Quality, Imputation & Leakage Audit Findings", "41"),
        ("5.3", "Kaplan-Meier Empirical Survival Analysis (Figures 5.1 – 5.3)", "41"),
        ("5.4", "Survival Model Benchmarking & Performance (Figure 5.4)", "43"),
        ("5.5", "Time-Dependent Probability Calibration (Figure 5.5)", "44"),
        ("5.6", "Customer Risk Stratification Tiers", "44"),
        ("5.7", "Treatment Effect & Uplift Modeling Findings (Figure 5.6)", "45"),
        ("5.8", "Prescriptive Ranking & Portfolio ERV Optimization (Figure 5.7)", "45"),
        ("5.9", "Model Explainability & Hazard Attribution (Figures 5.8 – 5.9)", "46"),
        ("5.10", "System Robustness & Adversarial Stress Testing (Figure 5.10)", "49"),
        ("5.11", "Population Drift & Monitoring Metrics (PSI)", "49"),
        ("5.12", "Interactive Dashboard & REST API Performance (Figures 5.11 – 5.12)", "50"),
        # Chapter 6
        ("6", "CHAPTER 6: CONCLUSION AND FUTURE WORK", "52–53"),
        ("6.1", "Project Synthesis and Empirical Conclusion", "52"),
        ("6.2", "Future Scope and Architectural Extensions", "52"),
        ("6.3", "Real-World Operational Limitations", "53"),
        # Chapter 7
        ("7", "CHAPTER 7: REFERENCES", "54–55")
    ]

    for sr, topic, page in toc_rows:
        row = toc_table.add_row()
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(1.0), Inches(4.7), Inches(0.8)
        set_cell_margins(c0, top=24, bottom=24, left=60, right=40)
        set_cell_margins(c1, top=24, bottom=24, left=60, right=40)
        set_cell_margins(c2, top=24, bottom=24, left=60, right=40)

        is_chap = (sr.isdigit() and len(sr) == 1) or "PRELIMINARY" in topic
        bg_col = "F4F6F9" if is_chap else "FFFFFF"
        for c in [c0, c1, c2]:
            set_cell_background(c, bg_col)

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.line_spacing = 1.05
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(sr)
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(8.5)
        if is_chap:
            r0.font.bold = True

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p1.paragraph_format.line_spacing = 1.05
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(topic)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(8.5)
        if is_chap:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(15, 30, 60)

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.line_spacing = 1.05
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(page)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(8.5)
        if is_chap:
            r2.font.bold = True

    doc.add_page_break()

    # ==========================================
    # 8. LIST OF FIGURES (2 FULL PAGES)
    # ==========================================
    p_lfh = doc.add_paragraph()
    format_paragraph(p_lfh, space_before=10, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_lfh = p_lfh.add_run("LIST OF FIGURES")
    r_lfh.font.name = "Times New Roman"
    r_lfh.font.size = Pt(16)
    r_lfh.font.bold = True
    r_lfh.font.color.rgb = RGBColor(15, 30, 65)

    lof_table = doc.add_table(rows=1, cols=3)
    lof_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(lof_table, color="A0B0C0", sz="4", val="single")

    h0, h1, h2 = lof_table.rows[0].cells[0], lof_table.rows[0].cells[1], lof_table.rows[0].cells[2]
    h0.width, h1.width, h2.width = Inches(1.2), Inches(4.4), Inches(0.9)
    for h, title in [(h0, "FIGURE NO."), (h1, "FIGURE CAPTION / DESCRIPTION"), (h2, "PAGE")]:
        set_cell_background(h, "152035")
        set_cell_margins(h, top=80, bottom=80, left=100, right=80)
        p = h.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if h != h1 else WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    figures_list = [
        ("Figure 3.1", "Overall RetainIQ Multi-Stage System Architecture", "25"),
        ("Figure 3.2", "RetainIQ Engineering Lifecycle and Development Gantt Chart", "26"),
        ("Figure 3.3", "Data Flow Diagrams: Level 0 Context and Level 1 Functional Decomposition", "27"),
        ("Figure 3.4", "UML Class Diagram of RetainIQ Domain Models and Pipeline Classes", "28"),
        ("Figure 3.5", "UML Sequence Diagram: Asynchronous Customer Scoring and Prescription Lifecycle", "29"),
        ("Figure 3.6", "UML State Machine and Enterprise Use-Case Interaction Architecture", "30"),
        ("Figure 5.1", "Kaplan-Meier Baseline Survival Probability Curve with 95% Confidence Intervals", "41"),
        ("Figure 5.2", "Kaplan-Meier Survival Curves Stratified by Customer Contract Type (Log-Rank p < 0.0001)", "42"),
        ("Figure 5.3", "Kaplan-Meier Survival Curves Stratified by Internet Service Architecture", "43"),
        ("Figure 5.4", "Benchmark Model Comparison: Concordance Index and Integrated Brier Score Across Splits", "43"),
        ("Figure 5.5", "Time-Dependent Probability Calibration Curve at the 90-Day Operational Horizon", "44"),
        ("Figure 5.6", "Conditional Average Treatment Effect (CATE) Distributions Across Four Interventions", "45"),
        ("Figure 5.7", "Portfolio Expected Retention Value (ERV) and Net ROI Frontier Under Budget Constraints", "46"),
        ("Figure 5.8", "Cox Proportional Hazards Forest Plot with 95% Hazard Ratio Confidence Intervals", "47"),
        ("Figure 5.9", "Global Permutation Feature Importance for Survival Hazard Attribution", "47"),
        ("Figure 5.10", "System Robustness Degradation Curves Under Feature Missingness and Gaussian Noise", "49"),
        ("Figure 5.11", "RetainIQ Streamlit Dashboard: Executive Portfolio Overview & Risk Stratification Cockpit", "50"),
        ("Figure 5.12", "RetainIQ Streamlit Dashboard: Treatment Studio & Campaign ROI Optimizer", "50")
    ]

    for f_no, caption, pg in figures_list:
        row = lof_table.add_row()
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(1.2), Inches(4.4), Inches(0.9)
        set_cell_margins(c0, top=50, bottom=50, left=80, right=60)
        set_cell_margins(c1, top=50, bottom=50, left=80, right=60)
        set_cell_margins(c2, top=50, bottom=50, left=80, right=60)

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f_no)
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(9.5)
        r0.font.bold = True

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(caption)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(9.5)

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(pg)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(9.5)

    doc.add_page_break()

    # ==========================================
    # 9. LIST OF TABLES & LIST OF ABBREVIATIONS
    # ==========================================
    p_lth = doc.add_paragraph()
    format_paragraph(p_lth, space_before=6, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_lth = p_lth.add_run("LIST OF TABLES")
    r_lth.font.name = "Times New Roman"
    r_lth.font.size = Pt(13)
    r_lth.font.bold = True
    r_lth.font.color.rgb = RGBColor(15, 30, 65)

    lot_table = doc.add_table(rows=1, cols=3)
    lot_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(lot_table, color="A0B0C0", sz="4", val="single")

    th0, th1, th2 = lot_table.rows[0].cells[0], lot_table.rows[0].cells[1], lot_table.rows[0].cells[2]
    th0.width, th1.width, th2.width = Inches(1.1), Inches(4.6), Inches(0.8)
    for h, title in [(th0, "TABLE NO."), (th1, "TABLE TITLE / DESCRIPTION"), (th2, "PAGE")]:
        set_cell_background(h, "152035")
        set_cell_margins(h, top=35, bottom=35, left=60, right=50)
        p = h.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if h != th1 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    tables_list = [
        ("Table 2.1", "RetainIQ Hardware Environment Specifications", "21"),
        ("Table 2.2", "Software Stack, Dependencies and Engineering Frameworks", "22"),
        ("Table 4.1", "Comprehensive Automated Test Results Matrix (27 Test Cases)", "39"),
        ("Table 5.1", "IBM Telco Cohort Demographic and Commercial Data Dictionary", "40"),
        ("Table 5.2", "Survival Benchmark Model Performance Matrix (C-Index, IBS, Horizon AUC)", "43"),
        ("Table 5.3", "Customer Churn Risk Stratification Tier Distribution", "44"),
        ("Table 5.4", "Causal Uplift Meta-Learner Performance Evaluation (PEHE RMSE & Qini)", "45"),
        ("Table 5.5", "Portfolio Retention Campaign Budget Sweep and ERV Optimization", "46"),
        ("Table 5.6", "Cox Proportional Hazards Multivariate Regression Summary", "47"),
        ("Table 5.7", "Population Stability Index (PSI) Feature Drift Audit", "50")
    ]

    for t_no, t_title, t_pg in tables_list:
        row = lot_table.add_row()
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(1.1), Inches(4.6), Inches(0.8)
        set_cell_margins(c0, top=18, bottom=18, left=60, right=50)
        set_cell_margins(c1, top=18, bottom=18, left=60, right=50)
        set_cell_margins(c2, top=18, bottom=18, left=60, right=50)

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(t_no)
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(8)
        r0.font.bold = True

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(t_title)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(8)

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(t_pg)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(8)

    # Abbreviations
    p_abh = doc.add_paragraph()
    format_paragraph(p_abh, space_before=8, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_abh = p_abh.add_run("LIST OF ABBREVIATIONS & ACRONYMS")
    r_abh.font.name = "Times New Roman"
    r_abh.font.size = Pt(13)
    r_abh.font.bold = True
    r_abh.font.color.rgb = RGBColor(15, 30, 65)

    abb_table = doc.add_table(rows=1, cols=4)
    abb_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(abb_table, color="B0B8C0", sz="4", val="single")

    ac_widths = [Inches(0.9), Inches(2.3), Inches(0.9), Inches(2.4)]
    for idx, width in enumerate(ac_widths):
        abb_table.rows[0].cells[idx].width = width

    for idx, (c, title) in enumerate(zip(abb_table.rows[0].cells, ["ABBR.", "EXPANSION", "ABBR.", "EXPANSION"])):
        set_cell_background(c, "152035")
        set_cell_margins(c, top=35, bottom=35, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(8)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    abbreviations = [
        ("AFT", "Accelerated Failure Time", "KM", "Kaplan-Meier Estimator"),
        ("API", "Application Programming Interface", "LTV", "Lifetime Value"),
        ("AUC", "Area Under the ROC Curve", "ML", "Machine Learning"),
        ("C-Index", "Harrell's Concordance Index", "PEHE", "Precision in Estimating Heterogeneous Effect"),
        ("CATE", "Conditional Average Treatment Effect", "PSI", "Population Stability Index"),
        ("CI", "Confidence Interval", "REST", "Representational State Transfer"),
        ("Cox PH", "Cox Proportional Hazards Model", "ROI", "Return on Investment"),
        ("DFD", "Data Flow Diagram", "RSF", "Random Survival Forest"),
        ("EDA", "Exploratory Data Analysis", "SaaS", "Software as a Service"),
        ("ERV", "Expected Retention Value", "SHAP", "Shapley Additive Explanations"),
        ("HR", "Hazard Ratio", "UML", "Unified Modeling Language"),
        ("IBS", "Integrated Brier Score", "YAML", "YAML Ain't Markup Language")
    ]

    for a1, e1, a2, e2 in abbreviations:
        row = abb_table.add_row()
        for idx, width in enumerate(ac_widths):
            row.cells[idx].width = width
            set_cell_margins(row.cells[idx], top=16, bottom=16, left=50, right=50)
        
        row.cells[0].paragraphs[0].add_run(a1).font.bold = True
        row.cells[1].paragraphs[0].add_run(e1)
        row.cells[2].paragraphs[0].add_run(a2).font.bold = True
        row.cells[3].paragraphs[0].add_run(e2)
        for i in range(4):
            p = row.cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(7.8)

    doc.add_page_break()
    print("Front matter generated successfully.")
