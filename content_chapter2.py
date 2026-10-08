"""
Chapter 2 Builder for RetainIQ Black Book.
Covers:
2.1 Existing System
2.2 Proposed System
2.3 Requirements Analysis (A to F)
2.4 Hardware Requirements
2.5 Software Requirements
2.6 Survey of Technology
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from report_helpers import (
    add_body_p, add_heading_1, add_heading_2, add_heading_3,
    set_cell_background, set_cell_margins, set_table_borders, format_paragraph
)

def build_chapter_2(doc):
    add_heading_1(doc, "CHAPTER 2: SYSTEM ANALYSIS", space_before=15, space_after=12)

    add_body_p(
        doc,
        "System analysis constitutes the critical diagnostic phase of software and data science engineering. In this phase, the "
        "structural deficiencies of prevailing industry practices are methodically examined, the theoretical and commercial foundations "
        "of the proposed architecture are established, and comprehensive functional and operational requirements are defined. This chapter "
        "presents an exhaustive comparative critique of existing commercial churn prediction systems, introduces the architectural paradigm "
        "of RetainIQ, details hardware and software dependencies, and provides an in-depth survey of the algorithmic foundations enabling "
        "survival-based retention modeling."
    )

    # ==========================================
    # 2.1 EXISTING SYSTEM
    # ==========================================
    add_heading_2(doc, "2.1 EXISTING SYSTEM AND THEORETICAL LIMITATIONS")

    add_body_p(
        doc,
        "In contemporary enterprise environments, customer churn prediction pipelines are almost universally framed as standard supervised "
        "binary classification tasks. In such architectures, customer account telemetry (e.g., tenure, monthly charges, contract type, and "
        "service usage) is collected at an arbitrary snapshot date t_0. A target label Y in {0, 1} is assigned based on whether the account "
        "formally terminates its contract within a fixed forward-looking observation window Delta t (typically 30, 60, or 90 days):"
    )

    add_body_p(
        doc,
        "Y_i = 1 if T_i <= t_0 + Delta t, else 0",
        italic_prefix="   Mathematical Definition: ",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=4, space_after=6
    )

    add_body_p(
        doc,
        "Supervised classifiers—such as Logistic Regression, Random Forests, XGBoost, or multilayer perceptrons—are trained to estimate the "
        "conditional probability P(Y_i = 1 | X_i). While widely adopted due to off-the-shelf software accessibility, this formulation suffers "
        "from four catastrophic theoretical and practical failures:"
    )

    flaws = [
        ("1. Survival Blindness and Temporal Information Loss: ",
         "Binary classification compresses continuous time into a single boolean state. It cannot differentiate between a customer who cancels "
         "immediately on Day 1 versus a customer who cancels on Day 89. This temporal blindness prevents retention teams from sequencing "
         "proactive touchpoints, understanding early-tenure onboarding vulnerability, or projecting multi-quarter revenue depletion."),
        
        ("2. Right-Censoring Distortion and Downward Survival Bias: ",
         "In real-world recurring business models, the vast majority of subscribers remain active at any given analysis snapshot. In the IBM Telco "
         "cohort, 73.46% of accounts (5,174 out of 7,043) are right-censored—meaning their true lifetime T exceeds the observed duration. Binary "
         "classification algorithms possess no mathematical capacity to incorporate censored likelihood contributions. Labeling all active "
         "subscribers as '0' introduces severe downward bias into churn risk scores, treating ongoing loyalty as a static negative class rather "
         "than an unresolved continuous duration."),
        
        ("3. The Association vs. Causation Fallacy in Treatment Targeting: ",
         "Commercial retention teams routinely use the predicted churn probability P(Y=1|X) as an actionable rank-ordering for marketing campaigns. "
         "This represents a catastrophic conceptual error. A high churn probability indicates that an account is at high risk, but it does NOT "
         "indicate that the account will respond favorably to an intervention. Modern uplift literature categorizes customers into four behavioral segments: "
         "(a) 'Persuadables' (retain only if treated), (b) 'Sure Things' (retain regardless of treatment), (c) 'Lost Causes' (churn regardless of "
         "treatment), and (d) 'Sleeping Dogs' or 'Do Not Disturb' (churn specifically because an intrusive intervention was triggered). Targeting the "
         "highest-risk deciles inevitably wastes retention capital on 'Lost Causes' while missing high-leverage 'Persuadables' who possess moderate "
         "baseline risk but high treatment sensitivity."),
        
        ("4. Total Absence of Intervention Unit Economics: ",
         "Binary classifiers output unitless probabilities unmoored from financial reality. They provide no mechanism to incorporate the unit cost "
         "of retention treatments (e.g., $5 check-in call vs. $35 hardware incentive) against the residual Lifetime Value (LTV) of the customer. "
         "Consequently, marketing teams frequently deploy expensive interventions on low-value subscribers whose future margins can never recoup "
         "the cost of the retention campaign, resulting in negative net enterprise ROI.")
    ]

    for f_title, f_desc in flaws:
        add_body_p(doc, f_desc, bold_prefix=f_title, space_before=3, space_after=4)

    # ==========================================
    # 2.2 PROPOSED SYSTEM
    # ==========================================
    add_heading_2(doc, "2.2 PROPOSED SYSTEM: RETAINIQ ARCHITECTURE")

    add_body_p(
        doc,
        "RetainIQ fundamentally reframes the customer retention problem by integrating continuous-time statistical survival analysis with causal "
        "uplift modeling and prescriptive economic optimization. Instead of asking 'Will this customer churn?', RetainIQ asks: 'What is the full "
        "survival trajectory S_i(t) of this customer, which specific intervention will causally maximize their incremental survival gain Delta S_ik(t), "
        "and what is the net Expected Retention Value (ERV) after accounting for treatment costs?'"
    )

    add_body_p(
        doc,
        "RetainIQ introduces four major architectural innovations over existing commercial pipelines:",
        space_after=4
    )

    innovations = [
        ("Continuous Survival Dynamics: ",
         "The system estimates individual survival curves S_i(t) = P(T_i > t) and hazard rates h_i(t) across arbitrary time horizons. By correctly "
         "incorporating right-censored observations via partial and full likelihood estimation, RetainIQ eliminates downward survival bias and "
         "provides granular visibility into 30, 60, 90, 180, and 365-day retention probabilities from a single model."),
        
        ("Heterogeneous Causal Treatment Uplift (CATE): ",
         "RetainIQ implements causal Meta-Learner frameworks (T-Learner and S-Learner) to isolate the counterfactual effect of specific retention "
         "treatments. By modeling the Conditional Average Treatment Effect tau_ik(x) = E[S_k(t) - S_0(t) | X=x], the system targets only those "
         "subscribers who demonstrate true incremental responsiveness ('Persuadables')."),
        
        ("Prescriptive Knapsack ERV Optimization: ",
         "The system couples treatment uplift with customer Lifetime Value (LTV) and intervention unit costs C_k to compute net Expected "
         "Retention Value: ERV_ik = Delta S_ik(t) * LTV_i - C_k. A knapsack allocation engine solves the budget-constrained portfolio problem, "
         "enforcing an automated zero-cost fallback ('No Intervention') whenever ERV is non-positive."),
        
        ("Enterprise Governance and Operational Trust: ",
         "RetainIQ incorporates an immutable SHA-256 cryptographic audit trail recording all model recommendations and manual human overrides. "
         "Continuous monitoring evaluates Population Stability Index (PSI) drift, while automated adversarial suites stress-test pipeline resilience "
         "under simulated missingness and feature noise.")
    ]

    for in_title, in_desc in innovations:
        add_body_p(doc, in_desc, bold_prefix=in_title, space_before=2, space_after=3)

    # ==========================================
    # 2.3 REQUIREMENTS ANALYSIS
    # ==========================================
    add_heading_2(doc, "2.3 REQUIREMENTS ANALYSIS")

    add_body_p(
        doc,
        "The RetainIQ platform was designed under rigorous software engineering and machine learning specifications. The requirements are "
        "categorized into six distinct operational dimensions:"
    )

    # A. Functional
    add_heading_3(doc, "A. Functional Requirements")
    func_reqs = [
        ("FR-1 Data Ingestion & Sanitization: ", "The system shall ingest raw tabular customer records, identify and coerce whitespace-blank values "
         "to floating-point representations, and apply median imputation without inducing data leakage."),
        ("FR-2 Multi-Horizon Survival Scoring: ", "The system shall compute predicted survival probabilities S(t) and cumulative hazard H(t) "
         "for each customer across standard operational horizons: 30, 60, 90, 180, and 365 days."),
        ("FR-3 Multi-Treatment CATE Estimation: ", "The system shall estimate heterogeneous treatment effects for four distinct interventions: "
         "Basic Check-in, Proactive Tech Support, 15% Bill Discount, and Long-Term Contract Incentive."),
        ("FR-4 Prescriptive Treatment Ranking: ", "The system shall compute ERV for all eligible treatments, rank candidates in descending order, "
         "and select the optimal treatment while enforcing the fallback rule: assign 'No Intervention' if ERV <= 0."),
        ("FR-5 Cryptographic Audit Trail: ", "The system shall append every scoring, prescription, and analyst override event to an immutable "
         "SHA-256 cryptographic hash chain, enabling zero-knowledge tamper verification.")
    ]
    for r_title, r_desc in func_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # B. Non-Functional
    add_heading_3(doc, "B. Non-Functional Requirements")
    nfunc_reqs = [
        ("NFR-1 Latency & Throughput: ", "Single-customer inference endpoints shall return complete survival curves, treatment rankings, and "
         "ERV computations with an average response latency of less than 50 milliseconds (demonstrating <15ms in production)."),
        ("NFR-2 Memory & Compute Footprint: ", "The entire production serving environment, including loaded models and feature pipelines, "
         "shall consume less than 2.0 GB of system RAM, enabling deployment on cost-effective commodity compute instances."),
        ("NFR-3 Deterministic Reproducibility: ", "All pipeline preprocessing splits, model training seeds, and simulation runs shall be strictly "
         "deterministic, yielding identical results across repeated executions."),
        ("NFR-4 Scalability: ", "The batch scoring engine shall process a portfolio of 10,000 customer accounts within 5 seconds on standard hardware.")
    ]
    for r_title, r_desc in nfunc_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # C. Machine Learning
    add_heading_3(doc, "C. Machine Learning Requirements")
    ml_reqs = [
        ("MLR-1 Discriminatory Accuracy: ", "The primary survival models (Weibull AFT, Cox PH, and RSF) shall achieve a test Harrell's Concordance "
         "Index (C-Index) exceeding 0.85, significantly outperforming naive classification baselines."),
        ("MLR-2 Probabilistic Calibration: ", "The survival models shall achieve an Integrated Brier Score (IBS) below 0.10 at the 90-day operational "
         "decision horizon, ensuring reliable risk estimates across deciles."),
        ("MLR-3 Monotonicity Guarantee: ", "Predicted survival probability curves must satisfy the strict mathematical axiom of non-increasing "
         "monotonicity: S(t_2) <= S(t_1) for all t_2 > t_1."),
        ("MLR-4 Uplift Superiority: ", "The causal Meta-Learner architecture shall achieve a positive mean Qini score (>0.40), confirming statistically "
         "significant counterfactual ranking power over random targeting.")
    ]
    for r_title, r_desc in ml_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # D. Dashboard Requirements
    add_heading_3(doc, "D. Dashboard Requirements")
    dash_reqs = [
        ("DR-1 Multi-Page Modular Console: ", "The UI shall provide 11 dedicated operational pages covering Executive Overview, Data Quality, "
         "Cohort Survival, Model Lab, Customer 360, Treatment Studio, Campaign Planner, Explainability, Robustness, Governance, and Documentation."),
        ("DR-2 Interactive Counterfactual Simulator: ", "Users shall be able to dynamically adjust individual customer attributes and treatment "
         "parameters via sliders and inspect immediate recalculations of S(t), hazard percentiles, and ERV in real time."),
        ("DR-3 Campaign Budget Optimizer: ", "The campaign planner shall allow marketing directors to specify aggregate budget caps and instantly "
         "visualize optimal treatment allocation, customer coverage, and net campaign ROI curves.")
    ]
    for r_title, r_desc in dash_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # E. API Requirements
    add_heading_3(doc, "E. API Requirements")
    api_reqs = [
        ("AR-1 RESTful Interface: ", "The backend shall expose standards-compliant REST endpoints supporting GET and POST operations with "
         "JSON request/response payloads conforming to OpenAPI 3.0 specifications."),
        ("AR-2 Pydantic Schema Validation: ", "All incoming API requests shall be validated through strict Pydantic schemas, returning HTTP 422 "
         "Unprocessable Entity errors with descriptive field messages upon boundary violations."),
        ("AR-3 Enterprise Health Probes: ", "The API shall expose dedicated /health, /model-info, and /metrics endpoints to support Kubernetes "
         "liveness/readiness probes and Prometheus telemetry scraping.")
    ]
    for r_title, r_desc in api_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # F. Monitoring & Governance Requirements
    add_heading_3(doc, "F. Monitoring and Governance Requirements")
    gov_reqs = [
        ("GR-1 Drift Detection: ", "The system shall continuously compute Population Stability Index (PSI) metrics across all input features, "
         "flagging warnings when PSI exceeds 0.10 and triggering alerts when PSI exceeds 0.25."),
        ("GR-2 Stress Testing & Robustness: ", "The pipeline shall include automated adversarial test suites measuring performance degradation "
         "under simulated missingness (up to 30%) and additive Gaussian feature noise, establishing resilience boundaries."),
        ("GR-3 Tamper Verification: ", "The audit system shall provide an automated endpoint verifying the cryptographic continuity of the "
         "SHA-256 chain, immediately detecting any unauthorized manual database alterations.")
    ]
    for r_title, r_desc in gov_reqs:
        add_body_p(doc, r_desc, bold_prefix=r_title, space_before=2, space_after=2)

    # ==========================================
    # 2.4 HARDWARE REQUIREMENTS
    # ==========================================
    add_heading_2(doc, "2.4 HARDWARE REQUIREMENTS")

    add_body_p(
        doc,
        "RetainIQ is architected for maximum computational efficiency, allowing complete end-to-end training, statistical cross-validation, "
        "and real-time inference on commodity enterprise hardware without requiring expensive GPU accelerators. Table 2.1 delineates the "
        "minimum and recommended hardware specifications:"
    )

    # Table 2.1
    hw_tbl = doc.add_table(rows=1, cols=3)
    hw_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(hw_tbl, color="A0B0C0", sz="4", val="single")

    hw_cols = [Inches(1.8), Inches(2.3), Inches(2.4)]
    for idx, (c, title) in enumerate(zip(hw_tbl.rows[0].cells, ["COMPONENT", "MINIMUM SPECIFICATION", "PRODUCTION RECOMMENDED"])):
        c.width = hw_cols[idx]
        set_cell_background(c, "152035")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    hw_rows = [
        ("Central Processing Unit (CPU)", "Quad-core x86_64 Processor (2.0 GHz)", "8-core Intel Core i7 / AMD Ryzen 7 / Xeon"),
        ("System Memory (RAM)", "8.0 GB DDR4 SDRAM", "16.0 GB DDR4 / DDR5 Dual-Channel"),
        ("Disk Storage", "20 GB Available SSD Storage", "50 GB NVMe PCIe SSD (Fast I/O)"),
        ("Network Interface", "100 Mbps Ethernet / Wi-Fi", "1.0 Gbps Enterprise Low-Latency NIC"),
        ("Graphics Processing (GPU)", "Not Required (CPU Optimized)", "Not Required (Optional for DeepSurv)")
    ]

    for comp, min_s, rec_s in hw_rows:
        row = hw_tbl.add_row()
        for idx, (val, width) in enumerate(zip([comp, min_s, rec_s], hw_cols)):
            c = row.cells[idx]
            c.width = width
            set_cell_margins(c, top=45, bottom=45, left=80, right=80)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(9)
            if idx == 0:
                r.font.bold = True

    p_cap1 = doc.add_paragraph()
    format_paragraph(p_cap1, space_before=4, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    rc1 = p_cap1.add_run("Table 2.1: RetainIQ Hardware Environment Specifications")
    rc1.font.name = "Times New Roman"
    rc1.font.size = Pt(9)
    rc1.font.bold = True
    rc1.font.italic = True

    # ==========================================
    # 2.5 SOFTWARE REQUIREMENTS
    # ==========================================
    add_heading_2(doc, "2.5 SOFTWARE REQUIREMENTS")

    add_body_p(
        doc,
        "The software architecture of RetainIQ leverages modern, industry-standard open-source libraries within the Python data science and "
        "microservices ecosystem. Table 2.2 details the complete software dependency stack:"
    )

    # Table 2.2
    sw_tbl = doc.add_table(rows=1, cols=3)
    sw_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sw_tbl, color="A0B0C0", sz="4", val="single")

    sw_cols = [Inches(1.8), Inches(1.8), Inches(2.9)]
    for idx, (c, title) in enumerate(zip(sw_tbl.rows[0].cells, ["LAYER / CATEGORY", "TECHNOLOGY / PACKAGE", "VERSION & OPERATIONAL ROLE"])):
        c.width = sw_cols[idx]
        set_cell_background(c, "152035")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    sw_rows = [
        ("Base Runtime Environment", "Python Core Interpreter", "v3.11 / v3.14 (64-bit CPython Engine)"),
        ("Operating System Support", "Microsoft Windows / Linux", "Windows 11 Enterprise / Ubuntu 22.04 LTS"),
        ("Data Ingestion & Matrix Ops", "Pandas & NumPy", "Pandas 2.2+, NumPy 1.26+ (High-Performance Arrays)"),
        ("Survival Analysis Core", "Lifelines & Scikit-survival", "Lifelines 0.28+ (Cox, KM, Weibull AFT), RSF 0.23+"),
        ("Machine Learning Utilities", "Scikit-Learn & Joblib", "Scikit-learn 1.4+ (Transforms, Metrics, Serializers)"),
        ("Causal Uplift Modeling", "CausalML & Custom Learners", "T-Learner, S-Learner, and X-Learner Implementations"),
        ("RESTful Microservice API", "FastAPI & Uvicorn", "FastAPI 0.110+, Uvicorn 0.28+ (ASGI Async Engine)"),
        ("Data Validation Schemas", "Pydantic v2", "Pydantic 2.6+ (Strict Type Checking & OpenAPI Schema)"),
        ("Operational UI Dashboard", "Streamlit & Plotly", "Streamlit 1.32+, Plotly 5.20+ (Interactive Dark Theme)"),
        ("Quality Assurance & Testing", "Pytest & AnyIO", "Pytest 8.1+, Pytest-asyncio (27 Automated Test Cases)"),
        ("Audit & Cryptography", "Python hashlib (SHA-256)", "Standard Library Cryptographic Hash Chaining")
    ]

    for cat, tech, ver in sw_rows:
        row = sw_tbl.add_row()
        for idx, (val, width) in enumerate(zip([cat, tech, ver], sw_cols)):
            c = row.cells[idx]
            c.width = width
            set_cell_margins(c, top=45, bottom=45, left=80, right=80)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(9)
            if idx == 0:
                r.font.bold = True

    p_cap2 = doc.add_paragraph()
    format_paragraph(p_cap2, space_before=4, space_after=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    rc2 = p_cap2.add_run("Table 2.2: Software Stack, Dependencies and Engineering Frameworks")
    rc2.font.name = "Times New Roman"
    rc2.font.size = Pt(9)
    rc2.font.bold = True
    rc2.font.italic = True

    # ==========================================
    # 2.6 SURVEY OF TECHNOLOGY
    # ==========================================
    add_heading_2(doc, "2.6 SURVEY OF TECHNOLOGY AND ALGORITHMIC FOUNDATIONS")

    add_body_p(
        doc,
        "To establish a solid technical foundation, the core algorithms and software paradigms incorporated into RetainIQ are reviewed "
        "with mathematical and computational rigor:"
    )

    add_body_p(
        doc,
        "Survival analysis models the distribution of time until an event occurs (T >= 0). The fundamental quantities of interest are the "
        "Survival Function S(t) = P(T > t) and the Hazard Function h(t). The hazard function h(t) represents the instantaneous rate of "
        "event occurrence given survival up to time t: h(t) = lim_{Delta t -> 0} P(t <= T < t + Delta t | T >= t) / Delta t. "
        "The cumulative hazard H(t) is defined as the integral of h(u) from 0 to t, yielding the exact mathematical identity: "
        "S(t) = exp(-H(t)). Unlike standard classification models that minimize cross-entropy loss over binary labels, survival models "
        "maximize a likelihood function that explicitly partitions the cohort into uncensored event contributions (f(T_i)) and right-censored "
        "survival contributions (S(T_i)).",
        bold_prefix="1. Theoretical Foundations of Survival Analysis: "
    )

    add_body_p(
        doc,
        "The semi-parametric Cox model assumes the hazard for individual i with covariate vector x_i can be factored into a non-parametric "
        "baseline hazard h_0(t) and a parametric exponential covariate multiplier: h(t | x_i) = h_0(t) exp(beta^T x_i). "
        "The principal advantage of Cox PH is that the baseline hazard h_0(t) cancels out when formulating the partial likelihood over the "
        "set of individuals at risk R(t_i) at each event time t_i: L(beta) = prod_{i: d_i=1} [exp(beta^T x_i) / sum_{j in R(t_i)} exp(beta^T x_j)]. "
        "This allows robust estimation of covariate hazard ratios (HR = exp(beta)) without making restrictive distributional assumptions "
        "regarding the underlying survival trajectory.",
        bold_prefix="2. Semi-Parametric Cox Proportional Hazards (Cox PH): "
    )

    add_body_p(
        doc,
        "In contrast to the semi-parametric formulation, Accelerated Failure Time (AFT) models assume that covariates act multiplicatively on "
        "the time scale itself, accelerating or decelerating the rate of aging: ln(T_i) = mu + alpha^T x_i + sigma W_i, where W_i represents an "
        "extreme value error distribution. Under the Weibull distribution, the baseline hazard follows h_0(t) = (gamma / lambda) * (t / lambda)^{gamma - 1}, "
        "where gamma denotes the shape parameter and lambda denotes the scale parameter. When gamma > 1, the hazard rate increases over tenure "
        "(aging effect); when gamma < 1, the hazard rate decreases (burn-in effect). This parametric formulation enables analytical calculation "
        "of restricted mean survival times and continuous quantile projections.",
        bold_prefix="3. Parametric Weibull Accelerated Failure Time (AFT): "
    )

    add_body_p(
        doc,
        "Random Survival Forests extend Breiman's ensemble random forest paradigm to right-censored data (Ishwaran et al., 2008). Unlike "
        "classification trees that split on Gini impurity or variance, survival trees recursively partition feature space by maximizing the "
        "log-rank split statistic between daughter nodes. In each terminal node, a Nelson-Aalen cumulative hazard estimator is computed. The "
        "forest ensemble aggregates these estimates across 100 trees to construct non-linear, non-parametric survival curves that naturally "
        "capture high-order covariate interactions without requiring manual feature engineering.",
        bold_prefix="4. Non-Linear Random Survival Forests (RSF): "
    )

    add_body_p(
        doc,
        "To estimate counterfactual treatment effects under multiple intervention candidates, RetainIQ deploys Meta-Learner architectures "
        "(Künzel et al., 2019). The T-Learner fits separate base survival estimators for each treatment arm: mu_k(x) = E[S_k(t) | X=x] "
        "and mu_0(x) = E[S_0(t) | X=x]. The Conditional Average Treatment Effect (CATE) is then obtained as the counterfactual contrast: "
        "tau_k(x) = mu_k(x) - mu_0(x). In contrast, the S-Learner includes the treatment indicator as an integrated feature within a single "
        "composite model: mu(x, w). By evaluating uplift directly in the probability domain, the system accurately isolates incremental "
        "treatment sensitivity.",
        bold_prefix="5. Causal Uplift Meta-Learners (S-Learner and T-Learner): "
    )

    doc.add_page_break()
    print("Chapter 2 generated successfully.")
