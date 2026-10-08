"""
Chapter 1 Builder for RetainIQ Black Book.
Covers:
1.1 Significance
1.2 Objectives
1.3 Purpose and Scope
    1.3.1 Purpose
    1.3.2 Scope
1.4 Applicability
1.5 Achievements
"""

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_helpers import add_body_p, add_heading_1, add_heading_2, add_heading_3, format_paragraph

def build_chapter_1(doc):
    add_heading_1(doc, "CHAPTER 1: INTRODUCTION", space_before=15, space_after=12)

    add_body_p(
        doc,
        "Customer churn—the voluntary or involuntary cessation of a contractual or recurring relationship between a subscriber "
        "and a service provider—represents an existential commercial challenge across modern subscription economies. In capital-intensive "
        "industries such as telecommunications, cloud Software-as-a-Service (SaaS), digital streaming media, and retail banking, the "
        "financial dynamics of customer acquisition versus customer retention have diverged dramatically. Contemporary market studies "
        "consistently demonstrate that acquiring a new enterprise or retail subscriber requires between five to seven times the capital "
        "expenditure required to retain an existing customer. Moreover, retained subscribers exhibit compounding lifetime profitability: "
        "tenured cohorts consume less frontline customer support, demonstrate higher openness to cross-selling and tier upgrades, "
        "and serve as organic advocates within their peer ecosystems. Consequently, a marginal reduction in annual attrition (even by 1 "
        "to 2 percentage points) translates directly into tens of millions of dollars in preserved Enterprise Value."
    )

    add_body_p(
        doc,
        "Despite the paramount strategic importance of customer preservation, contemporary enterprise churn management remains crippled "
        "by foundational methodological limitations. Standard industry practice treats churn prediction almost exclusively as a static, "
        "binary classification problem. In such setups, supervised learning algorithms—ranging from classic Logistic Regression to "
        "modern gradient-boosted decision trees (XGBoost, LightGBM, CatBoost)—are trained on historical customer feature snapshots to "
        "predict a single binary label: whether a subscriber will cancel their subscription within an arbitrary observation window (typically "
        "30, 60, or 90 days). While computationally straightforward to implement, this formulation introduces severe theoretical distortions "
        "that undermine both predictive accuracy and commercial decision-making."
    )

    add_body_p(
        doc,
        "The first fundamental flaw is survival blindness: binary classification collapses the entire temporal continuum into a single "
        "dichotomous outcome. A customer who cancels their account on Day 3 post-onboarding is treated as functionally identical to a highly "
        "profitable, multi-year subscriber who leaves on Day 89 of a 90-day observation window. This loss of temporal granularity prevents "
        "retention teams from understanding whether attrition represents an immediate onboarding friction, a contractual expiration shock, "
        "or a gradual long-term erosion of product engagement."
    )

    add_body_p(
        doc,
        "The second and more profound statistical failure is the mishandling of right censoring. In any operational customer database, "
        "the vast majority of active subscribers have not yet churned as of the observation cut-off date. In standard survival statistics, "
        "these observations are 'right-censored'—their true churn event time T is known only to exceed the current elapsed tenure t (i.e., T > t). "
        "Binary classification models have no native mathematical mechanism to process censored observations. Consequently, practitioners are "
        "forced either to drop active subscribers from historical training splits or to arbitrarily assign them a 'non-churn' label of zero. "
        "Dropping active accounts induces catastrophic selection bias, while labeling them as zero fundamentally misrepresents the data "
        "generating process, penalizing long-lived accounts and producing severe downward survival bias."
    )

    add_body_p(
        doc,
        "Most critically, existing commercial workflows suffer from an association-versus-causation fallacy. Marketing operations routinely "
        "treat the output of a churn classifier—the predicted probability of attrition—as a direct ranking of customer intervention "
        "priority. Retention marketing teams blast aggressive promotional discounts and retention credits to the top decile of highest-risk "
        "accounts. This strategy represents a massive misallocation of corporate capital. Groundbreaking work in causal inference and uplift "
        "modeling reveals that customers with the highest baseline churn risk are frequently 'Lost Causes'—individuals whose decision to "
        "depart is driven by exogenous factors (e.g., relocation, business bankruptcy, fundamental price incompatibility) that cannot be altered "
        "by promotional offers. Conversely, allocating discounts to highly tenured subscribers who would have renewed organically ('Sure Things') "
        "merely erodes recurring margins without generating incremental retention value. Worse still, intrusive outbound retention calls "
        "frequently disturb dormant subscribers ('Sleeping Dogs'), triggering unforced cancellations that would not have occurred otherwise."
    )

    add_body_p(
        doc,
        "To fundamentally resolve these structural shortcomings, this capstone project develops RetainIQ (Project ID: BDS-11): an "
        "integrated, production-grade enterprise decision-support ecosystem. RetainIQ replaces point-in-time binary classification with "
        "continuous-time statistical survival analysis, and replaces naive risk-based targeting with heterogeneous causal uplift modeling "
        "and constrained Expected Retention Value (ERV) optimization. By unifying statistical survival theory, modern causal meta-learners, "
        "tamper-evident audit logging, and low-latency API architecture, RetainIQ bridges the gap between machine learning research and "
        "prescriptive business economics."
    )

    # ==========================================
    # 1.1 SIGNIFICANCE
    # ==========================================
    add_heading_2(doc, "1.1 SIGNIFICANCE OF THE PROJECT")

    add_body_p(
        doc,
        "The technical and commercial significance of RetainIQ spans multiple dimensions of modern data science, statistical modeling, "
        "and strategic enterprise decision support:"
    )

    add_body_p(
        doc,
        "From an applied machine learning perspective, survival analysis provides a mathematically principled framework for time-to-event "
        "phenomena under right censoring. By modeling both the hazard rate h(t) and the survival probability function S(t) = P(T > t), "
        "RetainIQ allows data science teams to extract continuous risk trajectories across arbitrary operational horizons (e.g., 30, 60, "
        "90, 180, and 365 days) from a single unified model. This eliminates the technical debt and maintenance burden of training separate "
        "binary classifiers for every operational timeframe.",
        bold_prefix="1. Advanced Methodological Synthesis: "
    )

    add_body_p(
        doc,
        "Standard machine learning focuses exclusively on observational prediction (E[Y | X]), answering the question: 'What is the "
        "likelihood this customer churns?' In contrast, retention operations require counterfactual decision support: 'How will this "
        "customer's survival probability change if we administer treatment k versus no intervention?' By deploying causal Meta-Learners "
        "(T-Learners and S-Learners) to estimate Conditional Average Treatment Effects (CATE), RetainIQ shifts the paradigm from passive "
        "risk scoring to active causal prescription.",
        bold_prefix="2. Bridging Prediction and Causal Prescription: "
    )

    add_body_p(
        doc,
        "Enterprise retention marketing budgets are inherently finite. By formulating the decision problem as a knapsack-style Expected "
        "Retention Value (ERV) optimization, RetainIQ ensures that every marketing dollar spent produces measurable marginal return on "
        "investment (ROI). The system explicitly evaluates intervention unit costs against customer Lifetime Value (LTV) and counterfactual "
        "survival gains, enforcing an automated economic fallback to 'No Intervention' when the cost of treatment exceeds expected value.",
        bold_prefix="3. Capital Efficiency and Margin Protection: "
    )

    add_body_p(
        doc,
        "In enterprise environments governed by stringent compliance standards (such as GDPR, CCPA, and internal risk oversight), "
        "algorithmic recommendations cannot operate as unmonitored black boxes. RetainIQ establishes an immutable cryptographic "
        "SHA-256 audit ledger tracking every algorithmic recommendation and human override. Coupled with continuous Population Stability "
        "Index (PSI) tracking and multi-tier stress testing, the system sets a new standard for responsible, auditable AI governance.",
        bold_prefix="4. Enterprise Governance and Regulatory Auditability: "
    )

    # ==========================================
    # 1.2 OBJECTIVES
    # ==========================================
    add_heading_2(doc, "1.2 PROJECT OBJECTIVES")

    add_body_p(
        doc,
        "The overarching mission of RetainIQ is to engineer, validate, and deploy a complete prescriptive churn retention platform. "
        "The project is structured around six concrete, verifiable technical objectives:"
    )

    objectives = [
        ("Objective 1: Ingestion & Leakage-Free Preprocessing Pipeline. ",
         "Design and implement an end-to-end data processing pipeline for authentic telecommunications telemetry (IBM Telco cohort of 7,043 "
         "subscribers). The pipeline must handle whitespace missingness, enforce strict schema validation, perform one-hot categorical "
         "encoding with baseline drop, and isolate train/test splits prior to transformation to achieve a verified 0% data leakage status."),
        
        ("Objective 2: Multi-Paradigm Survival Modeling & Benchmarking. ",
         "Implement, tune, and systematically benchmark four foundational survival modeling paradigms: non-parametric Kaplan-Meier estimation, "
         "semi-parametric Cox Proportional Hazards (M1), parametric Weibull Accelerated Failure Time (M2), and non-linear Random Survival "
         "Forests (M3). Evaluate performance using Harrell's Concordance Index (C-Index) and Integrated Brier Score (IBS) across 30, 60, "
         "90, and 180-day operational horizons."),
        
        ("Objective 3: Heterogeneous Causal Treatment Effect (CATE) Estimation. ",
         "Formulate and deploy causal Meta-Learner architectures (S-Learner, T-Learner, and X-Learner) to estimate the counterfactual "
         "survival uplift tau_k(x) = Delta S(t|x) attributable to four candidate retention interventions: Basic Check-in ($5.00), "
         "Proactive Tech Support ($12.00), 15% Bill Discount ($18.50), and Long-Term Contract Incentive ($35.00). Validate ranking efficacy "
         "using cumulative Qini curves and PEHE RMSE."),
        
        ("Objective 4: Prescriptive Ranking & ERV Economic Engine. ",
         "Develop a personalized ranking engine that evaluates net Expected Retention Value (ERV = Delta S * LTV - Cost) for every customer-"
         "treatment pair. Implement constrained portfolio knapsack optimization that maximizes aggregate retained revenue under strict "
         "marketing budget caps while enforcing automated zero-cost fallbacks for non-responsive accounts."),
        
        ("Objective 5: Enterprise Governance, Cryptographic Audit & Stress Testing. ",
         "Engineer a tamper-evident audit trail using SHA-256 cryptographic hash chaining to record algorithmic decisions and manual "
         "analyst overrides. Implement automated Population Stability Index (PSI) drift monitoring and adversarial stress testing suites "
         "evaluating model resilience under feature noise and data missingness up to 30%."),
        
        ("Objective 6: Dual Production Interfaces (FastAPI & Streamlit). ",
         "Deliver a production-ready dual-interface system comprising: (a) a high-throughput, low-latency RESTful microservice built with "
         "FastAPI, supporting Pydantic-validated real-time scoring endpoints (<15ms latency), and (b) an interactive, 11-page decision cockpit "
         "built with Streamlit providing executive portfolio summaries, customer 360 exploration, and what-if counterfactual simulators.")
    ]

    for obj_title, obj_desc in objectives:
        add_body_p(doc, obj_desc, bold_prefix=obj_title, space_before=3, space_after=4)

    # ==========================================
    # 1.3 PURPOSE AND SCOPE
    # ==========================================
    add_heading_2(doc, "1.3 PURPOSE AND SCOPE")

    add_heading_3(doc, "1.3.1 Purpose")
    add_body_p(
        doc,
        "The primary purpose of RetainIQ is to replace intuitive, ad-hoc, and economically destructive retention targeting practices with a "
        "mathematically rigorous, causally sound, and capital-efficient decision system. By transforming raw customer telemetry into calibrated "
        "survival probability curves and counterfactual uplift surfaces, RetainIQ empowers retention directors, marketing strategists, and "
        "customer success teams to answer three critical operational questions with quantitative certainty: (1) Exactly when is a specific "
        "subscriber at risk of departing? (2) Which specific intervention will maximize that subscriber's incremental survival duration? and "
        "(3) Does the expected financial gain from saving this customer exceed the cost of the prescribed intervention?"
    )

    add_heading_3(doc, "1.3.2 Scope and Boundary Conditions")
    add_body_p(
        doc,
        "To ensure academic rigor and engineering excellence, the operational scope and boundary conditions of this project are strictly defined:",
        space_after=4
    )

    in_scope = [
        ("Survival Analysis: ", "Complete non-parametric, semi-parametric, and parametric survival curve estimation, hazard ratio attribution, "
         "time-dependent Brier score calibration, and Harrell's C-index evaluation across 7,043 telecommunications accounts."),
        ("Causal Uplift Modeling: ", "Implementation of S-Learner, T-Learner, and X-Learner meta-learning frameworks estimating individual-level "
         "heterogeneous treatment effects across four predefined operational retention interventions."),
        ("Prescriptive Economic Optimization: ", "Individual-level ERV ranking, portfolio knapsack budget optimization, margin-protecting "
         "fallback logic, and single-customer counterfactual what-if simulation."),
        ("Enterprise Governance: ", "Cryptographic SHA-256 audit logging, automated tamper verification, PSI drift monitoring, and adversarial "
         "missingness/noise stress testing."),
        ("Software Engineering & Delivery: ", "Full implementation of a 27-test automated pytest verification suite, FastAPI REST microservice, "
         "and an 11-page interactive Streamlit dashboard.")
    ]
    for sc_title, sc_desc in in_scope:
        add_body_p(doc, sc_desc, bold_prefix=sc_title, space_before=2, space_after=3)

    p_out = doc.add_paragraph()
    format_paragraph(p_out, space_before=4, space_after=4)
    r_out = p_out.add_run("Explicit Out-of-Scope Boundaries: ")
    r_out.font.name = "Times New Roman"
    r_out.font.size = Pt(11)
    r_out.font.bold = True
    r_outt = p_out.add_run(
        "This project does not encompass physical telecommunications hardware telemetry (e.g., real-time cell tower packet routing or optical "
        "fiber signal attenuation), online continuous reinforcement learning under non-stationary real-world streaming data, or outbound physical "
        "telephony call-center dialing automation. In addition, in accordance with the project instructions, treatment effects are estimated and "
        "benchmarked on observational and simulated intervention telemetry rather than an active, multi-million-dollar randomized clinical trial."
    )
    r_outt.font.name = "Times New Roman"
    r_outt.font.size = Pt(11)

    # ==========================================
    # 1.4 APPLICABILITY
    # ==========================================
    add_heading_2(doc, "1.4 ENTERPRISE APPLICABILITY")

    add_body_p(
        doc,
        "While demonstrated on an authentic telecommunications cohort, the underlying mathematical and architectural principles of RetainIQ "
        "are universally applicable across diverse recurring-revenue and contractual business models:"
    )

    domains = [
        ("Telecommunications and Internet Service Providers (ISPs): ",
         "Optimizing contract renewal incentives, equipment upgrades (e.g., fiber router migrations), and bundle discounts for high-value "
         "subscribers facing fiber-optic competitive pressures."),
        ("Enterprise Cloud Software-as-a-Service (SaaS): ",
         "Preserving recurring software licenses by tracking seat utilization and feature engagement decay, deploying proactive customer "
         "success training interventions prior to annual master service agreement (MSA) renewals."),
        ("Direct-to-Consumer (D2C) Subscription Media: ",
         "Preventing subscriber attrition on platforms such as Netflix, Spotify, or digital news publications by deploying personalized "
         "content recommendations or temporary billing pauses to disengaged accounts without cannibalizing full-rate renewals."),
        ("Retail Banking and Credit Card Services: ",
         "Predicting account dormancy and card inactivation, selectively allocating fee waivers, bonus rewards points, or dedicated "
         "relationship management check-ins to high-LTV cardholders."),
        ("Insurance Policy Underwriting and Auto-Renewal: ",
         "Forecasting policy lapse hazards at annual renewal cycles, prescribing tailored premium adjustments or deductible restructuring "
         "to retain safe, low-loss-ratio policyholders.")
    ]
    for d_title, d_desc in domains:
        add_body_p(doc, d_desc, bold_prefix=d_title, space_before=2, space_after=3)

    # ==========================================
    # 1.5 ACHIEVEMENTS
    # ==========================================
    add_heading_2(doc, "1.5 TECHNICAL AND EMPIRICAL ACHIEVEMENTS")

    add_body_p(
        doc,
        "The RetainIQ project has successfully accomplished all design goals and delivered verifiable, production-grade outputs:",
        space_after=4
    )

    achievements = [
        ("100% Verified Zero-Leakage Pipeline: ",
         "Constructed a clean, reproducible preprocessing pipeline across 7,043 customer accounts, resolving 11 whitespace-missing values "
         "in TotalCharges through median imputation, and establishing strict temporal isolation with an automated leakage test suite."),
        ("State-of-the-Art Survival Discrimination: ",
         "Achieved a test Concordance Index (C-Index) of 0.8697 with the Weibull AFT model and 0.8694 with Cox Proportional Hazards, "
         "representing a +4.14 percentage-point improvement over the standard naive binary logistic classifier baseline (0.8283)."),
        ("Superior Probabilistic Calibration: ",
         "Delivered an exceptional Integrated Brier Score (IBS) of 0.0818 at the 90-day operational horizon, outperforming the baseline (0.1336) "
         "by 38.8%, proving reliable, unskewed risk quantification."),
        ("High-Performance Uplift Ranking: ",
         "Trained and evaluated causal meta-learners, with the T-Learner achieving a mean Qini score of 0.6920 across four treatments, outperforming "
         "the S-Learner (0.4046) and enabling high-precision identification of responsive subscribers."),
        ("Demonstrated Portfolio ERV Optimization: ",
         "Engineered an optimization engine that allocated 1,248 accounts from a 1,873-candidate pool, maximizing Net Expected Retention Value "
         "at $48,920 under a $15,000 budget cap, delivering a verified 329.2% campaign ROI."),
        ("Comprehensive Automated Test Coverage: ",
         "Executed an automated test suite comprising 27 comprehensive pytest test cases covering data preprocessing, leakage guards, survival "
         "monotonicity, uplift evaluation, API contracts, and cryptographic tampering, achieving a 100% PASS rate."),
        ("Sub-15ms Enterprise API Microservice: ",
         "Engineered a production FastAPI microservice featuring Pydantic request/response schema validation, CORS security, health probes, "
         "and real-time survival scoring with an average response latency of under 15 milliseconds."),
        ("Full-Featured 11-Page Interactive Dashboard: ",
         "Developed an intuitive, dark-mode Streamlit decision console providing executive KPI overviews, customer 360 profiling, what-if "
         "treatment simulation, campaign budgeting, model explainability, and cryptographic audit ledger inspection.")
    ]

    for a_title, a_desc in achievements:
        add_body_p(doc, a_desc, bold_prefix=a_title, space_before=2, space_after=3)

    doc.add_page_break()
    print("Chapter 1 generated successfully.")
