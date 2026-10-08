"""
Chapter 3 Builder for RetainIQ Black Book.
Covers:
3.1 Module Division
3.2 System Architecture (Figure 3.1)
3.3 Gantt Chart (Figure 3.2)
3.4 Data Flow Representation (DFD Level 0 & Level 1) (Figure 3.3)
3.5 UML Diagrams:
    3.5.1 Class Diagram (Figure 3.4)
    3.5.2 Sequence Diagram (Figure 3.5)
    3.5.3 State Chart Diagram (Figure 3.6a)
    3.5.4 Use-Case Diagram (Figure 3.6b)
3.6 Data Model & Configuration Schemas
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_helpers import (
    add_body_p, add_heading_1, add_heading_2, add_heading_3,
    add_figure, format_paragraph
)

def build_chapter_3(doc):
    add_heading_1(doc, "CHAPTER 3: SYSTEM DESIGN", space_before=15, space_after=12)

    add_body_p(
        doc,
        "System design represents the architectural blueprint of RetainIQ, establishing how abstract statistical theories and machine "
        "learning algorithms are translated into a cohesive, decoupled, and production-grade software system. The design follows enterprise "
        "software engineering best practices: strict separation of concerns, modular pipeline organization, immutable data artifact frozen-states, "
        "cryptographic auditability, and standardized REST microservice contracts. This chapter provides a detailed examination of the functional "
        "modules, architectural diagrams, project lifecycle timelines, data flow diagrams, and Unified Modeling Language (UML) structural "
        "and behavioral specifications."
    )

    # ==========================================
    # 3.1 MODULE DIVISION
    # ==========================================
    add_heading_2(doc, "3.1 MODULE DIVISION AND FUNCTIONAL DECOMPOSITION")

    add_body_p(
        doc,
        "RetainIQ is architectured into eight distinct, decoupled functional modules located within the src/, api/, and app/ packages. "
        "Each module possesses dedicated responsibilities, strict interface boundaries, and independent unit test coverage:"
    )

    modules = [
        ("Module 1: Ingestion & Preprocessing (src/data_preprocessing.py): ",
         "Responsible for ingesting raw customer CSV records, verifying cryptographic data integrity hashes, identifying and coercing "
         "whitespace anomalies in TotalCharges to numeric formats, executing tenure-stratified median imputation, performing one-hot categorical "
         "encoding with reference category removal, and writing frozen train/test feature artifacts to disk."),
        
        ("Module 2: Leakage Prevention & Audit (src/leakage_checks.py): ",
         "Enforces strict data hygiene by inspecting feature columns against forbidden future-leakage lists, verifying that data transformers "
         "(scalers and imputers) are fitted exclusively on the training partition, and mathematically auditing censoring consistency."),
        
        ("Module 3: Survival Modeling Engine (src/survival_models.py): ",
         "Encapsulates multi-paradigm survival estimators: Kaplan-Meier non-parametric baselines, semi-parametric Cox Proportional Hazards "
         "fitted via Newton-Raphson partial likelihood optimization, parametric Weibull Accelerated Failure Time models, and Random Survival "
         "Forest ensembles. Provides standardized interfaces for extracting cumulative hazard H(t) and survival curves S(t)."),
        
        ("Module 4: Survival Prediction & Evaluation (src/survival_predictions.py, src/survival_evaluation.py): ",
         "Computes horizon-specific survival probabilities (30, 60, 90, 180 days) for individual subscribers, stratifies customer cohorts "
         "into discrete risk tiers (Low, Medium, High, Critical), and calculates time-dependent Brier scores and Harrell's Concordance Indices."),
        
        ("Module 5: Causal Uplift & Treatment Modeling (src/treatment_effects.py): ",
         "Implements Meta-Learner architectures (T-Learner, S-Learner, and X-Learner) across four candidate retention interventions. Computes "
         "the Conditional Average Treatment Effect (CATE) tau(x) = Delta S(t|x) for every customer-treatment pair and evaluates ranking accuracy "
         "via cumulative Qini curves and PEHE RMSE."),
        
        ("Module 6: Prescriptive Ranking & Budget Knapsack (src/treatment_ranking.py): ",
         "Evaluates net Expected Retention Value (ERV = Delta S * LTV - Cost) for every candidate intervention, implements greedy and knapsack "
         "portfolio budget allocation under aggregate spend caps, and enforces margin-protecting fallback logic ('No Intervention' if ERV <= 0)."),
        
        ("Module 7: Enterprise Governance & Cryptographic Audit (src/audit.py, src/monitoring.py, src/robustness.py): ",
         "Maintains an immutable SHA-256 cryptographic hash-chained audit ledger of all prescriptions and manual overrides. Computes continuous "
         "Population Stability Index (PSI) drift metrics across input features, and executes automated noise/missingness stress testing suites."),
        
        ("Module 8: Dual Delivery Interfaces (api/routes.py, app/streamlit_app.py): ",
         "Exposes low-latency (<15ms) RESTful endpoints for real-time programmatic scoring via FastAPI, and powers an interactive, dark-themed "
         "11-page decision cockpit via Streamlit for business analysts, campaign planners, and executive stakeholders.")
    ]

    for m_title, m_desc in modules:
        add_body_p(doc, m_desc, bold_prefix=m_title, space_before=2, space_after=3)

    # ==========================================
    # 3.2 SYSTEM ARCHITECTURE
    # ==========================================
    add_heading_2(doc, "3.2 SYSTEM ARCHITECTURE")

    add_body_p(
        doc,
        "The overall system architecture of RetainIQ follows a layered, pipelined dataflow pattern designed to decouple feature extraction "
        "and model training from online serving and decision optimization. The architectural layers comprise:"
    )

    layers = [
        ("Data Ingestion & Integrity Layer: ", "Ingests raw enterprise customer telemetry, performs schema validation against Pydantic models, "
         "handles whitespace-missing values, and computes SHA-256 data integrity fingerprints."),
        ("Leakage-Guarded Transformation Layer: ", "Isolates stratified train and test partitions prior to computing transformation parameters, "
         "preventing future target leakage and establishing a certified zero-leakage pipeline."),
        ("Multi-Paradigm Survival Engine: ", "Trains and evaluates Kaplan-Meier, Cox Proportional Hazards, Weibull AFT, and Random Survival "
         "Forest models, emitting calibrated continuous survival curves S_i(t) and hazard percentiles."),
        ("Causal Uplift & Decision Engine: ", "Combines survival predictions with T-Learner meta-models to estimate heterogeneous treatment "
         "uplift Delta S_ik(t), evaluating individual ERV values and executing constrained knapsack budget allocation."),
        ("Governance & Operational Interface Layer: ", "Records tamper-evident audit records via SHA-256 chaining, monitors feature PSI drift, "
         "and delivers real-time inferences through FastAPI endpoints and an 11-page interactive Streamlit dashboard.")
    ]
    for l_title, l_desc in layers:
        add_body_p(doc, l_desc, bold_prefix=l_title, space_before=2, space_after=3)

    add_body_p(
        doc,
        "Figure 3.1 illustrates the complete RetainIQ multi-stage system architecture, detailing the flow of telemetry from ingestion to "
        "prescriptive allocation and cryptographic audit logging:",
        space_before=4, space_after=6
    )

    add_figure(doc, "outputs/figures/architecture_diagram.png", 
               "Figure 3.1: Overall RetainIQ Multi-Stage System Architecture", width_in=5.4)

    add_body_p(
        doc,
        "As depicted in Figure 3.1, raw enterprise customer telemetry enters through the Ingestion & Cleaning Layer, where whitespace "
        "anomalies in TotalCharges are rectified via stratified median imputation. The pipeline splits data into strictly disjoint 80/20 "
        "train and test sets before fitting transformers, guaranteeing zero data leakage. In the Modeling Core, parallel survival estimators "
        "(Cox PH, Weibull AFT, RSF) generate multi-horizon survival curves S(t). These baseline survival probabilities feed directly into the "
        "Causal Uplift Engine, where T-Learner models evaluate the counterfactual survival gain Delta S_ik(t) across candidate treatments. "
        "The Prescriptive Ranking Engine evaluates net ERV, filtering non-positive interventions into zero-cost fallbacks. Finally, all "
        "decisions pass into the Cryptographic Audit Trail and are served to downstream consumers via FastAPI and Streamlit.",
        space_after=8
    )

    # ==========================================
    # 3.3 GANTT CHART
    # ==========================================
    add_heading_2(doc, "3.3 PROJECT DEVELOPMENT GANTT CHART")

    add_body_p(
        doc,
        "The engineering lifecycle of RetainIQ was organized across a 14-week schedule structured into five cohesive phases: "
        "(1) Phase I: Requirements Analysis & Data Pipeline Formulation (Weeks 1–3); (2) Phase II: Exploratory Analysis & Survival Modeling "
        "(Weeks 4–7); (3) Phase III: Causal Uplift & Prescriptive Ranking (Weeks 8–10); (4) Phase IV: Governance, Robustness & Interface "
        "Engineering (Weeks 11–12); and (5) Phase V: Testing, Verification & Black Book Documentation (Weeks 13–14)."
    )

    add_body_p(
        doc,
        "Figure 3.2 illustrates the project timeline and work breakdown structure across all five development phases:",
        space_before=4, space_after=6
    )

    add_figure(doc, "outputs/figures/gantt_chart.png", 
               "Figure 3.2: RetainIQ Engineering Lifecycle and Development Gantt Chart", width_in=5.4)

    add_body_p(
        doc,
        "Figure 3.2 demonstrates the milestone progression of RetainIQ. Key dependencies included completing the zero-leakage data "
        "partitioning pipeline before initiating survival model benchmarking, and freezing trained survival estimators before calibrating "
        "the T-Learner and S-Learner causal uplift models. The final two weeks were dedicated to end-to-end automated pytest execution "
        "(achieving 27/27 passing tests), API load testing, and comprehensive Black Book documentation synthesis.",
        space_after=8
    )

    # ==========================================
    # 3.4 DATA FLOW REPRESENTATION
    # ==========================================
    add_heading_2(doc, "3.4 DATA FLOW REPRESENTATION (DFD LEVEL 0 & LEVEL 1)")

    add_body_p(
        doc,
        "Data Flow Diagrams (DFDs) provide a structured graphical representation of information movement, process transformations, and "
        "data storage repositories within RetainIQ. Figure 3.3 presents both the Level 0 Context Diagram and the Level 1 Functional "
        "Decomposition Diagram:"
    )

    add_figure(doc, "outputs/figures/dfd_level0_level1.png", 
               "Figure 3.3: Data Flow Diagrams: Level 0 Context and Level 1 Functional Decomposition", width_in=5.4)

    add_body_p(
        doc,
        "In the Level 0 Context Diagram (top panel of Figure 3.3), the entire RetainIQ platform is represented as a single centralized "
        "prescriptive system interacting with four primary external entities: (1) Enterprise Data Warehouse (providing raw subscriber "
        "demographics and billing telemetry), (2) Retention Marketing Team (submitting campaign budget constraints and receiving priority "
        "action queues), (3) Customer Relationship Management (CRM) Systems (receiving real-time personalized treatment dispatches), and "
        "(4) Enterprise Governance & Audit Committee (retrieving cryptographic audit proofs and PSI drift reports)."
    )

    add_body_p(
        doc,
        "In the Level 1 Functional Decomposition (bottom panel of Figure 3.3), the system is partitioned into six sequential sub-processes: "
        "(1.0) Data Ingestion & Sanitization, (2.0) Leakage Guard & Stratified Partitioning, (3.0) Survival Modeling & Horizon Scoring, "
        "(4.0) Causal Uplift & Counterfactual Estimation, (5.0) Prescriptive Ranking & Knapsack Optimization, and (6.0) Cryptographic "
        "Audit Logging & Interface Serving. Data stores explicitly isolate Frozen Clean Data (D1), Model Artifacts (D2), Prescriptive "
        "Decisions (D3), and the Tamper-Evident SHA-256 Ledger (D4).",
        space_after=8
    )

    # ==========================================
    # 3.5 UML DIAGRAMS
    # ==========================================
    add_heading_2(doc, "3.5 UNIFIED MODELING LANGUAGE (UML) SPECIFICATIONS")

    add_body_p(
        doc,
        "To provide rigorous object-oriented and behavioral design specifications, four standardized UML diagrams were engineered: "
        "Class Diagram, Sequence Diagram, State Chart Diagram, and Use-Case Diagram."
    )

    # 3.5.1 Class Diagram
    add_heading_3(doc, "3.5.1 UML Class Diagram")
    add_body_p(
        doc,
        "The UML Class Diagram defines the static structure of RetainIQ, establishing class attributes, member methods, inheritance "
        "hierarchies, and composition relationships across the analytical engine. Figure 3.4 illustrates the core domain classes:"
    )

    add_figure(doc, "outputs/figures/uml_class_diagram.png", 
               "Figure 3.4: UML Class Diagram of RetainIQ Domain Models and Pipeline Classes", width_in=5.4)

    add_body_p(
        doc,
        "As depicted in Figure 3.4, the system is anchored by the abstract base class SurvivalModelBase, which defines pure virtual "
        "interfaces for fit(), predict_survival_function(), and predict_hazard(). The concrete implementations—CoxPHModel, WeibullAFTModel, "
        "and RandomSurvivalForestModel—inherit from this interface. The UpliftMetaLearner class encapsulates treatment effect modeling, "
        "invoking underlying survival estimators to compute CATE. The TreatmentRankingEngine aggregates customer features, survival curves, "
        "and candidate treatments, executing knapsack ERV optimization. Finally, the AuditTrail class maintains an immutable chain of "
        "AuditRecord instances, verifying cryptographic block integrity via SHA-256 hashing.",
        space_after=8
    )

    # 3.5.2 Sequence Diagram
    add_heading_3(doc, "3.5.2 UML Sequence Diagram")
    add_body_p(
        doc,
        "The UML Sequence Diagram models the dynamic, chronological message passing between system components during an online customer "
        "scoring and treatment prescription lifecycle. Figure 3.5 details the sequence of interactions:"
    )

    add_figure(doc, "outputs/figures/uml_sequence_diagram.png", 
               "Figure 3.5: UML Sequence Diagram: Asynchronous Customer Scoring and Prescription Lifecycle", width_in=5.4)

    add_body_p(
        doc,
        "As shown in Figure 3.5, the workflow initiates when an API Client (or Streamlit session) dispatches an HTTP POST request to the "
        "/api/v1/score endpoint. The APIRouter deserializes and validates the payload via CustomerInput Pydantic models before passing it to "
        "the PreprocessingPipeline for feature vector scaling. The SurvivalEngine computes baseline survival probabilities S(t) across horizons "
        "(30, 60, 90, 180 days). The UpliftLearner computes counterfactual gains Delta S_k for candidate interventions. The RankingEngine evaluates "
        "ERV_ik, applies budget knapsack constraints, and verifies fallback rules. Once the optimal treatment is selected, an immutable record "
        "is committed to the AuditTrail via SHA-256 chaining before returning a complete JSON prescription payload to the client in <15ms.",
        space_after=8
    )

    # 3.5.3 State Chart & Use Case
    add_heading_3(doc, "3.5.3 UML State Chart and Use-Case Architecture")
    add_body_p(
        doc,
        "Figure 3.6 presents the dual behavioral models of RetainIQ: the Customer Churn Lifecycle State Machine (top panel) and the "
        "Enterprise Use-Case Architecture (bottom panel):"
    )

    add_figure(doc, "outputs/figures/uml_state_and_usecase.png", 
               "Figure 3.6: UML State Machine and Enterprise Use-Case Interaction Architecture", width_in=5.4)

    add_body_p(
        doc,
        "The State Machine Diagram (Figure 3.6, top) models the operational transitions of an account through five distinct states: "
        "(1) Active Safe State (90-day churn risk < 15%), (2) Monitored State (risk 15–35%), (3) Escalated Risk State (risk 35–65%), "
        "(4) Prescribed Intervention State (where the customer is matched with an optimal treatment or assigned to zero-cost fallback), "
        "and (5) Terminal Absorbing States (either 'Successfully Retained' with restored survival trajectory, or 'Voluntary Churn' upon "
        "contract cancellation)."
    )

    add_body_p(
        doc,
        "The Use-Case Model (Figure 3.6, bottom) defines the interactions between four primary enterprise actors and twelve system use cases. "
        "The Retention Strategy Director analyzes portfolio ERV frontiers, sets campaign budget caps, and approves automated allocations. "
        "The Machine Learning Engineer inspects model benchmarking metrics, trains survival estimators, and reviews PSI feature drift. "
        "The Retention Marketing Analyst explores single-customer 360 profiles, runs counterfactual what-if simulations, and submits manual "
        "overrides. Finally, the Compliance Auditor reviews immutable SHA-256 audit ledgers, verifies hash chains, and exports regulatory reports.",
        space_after=8
    )

    # ==========================================
    # 3.6 DATA MODEL & CONFIGURATION
    # ==========================================
    add_heading_2(doc, "3.6 DATA MODEL AND CONFIGURATION SCHEMAS")

    add_body_p(
        doc,
        "To ensure modularity and operational flexibility, RetainIQ externalizes treatment catalogs and operational business rules into "
        "declarative YAML configuration schemas (configs/treatments.yaml). The schema defines unit costs, eligibility criteria, and "
        "expected uplift characteristics for all four candidate interventions:",
        space_after=4
    )

    tx_specs = [
        ("Treatment 1: Basic Retention Check-in: ", "Unit Cost: $5.00 | Eligible: All active accounts | Description: Proactive outbound "
         "customer success call to resolve friction points, confirm plan suitability, and offer courtesy service checkups."),
        ("Treatment 2: Proactive Tech Support & Router Upgrade: ", "Unit Cost: $12.00 | Eligible: Accounts with Internet Service (DSL or Fiber) | "
         "Description: Priority tech support assignment, complimentary router firmware upgrade, and dedicated line optimization."),
        ("Treatment 3: 15% Bill Discount (3-Month Loyalty): ", "Unit Cost: $18.50 | Eligible: Month-to-month contracts and high-risk accounts | "
         "Description: Direct financial concession reducing monthly billing by 15% for three consecutive billing cycles."),
        ("Treatment 4: Long-Term Contract Incentive: ", "Unit Cost: $35.00 | Eligible: Month-to-month subscribers with tenure >= 3 months | "
         "Description: Hardware upgrade credit and bill concession contingent on signing a 12-month or 24-month contract lock-in.")
    ]
    for ts_title, ts_desc in tx_specs:
        add_body_p(doc, ts_desc, bold_prefix=ts_title, space_before=2, space_after=2)

    doc.add_page_break()
    print("Chapter 3 generated successfully.")
