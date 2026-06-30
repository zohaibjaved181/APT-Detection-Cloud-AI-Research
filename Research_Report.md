# Research Report: Detection of Advanced Persistent Threats in Cloud Networks Using Artificial Intelligence

---

## 1. Introduction

### 1.1 Background

Cloud computing has fundamentally reshaped how organisations build and deliver digital services. From small startups to multinational enterprises, the migration to cloud infrastructure has accelerated over the past decade, driven by promises of scalability, cost efficiency, and operational flexibility. However, this shift has also introduced a new and complex threat surface that traditional perimeter-based security tools were never designed to handle.

Among the most dangerous threats facing cloud environments today are Advanced Persistent Threats (APTs). Unlike opportunistic attacks that aim for quick wins, APTs are carefully planned, well-funded campaigns — often state-sponsored or backed by organised criminal groups — that infiltrate target networks and maintain a covert presence for weeks, months, or even years. Their goal is typically long-term espionage, intellectual property theft, or strategic disruption of critical infrastructure.

What makes APTs particularly challenging is their multi-stage nature. A typical APT campaign progresses through reconnaissance, initial compromise, lateral movement, privilege escalation, and eventual data exfiltration — all while actively avoiding detection. Each stage may use different techniques and generate different network signatures, making it extremely difficult for rule-based systems to piece together the full picture.

In cloud environments, these challenges are compounded by factors such as multi-tenancy, ephemeral workloads, encrypted east-west traffic, dynamic scaling, and the shared responsibility model between providers and tenants. Traditional Intrusion Detection Systems (IDS) that rely on known attack signatures or static threshold rules struggle to keep pace with adversaries who continuously adapt their tactics.

### 1.2 Problem Statement

The core problem this research addresses is straightforward: existing security tools cannot reliably detect sophisticated APT campaigns in cloud network environments. Signature-based systems fail against zero-day exploits and custom malware. Anomaly detection systems produce excessive false positives that overwhelm security operations teams. And most critically, very few approaches have been designed specifically for the unique characteristics of cloud infrastructure.

Recent literature reviews confirm this gap. Studies published between 2022 and 2025 consistently highlight the lack of cloud-native detection frameworks, the absence of realistic cloud APT datasets for training and evaluation, and the limited real-time processing capability of deep learning approaches when deployed at cloud scale.

### 1.3 Research Aim

The primary aim of this research is to develop and evaluate an artificial intelligence-driven detection framework capable of identifying multi-stage Advanced Persistent Threat attacks within cloud network environments. The framework should operate in near real-time, provide explainable outputs that security analysts can act upon, and demonstrate measurable improvement over existing detection approaches.

### 1.4 Research Objectives

1. **Design a hybrid deep learning architecture** that combines spatial feature extraction (via convolutional layers) with temporal sequence modelling (via recurrent or attention-based layers) for cloud network traffic analysis.

2. **Benchmark the proposed system** against established detection methods using standardised datasets, providing a fair and transparent comparison of detection accuracy, false positive rates, and computational efficiency.

3. **Evaluate real-time detection feasibility** by measuring inference latency under realistic cloud traffic volumes and assessing scalability across distributed deployment scenarios.

4. **Incorporate explainability mechanisms** (such as SHAP or attention visualisation) that allow security operations teams to understand why a particular traffic pattern was flagged as suspicious.

5. **Address the cloud-specific data gap** by generating a simulated dataset using MITRE ATT&CK-mapped attack scenarios executed against a cloud testbed environment.

### 1.5 Significance of the Study

This research contributes to both the academic knowledge base and practical industry needs in several ways:

- **For the research community**, it provides a validated framework that addresses documented gaps in cloud-specific APT detection, along with a reproducible methodology that others can build upon.
- **For security practitioners**, it offers a practical system capable of detecting threats that traditional tools miss, with explanations that support human decision-making rather than replacing it.
- **For the broader field**, it contributes a cloud-native APT dataset generated through controlled simulation, addressing a widely acknowledged shortage of appropriate training data.

The study is timely given the accelerating migration to cloud infrastructure and the growing sophistication of threat actors targeting these environments. With APT attacks on cloud assets increasing by over 150% between 2022 and 2024, the need for intelligent, adaptive detection systems has never been more pressing.

---

## 2. Research Philosophy and Approach

### 2.1 Understanding Research Philosophies

Every research project is built upon philosophical assumptions about the nature of knowledge and how it can be acquired. These assumptions — whether made explicit or not — shape the entire research design, from the questions we ask to the methods we use to answer them. In this section, we critically evaluate four major research philosophies and explain why one of them provides the most coherent foundation for our study.

### 2.2 Positivism

Positivism holds that reliable knowledge comes exclusively from observable, measurable phenomena. It values objectivity, detachment, and the scientific method. Under this philosophy, the researcher remains separate from the subject of study, formulates testable hypotheses, collects empirical data, and uses statistical tools to draw conclusions.

**Relevance to our research:** Positivism aligns naturally with computational research. Our AI detection system produces quantifiable outputs — accuracy scores, processing times, false positive rates — that can be measured, compared, and statistically validated. The entire experimental process follows the classical scientific model: we formulate a hypothesis (our model outperforms baselines), design controlled experiments, collect numerical results, and apply statistical tests.

**Strengths:** Enables reproducible experiments, supports statistical generalisation, and provides clear criteria for evaluating success or failure.

**Limitations:** Cannot capture the strategic reasoning behind attacker behaviour, the organisational context in which detection systems operate, or the qualitative aspects of how security teams interact with AI-generated alerts.

### 2.3 Interpretivism

Interpretivism argues that knowledge is socially constructed and that understanding comes from exploring subjective human experiences, meanings, and contexts. Research under this paradigm typically employs qualitative methods such as interviews, ethnography, and phenomenological analysis.

**Relevance to our research:** Interpretivism could be valuable if our study focused on how security analysts make sense of threat information, or how organisational culture shapes incident response practices. These are legitimate and important questions — but they are not the questions we are asking.

**Strengths:** Provides deep contextual understanding of human behaviour and decision-making in security contexts.

**Limitations:** Entirely unsuitable for evaluating the technical performance of a computational system. Our research question — whether an AI model can detect attacks more effectively than existing tools — demands numerical evidence, not subjective interpretation.

### 2.4 Pragmatism

Pragmatism takes the position that the value of any approach depends on its practical usefulness in solving the problem at hand. It supports mixing methods and paradigms, arguing that researchers should use whatever works best for their specific research questions rather than adhering rigidly to a single philosophical stance.

**Relevance to our research:** A pragmatist approach would allow us to combine quantitative experiments with qualitative elements — for example, gathering expert feedback on the usability of detection alerts alongside numerical performance data.

**Strengths:** Offers flexibility and a problem-centred orientation. Could provide a richer overall picture by incorporating multiple perspectives.

**Limitations:** For our specific study, the addition of qualitative components would increase complexity and resource requirements without proportionate benefit. Our primary research questions are entirely answerable through computational experiments and statistical analysis. While expert feedback would be valuable in future work, it is not essential for establishing the core technical contribution.

### 2.5 Realism

Realism maintains that an objective reality exists independently of human observation, but acknowledges that our access to that reality is always imperfect and mediated by theory. Critical realism, in particular, distinguishes between observable events and the deeper causal mechanisms that generate them.

**Relevance to our research:** Realism recognises that APT behaviour involves underlying causal structures that may not be fully visible in network traffic alone. It supports investigating why certain detection methods succeed or fail, not just whether they do.

**Strengths:** Provides a sophisticated ontological position that acknowledges complexity and multiple layers of reality.

**Limitations:** More naturally suited to explanatory social research than to computational systems benchmarking. While intellectually appealing, it is harder to operationalise cleanly in a study primarily focused on measuring and comparing model performance.

### 2.6 Selected Philosophy: Positivism

After careful consideration, we adopt **positivism** as the guiding research philosophy for this study. The justification is grounded in the nature of our research problem:

- Our research produces **measurable, numerical outputs** that can be objectively evaluated.
- We follow a **hypothesis-testing approach** where specific predictions about model performance are tested against empirical evidence.
- The experimental design employs **controlled conditions** with clearly defined independent and dependent variables.
- Results are assessed using **statistical significance tests** that provide objective criteria for accepting or rejecting claims.
- The study aims for **generalisability** — findings that hold across different datasets and cloud environments, not just the specific conditions under which they were produced.

Positivism provides the most coherent philosophical foundation because it directly supports the kind of evidence our research generates and the standards by which that evidence will be judged.

### 2.7 Research Approach: Deductive

Consistent with our positivist stance, we adopt a **deductive research approach** — moving from established theory to specific hypotheses to empirical testing:

1. **Theory:** Existing research demonstrates that deep learning architectures can model complex, non-linear patterns in network traffic data.
2. **Hypothesis:** Our proposed hybrid model will achieve significantly higher detection accuracy than baseline methods on both benchmark and cloud-specific datasets.
3. **Experimentation:** We train the model, evaluate it under controlled conditions, and collect performance metrics.
4. **Validation:** Statistical tests determine whether observed improvements are statistically significant or attributable to chance.

The deductive approach is appropriate because we are not exploring an entirely new phenomenon — substantial prior work on ML-based threat detection provides a strong theoretical foundation. Our contribution is to extend and improve upon existing approaches for a specific, well-defined context (cloud APT detection), which is precisely the kind of work that benefits from structured hypothesis testing.

---

## 3. Research Methodology

### 3.1 Methodology Selection: Quantitative

Given our positivist philosophy and deductive approach, we adopt a **quantitative research methodology**. This decision follows logically from the nature of our research questions and the type of evidence needed to answer them.

**Why quantitative?**

Our study asks whether an AI system can detect Advanced Persistent Threats in cloud networks more effectively than existing methods. Answering this question requires:

- Training machine learning models on large-scale, structured datasets containing millions of network flow records.
- Measuring system performance through established metrics: accuracy, precision, recall, F1-score, AUC-ROC, and inference latency.
- Comparing multiple approaches under identical conditions using standardised experimental protocols.
- Applying statistical hypothesis tests (t-tests, ANOVA, Wilcoxon signed-rank) to determine whether observed differences are meaningful or due to random variation.

Every element of our research produces numerical data that demands quantitative analysis. There is no subjective interpretation involved — a model either detects an attack or it does not, and its performance is captured entirely through measurable quantities.

**Why not qualitative?**

Qualitative methodology focuses on understanding human experiences, meanings, and social constructions through methods like interviews, focus groups, and thematic analysis. While valuable for exploring how security analysts perceive and respond to threats, it cannot evaluate the technical performance of a computational system. Our research question is fundamentally about model accuracy, not human perception.

**Why not mixed methods?**

A mixed-methods approach would combine quantitative experiments with qualitative components — for instance, interviews with security professionals about the usability of detection outputs. While this could enrich a follow-up study, it would add substantial complexity to the current research without addressing the core technical questions. Our primary contribution is the detection framework itself, and establishing its effectiveness requires a focused quantitative evaluation.

### 3.2 Research Methods

Within our quantitative methodology, we employ two primary data collection methods:

#### 3.2.1 Secondary Data Analysis (Benchmark Datasets)

We utilise publicly available, peer-reviewed network intrusion detection datasets as our primary training and benchmarking resource. These include:

- **CICIDS 2017/2018** — Contains labelled benign and attack traffic across multiple protocols, widely used for IDS evaluation.
- **UNSW-NB15** — Provides a comprehensive set of modern attack types with detailed feature extraction.
- **Unraveled APT Dataset** — Specifically designed to capture long-term, stealthy APT behaviours often missed by conventional datasets.
- **DARPA Transparent Computing** — System-level provenance data capturing multi-host attack campaigns.

**Justification:** These datasets have been validated by hundreds of published studies, providing established baselines for fair comparison. Their pre-labelled nature eliminates annotation bias, their massive scale supports deep learning training requirements, and their public availability ensures full reproducibility.

#### 3.2.2 Controlled Experiment (Simulated Cloud Environment)

We construct a cloud-based testbed environment (using AWS and Azure infrastructure) and execute realistic APT attack campaigns against it using established adversary simulation tools:

- **MITRE ATT&CK Framework** — Maps all attack actions to standardised tactics, techniques, and procedures.
- **Atomic Red Team** — Provides modular, automated execution of individual ATT&CK techniques.
- **Custom multi-stage scenarios** — Designed to replicate full APT campaigns including reconnaissance, initial access, lateral movement, and data exfiltration.

**Justification:** This method addresses the well-documented gap in cloud-specific APT data. Existing benchmark datasets primarily capture traditional network attacks and lack the API abuse patterns, container escape attempts, and serverless manipulation characteristic of cloud-native threats. Our simulation produces ground-truth-labelled data under controlled conditions, with full variable control over attack timing, intensity, and technique selection.

### 3.3 Data Analysis Approach

All collected data is analysed using quantitative statistical methods:

- **Performance metrics:** Accuracy, Precision, Recall, F1-Score, AUC-ROC, and inference latency for each model configuration.
- **Cross-validation:** k-fold cross-validation (k=5 and k=10) ensures results are robust and not dependent on a particular train/test split.
- **Statistical hypothesis testing:** Paired t-tests and Wilcoxon signed-rank tests compare our proposed model against each baseline, with significance thresholds of p < 0.05.
- **Confusion matrix analysis:** Provides detailed breakdown of true positives, false positives, true negatives, and false negatives across attack categories.
- **Ablation studies:** Systematically remove components of the proposed architecture to understand each element's contribution.

### 3.4 Appropriateness and Coherence

The research design achieves internal coherence through the logical alignment of its components:

| Component | Choice | Rationale |
|-----------|--------|-----------|
| Philosophy | Positivism | Research produces measurable, objective outputs |
| Approach | Deductive | Testing specific hypotheses derived from existing theory |
| Methodology | Quantitative | Numerical data requiring statistical analysis |
| Methods | Benchmarks + Experiment | Breadth (fair comparison) + Depth (cloud validation) |
| Analysis | Statistical + ML metrics | Matches the numerical nature of collected data |

Each element supports and reinforces the others. We ask a quantitative question, collect quantitative data through two complementary channels, and analyse it using statistical tools that provide clear, defensible answers. The result is a research design that is both rigorous enough for academic publication and practical enough to produce systems that can be deployed in real cloud environments.

---

## References

1. Carnegie Mellon SEI (2024). *Toward the Use of Artificial Intelligence for Advanced Persistent Threat Detection*. Available at: https://sei.cmu.edu/library/toward-the-use-of-artificial-intelligence-ai-for-advanced-persistent-threat-detection/

2. Arxiv (2025). *APT-LLM: Embedding-Based Anomaly Detection of Cyber Advanced Persistent Threats Using Large Language Models*. Available at: https://arxiv.org/html/2502.09385v1

3. Arxiv (2024). *RAPID: Robust APT Detection and Investigation Using Context-Aware Deep Learning*. Available at: https://arxiv.org/abs/2406.05362

4. Nature Scientific Reports (2024). *A Novel Approach for APT Attack Detection Based on Advanced Computing*. Available at: https://www.nature.com/articles/s41598-024-72957-0

5. Springer (2024). *Explainable Deep Learning Approach for Advanced Persistent Threats Detection in Cybersecurity: A Review*. Available at: https://link.springer.com/article/10.1007/s10462-024-10890-4

6. Oxford Academic (2024). *Systematic Literature Review on Advanced Persistent Threat Behaviors and its Detection Strategy*. Available at: https://academic.oup.com/cybersecurity/article/10/1/tyad023/7504935

7. ACM Computing Surveys (2024). *Advanced Persistent Threat Attack Detection Systems: A Review of Approaches, Challenges, and Trends*. Available at: https://dl.acm.org/doi/10.1145/3696014

8. Journal of Big Data (2025). *An Optimized Hybrid Ensemble Machine Learning Model for Detecting Advanced Persistent Threats in Networks*. Available at: https://journalofbigdata.springeropen.com/articles/10.1186/s40537-025-01272-w

9. Springer (2026). *Collaborative APT Detection through Cloud-Edge Synergy of Graph Transformer and Lightweight Clustering*. Available at: https://link.springer.com/article/10.1186/s13677-026-00940-3

10. Journal of Big Data (2024). *Advancing Cybersecurity: A Comprehensive Review of AI-Driven Detection Techniques*. Available at: https://journalofbigdata.springeropen.com/articles/10.1186/s40537-024-00957-y

---

*Word count: approximately 2,500 words*
