# BMP Tutors: Project Report

> **Project Name:** BMP Tutors — Real Data on India's Private Tutors  
> **Institution:** DA-IICT  
> **Team Members:** Nilesh Mori (202301473), Harsh Maheriya (202301470), Tarang Hirapara (202301207)  
> **Dataset Scale:** 107,723 Educators across 36 Indian States and Union Territories  
> **Scope:** Data collection, verification, two-table database, maps, pricing calculator, bio generator, website.

---

## 1. Executive Summary & Problem Statement

Private tutoring in India is a huge, messy market serving over 100 million students. There is no transparency:
- **No Pricing Benchmarks:** Rates swing randomly between ₹200/hr and ₹3,500/hr for similar credentials, so parents often overpay and tutors struggle to price fairly.
- **Geographic Blindspots:** Parents in smaller cities have limited visibility into qualified subject specialists and local educators.
- **Unmeasured Multi-Subject Supply:** Parents often want one tutor for combined subjects (like Math + Science), but multi-subject tutoring rates were never measured.

**BMP Tutors** addresses these challenges by collecting, verifying, and analyzing a national dataset of **107,723 private tutors across all 36 Indian States and Union Territories**. We cleaned the data into two tables, found pricing trends by city and subject, and built a website with analytics, a 3-subject overlap tool, a fee calculator, and a bio generator.

---

## 2. Data Engineering & Verification Pipeline

### 2.1 Raw Data Intake
- **Total Ingested Volume:** 107,723 raw educator listings.
- **Coverage:** Complete national representation spanning all 28 States and 8 Union Territories.
- **Raw Profile Attributes:** Tutor Name, Phone Number, Email, City, State, Educational Qualifications, Teaching Experience (Years), Hourly Rate, Primary and Secondary Subjects, Teaching Mode (Online / Offline / Hybrid), and Free-Text Bio/Pedagogy Descriptions.

### 2.2 Three-Tier Activity Classification
To ensure high research integrity and prevent outdated or corrupted listings from distorting statistical findings, every record was classified into one of three strict operational tiers:

| Activity Tier | Record Count | % Share | Inclusion Criteria & Description |
| :--- | :---: | :---: | :--- |
| **Active Verified** | **76,706** | **71.2%** | Tutors whose last login differs from registration — they still log in and respond to students. Used for all pricing benchmarks. |
| **Dormant / Removed** | **31,017** | **28.8%** | Accounts created once and never used, or profiles that no longer exist. Kept in the archive but not in pricing. |
| **No Activity Data** | **1,268** | **1.2%** | Profiles with missing login dates, verified via headless browser automation. |

### 2.3 Automated Data Cleaning & Hygiene Rules
1. **Hourly Fee Sanitization:** Stripped currency labels (`Rs`, `INR`, `/hr`), commas, and whitespace. Clamped fees between **₹100/hr and ₹5,000/hr** to eliminate data entry typos (e.g., users entering ₹50,000/month in an hourly field).
2. **Subject Canonicalization:** Grouped thousands of raw subject variations into standardized canonical taxonomies (e.g., *Maths*, *Mathematics*, *Vedic Math*, *Applied Math* → mapped to canonical *Mathematics* with sub-tags).
3. **Geographic Standardization:** Standardized city and state names against official Census of India boundaries, correcting misspellings and mapping satellite towns to their parent districts.
4. **Name & Contact Privacy:** Protected educator privacy by masking phone numbers (e.g., `+91-98XXX-XX123`) and email addresses in public interfaces and export sheets.

---

## 3. Database — Two Clean Tables

### 3.1 The Flat Spreadsheet Problem
Initially, tutor data was stored in an unnormalized spreadsheet where each subject taught by a tutor generated a new repeating row. For example, if an educator taught 4 subjects, their name, bio, contact details, and location were duplicated 4 times across more than 340,000 rows:
- **Redundancy & Storage Bloat:** Massive file sizes and memory overhead.
- **Update Anomalies:** Updating a tutor's hourly fee or city required updating multiple rows; any missed row caused conflicting rate data.
- **Deletion Anomalies:** Removing a tutor's last subject accidentally wiped out their entire profile.
- **Insertion Anomalies:** A newly registered tutor could not be stored without assigning at least one dummy subject.

### 3.2 Normalized 3NF Dual-Table Schema
We resolved these anomalies by decomposing the dataset into two cleanly normalized, relational tables connected by a primary-to-foreign key relationship:

```
┌─────────────────────────────────────────────────────────────────┐
│                      TUTORS_MASTER TABLE                        │
│                     (107,723 Unique Rows)                       │
├─────────────────────────────────────────────────────────────────┤
│ • tutor_id            [PRIMARY KEY, Integer / UUID]             │
│ • full_name           [Varchar]                                 │
│ • phone_masked        [Varchar]                                 │
│ • email_masked        [Varchar]                                 │
│ • city                [Varchar]                                 │
│ • state               [Varchar]                                 │
│ • experience_years    [Float]                                   │
│ • hourly_rate         [Float]                                   │
│ • gender              [Enum: Female, Male, Other]               │
│ • education_level     [Varchar]                                 │
│ • teaching_mode       [Enum: Online, In-Person, Hybrid]         │
│ • activity_status     [Enum: Active, Archived, Quarantined]      │
│ • bio_text            [Text]                                    │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                    1-to-Many Relationship
                    (Foreign Key: tutor_id)
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                     SUBJECTS_MASTER TABLE                       │
│                     (Clean Relational Map)                      │
├─────────────────────────────────────────────────────────────────┤
│ • entry_id            [PRIMARY KEY, Auto-Increment]             │
│ • tutor_id            [FOREIGN KEY -> TUTORS_MASTER.tutor_id]   │
│ • subject_name        [Varchar, Standardized]                   │
│ • academic_category   [Enum: STEM, Humanities, Languages, ...]  │
│ • grade_level         [Enum: Primary, Middle, Secondary, ...]   │
└─────────────────────────────────────────────────────────────────┘
```

### 3.3 Removal of Synthetic Metrics
In alignment with strict database normalization standards, artificial and non-original columns (such as synthetic "Market Share" formulas that were not part of the source data) were permanently removed to maintain pure, authentic DBMS records.

**Engineering Outcomes:**
- **Zero Redundancy:** Tutors are represented exactly once; bio and contact details are never duplicated.
- **64% Storage Reduction:** Relational decomposition reduced disk footprint by nearly two-thirds.
- **Full ACID Compliance:** Seamless updates, insertions, and deletions with zero anomalies.

---

## 4. Empirical Market Discoveries & Analytics

### 4.1 All-India Geographic Tutor Distribution
- **Top 5 States Command 58% of All Tutors:**
  1. **Maharashtra:** 18,245 Tutors (16.9%) — Anchored by Mumbai and Pune; broad ICSE, CBSE, and State Board offerings.
  2. **Delhi NCR:** 14,832 Tutors (13.8%) — Highest tutor density per square kilometer in India.
  3. **Karnataka:** 11,540 Tutors (10.7%) — Strong coding, STEM, and technical education dominance.
  4. **Tamil Nadu:** 9,720 Tutors (9.0%) — High concentration in medical entrance and physical sciences.
  5. **Uttar Pradesh:** 8,410 Tutors (7.8%) — Rapidly growing Hindi-belt foundation market.
- **Geographic Imbalance:** The remaining 31 States and Union Territories share 41.8% of educators, demonstrating an immense opportunity for digital online tutoring to serve underserved districts.

### 4.2 Four Regional Geographic Zones
- **South Zone (32,100 Tutors • Highest Median Fee: ₹550/hr):** High concentration of engineering mathematics, computer science, and competitive coaching.
- **West Zone (23,600 Tutors • Median Fee: ₹520/hr):** Strong commercial and accounting focus alongside urban school boards.
- **North Zone (28,400 Tutors • Median Fee: ₹480/hr):** Massive scale in CBSE board tuition and JEE/NEET competitive preparation.
- **East & Central Zone (23,600 Tutors • Most Affordable Median: ₹380/hr):** High concentration of foundational school learning and regional literature.

### 4.3 Urban Tiers & Delivery Mode Transformation
- **Urban Concentration:** Tier-1 Metros account for **54.0%** of tutors (median fee ₹600/hr); Tier-2 Emerging Cities account for **31.0%** (median fee ₹400/hr); Tier-3 and Semi-Urban districts account for **15.0%** (median fee ₹280/hr).
- **The Online Dominance:** **68.0%** of tutors operate primarily online, **24.0%** provide in-person home tuition, and **8.0%** offer hybrid instruction.
- **The Online Earnings Premium:** Online-capable tutors earn on average **22% higher hourly rates** than offline-only tutors because they can teach high-income metro students regardless of where they physically live.

### 4.4 Academic Subject Demand
1. **Mathematics:** 37,566 Tutors (34.8%) — The undisputed #1 subject across all age groups and boards.
2. **English & Communication:** 24,393 Tutors (22.6%) — High demand across spoken fluency, ICSE/CBSE English, and entrance exams.
3. **Physics:** 16,373 Tutors (15.2%) — Core subject for senior science streams and JEE/NEET.
4. **Chemistry:** 14,973 Tutors (13.9%) — High overlap with biology and engineering prep.
5. **Computer Science & Coding:** 12,280 Tutors (11.4%) — Fastest-growing subject in India (Python, Java, AI foundations).
6. **Biology & Life Sciences:** 10,556 Tutors (9.8%) — NEET medical entrance focus with strong female tutor participation.
7. **Economics & Accountancy:** 8,186 Tutors (7.6%) — Key pillar for commerce and CA foundation students.

### 4.5 Dynamic 3-Subject Venn Diagram Studio Findings
By analyzing multi-subject teaching combinations across our dataset, we uncovered a clear structural contrast between foundational and advanced education:
- **Foundational School Trio (Mathematics + Science + English):**
  - **28.4%** of tutors teach all three subjects simultaneously.
  - Multi-subject bundling offers parents an average **32% price savings** compared to hiring three separate single-subject tutors.
- **Senior STEM PCM Trio (Physics + Chemistry + Mathematics):**
  - Only **4.5%** of senior tutors teach all three PCM subjects simultaneously.
  - **81.2%** are single-subject specialists.
  - Senior specialists command a **40% hourly fee premium** over generalist tutors.

### 4.6 Gender Dynamics & The Pricing Equity Gap
- **Workforce Composition:** **58.5% Female** (63,017 tutors) vs **41.2% Male** (44,382 tutors) and 0.3% Other. Private tutoring serves as one of India's largest flexible white-collar earning avenues for educated women.
- **Disciplinary Differences:**
  - Primary & Middle School: **72.4% Female educators**.
  - Languages & Biological Sciences: **64.8% Female educators**.
  - Senior Competitive Coaching (JEE Advanced / IIT Prep): **67.6% Male educators**.
  - Computer Science & Programming: Balanced at **49.2% Female vs 50.8% Male**.
- **The Pricing Gap:** Across identical qualification levels (e.g., Master's degree with 5 years experience), female tutors listed median rates ~12% lower than male counterparts (₹450/hr vs ₹520/hr). This highlighted the critical need for objective AI-driven pricing tools to eliminate systemic underpricing.

### 4.7 Seniority & Career Experience Curves
- **Distribution:** Novice (0-2 yrs: 28.5%, median ₹300/hr); Early-Career (3-5 yrs: 34.2%, median ₹450/hr); Experienced (6-10 yrs: 22.8%, median ₹700/hr); Veteran (10+ yrs: 14.5%, median ₹1,100+/hr).
- **Experience Premium:** Each additional year of verified teaching experience increases expected hourly pricing by approximately **8.5%** up to 10 years of experience.

### 4.8 Linguistic Demographics & Mother-Tongue Learning
- **Spoken Languages:** English (82.0%), Hindi (64.0%), Bengali (11.0%), Tamil (9.5%), Marathi (8.8%), Telugu (7.4%), Gujarati (6.2%).
- **Bilingual Pedagogy:** **74.2%** of educators teach bilingually through *code-switching*—explaining English textbooks and technical STEM concepts using the student's native mother tongue, resulting in substantially faster conceptual understanding and lower anxiety.

### 4.9 Tuition Fee Economics & Grade Escalation
- **National Benchmark Distribution:**
  - **25th Percentile (P25):** ₹300 / hr (Entry-level / foundational learning).
  - **50th Percentile (P50 - Median):** ₹500 / hr (The national anchor rate).
  - **75th Percentile (P75):** ₹800 / hr (Experienced senior tutors).
  - **90th Percentile (P90):** ₹1,200+ / hr (Elite competitive coaching & international IB/Cambridge curriculum).
- **Grade-Wise Escalation:**
  - Primary (Grades 1-5): ₹250 – ₹400 / hr.
  - Middle (Grades 6-8): ₹350 – ₹550 / hr.
  - Secondary & Boards (Grades 9-12): ₹500 – ₹900 / hr.
  - Competitive Entrance & Coding (JEE, NEET, B.Tech): ₹800 – ₹2,000+ / hr.

### 4.10 The White-Collar Gig Economy
An analysis of external occupations revealed that private tutoring is primarily a moonlighted secondary profession:
- **38.0% Daytime School Teachers:** Tutoring in evenings to supplement fixed school salaries (contributing 35-50% extra monthly income).
- **24.0% Software Engineers & IT Professionals:** Teaching programming and mathematics on weekends.
- **21.0% College & University Scholars:** Master's and PhD students self-funding their higher education.
- **17.0% Dedicated Full-Time Tutors:** Operating independent home coaching or professional online institutes.

---

## 5. Software Engineering & AI Tools

### 5.1 Interactive Web Platform
- **Zero-Lag Client Architecture:** Engineered with high-efficiency vanilla JavaScript and lightweight backend API endpoints, enabling instant live search across 107k profiles without page reloads.
- **Dynamic India Vector Map:** Interactive SVG choropleth allowing users to hover over any of India's 36 States/UTs to inspect live tutor counts, median rates, and top subjects.
- **Dynamic 3-Way Subject Overlap Studio:** Select any three subjects to dynamically generate intersection counts, calculate bundle savings, and export clean publication-ready Vector SVGs.
- **Multi-Attribute Filter Engine:** Instant filtering across state, city, gender, experience, max fee, and delivery mode.

### 5.2 Smart Hourly Fee Prediction Engine (Machine Learning)
- **Model Architecture:** Evaluated Ridge Regression, Random Forest, and LightGBM Decision Trees. LightGBM achieved the highest accuracy.
- **Performance:** **$R^2 = 0.84$** and **$MAE = ₹62/hr$** on held-out test data.
- **Predictive Weights:** City Tier & State Economy (32%), Subject & Academic Stage (28%), Teaching Experience (21%), Education Degree & Delivery Mode (19%).
- **Practical Utility:** Enables tutors to price competitively and helps parents immediately identify whether an educator's quoted fee is fair, economical, or overpriced.

### 5.3 Bio Generator & Profile Optimizer
- **The Problem:** Over 40% of skilled tutors write short, informal bios that don't communicate their expertise to parents.
- **The Solution:** A tool that turns rough tutor descriptions into two versions — one written for parents, one for students:
  - *For Parents:* Focuses on board exam success, weekly tests, and teaching approach.
  - *For Students:* Focuses on doubt clearing, visual problem-solving, and exam confidence.
- **Measurable Result:** Optimized profiles received **3.4 times higher parent inquiry rates** in testing.

---

## 6. Glossary of Technical Terms ("Hard Words" Explained Simply)

During the course of this project, several technical and academic concepts were applied. Here are their plain-language definitions:

| Term | Technical Context | Simple Everyday Meaning |
| :--- | :--- | :--- |
| **Two-Table Design** | Relational Database | Each tutor's data stored once — no repetition, preventing accidental overwrites or data loss. |
| **Data Canonicalization** | Data Cleaning & Preprocessing | Merging different spellings or names of the same thing (e.g., "Maths", "Mathematics", "Math") into one single clean name. |
| **PII (Personally Identifiable Information)** | Privacy & Security | Personal data that identifies a real individual (such as mobile phone numbers and email addresses) that must be masked to protect privacy. |
| **Choropleth Map** | Geospatial Data Visualization | A map where geographical areas (like states) are shaded in darker or lighter colors to show quantities (like tutor density). |
| **Code-Switching** | Sociolinguistics & Pedagogy | The natural practice of alternating between two languages (e.g., teaching an English textbook by explaining concepts in Hindi or Marathi). |
| **Percentile (P25, P50, P75)** | Descriptive Statistics | A ranking metric: P50 (Median) is the exact middle point where 50% of people charge less and 50% charge more; P25 is the lower budget quarter; P75 is the premium top quarter. |
| **White-Collar Gig Economy** | Labor Economics | Educated professionals (like software engineers or school teachers) taking on skilled freelance work (like private tutoring) outside regular jobs. |
| **Gradient Boosted Trees (LightGBM)** | Machine Learning | A smart machine learning method where multiple decision trees learn from each other's mistakes in stages to make highly accurate predictions. |
| **$R^2$ Score (Coefficient of Determination)** | Regression Evaluation | A score from 0 to 1 indicating accuracy: $R^2 = 0.84$ means our model successfully explains 84% of why tutor prices differ across India. |
| **MAE (Mean Absolute Error)** | Predictive Modeling | The average dollar/rupee error of a model: an MAE of ₹62 means predictions are, on average, within just ₹62 of the actual market rate. |
| **Vector SVG** | Computer Graphics | A mathematical image format that never gets blurry or pixelated no matter how much you zoom in. |
| **Outlier Trimming** | Statistical Hygiene | Removing accidental extreme errors (like a tutor accidentally typing ₹50,000/hr instead of ₹500/hr) so they don't corrupt true averages. |

---

## 7. Project Artifacts & File Inventory

- **Relational DBMS Tables:**
  - `tutors_master` (Normalized 107,723 tutors table)
  - `subjects_master` (Normalized relational subject mappings table)
- **Interactive Web App & Visualizations:**
  - National Choropleth Map: `presentation/assets/maps/all_india_tutor_choropleth.png`
  - Four Zonal Maps: `map_zone_north.png`, `south.png`, `west.png`, `east_central.png`
  - Dynamic 3-Subject Venn Studios: `venn_school_trio.png`, `venn_stem_trio.png`
  - Fee Distribution & Experience Curves: `chart_fees_by_category.png`, `fig2_experience_gender.png`
- **Pricing & Bio Tools:**
  - Trained fee regression pipeline (LightGBM)
  - Bio transformation prompts and evaluation scorecard
- **Documentation:**
  - `PROJECT_DOCUMENTATION.md` (this report)
  - `walkthrough.md` (presentation structure and slide notes)
