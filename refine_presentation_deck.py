import pptx
import shutil
import os

src_path = '/home/nilesh7757/Downloads/TutorMetric_Master_Presentation.pptx'
prs = pptx.Presentation(src_path)

print(f"Initial slide count: {len(prs.slides)}")

# 1. Delete Slide 2 (0-indexed position 1: "WHAT WE'LL COVER")
slide_to_delete_idx = 1
rId = prs.slides._sldIdLst[slide_to_delete_idx].rId
prs.part.drop_rel(rId)
del prs.slides._sldIdLst[slide_to_delete_idx]

print(f"Slide count after deleting Slide 2: {len(prs.slides)}")

# 2. Text replacements dictionary
REPLACEMENTS = {
    # Section header badges to academic topic headers
    "SECTION 01  ·  THE PROBLEM": "PROJECT MOTIVATION & PROBLEM STATEMENT",
    "SECTION 01  ·  WHY THIS IS DIFFERENT": "VALUE PROPOSITION & COMPARATIVE UTILITY",
    "SECTION 02  ·  THE DASHBOARD": "SEARCH & FILTERING DASHBOARD",
    "SECTION 02  ·  WHAT WE BUILT": "SYSTEM CAPABILITIES & ARCHITECTURE",
    "SECTION 02  ·  METHOD": "REPRODUCIBLE DATA PIPELINE",
    "SECTION 02  ·  DATABASE DESIGN": "RELATIONAL SCHEMA & 3NF DESIGN",
    "SECTION 02  ·  CLEANING": "CANONICAL TAXONOMY & DATA CLEANING",
    "SECTION 03  ·  GEOGRAPHY": "GEOGRAPHIC DISTRIBUTION & STATE SPREAD",
    "SECTION 03  ·  THE CENTRAL FINDING": "MARKET STRUCTURE: TIER-3 DECENTRALIZATION",
    "SECTION 04  ·  PRICING": "HOURLY FEE DISPERSION & LEVEL LADDER",
    "SECTION 03  ·  SUBJECT BUNDLING": "CURRICULAR CLUSTERING & MULTI-SUBJECT ANALYSIS",
    "SECTION 03  ·  WHO TEACHES": "EDUCATOR DEMOGRAPHICS & DEGREE PREMIUM",
    "SECTION 05  ·  THE MODEL": "ECONOMETRIC MODEL: RIDGE REGRESSION",
    "SECTION 06  ·  TOOLS": "INTERACTIVE UTILITIES: FEE AI & BIO STUDIO",
    "SECTION 06  ·  CONCLUSION": "KEY EMPIRICAL FINDINGS & FUTURE WORK",

    # Informal quotes & conversational phrasing to formal academic prose
    "\"My daughter's maths tutor charges ₹400. Her friend's charges ₹1,100. Both teach Class 10 CBSE. I still don't know who is overcharging.\"": 
        "In Class 10 CBSE Mathematics across identical urban localities, hourly rates vary from ₹200 to ₹2,000 with zero standardized pricing benchmarks or verifiable curriculum alignment.",
    "A parent, Ahmedabad": "EMPIRICAL MARKET OPACITY",
    "SO THE QUESTION WE SET OURSELVES": "RESEARCH MANDATE",

    "Generated profile bios with a quality score": "Personalized profile bios tailored for parent and student conversion",

    "Cleaning took longer than anything else. Most of the work was deciding what counted as a real fee and what was a monthly rate typed into the hourly column.":
        "Systematic fee validation: hourly rates filtered for monthly/weekly unit misclassifications, Winsorized at ₹4,000/hr (99th percentile), and PII anonymized across all 107,673 profiles.",
    "Nilesh & Pratik, on the cleaning pass": "DATA HYGIENE PROTOCOL",

    "A fee model that is honest about what it cannot do": "Econometric modeling of private tuition rates",
    "A linear model on 72,001 tutors, and the plain reason it only explains 9.84% of the price.":
        "Ridge regression ($L_2$ regularized) on 72,001 observations: observable metrics vs unobservable latent factors.",
    "ANYONE CAN REPEAT THIS": "INDEPENDENT VALIDATION",
    "We retrained from the script in the repository and got 9.67% and ₹371.20,\nso the saved model file is right.":
        "Model validation: independently reproducible training script confirms test R² = 9.84% and MAE = ₹370.67.",
    "WHAT IT CANNOT KNOW": "UNOBSERVABLE LATENT FACTORS",
    "What a parent really pays for.": "Primary economic drivers of tutoring rates.",
    "We can suggest a fair range within about ₹370\neither way. We cannot price one particular tutor,\nbecause the evidence that would settle it is not\nin the data. Price here runs on trust.":
        "Observable data establishes an empirical baseline (MAE = ₹370/hr). Unrecorded variance reflects individual reputation, parental trust, and pedagogical quality."
}

# 3. Update shapes, headers, quotes, and slide number badges
for idx, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            # Update slide badge (top right corner)
            if shape.top < 500000 and shape.left > 10000000:
                txt = shape.text_frame.text.strip()
                if txt.isdigit() and len(txt) <= 2:
                    shape.text_frame.text = f"{idx:02d}"
                    continue

            # Text replacements
            full_text = shape.text_frame.text
            for old_str, new_str in REPLACEMENTS.items():
                if old_str in full_text:
                    # Replace in paragraphs to maintain styling
                    for p in shape.text_frame.paragraphs:
                        if old_str in p.text:
                            p.text = p.text.replace(old_str, new_str)

# 4. Define rigorous 17-slide speaker notes
SPEAKER_NOTES_17 = {
    1: """[SLIDE 01: TITLE & CREDITS]
Target Time: 35s
Spoken Script:
"Respected professors, examiners, and fellow students, good morning. We are Nilesh Mori and Pratik Chauhan, presenting our B.Tech Mini Project for course SC452 under the mentorship of Dr. Parul Gupta: TutorMetric — What 107,673 private tutors in India actually charge.
Private tutoring in India is an unorganized shadow education market worth billions of rupees that operates in near-total informational secrecy. We took the raw scraped dataset our mentor provided, verified which profiles actively still teach, and built an open-access platform featuring fair-price benchmarks, geospatial analytics, dynamic subject bundling, and predictive fee tools. We engineered this entire system as joint equal co-engineers."
""",

    2: """[SLIDE 02: PROJECT MOTIVATION & PROBLEM STATEMENT]
Target Time: 45s
Spoken Script:
"We begin directly with the motivation of our project. In India's private tuition marketplace, severe information asymmetry harms both parents and educators. For instance, in Class 10 CBSE Mathematics within the same urban area, one tutor quotes ₹200 an hour while another quotes ₹2,000 for the exact same syllabus, with zero public price benchmarks available to assess fair value.
Furthermore, commercial listing portals monetize lead unlocking and keep stale listings online. Our empirical audit revealed that 30,990 profiles—28.8% of the dataset—were completely inactive ghost accounts. Our project mandate was to clean this master data, eliminate stale listings, standardize curriculum taxonomy, and publish an open empirical benchmark of tutoring fees across India."
Examiner Defense:
If asked how inactive accounts were identified, explain that automated Playwright headless scripts and activity signals audited response timestamps, last active indicators, and uncontactable channels.
""",

    3: """[SLIDE 03: VALUE PROPOSITION & COMPARATIVE UTILITY]
Target Time: 35s
Spoken Script:
"Traditional tutor websites merely display static advertisements: a photo, a name, and an unverified self-reported fee with zero context. TutorMetric transforms this into a transparent analytical utility.
We add three distinct capabilities: First, transparent price benchmarks—providing state medians and econometric baseline estimates. Second, active verification status so families never waste time contacting dormant profiles. Third, side-by-side comparability across 107,673 educators in an open, downloadable database for parents, tutors, and academic researchers."
""",

    4: """[SLIDE 04: SEARCH & FILTERING DASHBOARD]
Target Time: 40s
Spoken Script:
"This slide illustrates the primary directory dashboard of TutorMetric. Users can isolate relevant educators in under four clicks using nine reactive multi-parameter filters. Parents can filter by state, city tier, hourly fee range, delivery modality, years of experience, gender, and active verification status.
Crucially, all 386 clean canonical disciplines are indexed. The search executes with sub-50ms query latency against our SQLite backend, and any filtered view can be exported directly to Excel."
""",

    5: """[SLIDE 05: SYSTEM CAPABILITIES & ARCHITECTURE]
Target Time: 40s
Spoken Script:
"This slide summarizes our complete production deliverables, deployed live at bmptutors.onrender.com. The platform is backed by 107,673 verified tutor profiles and 403,221 verified subject offerings across 386 canonical disciplines.
The national mean hourly rate is ₹864, with a median of ₹650. Exactly 76,683 educators were verified as active. Around this database, we built five operational utilities: the multi-filter directory, visual analytics with Leaflet GIS maps, an ML fee estimator, an AI bio generator, and relational Excel exports."
Examiner Defense:
If asked why mean fee is ₹864 while median is ₹650, explain that fees follow a right-skewed Pareto distribution where top entrance-exam specialists charging ₹2,000+ pull up the mean.
""",

    6: """[SLIDE 06: REPRODUCIBLE DATA PIPELINE]
Target Time: 45s
Spoken Script:
"Our data engineering followed four reproducible steps tracked in ProjectTracker.xlsx. We merged the scraped (76,706) and not scraped (31,017) cohorts from TrueData.xlsx into 107,673 unique profiles. We stripped direct PII and assigned anonymous tutor IDs T1, T2...
We then implemented the exact taxonomy from Categories_DataCleaning.docx: standardizing locations into 5 macro-regions, mapping qualifications into 4 tiers, and extracting numeric experience. We kept three promises: zero subject entries were lost, 21,956 raw spellings were unified into 386 disciplines, and 95.6% of tutors were identified as offering online instruction."
""",

    7: """[SLIDE 07: RELATIONAL SCHEMA & 3NF DESIGN]
Target Time: 45s
Spoken Script:
"From a DBMS perspective, the original spreadsheet was heavily unnormalized, packing 20 repeated subject columns per tutor with over 60% NULL cells.
We decomposed the schema into strict Third Normal Form with two relational tables: 'teachers' holding 107,673 unique educator profiles, and 'teacher_subjects' holding all 403,221 offerings linked by tutor_id. Because tutors teach an average of 3.79 subjects, this 1-to-many junction table eliminated NULL waste, reduced storage footprint by 64%, and enabled indexed B-Tree joins that execute in under one millisecond."
Examiner Defense:
If asked why not MongoDB, explain that relational 3NF guarantees ACID transactional integrity, avoids massive duplicate writes when updating subjects, and executes indexed joins instantly.
""",

    8: """[SLIDE 08: CANONICAL TAXONOMY & DATA CLEANING]
Target Time: 45s
Spoken Script:
"A major challenge in ProjectTracker.xlsx was resolving subject text fragmentation. Educators entered free text, producing 21,956 distinct raw strings. For example, 'Math', 'Maths', 'Mathematics', and 'Pure Math' were entered separately.
We designed a four-stage NLP normalization pipeline: case folding, delimiter splitting, phonetic mapping, and dictionary matching against Indian curricula. This consolidated 21,956 variants into 386 canonical disciplines with zero row loss, preserving all 403,221 offerings and enabling rigorous multi-subject analytics."
""",

    9: """[SLIDE 09: GEOGRAPHIC DISTRIBUTION & STATE SPREAD]
Target Time: 45s
Spoken Script:
"Moving to our empirical findings: tutor supply is heavily geographically concentrated. Five states account for 54.4% of all private tutors in India. Tamil Nadu leads nationally with 16,806 tutors, followed by Kerala with 14,333, Maharashtra with 11,544, West Bengal with 8,024, and Delhi with 7,869.
This reveals a sharp regional divergence: Kerala represents 2.8% of India's population but 13.3% of its tutors, whereas Bihar represents 8.6% of the population but only 1.4% of tutors. Tutoring supply reflects regional educational formalization and private coaching propensity rather than raw population."
""",

    10: """[SLIDE 10: MARKET STRUCTURE: TIER-3 DECENTRALIZATION]
Target Time: 50s
Spoken Script:
"This slide presents the most critical empirical finding of our project: 74.1% of private tutors—79,820 educators—operate outside Tier-1 metropolitan areas in Tier-3 and semi-urban towns. Only 12.7% (13,622 tutors) are located in Tier-1 cities.
EdTech commercial discourse focuses almost entirely on metro cities, yet the vast reality of Indian private tutoring is a distributed, small-town phenomenon. Because 95.6% of these educators teach online, high-speed broadband has connected non-metro educators to students nationwide, flattening geographic boundaries."
Examiner Defense:
If asked why online percentage is so high (95.6%), clarify that post-2020 digital infrastructure adoption led tutors to list online capability alongside offline home tutoring.
""",

    11: """[SLIDE 11: HOURLY FEE DISPERSION & LEVEL LADDER]
Target Time: 45s
Spoken Script:
"Examining fee structure, we observe a monotonic price ladder across education levels. The national median hourly rate is ₹650, while the mean is ₹864 due to high-fee outliers.
Rates scale strictly with academic complexity: Pre-primary and Primary tutors average ₹710 per hour, Secondary tutors average ₹825, Higher Secondary averages ₹950, and Professional/College trainers average ₹1,257—a 77% premium over school foundations. Geographically, Tier-1 cities command an average premium of ₹199 per hour over Tier-3 towns (₹1,022 vs ₹823)."
""",

    12: """[SLIDE 12: CURRICULAR CLUSTERING & MULTI-SUBJECT ANALYSIS]
Target Time: 40s
Spoken Script:
"Our normalized 3NF database allows us to analyze subject co-occurrence through set theory. A central question in Indian coaching is: how many educators teach the entire Science trio of Mathematics, Physics, and Chemistry?
Across 107,673 tutors, 34,923 teach at least one of these three subjects. However, exactly 2,917 tutors—only 8.35% of STEM educators—teach all three simultaneously. In contrast, 16,929 tutors specialize strictly in Mathematics alone. Multi-subject generalists are rare; specialization dominates secondary and higher secondary coaching."
""",

    13: """[SLIDE 13: EDUCATOR DEMOGRAPHICS & DEGREE PREMIUM]
Target Time: 40s
Spoken Script:
"In educator demographics, women represent the majority of private tutors at 58.7% of profiles reporting gender. This female representation is highest in languages and biological sciences: English at 77%, Hindi at 79%, and Biology at 73%. Mathematics is balanced at 58% women, while Coding is the only male-dominated discipline at 61% men.
Regarding academic qualifications, tutors holding a PhD or Advanced Postgraduate degree command a 41.8% premium, averaging ₹1,179.5 per hour versus ₹832 for Undergraduate degree holders."
""",

    14: """[SLIDE 14: ECONOMETRIC MODEL: RIDGE REGRESSION]
Target Time: 55s
Spoken Script:
"Fulfilling Tasks 49 to 61 in ProjectTracker.xlsx, we developed an econometric pricing model on 72,001 observations using Ridge regression with L2 regularization. The model incorporates 38 dummy and continuous features, achieving an R² of 9.84% with a Mean Absolute Error of ₹370.67 per hour.
A key question is: why does an objective 38-feature model explain only 9.84% of fee variance? Econometrically, tutoring is an experience good governed by unobservable latent factors: word-of-mouth reputation, personal patience, institutional pedigree, and past student examination results. Observable metrics establish a fair baseline range of ±₹370/hr, but individual pricing fundamentally hinges on parental trust."
Examiner Defense:
If asked why Ridge rather than OLS or Gradient Boosting, explain that Ridge handles multicollinearity among one-hot encoded subject features through L2 regularization, preventing overfitting and yielding robust, interpretable coefficients.
""",

    15: """[SLIDE 15: INTERACTIVE UTILITIES: FEE AI & BIO STUDIO]
Target Time: 45s
Spoken Script:
"To translate our research into direct practical utilities, we engineered two interactive web tools:
First, the AI Fee Estimator. A tutor or parent enters seven parameters—subject, student level, city tier, experience, mode, and qualification. The backend model outputs a recommended hourly rate, a confidence range (±₹370), estimated monthly and annual income, and an itemized cost driver breakdown.
Second, the Bio Studio, powered by Google Gemini GenAI. It generates structured professional biographies tailored for parent and student conversions, producing engaging headlines and clear credentials."
""",

    16: """[SLIDE 16: KEY EMPIRICAL FINDINGS & FUTURE WORK]
Target Time: 45s
Spoken Script:
"In conclusion, our B.Tech Mini Project delivers five verified empirical conclusions:
First, 74.1% of tutoring operates in Tier-3 semi-urban markets.
Second, Southern states dominate national supply, led by Tamil Nadu and Kerala.
Third, women represent the majority of tutors at 58.7%, particularly in languages and sciences.
Fourth, 18.8% of tutors teach alongside another primary occupation.
Fifth, tutoring fees are anchored by trust and reputation, explaining the quantitative bounds of observable pricing models.
For future work, we propose longitudinal fee tracking across quarters, district-level GIS mapping within states, and integrating student reviews and ratings into the pricing model."
""",

    17: """[SLIDE 17: CONCLUSION & SYSTEM DEMONSTRATION]
Target Time: Transition to Web Browser (4-5 Minutes Live Demo)
Spoken Script:
"Thank you, professors. We invite questions on our database schema, cleaning methodology, or econometric findings. We will now switch to the live deployed application at bmptutors.onrender.com for our interactive system demonstration."
Examiner Defense:
Demonstration checklist ready:
1. Search directory: filter Ahmedabad -> Class 10 -> Mathematics -> Active only -> export Excel.
2. Analytics: show Tier-3 distribution pie chart, Leaflet state choropleth, and MPC 3-way Venn diagram.
3. Fee Estimator: input Class 10 Math, Tier-2, 5 yrs experience -> show ₹850 recommendation and itemized breakdown.
4. Bio Studio: generate student-converting bio via Gemini API.
"""
}

# 5. Inject speaker notes into all 17 slides
for idx, slide in enumerate(prs.slides, 1):
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = SPEAKER_NOTES_17[idx].strip()

# 6. Save modified presentation
dest_paths = [
    '/home/nilesh7757/Downloads/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/Downloads/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/BMP_Tutors_Deck_Full.pptx'
]

for d in dest_paths:
    prs.save(d)
    print(f"Saved: {d} (Size: {os.path.getsize(d):,} bytes)")

print("Refinement completed successfully!")
