import pptx
import shutil

src = '/home/nilesh7757/Downloads/BMP_Tutors_Deck_Full.pptx'
prs = pptx.Presentation(src)

# Define comprehensive speaker notes for all 18 slides
SPEAKER_NOTES = {
    1: """[SLIDE 01: TITLE & CREDITS]
Target Time: 35s
Spoken Script:
"Respected professors, examiners, and fellow students, good morning. We are Nilesh Mori and Pratik Chauhan, presenting our B.Tech Mini Project for course SC452 under the mentorship of Dr. Parul Gupta: TutorMetric — What 107,673 private tutors in India actually charge.
Private tutoring in India is an unorganized shadow education market worth billions of rupees that operates in near-total informational secrecy. We took the raw scraped dataset our mentor provided, verified which profiles actively still teach, and built an open-access platform featuring fair-price benchmarks, geospatial analytics, dynamic subject bundling, and predictive fee tools. We engineered this entire system as joint equal co-engineers."
""",
    2: """[SLIDE 02: WHAT WE'LL COVER]
Target Time: 25s
Spoken Script:
"Our presentation is structured into six concise chapters following our project tracker roadmap. We start with the core problem of pricing opacity and stale listings. Then, we walk through our data cleaning pipeline—reducing 78 messy columns into two 3NF normalized tables. Next, we present our empirical findings on where tutors live, what they teach, and how fees scale. Finally, we discuss our fee regression model before keeping the last 5 minutes for a live demonstration of our deployed web platform."
""",
    3: """[SLIDE 03: THE PROBLEM & RAW DATASET]
Target Time: 45s
Spoken Script:
"Our project began with raw TeacherOn data provided by Dr. Parul Gupta in TrueData.xlsx. In this market, two parents in Ahmedabad seeking a Class 10 CBSE Maths tutor can be quoted ₹200 an hour by one teacher and ₹2,000 by another for the exact same syllabus. Neither parent has any public price benchmark to know what is fair.
Furthermore, commercial platforms monetize phone number unlocks and leave inactive listings online. Our audit discovered that 30,990 profiles—28.8% of the dataset—were completely dormant ghost accounts. Our mandate was to clean this data, verify active teachers, and publish transparent pricing benchmarks."
Examiner Defense:
If asked how you identified the 28.8% dormant profiles, state that automated Playwright headless test scripts audited profile response signals, activity timestamps, and uncontactable channels.
""",
    4: """[SLIDE 04: WHY TUTORMETRIC IS DIFFERENT]
Target Time: 35s
Spoken Script:
"Traditional tutor websites merely display static advertisements: a photo, a name, and a self-reported fee with zero context. TutorMetric transforms this into a true public utility.
We add three crucial capabilities: First, transparent price benchmarks—showing state medians and predicted fair rates. Second, active verification status so parents never waste time chasing dormant accounts. Third, side-by-side comparability across 107,673 educators in an open, downloadable dataset for parents, tutors, and educational researchers."
""",
    5: """[SLIDE 05: THE DIRECTORY DASHBOARD]
Target Time: 40s
Spoken Script:
"Here is the core directory dashboard of TutorMetric. Finding the right tutor takes four clicks because the interface is built around nine reactive multi-parameter filters. A parent or student can filter by state, city tier, hourly fee range, delivery modality, years of experience, gender, and active verification status.
Crucially, all 386 clean canonical disciplines are indexed. The search runs with sub-50ms query latency against our SQLite backend, and any filtered view can be exported directly to Excel."
""",
    6: """[SLIDE 06: WHAT WE BUILT]
Target Time: 40s
Spoken Script:
"This slide summarizes our complete production deliverables, deployed live at bmptutors.onrender.com. The platform is backed by 107,673 verified tutor profiles and 403,221 verified subject offerings across 386 canonical disciplines.
The national mean hourly rate is ₹864, with a median of ₹650. Exactly 76,683 educators were verified as active. Around this database, we built five operational tools: the multi-filter directory, visual analytics with Leaflet GIS maps, an ML fee estimator, an AI bio generator, and relational Excel exports."
Examiner Defense:
If asked why mean fee is ₹864 while median is ₹650, explain that fees follow a right-skewed Pareto distribution where top entrance-exam specialists charging ₹2,000+ pull up the mean.
""",
    7: """[SLIDE 07: DATA ENGINEERING METHOD]
Target Time: 45s
Spoken Script:
"Our data engineering followed four reproducible steps tracked in ProjectTracker.xlsx. We merged the scraped (76,706) and not scraped (31,017) cohorts from TrueData.xlsx into 107,673 unique profiles. We stripped direct PII and assigned anonymous tutor IDs T1, T2...
We then implemented the exact taxonomy from Categories_DataCleaning.docx: standardizing locations into 5 macro-regions, mapping qualifications into 4 tiers, and extracting numeric experience. We kept three promises: zero subject entries were lost, 21,956 raw spellings were unified into 386 disciplines, and 95.6% of tutors were identified as offering online instruction."
""",
    8: """[SLIDE 08: 3NF RELATIONAL DATABASE DESIGN]
Target Time: 45s
Spoken Script:
"From a DBMS perspective, the original spreadsheet was heavily unnormalized, packing 20 repeated subject columns per tutor with over 60% NULL cells.
We decomposed the schema into strict Third Normal Form with two relational tables: 'teachers' holding 107,673 unique educator profiles, and 'teacher_subjects' holding all 403,221 offerings linked by tutor_id. Because tutors teach an average of 3.79 subjects, this 1-to-many junction table eliminated NULL waste, reduced storage footprint by 64%, and enabled indexed B-Tree joins that execute in under one millisecond."
Examiner Defense:
If asked why not MongoDB, explain that relational 3NF guarantees ACID transactional integrity, avoids massive duplicate writes when updating subjects, and executes indexed joins instantly.
""",
    9: """[SLIDE 09: NLP SUBJECT CANONICALIZATION]
Target Time: 40s
Spoken Script:
"When teachers register online, free-text entries produce enormous semantic entropy. We uncovered 21,956 unique raw spelling variants. For example, 'Maths', 'MATHEMATICS', and 'Math (CBSE)' were separate strings. A search for 'Mathematics' would have missed 70% of available teachers.
Using a custom 4-pass NLP pipeline—lowercasing, punctuation stripping, regex discipline extraction, and hierarchical mapping—we reduced vocabulary clutter by 98.24% down to 386 clean canonical disciplines. A strict database assertion confirmed that row counts remained exactly 403,221 before and after mapping with zero data loss."
""",
    10: """[SLIDE 10: GEOGRAPHY & SOUTHERN SUPPLY HUB]
Target Time: 45s
Spoken Script:
"Our geospatial analysis overturned common assumptions. It is commonly assumed that Delhi or Mumbai dominates private tutoring. In reality, Southern India leads the country: Tamil Nadu is number one with 16,806 tutors, followed by Kerala with 14,333. The top five states—Tamil Nadu, Kerala, Maharashtra, Karnataka, and Delhi—account for 54.4% of India's entire private tutoring workforce.
Crucially, state population does not explain this. Bihar holds 8.6% of India's people but only 1.4% of tutors, whereas Kerala holds 2.8% of the population but 13.3% of tutors, showing that tutor supply is driven by literacy rates and regional demand for shadow coaching."
""",
    11: """[SLIDE 11: THE CENTRAL FINDING - GRASSROOTS DOMINANCE]
Target Time: 45s
Spoken Script:
"Slide 11 represents the central empirical discovery of our project. Commercial media portrays private tutoring as an elite metropolitan service. Our empirical data proves the exact opposite: 74.1% of educators—79,820 tutors—live and work in Tier-3 and semi-urban towns.
Tier-1 Metros represent only 12.7% of tutors. Combined with our finding that 95.6% of tutors teach online, the implication is profound: affordable high-speed broadband has enabled small-town educators to build nationwide digital tutoring practices, serving students in major metros at competitive rates."
""",
    12: """[SLIDE 12: PRICING ECONOMICS - THE ACADEMIC LADDER]
Target Time: 40s
Spoken Script:
"When we examine hourly fees across education levels, we observe a strictly monotonic progression up the academic ladder. Middle school tutoring (Classes 6-8) starts at ₹710/hr, rises to ₹885/hr for Class 10 board prep, reaches ₹1,045/hr for senior secondary science, and peaks at ₹1,257/hr for professional and university upskilling.
Location adds a consistent premium: Tier-1 metro tutors charge an average of ₹1,022/hr compared to ₹823/hr in semi-urban towns. The ₹214 difference between mean (₹864) and median (₹650) is driven by the 8.6% of tutors charging over ₹2,000/hr."
""",
    13: """[SLIDE 13: SUBJECT BUNDLING & THE MPC TRIO]
Target Time: 40s
Spoken Script:
"Our relational schema allows us to analyze multi-subject co-teaching through our interactive Venn Studio. Because teachers offer an average of 3.79 subjects, single-subject tutoring is the exception rather than the rule.
Analyzing the core science trio—Mathematics, Physics, and Chemistry—a total of 47,417 tutors teach at least one of these subjects. Exactly 2,917 educators teach all three simultaneously. For secondary students, these multi-subject educators serve as comprehensive coaching mentors, saving parents from hiring multiple separate tutors."
""",
    14: """[SLIDE 14: WHO TEACHES - GENDER & DEGREE PREMIUM]
Target Time: 40s
Spoken Script:
"In terms of educator demographics, women represent the majority of India's private tutoring workforce: 58.7% female (62,440) versus 41.3% male (43,947) among stating educators. However, there is clear subject segregation: women teach 77% of English and 73% of Biology, whereas Coding is the sole subject where men form the majority at 61%.
In terms of degrees, holding a PhD commands the highest wage premium: doctorate tutors average ₹1,179.5/hr compared to ₹832/hr for bachelor's degree holders—a 41.8% premium. However, only 3,777 tutors hold a PhD, so the market remains predominantly led by bachelor's and master's graduates."
""",
    15: """[SLIDE 15: THE REGRESSION MODEL - REPUTATION VS METRICS]
Target Time: 55s
Spoken Script:
"Fulfilling Tasks 49 to 61 in ProjectTracker.xlsx, we developed an econometric pricing model using L2-regularized Ridge regression (lambda=30) across 38 features on 72,001 outlier-filtered profiles. On an independent held-out split, the model achieves an R-squared of 9.84% with an MAE of ₹370.67/hr, beating naive median guessing by about ₹6/hr.
We present this finding with complete academic honesty. In labor econometrics, human service pricing cannot be accurately predicted purely from categorical profile fields like degrees and city tiers. An educator charging ₹1,500/hr does so because of personal word-of-mouth reputation, school affiliations, past student board ranks, and parental trust—qualities not recorded in public web profiles. Our model successfully provides a realistic baseline floor while proving that private tuition in India runs on trust."
Examiner Defense:
If asked why Ridge instead of Random Forest, explain that decision trees produce erratic step discontinuities and negative pricing artifacts at edge cases, whereas Ridge regression guarantees monotonic, stable, and interpretable coefficients.
""",
    16: """[SLIDE 16: INTERACTIVE TOOLS - FEE AI & BIO STUDIO]
Target Time: 45s
Spoken Script:
"To translate our research into direct practical utilities, we engineered two live interactive tools. The AI Fee Estimator takes seven tutor parameters—experience, tier, qualification, subject, level, and modality—and computes a recommended rate with dynamic confidence bands and monthly earning projections.
The Bio Studio addresses a widespread marketplace problem: over 40% of educators write weak or incomplete profile bios. Using Google Gemini 2.5 Flash with structured system prompting, the studio ingests teacher metadata and generates dual-perspective copy—one tailored for reassuring parents and another engaging students—alongside click-worthy headline hooks."
""",
    17: """[SLIDE 17: CONCLUSION & SUMMARY OF DELIVERABLES]
Target Time: 45s
Spoken Script:
"In conclusion, our B.Tech Mini Project delivers five verified empirical contributions: First, 74.1% of tutors are semi-urban, proving that shadow education in India is a grassroots service. Second, Southern India leads national supply with Tamil Nadu and Kerala holding nearly 30% of tutors. Third, women represent 58.7% of educators. Fourth, one in five tutors (20,275) is a working professional in IT or finance. Fifth, tuition pricing is an informal economy governed by trust and reputation.
We have fully cleaned, normalized, and deployed this 107k-tutor dataset on TutorMetric, and we will now transition directly into our live demonstration."
""",
    18: """[SLIDE 18: THANK YOU & LIVE DEMO PROTOCOL]
Target Time: Transition to Web Browser (4-5 Minutes Live Demo)
Spoken Script:
"Thank you, professors. We invite you to test any subject, city, or fee combination live on our deployed platform at bmptutors.onrender.com."
Live Demo Checklist:
1. Directory Filters: Filter State = Tamil Nadu, Mode = Online (show sub-50ms speed).
2. Profile Modal: Open Tutor T1 (show verified subjects and zero paywall contact).
3. Visual Analytics: Show Leaflet choropleth map & hover tooltips.
4. Venn Studio: Select Maths, Physics, Chemistry to show the 2,917 triple-subject intersection.
5. Fee Estimator & Bio Studio: Run live prediction and Gemini bio generation.
"""
}

# Attach speaker notes to each slide
for idx, slide in enumerate(prs.slides, 1):
    notes_text = SPEAKER_NOTES.get(idx, "")
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = notes_text
    print(f"Attached speaker notes to Slide {idx:02d} ({len(notes_text)} chars)")

# Save updated presentations
out_paths = [
    '/home/nilesh7757/Downloads/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/Downloads/BMP_Tutors_Deck_Full.pptx',
    '/home/nilesh7757/BMPTUTORS1/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/TutorMetric_Master_Presentation.pptx',
    '/home/nilesh7757/BMPTUTORS1/project_showcase_bundle/BMP_Tutors_Deck_Full.pptx'
]

for p in out_paths:
    prs.save(p)
    print(f"Saved master deck to: {p}")

print("Master deck update completed successfully!")
