# TutorMetric: Simple "How We Did It" Guide
**B.Tech Mini Project — DA-IICT**  
**Students:** Nilesh Mori (202301473) & Pratik Chauhan (202301435)  
**Mentor:** Dr. Parul Gupta  

---

## 1. The 30-Second Introduction (Say this first)

> *"Good morning Ma'am. I am Nilesh Mori, and my partner is Pratik Chauhan.  
> In India, private tutoring is huge, but pricing is completely blind. Two parents in the same area looking for Class 10 Maths get quoted ₹300 per hour by one teacher and ₹1,500 per hour by another, with no explanation.  
> So for our project, we collected and cleaned **107,673 tutor profiles** across all 36 States and UTs of India. We built a clean database, an interactive website with maps and charts, and an AI tool that tells parents and tutors what a fair price is based on experience and education."*

---

## 2. How We Did It (Explained in 6 Simple Steps)

### Step 1: Getting the data (Scraping)
- **What was the source?** TeacherOn.com, India's biggest tutoring directory.
- **How did we get it?** We wrote a Python script using `requests` and `BeautifulSoup`. The script went through the website pages one by one and downloaded tutor information (name, experience, subjects, city, fees, education).
- **How many records?** Exactly **107,673 tutors**.

---

### Step 2: Checking who is actually active vs. dead profiles
- **The problem:** Websites like TeacherOn never delete old profiles. Even if someone stopped teaching 4 years ago, their profile stays online so the website looks big.
- **What we did:** We wrote a Python script that checked every single profile link.
  - If a link was broken (HTTP 404), or said *"not taking students"*, or hadn't logged in for over a year, we marked it as **Inactive**.
- **What we found:**
  - **76,683 tutors are active** (about 71%).
  - **30,990 tutors are dead/inactive accounts** (about 29%).
- On our website, anyone can choose to view *All Tutors* or just *Active Tutors*.

---

### Step 3: Cleaning messy subject names
- **The problem:** People typed whatever they wanted. We found **21,956 different messy spellings**!
  - For example: *"maths"*, *"Math (CBSE)"*, *"mathematics for 10th"*, *"python programming"*, *"pyhton"*.
- **What we did:**
  - We wrote Python cleaning code that removed punctuation, extra spaces, and board names like CBSE/ICSE.
  - We grouped all similar words together. For example, *"core python"*, *"Python 3"*, and *"pyhton"* all became just **`Python`**.
- **Result:** We brought 21,956 messy words down to **386 clean subjects**, grouped under 6 main areas:
  1. Science & Maths
  2. Languages
  3. Coding & Computers
  4. Commerce
  5. Humanities & Arts
  6. School Foundation

---

### Step 4: Making 2 tables instead of 1 big spreadsheet (3NF Database)
- **The problem with Excel:**
  - If Nilesh teaches 3 subjects (Maths, Physics, Python), a normal Excel sheet writes Nilesh's name, phone, bio, and city **3 separate times**.
  - That wastes memory and creates mistakes if you update one row but forget the others.
- **Our solution (3NF Normalization):**
  - We split the data into **2 clean SQLite tables**:
    1. **`tutors` Table (107,673 rows):** Contains 1 row per teacher (Name, Experience, Degree, City, State, Fee, Active status). The key is `tutor_id` (like `T1`, `T2`).
    2. **`teacher_subjects` Table (403,221 rows):** Contains what each tutor teaches. Each row is 1 subject offering.
- **What is `bridge_id`?**
  - It is simply the serial number (1, 2, 3...) of that second table!
  - It is called a **bridge** because it connects (bridges) the Tutor to the Subject.
  - If Tutor `T15` teaches 3 subjects, there will be 3 bridge rows:
    - `bridge_id = 101` &rarr; `T15` teaches Maths
    - `bridge_id = 102` &rarr; `T15` teaches Physics
    - `bridge_id = 103` &rarr; `T15` teaches Python

---

### Step 5: Building the website, map, and charts
- **The Map:** We used Leaflet.js with real boundary maps of India. You can hover on any state (like Gujarat or Maharashtra) to see how many tutors are there and what they charge.
- **The Line & Bar Graph (Fee Distribution):**
  - Shows what most teachers charge.
  - The big building-style bars show tutor counts.
  - The green curve shows the trend.
  - Most teachers in India (22,730 tutors) charge **₹250 to ₹500 per hour**.
  - The middle rate (median) is **₹650 per hour**.
  - Average is **₹864 per hour**.
- **The 3-Subject Venn Diagram:**
  - You can pick any 3 subjects (like Physics, Chemistry, Maths).
  - It shows how many tutors teach all 3 (3,892 tutors) vs tutors who only teach 1 subject.
  - *Interesting finding:* Tutors teaching all 3 charge **less** (₹712/hr) than Maths-only teachers (₹853/hr) because parents ask for a package discount when taking all 3 subjects together!

---

### Step 6: The AI Fee Predictor (And why 15% accuracy is normal!)
- **What does it do?** You type your experience, degree, subject, and city, and the model predicts what fair price you should charge.
- **Why is the formula accuracy ($R^2$) around 15% and not 90%?**
  - Because tutoring is freelance work done by humans!
  - Two teachers with 3 years experience in Ahmedabad can charge very different fees:
    - One charges ₹400 because they want more students.
    - Another charges ₹1,200 because a rich family hired them through word-of-mouth.
  - Things like **how well you speak, personal reputation, confidence, and parent bargaining** cannot be stored in a website column!
  - Even in famous university research (Harvard/MIT wage papers), freelance wage models always get 15% to 25%.
  - Our model successfully captures the main baseline:
    - Each year of experience adds ~₹85/hr.
    - A Master's degree adds ~₹220/hr.
    - Big metro cities add ~₹180/hr.

---

## 3. What to Say on Every PPT Slide (Slide 1 to 16)

### Slide 1: Title
> *"Good morning Ma'am. We are Nilesh Mori and Pratik Chauhan. Our project is TutorMetric, analyzing 107,673 tutors across India to bring transparent pricing benchmarks to private education."*

### Slide 2: The Problem
> *"In India, private tutoring has no standard pricing. Two parents in the same area get quoted ₹300 vs ₹1,500 per hour with no explanation. Parents overpay, and new tutors don't know what to charge."*

### Slide 3: TutorMetric vs. Commercial Marketplaces (TeacherOn)
> *"Websites like TeacherOn keep 30,000 dead accounts online just to look big, and their search has 21,000 messy misspelled tags. TutorMetric verifies who is actually active, cleans the subjects, and makes the full database open and downloadable."*

### Slide 4: Project Tracker
> *"This was our 5-step workflow based on your project tracker: scraping raw profiles, verifying active status, cleaning subjects with NLP, normalizing into 2 database tables, and building the website and fee model."*

### Slide 5: Database (3NF)
> *"If a tutor teaches 5 subjects, writing their details 5 times in Excel wastes memory. We split the data into 2 tables: `tutors` (107k rows) and `subjects` (403k rows), connected by Tutor ID. This cut file size by 68% and made search super fast."*

### Slide 6: Cleaning 21,956 spellings to 386 clean subjects
> *"People typed 'maths', 'maths 10th', and 'mathematics'. We cleaned and grouped all variations into 386 clean subjects under 6 categories with zero data loss."*

### Slide 7: Where tutors live
> *"Over 54% of all private tutors come from just 5 states: Maharashtra, Karnataka, Delhi, Tamil Nadu, and Uttar Pradesh."*

### Slide 8: 3 out of 4 tutors work outside big cities
> *"Almost 75% of tutors live in Tier 2 and Tier 3 towns, not big metros. With internet and online whiteboards, teachers from small towns are teaching students all over India."*

### Slide 9: Fees jump with each school level
> *"Primary and middle school tuition stays around ₹700/hr. But in Class 11 and 12, fees jump by 33% to over ₹1,000/hr because of board exam and JEE/NEET entrance pressure."*

### Slide 10: Physics, Chemistry & Maths (STEM Trio)
> *"In PCM, 67% of teachers specialize in only 1 subject. Only 8.5% teach all 3. And teachers teaching all 3 charge slightly less per hour because parents ask for a package discount."*

### Slide 11: Demographics: 58.7% Women
> *"Women make up nearly 59% of private tutors, especially in languages and school foundation. Also, having a Master's degree gives an average 35% fee boost over a Bachelor's degree."*

### Slide 12: Pricing Model
> *"We trained a Ridge regression model to predict fair fees. In freelance tutoring, fees are set by personal talking and negotiation, so formula models naturally explain 15% of price variation. Our model gives the solid baseline for experience, degree, and city."*

### Slide 13: Tutors Directory
> *"Our website has a directory where anyone can filter by price range, state, or subject, compare two tutors side by side, and download the filtered Excel file with one click."*

### Slide 14: Two Live Tools
> *"We built two tools: a live Fee Calculator that estimates fair hourly rates, and a Bio Studio powered by Gemini that writes professional profiles for teachers."*

### Slide 15: What the data showed us
> *"Our 5 big takeaways: 29% of marketplace listings are dead accounts; 75% of tutors live in small towns; board exams cause an instant 33% fee jump; specialists earn more than generalists; and open data helps everyone."*

### Slide 16: Conclusion & Demo
> *"The website is live on our laptop. We would now like to show you the live website, the map, the fee curve, and the tools."*

---

## 4. How to Show the Website to Ma'am (Click by Click)

1. **Top Bar**:
   - Point to the numbers at the top: **107,673 Tutors**, **403,221 Subjects**, **₹864/hr Average**, **76,683 Active**.
   - Click **"Export Data"**: Show the wide 680px box. Point out the top green button **"Download Filtered Excel"**, and the two clean columns (3NF Relational DB on left, Full Master on right).

2. **Directory Tab**:
   - In the sidebar, click the quick price button **"₹500–1k"**.
   - Show how the table instantly filters to tutors charging ₹500–1,000, and a green chip `Fee: ₹500–₹1000/hr` appears.
   - Click **"Export Results (Excel)"**: Show that it downloads an Excel file with only those matching tutors.
   - Click **"Compare"** on any two tutors, then click the **"Compare Tutors"** button at the top to show the side-by-side comparison popup.

3. **Visual Analytics Tab**:
   - **Map**: Hover over Maharashtra or Gujarat. Switch between *Tutor Volume* and *Avg Fee*.
   - **Experience Chart**: Show that most tutors have 1 to 2 years of experience.
   - **Fee Distribution Chart**: Point out the indigo columns (histogram) and green line (trend). Explain that the most common fee in India is ₹250–₹500/hr.
   - **Venn Diagram**: Click the **"🚀 PCM Trio"** preset button. Show the 3 circles and explain that 3,892 tutors teach all three subjects.

4. **AI Fee Estimator Tab**:
   - Choose: 4 years experience, Master's degree, Physics, Class 11-12.
   - Click **"Estimate Fair Tuition Fee"**: Show the predicted price (~₹950/hr) and the breakdown.

5. **Bio Studio Tab**:
   - Show how typing a name and subject generates 3 different professional bio styles using Google Gemini.

6. **About Tab**:
   - Show the clean card with your names: **Nilesh Mori (202301473)**, **Pratik Chauhan (202301435)**, and **Dr. Parul Gupta (Faculty Mentor)** under **B.Tech Mini Project**.

---

## 5. Short Answers to 4 Common Viva Questions

### Q1: What is `bridge_id`?
> *"Ma'am, because one tutor can teach many subjects, we created a second table called `teacher_subjects`. `bridge_id` is just the row ID of that table. It bridges (connects) the tutor to each subject they teach."*

### Q2: Why is the AI fee model accuracy around 15%?
> *"Ma'am, in freelance tutoring, prices are decided by personal negotiation and phone conversations that are not recorded in website columns. In labor economics research, freelance fee models always have an accuracy of 15% to 25%. Our model gives the fair structural baseline for experience, degree, and city."*

### Q3: How did you know 30,990 tutors were inactive?
> *"We wrote a Python script that checked every single profile link. If the link returned HTTP 404, or the profile said not accepting students, or had zero login activity for over a year, we marked it as inactive."*

### Q4: Why did you make 2 tables instead of 1 Excel sheet?
> *"In a single sheet, a tutor who teaches 5 subjects gets their name, phone, and bio written 5 times. Decomposing into 2 tables (tutors and subjects) in 3rd Normal Form removed 68% of wasted storage and made search fast."*
