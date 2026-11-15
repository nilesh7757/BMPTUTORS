#!/usr/bin/env python3
"""
Update SQLite Database with Mentor Fixes:
1. Clean 'speaks_languages' column where scraping errors placed Fee details.
2. Populate 'fee_raw', 'fee_unit', 'hourly_fee_...', and 'monthly_fee_...' for shifted rows.
3. Add 'monthly_fee_min', 'monthly_fee_max', 'monthly_fee_avg' columns and populate for all monthly-fee tutors.
4. Anonymise tutor names to protect privacy.
"""

import sqlite3
import re
import hashlib

DB_PATH = "cleaned_tutors_data.db"

# Realistic first names for gender-appropriate fictitious names
FIRST_NAMES_FEMALE = [
    "Aanya", "Aarohi", "Aditi", "Ananya", "Anushka", "Avani", "Deepika", "Divya", "Isha", "Kavya",
    "Khushi", "Meera", "Neha", "Nisha", "Pooja", "Prerna", "Priya", "Radhika", "Rhea", "Riya",
    "Sakshi", "Sanya", "Shreya", "Sneha", "Sonali", "Tanvi", "Vandana", "Varsha", "Vidya", "Zoya"
]
FIRST_NAMES_MALE = [
    "Aarav", "Aditya", "Akhil", "Alok", "Aman", "Amit", "Anand", "Ankit", "Arjun", "Ashish",
    "Ayush", "Dev", "Gaurav", "Harsh", "Karan", "Manish", "Mohit", "Nikhil", "Pranav", "Rahul",
    "Rajesh", "Rohan", "Rohit", "Sachin", "Sanjay", "Siddharth", "Sumit", "Varun", "Vikas", "Vivek"
]
LAST_INITIALS = ["S.", "P.", "K.", "M.", "R.", "V.", "G.", "T.", "N.", "D.", "J.", "B.", "C.", "A.", "L."]

def get_fictitious_name(tutor_id, gender):
    h = int(hashlib.md5(tutor_id.encode('utf-8')).hexdigest()[:8], 16)
    last_init = LAST_INITIALS[h % len(LAST_INITIALS)]
    if gender == 'Female':
        first = FIRST_NAMES_FEMALE[(h // len(LAST_INITIALS)) % len(FIRST_NAMES_FEMALE)]
    elif gender == 'Male':
        first = FIRST_NAMES_MALE[(h // len(LAST_INITIALS)) % len(FIRST_NAMES_MALE)]
    else:
        # Neutral / Not specified
        pool = FIRST_NAMES_FEMALE if (h % 2 == 0) else FIRST_NAMES_MALE
        first = pool[(h // len(LAST_INITIALS)) % len(pool)]
    return f"{first} {last_init} ({tutor_id})"

def parse_fee(text):
    if not text:
        return None, None, None, None, None
    t = text.strip()
    if t.startswith('Fee:'):
        t = t[4:].strip()
    
    # Unit
    unit = 'other'
    if re.search(r'\/hour|\/hr|\bhr\b|\bhour\b', t, re.I):
        unit = 'hour'
    elif re.search(r'\/month|\/mo|\bmonth\b|\bmonthly\b', t, re.I):
        unit = 'month'
    elif re.search(r'\/day|\bday\b|\bdaily\b', t, re.I):
        unit = 'day'
    elif re.search(r'\/week|\bweek\b|\bweekly\b', t, re.I):
        unit = 'week'
    elif re.search(r'\/class|\bclass\b|\bsession\b', t, re.I):
        unit = 'class'

    inr_part = t.split('(USD')[0] if '(USD' in t else t
    cleaned_num_text = re.sub(r'(\d+)\s*,\s*(\d+)', r'\1\2', inr_part)
    nums = [float(x) for x in re.findall(r'\b\d+(?:\.\d+)?\b', cleaned_num_text)]
    if not nums:
        return t, unit, None, None, None
    elif len(nums) == 1:
        val = nums[0]
        return t, unit, val, val, val
    else:
        min_v = min(nums[:2])
        max_v = max(nums[:2])
        avg_v = (min_v + max_v) / 2.0
        return t, unit, min_v, max_v, avg_v

def run_updates():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # 1. Add columns if not existing
    cur.execute("PRAGMA table_info(teachers);")
    existing_cols = {row[1] for row in cur.fetchall()}

    for col in ["monthly_fee_min", "monthly_fee_max", "monthly_fee_avg"]:
        if col not in existing_cols:
            print(f"Adding column {col} to teachers...")
            cur.execute(f"ALTER TABLE teachers ADD COLUMN {col} REAL;")

    if "original_name" not in existing_cols:
        print("Adding column original_name to teachers...")
        cur.execute("ALTER TABLE teachers ADD COLUMN original_name TEXT;")
        cur.execute("UPDATE teachers SET original_name = teacher_name;")

    conn.commit()

    # 2. Fix rows where speaks_languages has fee text
    print("\n[Step 2] Cleaning speaks_languages shifted fee rows...")
    cur.execute("SELECT tutor_id, speaks_languages FROM teachers WHERE speaks_languages LIKE 'Fee:%';")
    shifted_rows = cur.fetchall()
    print(f"  Found {len(shifted_rows)} rows with 'Fee:' in speaks_languages.")

    cleaned_shifted_count = 0
    for tid, lang_text in shifted_rows:
        raw_fee, unit, min_v, max_v, avg_v = parse_fee(lang_text)
        
        # update query
        if unit == 'hour':
            cur.execute("""
                UPDATE teachers
                SET speaks_languages = 'Not Specified',
                    num_languages = 0,
                    fee_raw = ?,
                    fee_unit = 'hour',
                    hourly_fee_min = ?,
                    hourly_fee_max = ?,
                    hourly_fee_avg = ?
                WHERE tutor_id = ?
            """, (raw_fee, min_v, max_v, avg_v, tid))
        elif unit == 'month':
            cur.execute("""
                UPDATE teachers
                SET speaks_languages = 'Not Specified',
                    num_languages = 0,
                    fee_raw = ?,
                    fee_unit = 'month',
                    monthly_fee_min = ?,
                    monthly_fee_max = ?,
                    monthly_fee_avg = ?
                WHERE tutor_id = ?
            """, (raw_fee, min_v, max_v, avg_v, tid))
        else:
            cur.execute("""
                UPDATE teachers
                SET speaks_languages = 'Not Specified',
                    num_languages = 0,
                    fee_raw = ?,
                    fee_unit = ?
                WHERE tutor_id = ?
            """, (raw_fee, unit, tid))
        cleaned_shifted_count += 1

    print(f"  Cleaned {cleaned_shifted_count} shifted fee rows successfully.")
    conn.commit()

    # 3. Populate monthly_fee_min/max/avg for all monthly-fee tutors
    print("\n[Step 3] Populating monthly_fee_min, monthly_fee_max, monthly_fee_avg for all monthly-fee tutors...")
    cur.execute("SELECT tutor_id, fee_raw FROM teachers WHERE fee_unit = 'month' AND monthly_fee_avg IS NULL;")
    monthly_rows = cur.fetchall()
    print(f"  Found {len(monthly_rows)} monthly fee rows needing parsing.")

    parsed_monthly_count = 0
    for tid, f_raw in monthly_rows:
        _, _, min_v, max_v, avg_v = parse_fee(f_raw)
        if avg_v is not None:
            cur.execute("""
                UPDATE teachers
                SET monthly_fee_min = ?,
                    monthly_fee_max = ?,
                    monthly_fee_avg = ?
                WHERE tutor_id = ?
            """, (min_v, max_v, avg_v, tid))
            parsed_monthly_count += 1

    print(f"  Populated monthly fees for {parsed_monthly_count} tutors.")
    conn.commit()

    # 4. Anonymise tutor names
    print("\n[Step 4] Anonymising tutor names to protect privacy...")
    cur.execute("SELECT tutor_id, gender FROM teachers;")
    all_teachers = cur.fetchall()
    
    anonymised_updates = []
    for tid, gender in all_teachers:
        fict_name = get_fictitious_name(tid, gender)
        anonymised_updates.append((fict_name, tid))

    cur.executemany("UPDATE teachers SET teacher_name = ? WHERE tutor_id = ?;", anonymised_updates)
    print(f"  Anonymised names for all {len(anonymised_updates):,} educators.")
    conn.commit()

    # Verification queries
    print("\n[Verification Summary]")
    cur.execute("SELECT COUNT(*) FROM teachers WHERE speaks_languages LIKE 'Fee:%';")
    print("  Remaining 'Fee:' in speaks_languages:", cur.fetchone()[0])

    cur.execute("SELECT COUNT(*) FROM teachers WHERE monthly_fee_avg IS NOT NULL;")
    print("  Total tutors with monthly_fee_avg populated:", cur.fetchone()[0])

    cur.execute("SELECT COUNT(*) FROM teachers WHERE hourly_fee_avg IS NOT NULL;")
    print("  Total tutors with hourly_fee_avg populated:", cur.fetchone()[0])

    cur.execute("SELECT tutor_id, teacher_name, speaks_languages, fee_raw, fee_unit, hourly_fee_avg, monthly_fee_avg FROM teachers LIMIT 5;")
    for row in cur.fetchall():
        print("  Sample row:", row)

    conn.close()
    print("\n✅ Database update completed successfully!")

if __name__ == '__main__':
    run_updates()
