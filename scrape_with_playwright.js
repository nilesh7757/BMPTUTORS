#!/usr/bin/env node
/**
 * TeacherOn Playwright Headless Browser Scraper
 * Uses headless Firefox to bypass Cloudflare and verify live tutor profiles.
 *
 * Usage:
 *   node scrape_with_playwright.js --limit 20
 *   node scrape_with_playwright.js --tutor-id T76707
 */

const fs = require('fs');
const path = require('path');
const { firefox } = require('/home/nilesh7757/.npm/_npx/e41f203b7505f1fb/node_modules/playwright');
const { execSync } = require('child_process');

const DB_PATH = path.join(__dirname, 'cleaned_tutors_data.db');

function parseArgs() {
  const args = process.argv.slice(2);
  let limit = 10;
  let tutorId = null;
  let delay = 3000;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--limit' && args[i + 1]) {
      limit = parseInt(args[i + 1], 10);
      i++;
    } else if (args[i] === '--tutor-id' && args[i + 1]) {
      tutorId = args[i + 1];
      i++;
    } else if (args[i] === '--delay' && args[i + 1]) {
      delay = parseInt(args[i + 1], 10);
      i++;
    }
  }
  return { limit, tutorId, delay };
}

function querySql(query) {
  const safeQuery = query.replace(/"/g, '\\"');
  const cmd = `python3 -c "import sqlite3, json; conn = sqlite3.connect('${DB_PATH}'); c = conn.cursor(); c.execute(\\"${safeQuery}\\"); print(json.dumps(c.fetchall()))"`;
  try {
    const out = execSync(cmd, { encoding: 'utf-8' });
    return JSON.parse(out);
  } catch (err) {
    console.error('SQL query error:', err.message);
    return [];
  }
}

function updateTutorStatus(tutorId, isActive, reason) {
  const safeReason = reason.replace(/'/g, "''");
  const cmd = `python3 -c "import sqlite3; conn = sqlite3.connect('${DB_PATH}'); c = conn.cursor(); c.execute('UPDATE teachers SET is_active = ?, validation_reason = ? WHERE tutor_id = ?', (${isActive}, '${safeReason}', '${tutorId}')); conn.commit(); conn.close()"`;
  try {
    execSync(cmd, { encoding: 'utf-8' });
  } catch (err) {
    console.error(`Failed to update ${tutorId}:`, err.message);
  }
}

async function run() {
  const { limit, tutorId, delay } = parseArgs();

  let tutors = [];
  if (tutorId) {
    tutors = querySql(`SELECT tutor_id, profile_url, teacher_name FROM teachers WHERE tutor_id = '${tutorId}'`);
  } else {
    tutors = querySql(`SELECT tutor_id, profile_url, teacher_name FROM teachers WHERE source_tab = 'not scraped' LIMIT ${limit}`);
  }

  console.log('=======================================================');
  console.log(` BMP Tutors - Playwright Live Profile Verifier`);
  console.log(` Queue Size: ${tutors.length} tutors | Delay: ${delay}ms`);
  console.log('=======================================================\n');

  const browser = await firefox.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:130.0) Gecko/20100101 Firefox/130.0',
    viewport: { width: 1366, height: 768 }
  });

  const page = await context.newPage();

  let activeCount = 0;
  let deletedCount = 0;
  let blockedCount = 0;

  for (let i = 0; i < tutors.length; i++) {
    const [tid, url, name] = tutors[i];
    process.stdout.write(`[${i + 1}/${tutors.length}] Checking ${tid} (${name || 'Anonymous'}) -> ${url} ... `);

    try {
      const resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 25000 });
      await page.waitForTimeout(3000);

      const status = resp ? resp.status() : 0;
      const bodyText = await page.innerText('body').catch(() => '');

      if (bodyText.includes('This profile has been deleted') || status === 404) {
        console.log('❌ DELETED PROFILE');
        updateTutorStatus(tid, 0, 'No Profile on TeacherOn (Deleted)');
        deletedCount++;
      } else if (bodyText.includes('Performing security verification') || bodyText.includes('Attention Required')) {
        console.log('🛡️ CLOUDFLARE CHALLENGE');
        blockedCount++;
      } else if (bodyText.includes('Last login:') || bodyText.includes('Hourly fee') || bodyText.includes('Send message')) {
        console.log('✅ LIVE ACTIVE PROFILE');
        updateTutorStatus(tid, 1, 'Verified Active on TeacherOn');
        activeCount++;
      } else {
        console.log(`ℹ️ HTTP ${status} (Profile Inactive / Archived)`);
        updateTutorStatus(tid, 0, 'Inactive on TeacherOn');
      }
    } catch (err) {
      console.log(`⚠️ TIMEOUT / ERROR: ${err.message.slice(0, 35)}`);
    }

    await page.waitForTimeout(delay);
  }

  await browser.close();

  console.log('\n=======================================================');
  console.log(` Playwright Verification Completed`);
  console.log(` Live Active: ${activeCount} | Deleted: ${deletedCount} | Cloudflare Protected: ${blockedCount}`);
  console.log('=======================================================\n');
}

run().catch(console.error);
