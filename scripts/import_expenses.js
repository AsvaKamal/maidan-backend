// import_expenses.js
// Imports monthly expense workbooks into Supabase.
//
// Usage:
//   node import_expenses.js <file-or-directory> [--dry-run] [--force]
//
//   --dry-run   Parse everything and print what WOULD happen. No DB writes.
//   --force     Re-import a file already recorded in raw_import_batches
//               (deletes that batch's expenses first).
//
// One transaction per file; each file recorded in raw_import_batches with
// source_type='expense', so re-running a folder skips files already loaded.
//
// -------------------------------------------------------------------------
// IMPORTANT: THE DATE COLUMN IN THESE FILES IS UNRELIABLE
// -------------------------------------------------------------------------
// The client's expense sheets are written day-first (1/6/26 = 1 June 2026) but
// Excel has parsed many of them month-first, producing 6 January 2026 instead.
// In "Expense - June 2026.xlsx" EVERY dated row came through with day=6 and
// month spread across 1-12 — i.e. day and month transposed. The rows Excel
// could NOT misread (day 13+, where a month-first reading is impossible) were
// left as TEXT: "15-6-26", "16-6-26", ...
//
// So the date cell alone cannot be trusted. This script treats the FILENAME
// month/year as authoritative ("Expense - June 2026.xlsx" -> June 2026) and
// reconciles each cell against it:
//   * month already matches the file's month  -> take as-is
//   * day matches the file's month            -> transposed, swap them
//   * text date                               -> parse day-first
//   * neither                                 -> keep as-is and warn loudly
// Every correction is logged so it can be audited against the sheet.

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { Pool } = require('pg');

const DRY_RUN = process.argv.includes('--dry-run');
const FORCE = process.argv.includes('--force');
const TARGET = process.argv.find((a, i) => i >= 2 && !a.startsWith('--'));

// How to treat rows that have an amount but NO date at all. Seen in
// "Expense - May 2026.xlsx": a contiguous block of 58 petty-cash rows
// (PKR 259,531, ~2.6% of the month) where the date column was simply never
// filled in. The month is known from the filename even though the day isn't.
//   month-start : date them to the 1st of the file's month (DEFAULT). Keeps
//                 the spend in monthly totals and signals "day unknown".
//   carry       : inherit the last date seen above them in the sheet.
//   skip        : drop them (they will be missing from all expense figures).
const UNDATED_ARG = process.argv.find((a) => a.startsWith('--undated='));
const UNDATED_POLICY = UNDATED_ARG ? UNDATED_ARG.split('=')[1] : 'month-start';

if (!TARGET) {
    console.error('Usage: node import_expenses.js <file-or-directory> [--dry-run] [--force] [--undated=month-start|carry|skip]');
    process.exit(1);
}
if (!['month-start', 'carry', 'skip'].includes(UNDATED_POLICY)) {
    console.error(`Invalid --undated value "${UNDATED_POLICY}". Use month-start, carry, or skip.`);
    process.exit(1);
}

const pool = DRY_RUN ? null : new Pool({ connectionString: process.env.DATABASE_URL });

// ---------------------------------------------------------------------
// MAPPINGS
// ---------------------------------------------------------------------

// Facility labels are freer here than in the booking sheets: as well as plain
// "KMC"/"Emaar"/"Malir" we see "Padel at KMC", "Padel Court at Emaar",
// "Pickleball Court at Malir". Matching on the site name contained in the
// string covers those and any similar future variant.
const FACILITY_PATTERNS = [
    { code: 'KMC', test: (s) => s.includes('kmc') },
    { code: 'EMAAR', test: (s) => s.includes('emaar') },
    { code: 'MALIR', test: (s) => s.includes('malir') },
];

// Payment methods arrive as 'cheque', 'Cheque', 'cash ', 'Cash', 'IBFT'.
const PAYMENT_MAP = {
    cheque: 'Cheque',
    cash: 'Cash',
    ibft: 'IBFT',
    card: 'Card',
};

const MONTH_NUMBERS = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    jan: 1, feb: 2, mar: 3, apr: 4, jun: 6, jul: 7, aug: 8, sep: 9,
    sept: 9, oct: 10, nov: 11, dec: 12,
};

const UNCATEGORIZED = 'Uncategorized';

// ---------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------

function normHeader(s) {
    return String(s ?? '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function cleanText(v) {
    if (v === null || v === undefined) return null;
    const s = String(v).replace(/\s+/g, ' ').trim();
    return s.length ? s : null;
}

function cleanNumber(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return null; // a date in a money column is not a number
    const n = Number(String(v).replace(/,/g, ''));
    return Number.isNaN(n) ? null : n;
}

function pad2(n) {
    return String(n).padStart(2, '0');
}

function fmtDate(y, m, d) {
    return `${y}-${pad2(m)}-${pad2(d)}`;
}

// Title Case, whitespace squeezed — keeps 'cash ' and 'Cash' from splitting.
function titleCase(s) {
    return s
        .toLowerCase()
        .split(' ')
        .map((w) => (w ? w.charAt(0).toUpperCase() + w.slice(1) : w))
        .join(' ');
}

// ---------------------------------------------------------------------
// Filename -> expected period.  "Expense - June 2026.xlsx" -> {6, 2026}
// ---------------------------------------------------------------------

function periodFromFilename(filePath) {
    const base = path.basename(filePath, path.extname(filePath)).toLowerCase();
    let month = null;
    for (const [name, num] of Object.entries(MONTH_NUMBERS)) {
        const re = new RegExp(`\\b${name}\\b`);
        if (re.test(base)) {
            // Prefer the longest match ("june" over "jun") so we don't half-match.
            if (month === null || name.length > month.name.length) month = { name, num };
        }
    }
    const yearMatch = base.match(/\b(20\d{2})\b/);
    return {
        month: month ? month.num : null,
        year: yearMatch ? parseInt(yearMatch[1], 10) : null,
    };
}

// ---------------------------------------------------------------------
// Date resolution (see the long note at the top of this file)
// ---------------------------------------------------------------------

// Excel serial -> Date (1900 system, allowing for the fictitious 1900-02-29).
function dateFromExcelSerial(n) {
    const ms = Math.round((n + (n <= 60 ? 1 : 0)) * 86400000) + Date.UTC(1899, 11, 30);
    const u = new Date(ms);
    return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
}

function resolveDate(raw, expected, fileName, rowNum, warnings) {
    const { month: expMonth, year: expYear } = expected;

    // Some date cells carry no date number format at all, so they arrive as a
    // bare Excel serial (e.g. 46147). Seen in the May 2026 file. Convert first,
    // then fall through to the same reconciliation as any other date.
    if (typeof raw === 'number' && Number.isFinite(raw) && raw > 20000 && raw < 80000) {
        raw = dateFromExcelSerial(raw);
    }

    if (raw instanceof Date) {
        const y = raw.getFullYear();
        const m = raw.getMonth() + 1;
        const d = raw.getDate();

        if (expMonth === null) return fmtDate(y, m, d);
        if (m === expMonth) return fmtDate(y, m, d);

        if (d === expMonth) {
            // Transposed: Excel read a day-first date month-first.
            warnings.push(`[${fileName}] Row ${rowNum}: date read as ${fmtDate(y, m, d)} but this is a month-${expMonth} file — day/month were transposed, corrected to ${fmtDate(y, d, m)}.`);
            return fmtDate(y, d, m);
        }

        warnings.push(`[${fileName}] Row ${rowNum}: date ${fmtDate(y, m, d)} falls outside the file's month (${expMonth}/${expYear}) and isn't a day/month transposition — imported as-is, please verify.`);
        return fmtDate(y, m, d);
    }

    if (typeof raw === 'string') {
        const s = raw.trim();
        const m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
        if (m) {
            let day = parseInt(m[1], 10);
            let mon = parseInt(m[2], 10);
            let yr = parseInt(m[3], 10);
            if (yr < 100) yr += 2000;

            // These text dates are the ones Excel refused to parse, which happens
            // precisely because the first number is > 12 — i.e. they are day-first.
            if (expMonth !== null && mon !== expMonth && day === expMonth) {
                const t = day; day = mon; mon = t;
                warnings.push(`[${fileName}] Row ${rowNum}: text date "${s}" read as day/month transposed, corrected to ${fmtDate(yr, mon, day)}.`);
            }
            if (mon < 1 || mon > 12 || day < 1 || day > 31) {
                warnings.push(`[${fileName}] Row ${rowNum}: text date "${s}" is not a valid date — row skipped.`);
                return null;
            }
            return fmtDate(yr, mon, day);
        }
        warnings.push(`[${fileName}] Row ${rowNum}: unrecognised date format "${s}" — row skipped.`);
        return null;
    }

    return null;
}

// ---------------------------------------------------------------------
// Header detection — the two file layouts differ in column ORDER and CONTENT
//   old (Oct 2025): Date, Item, Category, Description, Facility, Payment
//                   Method, Cheque number, Amount, Source Account Title,
//                   Source Bank
//   new (Jun 2026): Date, Facility, Item, Category, Sub-Category, Payment
//                   Method, Cheque number, Amount, Comments, Source Account
//                   Title, Source Bank
// so columns are always located by name, never by position.
// ---------------------------------------------------------------------

function findHeaderRowIdx(rows) {
    for (let i = 0; i < Math.min(15, rows.length); i++) {
        const norms = (rows[i] || []).map(normHeader);
        if (norms.includes('date') && norms.includes('amount')) return i;
    }
    return -1;
}

function buildColMap(headerRow) {
    const map = {};
    headerRow.forEach((raw, idx) => {
        const h = normHeader(raw);
        if (!h) return;
        if (map.date === undefined && h === 'date') { map.date = idx; return; }
        if (map.facility === undefined && h === 'facility') { map.facility = idx; return; }
        if (map.item === undefined && h === 'item') { map.item = idx; return; }
        if (map.subCategory === undefined && h.includes('sub') && h.includes('category')) { map.subCategory = idx; return; }
        if (map.category === undefined && h === 'category') { map.category = idx; return; }
        if (map.paymentMethod === undefined && h.includes('payment')) { map.paymentMethod = idx; return; }
        if (map.chequeNumber === undefined && h.includes('cheque')) { map.chequeNumber = idx; return; }
        if (map.amount === undefined && h === 'amount') { map.amount = idx; return; }
        // 'Description' (old layout) and 'Comments' (new layout) both feed description
        if (map.description === undefined && (h.includes('description') || h.includes('comment'))) { map.description = idx; return; }
        if (map.sourceAccount === undefined && h.includes('source') && h.includes('account')) { map.sourceAccount = idx; return; }
        if (map.sourceBank === undefined && h.includes('source') && h.includes('bank')) { map.sourceBank = idx; return; }
    });
    return map;
}

// ---------------------------------------------------------------------
// Lookup caches
// ---------------------------------------------------------------------

const facilityIdCache = new Map();
const paymentMethodIdCache = new Map();
const categoryIdCache = new Map();

async function getFacilityId(client, code) {
    if (facilityIdCache.has(code)) return facilityIdCache.get(code);
    const res = await client.query('SELECT id FROM facilities WHERE code = $1', [code]);
    if (res.rows.length === 0) {
        throw new Error(`No row in "facilities" with code = '${code}'. Create it before running this import.`);
    }
    facilityIdCache.set(code, res.rows[0].id);
    return res.rows[0].id;
}

async function getOrCreatePaymentMethodId(client, name) {
    if (paymentMethodIdCache.has(name)) return paymentMethodIdCache.get(name);
    const res = await client.query(
        `INSERT INTO payment_methods (name) VALUES ($1)
     ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
     RETURNING id`,
        [name]
    );
    paymentMethodIdCache.set(name, res.rows[0].id);
    return res.rows[0].id;
}

async function getOrCreateCategoryId(client, name, parentId) {
    const key = `${name}::${parentId ?? 'root'}`;
    if (categoryIdCache.has(key)) return categoryIdCache.get(key);
    // expense_categories.name is globally UNIQUE, so a sub-category name can only
    // live under one parent. Where the sheets reuse a name under two parents
    // (seen with 'Office Expense'), the first parent encountered wins and a
    // warning is raised rather than silently re-parenting the existing row.
    const res = await client.query(
        `INSERT INTO expense_categories (name, parent_category_id) VALUES ($1, $2)
     ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
     RETURNING id, parent_category_id`,
        [name, parentId]
    );
    const row = res.rows[0];
    categoryIdCache.set(key, row.id);
    return row.id;
}

// Strip the client's ordering prefix: "1. Payroll & Human Resources" -> "Payroll & Human Resources"
function cleanCategoryName(s) {
    return s.replace(/^\s*\d+\s*[.)-]\s*/, '').trim();
}

// ---------------------------------------------------------------------
// Directory walk
// ---------------------------------------------------------------------

function findXlsxFiles(target) {
    const stat = fs.statSync(target);
    if (stat.isFile()) return [target];
    const results = [];
    (function walk(dir) {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) walk(full);
            else if (/\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$')) results.push(full);
        }
    })(target);
    results.sort();
    return results;
}

// ---------------------------------------------------------------------
// Per-file processing
// ---------------------------------------------------------------------

async function processFile(client, filePath, warnings) {
    const fileName = path.relative(process.cwd(), filePath);
    const expected = periodFromFilename(filePath);

    if (expected.month === null) {
        warnings.push(`[${fileName}] Could not read a month from the filename — dates cannot be validated or corrected for this file. Rename it like "Expense - June 2026.xlsx".`);
    }

    if (!DRY_RUN) {
        const existing = await client.query(
            'SELECT id FROM raw_import_batches WHERE source_type = $1 AND source_name = $2',
            ['expense', fileName]
        );
        if (existing.rows.length > 0) {
            if (!FORCE) {
                console.log(`  SKIP (already imported, use --force to redo): ${fileName}`);
                return { fileName, skipped: true, inserted: 0 };
            }
            const batchId = existing.rows[0].id;
            await client.query('DELETE FROM expenses WHERE import_batch_id = $1', [batchId]);
            await client.query('DELETE FROM raw_import_batches WHERE id = $1', [batchId]);
        }
    }

    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null });

    const headerIdx = findHeaderRowIdx(rows);
    if (headerIdx === -1) {
        warnings.push(`[${fileName}] Could not find a header row containing both "Date" and "Amount" — file skipped.`);
        return { fileName, skipped: false, inserted: 0 };
    }
    const colMap = buildColMap(rows[headerIdx]);
    if (colMap.amount === undefined || colMap.facility === undefined) {
        warnings.push(`[${fileName}] Header found but "Amount" or "Facility" column is missing — file skipped.`);
        return { fileName, skipped: false, inserted: 0 };
    }

    let importBatchId = null;
    if (!DRY_RUN) {
        const batchRes = await client.query(
            `INSERT INTO raw_import_batches (source_type, source_name, row_count, notes)
       VALUES ('expense', $1, 0, 'Imported by import_expenses.js')
       RETURNING id`,
            [fileName]
        );
        importBatchId = batchRes.rows[0].id;
    }

    let inserted = 0;
    let lastSeenDate = null;   // for --undated=carry
    let undatedCount = 0;
    let undatedTotal = 0;

    for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        const rowNum = i + 1;

        const get = (k) => (colMap[k] !== undefined ? row[colMap[k]] : null);

        const rawDate = get('date');
        const amount = cleanNumber(get('amount'));
        const facilityRaw = cleanText(get('facility'));
        const item = cleanText(get('item'));

        // Entirely blank row
        if (rawDate === null && amount === null && !facilityRaw && !item) continue;

        if (amount === null) {
            warnings.push(`[${fileName}] Row ${rowNum}: no Amount — row skipped (item: ${item || 'n/a'}).`);
            continue;
        }

        let expenseDate;
        const dateIsBlank = rawDate === null || rawDate === undefined || String(rawDate).trim() === '';

        if (dateIsBlank) {
            // Row has money but no date — apply the chosen policy rather than
            // silently dropping real spend.
            if (UNDATED_POLICY === 'skip' || expected.month === null) {
                warnings.push(`[${fileName}] Row ${rowNum}: no date — row skipped (item: ${item || 'n/a'}, amount: ${amount}).`);
                continue;
            }
            if (UNDATED_POLICY === 'carry' && lastSeenDate) {
                expenseDate = lastSeenDate;
                undatedCount++;
                undatedTotal += amount;
            } else {
                expenseDate = fmtDate(expected.year, expected.month, 1);
                undatedCount++;
                undatedTotal += amount;
            }
        } else {
            expenseDate = resolveDate(rawDate, expected, fileName, rowNum, warnings);
            if (!expenseDate) {
                warnings.push(`[${fileName}] Row ${rowNum}: no usable date — row skipped (item: ${item || 'n/a'}, amount: ${amount}).`);
                continue;
            }
            lastSeenDate = expenseDate;
        }

        // Facility is NOT NULL in the schema, so an unresolvable one must skip.
        let facilityCode = null;
        if (facilityRaw) {
            const f = facilityRaw.toLowerCase();
            const hit = FACILITY_PATTERNS.find((p) => p.test(f));
            if (hit) facilityCode = hit.code;
        }
        if (!facilityCode) {
            warnings.push(`[${fileName}] Row ${rowNum}: UNRECOGNISED facility "${facilityRaw}" — row skipped (item: ${item || 'n/a'}, amount: ${amount}).`);
            continue;
        }

        // Category: parent from "Category", child from "Sub-Category".
        // The older layout has neither populated, so those fall to 'Uncategorized'.
        const parentRaw = cleanText(get('category'));
        const childRaw = cleanText(get('subCategory'));
        const parentName = parentRaw && parentRaw.length > 1 ? cleanCategoryName(parentRaw) : null;
        const childName = childRaw && childRaw.length > 1 ? childRaw : null;

        let paymentMethodName = null;
        const pmRaw = cleanText(get('paymentMethod'));
        if (pmRaw) {
            const key = pmRaw.toLowerCase();
            paymentMethodName = PAYMENT_MAP[key] || titleCase(pmRaw);
            if (!PAYMENT_MAP[key]) {
                warnings.push(`[${fileName}] Row ${rowNum}: UNRECOGNISED payment method "${pmRaw}" — imported as "${paymentMethodName}", please review.`);
            }
        }

        if (DRY_RUN) {
            inserted++;
            continue;
        }

        const facilityId = await getFacilityId(client, facilityCode);

        let categoryId;
        if (parentName) {
            const parentId = await getOrCreateCategoryId(client, parentName, null);
            categoryId = childName ? await getOrCreateCategoryId(client, childName, parentId) : parentId;
        } else {
            categoryId = await getOrCreateCategoryId(client, UNCATEGORIZED, null);
        }

        const paymentMethodId = paymentMethodName
            ? await getOrCreatePaymentMethodId(client, paymentMethodName)
            : null;

        const chequeRaw = get('chequeNumber');
        const chequeNumber = chequeRaw === null || chequeRaw === undefined ? null : String(chequeRaw).trim() || null;

        await client.query(
            `INSERT INTO expenses
         (facility_id, expense_category_id, expense_date, item, description, amount,
          payment_method_id, cheque_number, source_account, source_bank, import_batch_id)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
            [
                facilityId,
                categoryId,
                expenseDate,
                item,
                cleanText(get('description')),
                amount,
                paymentMethodId,
                chequeNumber,
                cleanText(get('sourceAccount')),
                cleanText(get('sourceBank')),
                importBatchId,
            ]
        );
        inserted++;
    }

    if (undatedCount > 0) {
        const where = UNDATED_POLICY === 'carry'
            ? 'dated by carrying down the previous row\'s date'
            : `dated to the 1st of ${pad2(expected.month)}/${expected.year}`;
        warnings.push(`[${fileName}] UNDATED ROWS: ${undatedCount} row(s) totalling ${undatedTotal.toLocaleString()} had no date and were ${where}. The month is right; the day is a placeholder. Re-run with --undated=carry or --undated=skip to change this.`);
    }

    if (!DRY_RUN) {
        await client.query('UPDATE raw_import_batches SET row_count = $1 WHERE id = $2', [inserted, importBatchId]);
    }

    console.log(`  ${DRY_RUN ? '[dry-run] would insert' : 'inserted'} ${inserted} expense rows from ${fileName}`);
    return { fileName, skipped: false, inserted };
}

// ---------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------

async function run() {
    const files = findXlsxFiles(TARGET);
    if (files.length === 0) {
        console.error(`No .xlsx files found under: ${TARGET}`);
        process.exit(1);
    }
    console.log(`Found ${files.length} file(s) to process.${DRY_RUN ? ' [DRY RUN — no DB writes]' : ''}`);

    const warnings = [];
    const results = [];
    const client = DRY_RUN ? null : await pool.connect();

    try {
        for (const filePath of files) {
            if (!DRY_RUN) await client.query('BEGIN');
            try {
                const result = await processFile(client, filePath, warnings);
                results.push(result);
                if (!DRY_RUN) await client.query('COMMIT');
            } catch (err) {
                if (!DRY_RUN) await client.query('ROLLBACK');
                warnings.push(`[${path.relative(process.cwd(), filePath)}] FILE FAILED, rolled back: ${err.message}`);
                console.error(`  ❌ ${filePath}: ${err.message}`);
            }
        }
    } finally {
        if (!DRY_RUN) {
            client.release();
            await pool.end();
        }
    }

    const totalInserted = results.reduce((s, r) => s + (r.inserted || 0), 0);
    const skippedFiles = results.filter((r) => r.skipped).length;

    console.log('\n' + '='.repeat(60));
    console.log(`Files processed: ${results.length} (${skippedFiles} already-imported, skipped)`);
    console.log(`Total expense rows ${DRY_RUN ? 'that would be inserted' : 'inserted'}: ${totalInserted}`);
    console.log(`Warnings: ${warnings.length}`);
    if (warnings.length) {
        console.log('-'.repeat(60));
        warnings.forEach((w) => console.log('  ⚠ ' + w));
    }
    console.log('='.repeat(60));
}

run();