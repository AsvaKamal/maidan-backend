// import_bookings.js
// Recursively walks a directory of daily booking Excel files (Sales + Bank Details
// sheets only — all other sheets are legacy/ignored) and loads them into Supabase.
//
// Usage:
//   node import_bookings.js <file-or-directory> [--dry-run] [--force]
//
//   --dry-run   Parse everything and print what WOULD happen. No DB connection,
//               no writes. Safe to run against your whole data/bookings folder.
//   --force     Re-import a file even if it was already recorded in
//               raw_import_batches (deletes the old batch's bookings first).
//
// Safe to re-run per file: each file is one transaction, and a raw_import_batches
// row (source_type='booking') records that the file was processed, so re-running
// the whole folder just skips files already loaded (unless --force).
//
// Filename is NOT used to determine the booking date — the Date column inside
// the Sales sheet is authoritative. Filename (e.g. "5_APR.xlsx") is only used
// for logging and a soft sanity-check warning if it looks wildly inconsistent
// with the dates actually found inside the file.

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { Pool } = require('pg');

const DRY_RUN = process.argv.includes('--dry-run');
const FORCE = process.argv.includes('--force');
const TARGET = process.argv.find((a, i) => i >= 2 && !a.startsWith('--'));

if (!TARGET) {
  console.error('Usage: node import_bookings.js <file-or-directory> [--dry-run] [--force]');
  process.exit(1);
}

const pool = DRY_RUN ? null : new Pool({ connectionString: process.env.DATABASE_URL });

// ---------------------------------------------------------------------
// KNOWN BUSINESS-LOGIC MAPPINGS
// These are judgment calls made from inspecting real Sales sheet data.
// If Maidaan adds a new agent/sport/payment code we don't recognize, the row
// is NOT silently dropped — it's imported with a warning so it can be reviewed
// (see "UNRECOGNIZED" warnings in the run summary), except facility, which is
// a hard requirement (facility_id is NOT NULL) and causes the row to be skipped.
// ---------------------------------------------------------------------

// Agent Name -> facility code (facilities.code must already exist in DB)
const AGENT_FACILITY_MAP = {
  'supervisor-maidan kmc': 'KMC',
  'supervisor-emaar': 'EMAAR',
  'supervisor-maidan malir': 'MALIR',
};

// Income label (after stripping "Per hour Booking- ") -> canonical sport_types.name
// "Padel NEO" and "Padel KMC" are both just Padel — facility is already captured
// via Agent Name, so the KMC/NEO suffix here is redundant, not a different sport.
const SPORT_MAP = {
  'futsal': 'Futsal',
  'padel': 'Padel',
  'padel neo': 'Padel',
  'padel kmc': 'Padel',
  'pickle ball': 'Pickleball',
  'multi purpose': 'Multi Purpose',
  'football': 'Football',
  'cricket': 'Cricket',
};

// BIN column literal text values -> canonical payment_methods.name.
// A numeric BIN always means 'Card' (bank_bin FK is also set in that case).
const PAYMENT_TEXT_MAP = {
  'cash': 'Cash',
  'card': 'Card',
  'ibft': 'IBFT',
  'gtcash': 'GT Cash',
  'gtibft': 'GT IBFT',
  'na': null, // no payment method yet — usually Pending/Cancel rows
};

// Columns we deliberately IGNORE from the Sales sheet and why:
//   Receiveable   - duplicates Amount in practice, no schema field for it
//   Cap           - discount cap threshold, informational only (lives in
//                   bank_bin_reference already), not stored per-booking
//   Revenue / 1st Shift / 2nd Shift - these are NOT per-row transaction data.
//   They're a day-level shift-summary block that happens to sit in the same
//   sheet a few columns to the right, landing on whichever row it visually
//   lines up with. Treating them as row data would attach random totals to
//   unrelated bookings.

// ---------------------------------------------------------------------
// Header + row parsing helpers (dynamic — no hardcoded column indices)
// ---------------------------------------------------------------------

function normHeader(s) {
  return String(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9%]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function findSalesHeaderRowIdx(rows) {
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const norms = (rows[i] || []).map(normHeader);
    if (norms.includes('date') && norms.some((h) => h.includes('customer'))) return i;
  }
  return 0;
}

function buildSalesColMap(headerRow) {
  const map = {};
  headerRow.forEach((raw, idx) => {
    const h = normHeader(raw);
    if (!h) return;
    if (map.date === undefined && h === 'date') { map.date = idx; return; }
    if (map.customerName === undefined && h.includes('customer')) { map.customerName = idx; return; }
    if (map.timings === undefined && h.includes('timing')) { map.timings = idx; return; }
    if (map.status === undefined && h === 'status') { map.status = idx; return; }
    if (map.perHourRate === undefined && h.includes('rate')) { map.perHourRate = idx; return; }
    if (map.hours === undefined && h.includes('hour')) { map.hours = idx; return; }
    if (map.discPct === undefined && h.includes('dis') && h.includes('%')) { map.discPct = idx; return; }
    if (map.discAmount === undefined && h.includes('disc') && (h.includes('amount') || h.includes('smount') || h.includes('amt'))) { map.discAmount = idx; return; }
    if (map.netAmount === undefined && h.includes('net') && (h.includes('amount') || h.includes('smount'))) { map.netAmount = idx; return; }
    if (map.amount === undefined && h === 'amount') { map.amount = idx; return; }
    if (map.agentName === undefined && h.includes('agent')) { map.agentName = idx; return; }
    if (map.income === undefined && h === 'income') { map.income = idx; return; }
    if (map.bin === undefined && h === 'bin') { map.bin = idx; return; }
    if (map.bankName === undefined && h.includes('bank')) { map.bankName = idx; return; }
    if (map.remarks === undefined && h.includes('remark')) { map.remarks = idx; return; }
  });
  return map;
}

function findBankDetailsDataRowIdx(rows) {
  // The real header row is inconsistent between files (sometimes row 1,
  // sometimes row 2, and the two rows can be misaligned with each other —
  // seen in real files). The reliable anchor is: the first data row is the
  // one where column A is an actual number (S.No. = 1, 2, 3...). The header
  // is whatever row sits directly above that.
  return rows.findIndex((r) => typeof r[0] === 'number');
}

function buildBankDetailsColMap(headerRow) {
  const map = {};
  headerRow.forEach((raw, idx) => {
    const h = normHeader(raw);
    if (!h) return;
    if (map.bankName === undefined && h === 'bank') { map.bankName = idx; return; }
    if (map.cardType === undefined && h.includes('card') && h.includes('type')) { map.cardType = idx; return; }
    if (map.bin === undefined && h === 'bin') { map.bin = idx; return; }
    if (map.discountPct === undefined && h.includes('discount') && h.includes('%')) { map.discountPct = idx; return; }
    if (map.discountCap === undefined && h.includes('discount') && h.includes('cap')) { map.discountCap = idx; return; }
    if (map.monthlyLimit === undefined && h.includes('monthly')) { map.monthlyLimit = idx; return; }
    if (map.dailyLimit === undefined && h.includes('daily')) { map.dailyLimit = idx; return; }
  });
  return map;
}

function toDateOnly(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return null;
}

function cleanText(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s.length ? s : null;
}

// Reverse of the xlsx library's serial -> Date conversion (1900 date system,
// including Excel's fictitious 1900-02-29, hence the <= 60 correction).
// 1900-01-01 -> 1, 1900-01-02 -> 2, 1900-01-01T12:00 -> 1.5
function excelSerialFromDate(d) {
  const epoch = Date.UTC(1899, 11, 30);
  const ms = Date.UTC(
    d.getFullYear(), d.getMonth(), d.getDate(),
    d.getHours(), d.getMinutes(), d.getSeconds()
  );
  let serial = (ms - epoch) / 86400000;
  if (serial <= 60) serial -= 1;
  return serial;
}

// Numeric cells in these sheets are sometimes DATE-FORMATTED by accident: a
// supervisor types "1" into a cell Excel has formatted as a date, so Excel
// stores 1 but renders it "1-Jan" and the xlsx library hands us a Date object.
// Number(Date) would yield epoch milliseconds (~-2.2e12), which overflows the
// numeric column and rolls back the WHOLE file. Converting back to the Excel
// serial recovers the number the person actually typed.
function cleanNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    const serial = excelSerialFromDate(v);
    return Number.isFinite(serial) ? serial : null;
  }
  const n = Number(v);
  return Number.isNaN(n) ? null : n;
}

// ---------------------------------------------------------------------
// Directory walk
// ---------------------------------------------------------------------

function findXlsxFiles(target) {
  const stat = fs.statSync(target);
  if (stat.isFile()) return [target];

  const results = [];
  function walk(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(full);
      } else if (/\.xlsx$/i.test(entry.name) && !entry.name.startsWith('~$')) {
        results.push(full);
      }
    }
  }
  walk(target);
  results.sort();
  return results;
}

// ---------------------------------------------------------------------
// Per-lookup ID caches (populated per DB run; unused in dry-run)
// ---------------------------------------------------------------------

const facilityIdCache = new Map();
const sportTypeIdCache = new Map();
const paymentMethodIdCache = new Map();
let bankBinRefSet = new Set(); // BINs currently in bank_bin_reference

async function getFacilityId(client, code) {
  if (facilityIdCache.has(code)) return facilityIdCache.get(code);
  const res = await client.query('SELECT id FROM facilities WHERE code = $1', [code]);
  if (res.rows.length === 0) {
    throw new Error(`No row in "facilities" with code = '${code}'. Create it before running this import.`);
  }
  facilityIdCache.set(code, res.rows[0].id);
  return res.rows[0].id;
}

async function getOrCreateSportTypeId(client, name) {
  if (sportTypeIdCache.has(name)) return sportTypeIdCache.get(name);
  const res = await client.query(
    `INSERT INTO sport_types (name) VALUES ($1)
     ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
     RETURNING id`,
    [name]
  );
  sportTypeIdCache.set(name, res.rows[0].id);
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

// ---------------------------------------------------------------------
// Bank Details -> bank_bin_reference upsert
// ---------------------------------------------------------------------

async function upsertBankDetails(client, workbook, warnings) {
  const sheet = workbook.Sheets['Bank Details'];
  if (!sheet) {
    warnings.push('No "Bank Details" sheet found — BIN lookups may fail for numeric BINs in this file.');
    return;
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const dataStart = findBankDetailsDataRowIdx(rows);
  if (dataStart < 1) {
    warnings.push('Could not locate a data row in "Bank Details" (expected a numeric S.No. in column A). Skipping bank details for this file.');
    return;
  }
  const colMap = buildBankDetailsColMap(rows[dataStart - 1]);
  if (colMap.bin === undefined || colMap.bankName === undefined) {
    warnings.push('Could not confidently map BIN/Bank columns in "Bank Details". Skipping bank details for this file.');
    return;
  }

  let count = 0;
  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const binNumber = cleanNumber(row[colMap.bin]);
    const bankName = cleanText(row[colMap.bankName]);
    if (binNumber === null || !bankName) continue;

    if (!DRY_RUN) {
      await client.query(
        `INSERT INTO bank_bin_reference
           (bin_number, bank_name, card_type, discount_pct, discount_cap, monthly_txn_limit, daily_txn_limit)
         VALUES ($1, $2, $3, $4, $5, $6, $7)
         ON CONFLICT (bin_number) DO UPDATE SET
           bank_name = EXCLUDED.bank_name,
           card_type = EXCLUDED.card_type,
           discount_pct = EXCLUDED.discount_pct,
           discount_cap = EXCLUDED.discount_cap,
           monthly_txn_limit = EXCLUDED.monthly_txn_limit,
           daily_txn_limit = EXCLUDED.daily_txn_limit,
           updated_at = now()`,
        [
          binNumber,
          bankName,
          cleanText(row[colMap.cardType]),
          cleanNumber(colMap.discountPct !== undefined ? row[colMap.discountPct] : null),
          cleanNumber(colMap.discountCap !== undefined ? row[colMap.discountCap] : null),
          cleanNumber(colMap.monthlyLimit !== undefined ? row[colMap.monthlyLimit] : null),
          cleanNumber(colMap.dailyLimit !== undefined ? row[colMap.dailyLimit] : null),
        ]
      );
    }
    bankBinRefSet.add(binNumber);
    count++;
  }
  return count;
}

// ---------------------------------------------------------------------
// Sales -> bookings
// ---------------------------------------------------------------------

async function processSales(client, workbook, importBatchId, fileName, warnings, stats) {
  const sheet = workbook.Sheets['Sales'];
  if (!sheet) {
    warnings.push(`[${fileName}] No "Sales" sheet found — nothing imported from this file.`);
    return 0;
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const headerIdx = findSalesHeaderRowIdx(rows);
  const colMap = buildSalesColMap(rows[headerIdx]);

  if (colMap.date === undefined || colMap.agentName === undefined) {
    warnings.push(`[${fileName}] Could not confidently map Date/Agent Name columns in "Sales". Skipping file.`);
    return 0;
  }

  let inserted = 0;
  const seenDates = new Set();

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const date = row[colMap.date];
    const agentRaw = colMap.agentName !== undefined ? row[colMap.agentName] : null;
    const customerRaw = colMap.customerName !== undefined ? row[colMap.customerName] : null;

    // Blank separator rows between sport sections
    if (date === null && customerRaw === null && agentRaw === null) continue;

    const bookingDate = toDateOnly(date);
    const agentName = cleanText(agentRaw);
    const facilityCode = agentName ? AGENT_FACILITY_MAP[agentName.trim().toLowerCase()] : null;

    if (!bookingDate) {
      warnings.push(`[${fileName}] Row ${i + 1}: missing/unparseable Date — skipped.`);
      continue;
    }
    if (bookingDate) seenDates.add(bookingDate);

    if (!facilityCode) {
      warnings.push(`[${fileName}] Row ${i + 1}: UNRECOGNIZED agent name "${agentName}" — no facility mapping, row skipped.`);
      continue;
    }

    const incomeRaw = cleanText(colMap.income !== undefined ? row[colMap.income] : null);
    const sportKey = incomeRaw ? incomeRaw.replace(/^per hour booking-\s*/i, '').trim().toLowerCase() : '';
    let sportName = SPORT_MAP[sportKey];
    if (!sportName) {
      sportName = incomeRaw || 'Unknown';
      warnings.push(`[${fileName}] Row ${i + 1}: UNRECOGNIZED income/sport label "${incomeRaw}" — imported as sport_types."${sportName}", please review.`);
    }

    const hoursRaw = colMap.hours !== undefined ? row[colMap.hours] : null;
    const grossRaw = colMap.amount !== undefined ? row[colMap.amount] : null;
    const netRaw = colMap.netAmount !== undefined ? row[colMap.netAmount] : null;

    const hours = cleanNumber(hoursRaw);
    const grossAmount = cleanNumber(grossRaw);
    const netAmount = cleanNumber(netRaw);

    // Flag (but still import) any numeric cell that arrived date-formatted, so
    // the recovered value can be eyeballed against the source sheet.
    for (const [label, raw, val] of [
      ['No of Hours', hoursRaw, hours],
      ['Amount', grossRaw, grossAmount],
      ['Net amount', netRaw, netAmount],
    ]) {
      if (raw instanceof Date) {
        warnings.push(`[${fileName}] Row ${i + 1}: "${label}" was DATE-FORMATTED in Excel (shows as a date, not a number) — recovered underlying value ${val}. Please verify against the sheet.`);
      }
    }

    if (hours === null || grossAmount === null || netAmount === null) {
      warnings.push(`[${fileName}] Row ${i + 1}: missing hours/amount/net amount — row skipped (customer: ${cleanText(customerRaw) || 'n/a'}).`);
      continue;
    }

    // Guard against absurd values reaching the DB and rolling back a whole file.
    // A single booking can't plausibly run 24+ hours; anything past that is a
    // data-entry artifact, so skip the row loudly rather than kill the file.
    if (hours <= 0 || hours > 24) {
      warnings.push(`[${fileName}] Row ${i + 1}: implausible duration_hours (${hours}) — row skipped (customer: ${cleanText(customerRaw) || 'n/a'}, status: ${cleanText(colMap.status !== undefined ? row[colMap.status] : null) || 'n/a'}).`);
      continue;
    }

    // BIN column: numeric -> Card + bank_bin FK; text -> mapped payment method
    const binRaw = colMap.bin !== undefined ? row[colMap.bin] : null;
    const bankNameRaw = colMap.bankName !== undefined ? row[colMap.bankName] : null;
    let bankBin = null;
    let paymentMethodName = null;

    if (typeof binRaw === 'number') {
      bankBin = binRaw;
      paymentMethodName = 'Card';
      if (!bankBinRefSet.has(binRaw)) {
        warnings.push(`[${fileName}] Row ${i + 1}: BIN ${binRaw} not found in this file's Bank Details reference — bank_bin will be left blank.`);
        bankBin = null;
      }
    } else {
      const textKey = (cleanText(binRaw) || cleanText(bankNameRaw) || '').toLowerCase();
      if (textKey && Object.prototype.hasOwnProperty.call(PAYMENT_TEXT_MAP, textKey)) {
        paymentMethodName = PAYMENT_TEXT_MAP[textKey];
      } else if (textKey) {
        paymentMethodName = cleanText(binRaw) || cleanText(bankNameRaw);
        warnings.push(`[${fileName}] Row ${i + 1}: UNRECOGNIZED payment code "${textKey}" — imported as payment_methods."${paymentMethodName}", please review.`);
      }
    }

    if (DRY_RUN) {
      inserted++;
      continue;
    }

    const facilityId = await getFacilityId(client, facilityCode);
    const sportTypeId = await getOrCreateSportTypeId(client, sportName);
    const paymentMethodId = paymentMethodName ? await getOrCreatePaymentMethodId(client, paymentMethodName) : null;

    await client.query(
      `INSERT INTO bookings
         (facility_id, sport_type_id, booking_date, time_slot_raw, duration_hours, customer_name,
          per_hour_rate, gross_amount, discount_pct, discount_amount, net_amount,
          payment_method_id, bank_bin, status, remarks, agent_name_raw, income_raw, import_batch_id)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)`,
      [
        facilityId,
        sportTypeId,
        bookingDate,
        cleanText(colMap.timings !== undefined ? row[colMap.timings] : null),
        hours,
        cleanText(customerRaw),
        cleanNumber(colMap.perHourRate !== undefined ? row[colMap.perHourRate] : null),
        grossAmount,
        cleanNumber(colMap.discPct !== undefined ? row[colMap.discPct] : null) ?? 0,
        cleanNumber(colMap.discAmount !== undefined ? row[colMap.discAmount] : null) ?? 0,
        netAmount,
        paymentMethodId,
        bankBin,
        cleanText(colMap.status !== undefined ? row[colMap.status] : null) || 'Unknown',
        cleanText(colMap.remarks !== undefined ? row[colMap.remarks] : null),
        agentName,
        incomeRaw,
        importBatchId,
      ]
    );
    inserted++;
  }

  if (seenDates.size > 1) {
    warnings.push(`[${fileName}] Sales rows span ${seenDates.size} distinct dates (${[...seenDates].join(', ')}) — expected a single-day file.`);
  } else if (seenDates.size === 1) {
    const [onlyDate] = [...seenDates];
    const base = path.basename(fileName, path.extname(fileName));
    const m = base.match(/^(\d{1,2})_([A-Za-z]{3,})/);
    if (m) {
      const day = parseInt(m[1], 10);
      const monthName = m[2].toLowerCase();
      const cellMonth = new Date(onlyDate + 'T00:00:00').toLocaleString('en-US', { month: 'short' }).toLowerCase();
      const cellDay = new Date(onlyDate + 'T00:00:00').getDate();
      if (day !== cellDay || !cellMonth.startsWith(monthName.slice(0, 3))) {
        warnings.push(`[${fileName}] Filename suggests ${m[1]}-${m[2]}, but Sales data is dated ${onlyDate} — double check this file is filed correctly.`);
      }
    }
  }

  return inserted;
}

// ---------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------

async function processFile(client, filePath, warnings) {
  const fileName = path.relative(process.cwd(), filePath);
  const workbook = XLSX.readFile(filePath, { cellDates: true });

  if (!DRY_RUN) {
    const existing = await client.query(
      'SELECT id FROM raw_import_batches WHERE source_type = $1 AND source_name = $2',
      ['booking', fileName]
    );
    if (existing.rows.length > 0) {
      if (!FORCE) {
        console.log(`  SKIP (already imported, use --force to redo): ${fileName}`);
        return { fileName, skipped: true, inserted: 0 };
      }
      const batchId = existing.rows[0].id;
      await client.query('DELETE FROM bookings WHERE import_batch_id = $1', [batchId]);
      await client.query('DELETE FROM raw_import_batches WHERE id = $1', [batchId]);
    }
  }

  await upsertBankDetails(client, workbook, warnings);

  let importBatchId = null;
  if (!DRY_RUN) {
    const batchRes = await client.query(
      `INSERT INTO raw_import_batches (source_type, source_name, row_count, notes)
       VALUES ('booking', $1, 0, 'Imported by import_bookings.js')
       RETURNING id`,
      [fileName]
    );
    importBatchId = batchRes.rows[0].id;
  }

  const inserted = await processSales(client, workbook, importBatchId, fileName, warnings, {});

  if (!DRY_RUN) {
    await client.query('UPDATE raw_import_batches SET row_count = $1 WHERE id = $2', [inserted, importBatchId]);
  }

  console.log(`  ${DRY_RUN ? '[dry-run] would insert' : 'inserted'} ${inserted} booking rows from ${fileName}`);
  return { fileName, skipped: false, inserted };
}

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
      // Each file gets its own in-memory BIN set scoped to that file's Bank Details tab
      bankBinRefSet = new Set();
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

  const totalInserted = results.reduce((sum, r) => sum + (r.inserted || 0), 0);
  const skippedFiles = results.filter((r) => r.skipped).length;

  console.log('\n' + '='.repeat(60));
  console.log(`Files processed: ${results.length} (${skippedFiles} already-imported, skipped)`);
  console.log(`Total booking rows ${DRY_RUN ? 'that would be inserted' : 'inserted'}: ${totalInserted}`);
  console.log(`Warnings: ${warnings.length}`);
  if (warnings.length) {
    console.log('-'.repeat(60));
    warnings.forEach((w) => console.log('  ⚠ ' + w));
  }
  console.log('='.repeat(60));
}

run();