// import_investors.js
// One-time backfill: reads Maidan_Investor_Disbursements.xlsx and inserts
// investors, facility_investors, monthly_profit, and disbursements into Supabase.
// Safe to re-run — uses ON CONFLICT so re-running won't create duplicates.

require('dotenv').config();
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { Pool } = require('pg');

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

// ---------------------------------------------------------------------
// CONFIG — one entry per facility tab, hand-mapped from the real file.
// Row/column indices are 0-indexed, matching XLSX's sheet_to_json({header:1}) output.
// ---------------------------------------------------------------------
const FACILITY_CONFIGS = [
  {
    facilityCode: 'KMC',
    sheetName: 'KMC',
    cols: { month: 0, year: 1, amount: 2 },
    investmentRow: 1,
    percentageRow: 2,
    dataStartRow: 3,
    dataEndRow: 31, // exclusive — row 31 is blank, before the Totals section
    investors: [
      { name: 'Iqbal', amountCol: 3, statusCol: 4 },
      { name: 'Haris Aslam', amountCol: 5, statusCol: 6 },
      { name: 'Muzammil', amountCol: 7, statusCol: 8 },
      { name: 'Umair', amountCol: 9, statusCol: 10 },
      { name: 'Haris Lakhani', amountCol: 11, statusCol: 12 },
      { name: 'Usman Ladha', amountCol: 13, statusCol: 14 },
      { name: 'Maidan', amountCol: 15, statusCol: 16 },
    ],
  },
  {
    facilityCode: 'EMAAR',
    sheetName: 'Emaar',
    cols: { month: 0, year: 1, amount: 2 },
    investmentRow: 1,
    percentageRow: 2,
    dataStartRow: 3,
    dataEndRow: 20,
    investors: [
      { name: 'Ahsan', amountCol: 3, statusCol: 4 },
      { name: 'Haris', amountCol: 5, statusCol: 6 },
      { name: 'Abdullah', amountCol: 7, statusCol: 8 },
      { name: 'Suleman', amountCol: 9, statusCol: 10 },
    ],
  },
  {
    facilityCode: 'MALIR',
    sheetName: 'Malir',
    cols: { month: 0, year: 1, amount: 2 },
    investmentRow: 1,
    percentageRow: 2,
    dataStartRow: 3,
    dataEndRow: 20,
    investors: [
      { name: 'STA', amountCol: 3, statusCol: 4 },
      { name: 'UC', amountCol: 5, statusCol: 6 },
      { name: 'Entertainer Asia', amountCol: 7, statusCol: 8 },
      { name: 'MC', amountCol: 9, statusCol: 10 },
    ],
  },
];

const MONTH_NUMBERS = {
  january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
  july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
};

// Combined-month rows (e.g. "August/September") are filed under the SECOND
// month, per the decision made — original label is preserved in the notes field.
function parseMonthLabel(rawLabel) {
  const isCombined = rawLabel.includes('/');
  const parts = rawLabel.split('/').map((s) => s.trim());
  const monthToUse = parts[parts.length - 1]; // second month if combined
  return { monthName: monthToUse, isCombined, rawLabel };
}

// Known data-fix: KMC's "January 2024" row is a confirmed typo for January 2025
// (it appears out of sequence, right after December 2024).
function applyKnownFixes(facilityCode, monthName, year) {
  if (facilityCode === 'KMC' && monthName === 'January' && year === 2024) {
    return 2025;
  }
  return year;
}

function toPeriodDate(monthName, year) {
  const monthNum = MONTH_NUMBERS[monthName.toLowerCase()];
  if (!monthNum) throw new Error(`Unrecognized month name: ${monthName}`);
  return `${year}-${String(monthNum).padStart(2, '0')}-01`;
}

// Resolve the workbook path. Override with XLSX_PATH=/some/path node import_investors.js
// if your file isn't at the default location — the previous hardcoded path used spaces
// and a '../data/' folder that didn't match the actual filename (underscores, no spaces).
const XLSX_PATH = process.env.XLSX_PATH || path.join(__dirname, '..', 'data', 'Maidan Investor Disbursements.xlsx');

async function run() {
  if (!fs.existsSync(XLSX_PATH)) {
    console.error(`❌ Workbook not found at: ${XLSX_PATH}`);
    console.error(`   Set XLSX_PATH env var to the correct location, or place the file next to this script.`);
    process.exit(1);
  }
  console.log(`Reading workbook: ${XLSX_PATH}`);
  const workbook = XLSX.readFile(XLSX_PATH);
  const client = await pool.connect();

  try {
    await client.query('BEGIN');

    for (const config of FACILITY_CONFIGS) {
      console.log(`\n=== Processing ${config.facilityCode} ===`);
      const sheet = workbook.Sheets[config.sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

      // Get (or confirm) the facility row
      const facilityRes = await client.query(
        'SELECT id FROM facilities WHERE code = $1',
        [config.facilityCode]
      );
      if (facilityRes.rows.length === 0) {
        throw new Error(`No row in "facilities" with code = '${config.facilityCode}'. Create it before running this import.`);
      }
      const facilityId = facilityRes.rows[0].id;

      const investmentRowData = rows[config.investmentRow];
      const percentageRowData = rows[config.percentageRow];

      // -------------------------------------------------------------
      // Step 1: upsert investors + facility_investors (ownership %)
      // -------------------------------------------------------------
      const facilityInvestorIds = {}; // investor name -> facility_investors.id

      for (const inv of config.investors) {
        const investmentAmount = investmentRowData[inv.amountCol];
        const ownershipPct = percentageRowData[inv.amountCol];

        if (ownershipPct === null || ownershipPct === undefined) continue; // no data for this investor here

        // Upsert investor
        const investorRes = await client.query(
          `INSERT INTO investors (name) VALUES ($1)
           ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
           RETURNING id`,
          [inv.name]
        );
        const investorId = investorRes.rows[0].id;

        // Upsert facility_investors
        const fiRes = await client.query(
          `INSERT INTO facility_investors (facility_id, investor_id, investment_amount, ownership_percentage, effective_from)
           VALUES ($1, $2, $3, $4, '2024-08-01')
           ON CONFLICT (facility_id, investor_id) WHERE effective_to IS NULL DO NOTHING
           RETURNING id`,
          [facilityId, investorId, investmentAmount || null, ownershipPct]
        );

        let facilityInvestorId;
        if (fiRes.rows.length > 0) {
          facilityInvestorId = fiRes.rows[0].id;
        } else {
          const existing = await client.query(
            `SELECT id FROM facility_investors WHERE facility_id = $1 AND investor_id = $2 AND effective_to IS NULL`,
            [facilityId, investorId]
          );
          facilityInvestorId = existing.rows[0].id;
        }

        facilityInvestorIds[inv.name] = facilityInvestorId;
      }

      // -------------------------------------------------------------
      // Step 2: monthly_profit + disbursements, row by row
      // -------------------------------------------------------------
      for (let i = config.dataStartRow; i < config.dataEndRow; i++) {
        const row = rows[i];
        if (!row || row[config.cols.month] === null) continue;

        const amount = row[config.cols.amount];
        if (amount === null || amount === undefined) continue; // future/unfilled month — skip

        const rawMonthLabel = row[config.cols.month];
        const rawYear = row[config.cols.year];
        const { monthName, isCombined, rawLabel } = parseMonthLabel(String(rawMonthLabel));
        const year = applyKnownFixes(config.facilityCode, monthName, rawYear);
        const periodMonth = toPeriodDate(monthName, year);

        const notes = isCombined ? `Original label in source sheet: "${rawLabel}"` : null;

        const mpRes = await client.query(
          `INSERT INTO monthly_profit (facility_id, period_month, profit_amount, entered_by, notes)
           VALUES ($1, $2, $3, $4, $5)
           ON CONFLICT (facility_id, period_month)
           DO UPDATE SET profit_amount = EXCLUDED.profit_amount, notes = EXCLUDED.notes
           RETURNING id`,
          [facilityId, periodMonth, amount, 'Migrated from historical records', notes]
        );
        const monthlyProfitId = mpRes.rows[0].id;

        for (const inv of config.investors) {
          const facilityInvestorId = facilityInvestorIds[inv.name];
          if (!facilityInvestorId) continue;

          const disbursedAmount = row[inv.amountCol];
          const status = row[inv.statusCol] || 'Pending';
          if (disbursedAmount === null || disbursedAmount === undefined) continue;

          // Look up the ownership % actually used, for the snapshot field
          const pctRes = await client.query(
            'SELECT ownership_percentage FROM facility_investors WHERE id = $1',
            [facilityInvestorId]
          );
          const pctUsed = pctRes.rows[0].ownership_percentage;

          await client.query(
            `INSERT INTO disbursements (facility_investor_id, monthly_profit_id, period_month, ownership_pct_used, disbursed_amount, status)
             VALUES ($1, $2, $3, $4, $5, $6)
             ON CONFLICT (facility_investor_id, period_month)
             DO UPDATE SET disbursed_amount = EXCLUDED.disbursed_amount, status = EXCLUDED.status`,
            [facilityInvestorId, monthlyProfitId, periodMonth, pctUsed, disbursedAmount, status]
          );
        }

        console.log(`  Inserted ${periodMonth}${notes ? ' (combined month, see notes)' : ''}`);
      }
    }

    await client.query('COMMIT');
    console.log('\n✅ Import complete.');
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('❌ Import failed, rolled back:', err);
  } finally {
    client.release();
    await pool.end();
  }
}

run();