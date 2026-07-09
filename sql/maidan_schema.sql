-- =====================================================================
-- MAIDAN SPORTS COMPLEX — REPORTING & DASHBOARD DATABASE SCHEMA
-- Target: Supabase (Postgres)
-- Domains covered: Investor Disbursements | Expenses | Bookings/Revenue
-- =====================================================================


-- =====================================================================
-- 1. CORE REFERENCE TABLES
-- =====================================================================

-- Facilities: KMC, Emaar, Malir Cantt
CREATE TABLE facilities (
    id              SERIAL PRIMARY KEY,
    code            TEXT NOT NULL UNIQUE,      -- e.g. 'KMC', 'EMAAR', 'MALIR'
    name            TEXT NOT NULL,             -- e.g. 'Maidan KMC'
    location        TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Sport types: Futsal, Padel (extensible if new sports get added later)
CREATE TABLE sport_types (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL UNIQUE       -- 'Futsal', 'Padel'
);

-- Expense categories: kept as a lookup table (not free text) so dashboard
-- filters/dropdowns stay clean. Add new categories via INSERT any time.
CREATE TABLE expense_categories (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL UNIQUE       -- 'Salary', 'Investor', 'Internet', etc.
);

-- Payment methods used across bookings + expenses
CREATE TABLE payment_methods (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL UNIQUE       -- 'Cash', 'Card', 'IBFT', 'Cheque', 'NA'
);

-- Bank BIN lookup table — static reference data from the "Bank Details" tab.
-- Seeded once, updated only when a new card program is added.
CREATE TABLE bank_bin_reference (
    bin_number              BIGINT PRIMARY KEY,
    bank_name               TEXT NOT NULL,
    card_type               TEXT,
    discount_pct            NUMERIC(5,4),      -- e.g. 0.4000 = 40%
    discount_cap            NUMERIC(12,2),
    monthly_txn_limit       INTEGER,
    daily_txn_limit         INTEGER,
    bank_discount_share     NUMERIC(5,4),      -- portion of discount borne by bank
    updated_at              TIMESTAMPTZ NOT NULL DEFAULT now()
);


-- =====================================================================
-- 2. BOOKINGS / REVENUE
-- =====================================================================

CREATE TABLE bookings (
    id                  BIGSERIAL PRIMARY KEY,
    facility_id         INTEGER NOT NULL REFERENCES facilities(id),
    sport_type_id       INTEGER NOT NULL REFERENCES sport_types(id),

    booking_date        DATE NOT NULL,
    time_slot_raw        TEXT,                 -- original text e.g. '7 am to 10 am'
    duration_hours       NUMERIC(5,2) NOT NULL,

    customer_name        TEXT,

    per_hour_rate         NUMERIC(12,2),
    gross_amount          NUMERIC(12,2) NOT NULL,   -- Amount before discount
    discount_pct          NUMERIC(6,4) DEFAULT 0,
    discount_amount        NUMERIC(12,2) DEFAULT 0,
    net_amount             NUMERIC(12,2) NOT NULL,   -- Amount actually received (= Revenue)

    payment_method_id      INTEGER REFERENCES payment_methods(id),
    bank_bin               BIGINT REFERENCES bank_bin_reference(bin_number),

    status                 TEXT NOT NULL,      -- 'Done', 'Cancel', 'Academy', etc.
    remarks                TEXT,

    -- Raw/original fields kept for traceability — if our parsing logic
    -- ever misreads a facility or sport type, you can see exactly what
    -- it parsed from and fix it without losing the source data.
    agent_name_raw          TEXT,
    income_raw              TEXT,

    import_batch_id          BIGINT,           -- links to raw_import_batches
    created_at               TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_bookings_facility_date ON bookings(facility_id, booking_date);
CREATE INDEX idx_bookings_date ON bookings(booking_date);


-- =====================================================================
-- 3. EXPENSES
-- =====================================================================

CREATE TABLE expenses (
    id                  BIGSERIAL PRIMARY KEY,
    facility_id          INTEGER NOT NULL REFERENCES facilities(id),
    expense_category_id  INTEGER NOT NULL REFERENCES expense_categories(id),

    expense_date          DATE NOT NULL,
    item                  TEXT,
    description           TEXT,

    amount                NUMERIC(12,2) NOT NULL,

    payment_method_id      INTEGER REFERENCES payment_methods(id),
    cheque_number           TEXT,
    source_account          TEXT,
    source_bank             TEXT,

    import_batch_id          BIGINT,
    created_at                TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_expenses_facility_date ON expenses(facility_id, expense_date);
CREATE INDEX idx_expenses_category ON expenses(expense_category_id);


-- =====================================================================
-- 4. INVESTORS & DISBURSEMENTS
-- =====================================================================

CREATE TABLE investors (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL,
    contact_info    TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Ownership % per facility, with date-ranged validity.
-- Handles BOTH cases without needing to know the answer up front:
--   - Fixed forever  -> just one row per investor/facility, effective_to = NULL
--   - Changes later  -> close old row (set effective_to), insert new row
CREATE TABLE facility_investors (
    id                      SERIAL PRIMARY KEY,
    facility_id             INTEGER NOT NULL REFERENCES facilities(id),
    investor_id             INTEGER NOT NULL REFERENCES investors(id),
    investment_amount       NUMERIC(14,2),
    ownership_percentage    NUMERIC(6,4) NOT NULL,   -- e.g. 0.2500 = 25%
    effective_from          DATE NOT NULL DEFAULT '2024-01-01',
    effective_to            DATE,                     -- NULL = currently active
    created_at              TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_facility_investors_active
    ON facility_investors(facility_id, investor_id)
    WHERE effective_to IS NULL;

-- Monthly profit pool per facility — manually entered for now (Option A).
-- Later this can be cross-checked against SUM(bookings.net_amount) - SUM(expenses.amount)
-- for that facility/month as a reconciliation step, without changing this table.
CREATE TABLE monthly_profit (
    id                  SERIAL PRIMARY KEY,
    facility_id         INTEGER NOT NULL REFERENCES facilities(id),
    period_month        DATE NOT NULL,            -- store as first-of-month, e.g. 2026-05-01
    profit_amount       NUMERIC(14,2) NOT NULL,
    entered_by           TEXT,
    notes                 TEXT,
    created_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE(facility_id, period_month)
);

-- Actual disbursement per investor per month.
-- disbursed_amount is a SNAPSHOT (computed once, stored) — so a later change
-- to ownership % never silently rewrites a past month's disbursed amount.
CREATE TABLE disbursements (
    id                       BIGSERIAL PRIMARY KEY,
    facility_investor_id     INTEGER NOT NULL REFERENCES facility_investors(id),
    monthly_profit_id        INTEGER NOT NULL REFERENCES monthly_profit(id),

    period_month             DATE NOT NULL,
    ownership_pct_used       NUMERIC(6,4) NOT NULL,   -- % actually applied, snapshot
    disbursed_amount         NUMERIC(14,2) NOT NULL,

    status                   TEXT NOT NULL DEFAULT 'Pending', -- 'Pending' / 'Disbursed'
    disbursed_date            DATE,

    created_at                 TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE(facility_investor_id, period_month)
);

CREATE INDEX idx_disbursements_period ON disbursements(period_month);
CREATE INDEX idx_disbursements_status ON disbursements(status);


-- =====================================================================
-- 5. IMPORT TRACKING (for the Google Sheets sync pipeline)
-- =====================================================================

-- Every sync run logs itself here — lets the sync script avoid
-- re-importing the same rows twice (idempotency), and gives you an
-- audit trail of when data last landed.
CREATE TABLE raw_import_batches (
    id              BIGSERIAL PRIMARY KEY,
    source_type     TEXT NOT NULL,     -- 'booking', 'expense', 'investor'
    source_name     TEXT,              -- sheet/file name or tab name
    row_count       INTEGER,
    imported_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
    notes           TEXT
);


-- =====================================================================
-- 6. SEED DATA — facilities (safe to run once)
-- =====================================================================

INSERT INTO facilities (code, name) VALUES
    ('KMC', 'Maidan KMC'),
    ('EMAAR', 'Maidan Emaar'),
    ('MALIR', 'Maidan Malir Cantt')
ON CONFLICT (code) DO NOTHING;

INSERT INTO sport_types (name) VALUES
    ('Futsal'),
    ('Padel')
ON CONFLICT (name) DO NOTHING;

INSERT INTO payment_methods (name) VALUES
    ('Cash'), ('Card'), ('IBFT'), ('Cheque'), ('NA')
ON CONFLICT (name) DO NOTHING;
