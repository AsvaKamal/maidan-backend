-- =====================================================================
-- SAMPLE DATA — for testing the schema manually via SQL Editor
-- Run these AFTER maidan_schema.sql has already been executed once.
-- =====================================================================

-- ---------------------------------------------------------------------
-- 1. INVESTORS
-- ---------------------------------------------------------------------
INSERT INTO investors (name, contact_info) VALUES
    ('Taha Bin Nadeem', 'taha@example.com'),
    ('Investor Two', NULL)
RETURNING id, name;  -- note the returned ids, you'll need them below


-- ---------------------------------------------------------------------
-- 2. FACILITY INVESTORS (ownership %) — using facility 'KMC'
-- Replace investor_id values with the actual ids returned above.
-- ---------------------------------------------------------------------
INSERT INTO facility_investors (facility_id, investor_id, investment_amount, ownership_percentage, effective_from)
VALUES
    ((SELECT id FROM facilities WHERE code = 'KMC'), 1, 500000.00, 0.25, '2024-08-01'),
    ((SELECT id FROM facilities WHERE code = 'KMC'), 2, 1500000.00, 0.75, '2024-08-01');


-- ---------------------------------------------------------------------
-- 3. EXPENSE CATEGORY (add one if it's not already seeded)
-- ---------------------------------------------------------------------
INSERT INTO expense_categories (name) VALUES
    ('Salary'), ('Internet'), ('Refreshment'), ('Investor')
ON CONFLICT (name) DO NOTHING;


-- ---------------------------------------------------------------------
-- 4. A SAMPLE EXPENSE ROW
-- ---------------------------------------------------------------------
INSERT INTO expenses (
    facility_id, expense_category_id, expense_date, item, description,
    amount, payment_method_id, source_account, source_bank
) VALUES (
    (SELECT id FROM facilities WHERE code = 'KMC'),
    (SELECT id FROM expense_categories WHERE name = 'Internet'),
    '2026-05-10',
    'Monthly internet bill',
    'PTCL fiber connection',
    3500.00,
    (SELECT id FROM payment_methods WHERE name = 'Cash'),
    NULL, NULL
);


-- ---------------------------------------------------------------------
-- 5. A SAMPLE BOOKING ROW
-- ---------------------------------------------------------------------
INSERT INTO bookings (
    facility_id, sport_type_id, booking_date, time_slot_raw, duration_hours,
    customer_name, per_hour_rate, gross_amount, net_amount,
    payment_method_id, status, agent_name_raw, income_raw
) VALUES (
    (SELECT id FROM facilities WHERE code = 'KMC'),
    (SELECT id FROM sport_types WHERE name = 'Futsal'),
    '2026-05-10',
    '7pm to 9pm',
    2.0,
    'Ali Raza',
    2000.00,
    4000.00,
    4000.00,
    (SELECT id FROM payment_methods WHERE name = 'Cash'),
    'Done',
    'Supervisor-Maidan KMC',
    'Per hour Booking- Futsal'
);


-- ---------------------------------------------------------------------
-- 6. MONTHLY PROFIT + DISBURSEMENT (the calculation you asked about)
-- ---------------------------------------------------------------------
INSERT INTO monthly_profit (facility_id, period_month, profit_amount, entered_by)
VALUES (
    (SELECT id FROM facilities WHERE code = 'KMC'),
    '2026-05-01',
    250000.00,
    'Anaya'
)
RETURNING id;  -- note this id, used below as monthly_profit_id

-- Now generate disbursements for every active investor of that facility,
-- for that month, in one shot:
INSERT INTO disbursements (facility_investor_id, monthly_profit_id, period_month, ownership_pct_used, disbursed_amount)
SELECT
    fi.id,
    mp.id,
    mp.period_month,
    fi.ownership_percentage,
    mp.profit_amount * fi.ownership_percentage
FROM facility_investors fi
JOIN monthly_profit mp ON mp.facility_id = fi.facility_id
WHERE mp.facility_id = (SELECT id FROM facilities WHERE code = 'KMC')
  AND mp.period_month = '2026-05-01'
  AND fi.effective_to IS NULL;  -- only currently-active ownership rows
