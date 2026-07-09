-- =====================================================================
-- SCHEMA UPDATE v2 — based on Taha's answers
-- Run this AFTER maidan_schema.sql has already been executed.
-- Safe to run on a table that already has data (all additive changes).
-- =====================================================================

-- ---------------------------------------------------------------------
-- 1. Expense sub-categories
-- Taha confirmed categories/subcategories change monthly. Rather than
-- guess the exact hierarchy, this lets a category optionally point to
-- a parent category — e.g. 'Internet' -> parent 'Utilities'.
-- Top-level categories just leave parent_category_id as NULL.
-- ---------------------------------------------------------------------
ALTER TABLE expense_categories
    ADD COLUMN parent_category_id INTEGER REFERENCES expense_categories(id);

CREATE INDEX idx_expense_categories_parent ON expense_categories(parent_category_id);


-- ---------------------------------------------------------------------
-- 2. Index to support "which BIN numbers are used most" reporting
-- (Taha specifically wants this comparison on the dashboard)
-- ---------------------------------------------------------------------
CREATE INDEX idx_bookings_bank_bin ON bookings(bank_bin);
