-- TAS payroll tables.
-- Holiday letter H is a special holiday. Holiday letter L is a legal holiday.

CREATE TABLE IF NOT EXISTS dataset (
  id integer PRIMARY KEY CHECK (id = 1),
  initialized boolean NOT NULL
);

CREATE TABLE IF NOT EXISTS agencies (
  code text PRIMARY KEY,
  name text NOT NULL,
  database_name text,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
  id text PRIMARY KEY,
  username text NOT NULL UNIQUE,
  name text NOT NULL,
  account_type text NOT NULL,
  active_login boolean NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS company (
  id integer PRIMARY KEY CHECK (id = 1),
  name text NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS clients (
  id text PRIMARY KEY,
  code text NOT NULL,
  code2 text NOT NULL,
  name text NOT NULL,
  active boolean NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS shifts (
  id text PRIMARY KEY,
  code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS holidays (
  id text PRIMARY KEY,
  holiday_date text NOT NULL,
  name text NOT NULL,
  letter text NOT NULL CHECK (letter IN ('H', 'L')),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS employees (
  id text PRIMARY KEY,
  emp_no text NOT NULL UNIQUE,
  dtr_id text NOT NULL,
  client_id text NOT NULL REFERENCES clients (id),
  payroll_status text NOT NULL,
  role text NOT NULL,
  active boolean NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS leaves (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS overtime_filings (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS punches (
  id text PRIMARY KEY,
  dtr_id text NOT NULL,
  punch_date text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS staff_days (
  emp_id text NOT NULL REFERENCES employees (id),
  dtr_key text NOT NULL,
  days jsonb NOT NULL,
  PRIMARY KEY (emp_id, dtr_key)
);

CREATE TABLE IF NOT EXISTS fms_days (
  emp_id text NOT NULL REFERENCES employees (id),
  dtr_key text NOT NULL,
  days jsonb NOT NULL,
  PRIMARY KEY (emp_id, dtr_key)
);

CREATE TABLE IF NOT EXISTS payroll_attendance (
  emp_id text NOT NULL REFERENCES employees (id),
  dtr_key text NOT NULL,
  kind text NOT NULL CHECK (kind IN ('staff', 'fms')),
  finalized_at text NOT NULL,
  doc jsonb NOT NULL,
  PRIMARY KEY (emp_id, dtr_key)
);

CREATE TABLE IF NOT EXISTS cutoffs (
  client_id text NOT NULL REFERENCES clients (id),
  dtr_key text NOT NULL,
  start_date text NOT NULL,
  end_date text NOT NULL,
  billed boolean NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL,
  PRIMARY KEY (client_id, dtr_key)
);

CREATE TABLE IF NOT EXISTS deductions (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  code integer NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS payroll_records (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  client_id text NOT NULL REFERENCES clients (id),
  dtr_key text NOT NULL,
  half smallint NOT NULL,
  open boolean NOT NULL,
  gross numeric NOT NULL,
  net numeric NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS trips (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  trip_date text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS commission_payments (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  pay_year integer NOT NULL,
  pay_month integer NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS thirteenth_month (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  client_id text NOT NULL REFERENCES clients (id),
  pay_year integer NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS banks (
  id text PRIMARY KEY,
  code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS branches (
  id text PRIMARY KEY,
  code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS departments (
  id text PRIMARY KEY,
  code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS sss_table (
  id text PRIMARY KEY,
  wage_from numeric NOT NULL,
  wage_to numeric NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS philhealth_table (
  id text PRIMARY KEY,
  wage_from numeric NOT NULL,
  wage_to numeric NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS hdmf_table (
  id text PRIMARY KEY,
  wage_from numeric NOT NULL,
  wage_to numeric NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS tax_table (
  id text PRIMARY KEY,
  tax_set text NOT NULL,
  tax_code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS tax_codes (
  id text PRIMARY KEY,
  code text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS user_logs (
  id text PRIMARY KEY,
  at text NOT NULL,
  username text NOT NULL,
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS benefits (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS assignments (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS trainings (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS offenses (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS leave_balances (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS appraisals (
  id text PRIMARY KEY,
  emp_id text NOT NULL REFERENCES employees (id),
  ord integer NOT NULL,
  doc jsonb NOT NULL
);

CREATE TABLE IF NOT EXISTS desk_session (
  id integer PRIMARY KEY CHECK (id = 1),
  session jsonb,
  flash jsonb
);

CREATE INDEX IF NOT EXISTS employees_client_idx ON employees (client_id);
CREATE INDEX IF NOT EXISTS punches_dtr_idx ON punches (dtr_id, punch_date);
CREATE INDEX IF NOT EXISTS payroll_emp_idx ON payroll_records (emp_id, dtr_key);
CREATE INDEX IF NOT EXISTS deductions_emp_idx ON deductions (emp_id);
CREATE INDEX IF NOT EXISTS cutoffs_billed_idx ON cutoffs (billed);
