-- Supabase schema for BM2
-- Run this in Supabase SQL Editor:
-- Dashboard -> SQL Editor -> New query
--
-- These tables mirror the local SQLite schema used by the app.
-- Local source: SQLite
-- Online/mobile sync source: Supabase
-- Sync path: SQLite <-> SyncService <-> Supabase

-- ============================================================
-- members
-- ============================================================
-- id TEXT PRIMARY KEY implies UNIQUE in Postgres.
-- Upsert uses on_conflict="id"; no separate UNIQUE index needed.
CREATE TABLE IF NOT EXISTS public.members (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    status      TEXT NOT NULL,
    note        TEXT NOT NULL DEFAULT '',
    sort_order  INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL DEFAULT '',
    disabled_at TEXT NOT NULL DEFAULT '',
    updated_at  TEXT NOT NULL,
    version     INTEGER NOT NULL DEFAULT 1,
    deleted     INTEGER NOT NULL DEFAULT 0,
    UNIQUE (name)
);

-- ============================================================
-- score_entries
-- ============================================================
-- id TEXT PRIMARY KEY implies UNIQUE; upsert uses on_conflict="id".
CREATE TABLE IF NOT EXISTS public.score_entries (
    id          TEXT PRIMARY KEY,
    member_id   TEXT NOT NULL,
    member_name TEXT NOT NULL,
    score_date  TEXT NOT NULL,
    score       INTEGER NOT NULL,
    before_balance TEXT NOT NULL DEFAULT '',
    after_balance  TEXT NOT NULL DEFAULT '',
    manual_wear    TEXT NOT NULL DEFAULT '',
    income         TEXT NOT NULL DEFAULT '',
    other_expense  TEXT NOT NULL DEFAULT '',
    profit         TEXT NOT NULL DEFAULT '0',
    updated_at  TEXT NOT NULL,
    version     INTEGER NOT NULL DEFAULT 1,
    deleted     INTEGER NOT NULL DEFAULT 0,
    source      TEXT NOT NULL DEFAULT 'local',
    UNIQUE (member_name, score_date)
);

CREATE INDEX IF NOT EXISTS idx_score_entries_date ON public.score_entries (score_date);

ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS before_balance TEXT NOT NULL DEFAULT '';
ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS after_balance  TEXT NOT NULL DEFAULT '';
ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS manual_wear    TEXT NOT NULL DEFAULT '';
ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS income         TEXT NOT NULL DEFAULT '';
ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS other_expense  TEXT NOT NULL DEFAULT '';
ALTER TABLE public.score_entries ADD COLUMN IF NOT EXISTS profit         TEXT NOT NULL DEFAULT '0';

-- ============================================================
-- settlement_cycles
-- ============================================================
-- One row per cycle. Cycle data is synced so the online/mobile views
-- can scope wear / income / expense to "this cycle" instead of a
-- rolling 15-day window.
-- Upsert uses on_conflict="id".
CREATE TABLE IF NOT EXISTS public.settlement_cycles (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    start_date  TEXT NOT NULL DEFAULT '',
    settle_date TEXT NOT NULL DEFAULT '',
    settled     INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT NOT NULL,
    updated_at  TEXT NOT NULL,
    version     INTEGER NOT NULL DEFAULT 1,
    deleted     INTEGER NOT NULL DEFAULT 0,
    source      TEXT NOT NULL DEFAULT 'local'
);

ALTER TABLE public.settlement_cycles ADD COLUMN IF NOT EXISTS version INTEGER NOT NULL DEFAULT 1;
ALTER TABLE public.settlement_cycles ADD COLUMN IF NOT EXISTS source  TEXT    NOT NULL DEFAULT 'local';

-- ============================================================
-- settlement_entries
-- ============================================================
-- Per-member start/end balance rows scoped to a cycle. Composite PK
-- (cycle_id, member_name); upsert uses on_conflict="cycle_id,member_name".
CREATE TABLE IF NOT EXISTS public.settlement_entries (
    cycle_id      TEXT NOT NULL,
    member_name   TEXT NOT NULL,
    start_date    TEXT NOT NULL DEFAULT '',
    start_balance TEXT NOT NULL DEFAULT '',
    settle_date   TEXT NOT NULL DEFAULT '',
    end_balance   TEXT NOT NULL DEFAULT '',
    is_extra      INTEGER NOT NULL DEFAULT 0,
    sort_order    INTEGER NOT NULL DEFAULT 0,
    updated_at    TEXT NOT NULL,
    version       INTEGER NOT NULL DEFAULT 1,
    deleted       INTEGER NOT NULL DEFAULT 0,
    source        TEXT NOT NULL DEFAULT 'local',
    PRIMARY KEY (cycle_id, member_name)
);

CREATE INDEX IF NOT EXISTS idx_settlement_entries_cycle
    ON public.settlement_entries (cycle_id);

-- ============================================================
-- Row Level Security
-- Service role key bypasses RLS automatically.
-- Enable RLS so anon key cannot read/write without policy.
-- ============================================================
ALTER TABLE public.members             ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.score_entries       ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.settlement_cycles   ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.settlement_entries  ENABLE ROW LEVEL SECURITY;

-- Allow service role full access (already implicit, but explicit is clearer)
CREATE POLICY "service_role_members" ON public.members
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_score_entries" ON public.score_entries
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_settlement_cycles" ON public.settlement_cycles
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_settlement_entries" ON public.settlement_entries
    FOR ALL TO service_role USING (true) WITH CHECK (true);
