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
    updated_at  TEXT NOT NULL,
    version     INTEGER NOT NULL DEFAULT 1,
    deleted     INTEGER NOT NULL DEFAULT 0,
    source      TEXT NOT NULL DEFAULT 'local',
    UNIQUE (member_name, score_date)
);

CREATE INDEX IF NOT EXISTS idx_score_entries_date ON public.score_entries (score_date);

-- ============================================================
-- Row Level Security
-- Service role key bypasses RLS automatically.
-- Enable RLS so anon key cannot read/write without policy.
-- ============================================================
ALTER TABLE public.members      ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.score_entries ENABLE ROW LEVEL SECURITY;

-- Allow service role full access (already implicit, but explicit is clearer)
CREATE POLICY "service_role_members" ON public.members
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_score_entries" ON public.score_entries
    FOR ALL TO service_role USING (true) WITH CHECK (true);
