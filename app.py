"""
Excel Automation Studio — AB InBev
===================================
Professional Streamlit dashboard for the 12 automation scripts in this folder.
Black-and-gold design system with custom CSS injection.
"""

from __future__ import annotations

import os
import sys
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent

# The filename every relative-path script expects to read.
INPUT_NAME = "PRP Sample Jun (2).xlsx"

# The one script that hardcodes an absolute path; we rewrite this prefix to the
# temp dir so it reads/writes locally and never touches the real location.
ABS_PREFIX = (
    r"C:\Users\C915662\OneDrive - Anheuser-Busch InBev"
    r"\AB-InBev Automations\Assessment_Automation"
)


@dataclass
class ScriptEntry:
    id: str
    group: str
    label: str
    rel_path: str
    sheet: str
    notes: str = ""
    patch_abs: bool = False

    @property
    def path(self) -> Path:
        return APP_DIR / self.rel_path


# All 11 scripts, grouped by folder, with friendly labels.
REGISTRY: list[ScriptEntry] = [
    ScriptEntry(
        "d1", "Daigram 1 — Suppliers",
        "Zone-wise Tier-1 Suppliers (2026 Technology filter) + chart",
        "daigram 1 automation/automation.py",
        "TPRM Web-Portal Export",
        "Filters 2026 + TECHNOLOGY, merges static Tier-1 data, embeds a matplotlib chart.",
    ),
    ScriptEntry(
        "d2", "Daigram 2 — Assessments",
        "Cyber assessments 2026: Open vs Closed by zone + KPI charts",
        "daigram 2 automation/automation.py",
        "OneTrust Assessment",
        "Tags=Cyber & year 2026, Open/Closed pivot, Q2 KPI, two native Excel charts.",
    ),
    ScriptEntry(
        "d3", "Daigram 3 — Assessments",
        "Beyond-1-year-overdue assessments: Completed vs Open pivot + chart",
        "daigram 3 automation/automation.py",
        "OneTrust Assessment",
        "Has an absolute path baked in — the app rewrites it to run locally.",
        patch_abs=True,
    ),
    ScriptEntry(
        "s1a", "Slide 12 · 1st — Suppliers",
        "Assessment Status Overview (Active/Deprioritized/Duplicate)",
        "Slide 12 1st daigram Automation/automation.py",
        "TPRM Web-Portal Export",
        "Single status-overview chart.",
    ),
    ScriptEntry(
        "s1b", "Slide 12 · 1st — Suppliers",
        "Status Overview + Active-by-Zone (two charts)",
        "Slide 12 1st daigram Automation/automation2.py",
        "TPRM Web-Portal Export",
        "Adds a per-zone table and a second chart.",
    ),
    ScriptEntry(
        "s1c", "Slide 12 · 1st — Suppliers",
        "Status Overview + Active-by-Zone (combines ACTIVE/Active)",
        "Slide 12 1st daigram Automation/automation3.py",
        "TPRM Web-Portal Export",
        "Most robust 1st-diagram variant: merges case-variant Active columns.",
    ),
    ScriptEntry(
        "s2a", "Slide 12 · 2nd — Assessments",
        "Open vs Closed by Zone (pivot + stacked chart)",
        "Slide 12 2nd daigram Automation/automation.py",
        "OneTrust Assessment",
        "Appends two summary sheets to a copy of the workbook.",
    ),
    ScriptEntry(
        "s2b", "Slide 12 · 2nd — Assessments",
        "Open/Closed + Overdue (90d) — standalone report, 3 sheets",
        "Slide 12 2nd daigram Automation/automation2.py",
        "OneTrust Assessment",
        "Most complete 2nd-diagram variant: adds an overdue-by-zone report.",
    ),
    ScriptEntry(
        "s3a", "Slide 12 · 3rd — Risks",
        "Zone-wise Total vs Open Risks (dark stacked chart)",
        "Slide 12 3rd daigram Automation/automation.py",
        "OneTrust - Risk Export",
        "Total = Eval/Identified/Treatment/Monitoring; Open = first three.",
    ),
    ScriptEntry(
        "s3b", "Slide 12 · 3rd — Risks",
        "Zone Summary + Open vs Overdue (uses Aging buckets)",
        "Slide 12 3rd daigram Automation/automation2.py",
        "OneTrust - Risk Export",
        "Adds an Open-vs-Overdue chart from the Aging column text buckets.",
    ),
    ScriptEntry(
        "s3c", "Slide 12 · 3rd — Risks",
        "Formatted Zone Summary + Open vs Overdue (styled headers)",
        "Slide 12 3rd daigram Automation/automation3.py",
        "OneTrust - Risk Export",
        "Same as above with full cell formatting.",
    ),
    ScriptEntry(
        "s3d", "Slide 12 · 3rd — Risks",
        "Open/Overdue pivot (Aging > 90 days) — most complete risk report",
        "Slide 12 3rd daigram Automation/automation4.py",
        "OneTrust - Risk Export",
        "Most robust 3rd-diagram variant: numeric Aging>90 rule, column auto-detect.",
    ),
]

REGISTRY_BY_ID = {e.id: e for e in REGISTRY}


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

@dataclass
class RunResult:
    ok: bool
    returncode: int
    stdout: str
    stderr: str
    outputs: list[tuple[str, bytes]] = field(default_factory=list)   # (filename, bytes)
    tables: dict[str, dict[str, pd.DataFrame]] = field(default_factory=dict)  # file -> {sheet: df}


def _snapshot(folder: Path) -> dict[str, float]:
    snap = {}
    for p in folder.iterdir():
        if p.is_file():
            snap[p.name] = p.stat().st_mtime
    return snap


def run_script(entry: ScriptEntry, uploaded_bytes: bytes) -> RunResult:
    """Run one automation script in an isolated temp dir and collect its outputs."""
    tmpdir = Path(tempfile.mkdtemp(prefix="excel_auto_"))
    try:
        # 1. Place the uploaded workbook under the name the scripts expect.
        (tmpdir / INPUT_NAME).write_bytes(uploaded_bytes)

        # 2. Copy (and possibly patch) the script into the temp dir.
        source = entry.path.read_text(encoding="utf-8")
        if entry.patch_abs:
            source = source.replace(ABS_PREFIX, str(tmpdir))
        script_copy = tmpdir / "script_to_run.py"
        script_copy.write_text(source, encoding="utf-8")

        # 3. Snapshot existing files so we can detect what the script creates.
        before = _snapshot(tmpdir)

        # 4. Run it. Force UTF-8 (scripts print emoji that crash cp1252 on Windows)
        #    and a headless matplotlib backend.
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["MPLBACKEND"] = "Agg"
        proc = subprocess.run(
            [sys.executable, script_copy.name],
            cwd=str(tmpdir),
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )

        # 5. Diff to find new/modified output files (ignore the input + the copy).
        after = _snapshot(tmpdir)
        ignore = {INPUT_NAME, script_copy.name}
        produced = [
            name for name, mtime in after.items()
            if name not in ignore and (name not in before or before[name] != mtime)
        ]

        result = RunResult(
            ok=proc.returncode == 0,
            returncode=proc.returncode,
            stdout=proc.stdout or "",
            stderr=proc.stderr or "",
        )

        for name in sorted(produced):
            data = (tmpdir / name).read_bytes()
            result.outputs.append((name, data))
            if name.lower().endswith(".xlsx"):
                try:
                    result.tables[name] = pd.read_excel(
                        tmpdir / name, sheet_name=None, engine="openpyxl"
                    )
                except Exception as exc:  # preview is best-effort
                    result.tables[name] = {
                        "(could not read)": pd.DataFrame({"error": [str(exc)]})
                    }
        return result
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Preview helpers
# ---------------------------------------------------------------------------

def _chartable(df: pd.DataFrame):
    """Return (index_col, numeric_df) suitable for st.bar_chart, or None."""
    if df is None or df.empty:
        return None
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    label_cols = [c for c in df.columns if c not in numeric_cols]
    if not numeric_cols or not label_cols:
        return None
    label = label_cols[0]
    chart_df = df[[label] + numeric_cols].copy()
    chart_df = chart_df.dropna(subset=[label])
    # Drop obvious total rows so the chart isn't dominated by them.
    mask = ~chart_df[label].astype(str).str.strip().str.lower().isin(
        {"grand total", "total"}
    )
    chart_df = chart_df[mask]
    if chart_df.empty:
        return None
    return chart_df.set_index(label)


def render_results(entry: ScriptEntry, result: RunResult, fmt: str) -> None:
    # ── Status banner ──────────────────────────────────────────────────────
    if result.ok:
        st.markdown(
            f'<div style="background:#0b1f12;border:1px solid #2a6640;border-left:3px solid #3ecf6e;'
            f'border-radius:8px;padding:.75rem 1.1rem;color:#3ecf6e;font-size:.88rem;font-weight:600;">'
            f'&#10003; &nbsp;Completed &mdash; {entry.label}</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div style="background:#1f0b0b;border:1px solid #662a2a;border-left:3px solid #f55;'
            f'border-radius:8px;padding:.75rem 1.1rem;color:#ff6666;font-size:.88rem;font-weight:600;">'
            f'&#10007; &nbsp;Script exited with error (code {result.returncode}) &mdash; '
            f'check sheet name / column headers</div>',
            unsafe_allow_html=True,
        )

    with st.expander("Script log  (stdout / stderr)", expanded=not result.ok):
        if result.stdout.strip():
            st.code(result.stdout, language="text")
        if result.stderr.strip():
            st.markdown(
                '<span style="font-size:.72rem;color:#D4AF37;text-transform:uppercase;'
                'letter-spacing:.1em;">stderr</span>',
                unsafe_allow_html=True,
            )
            st.code(result.stderr, language="text")
        if not result.stdout.strip() and not result.stderr.strip():
            st.markdown('<span style="color:#555;font-size:.8rem;">No console output</span>', unsafe_allow_html=True)

    if not result.outputs:
        st.warning("The script produced no output file.")
        return

    # ── Downloads ──────────────────────────────────────────────────────────
    st.markdown(_section_label("Downloads"), unsafe_allow_html=True)
    dl_cols = st.columns(min(4, len(result.outputs)))
    for i, (name, data) in enumerate(result.outputs):
        mime = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if name.lower().endswith(".xlsx")
            else "image/png" if name.lower().endswith(".png")
            else "application/octet-stream"
        )
        icon = "&#128202;" if name.endswith(".xlsx") else "&#128444;"
        dl_cols[i % len(dl_cols)].download_button(
            f"{icon}  {name}",
            data=data,
            file_name=name,
            mime=mime,
            key=f"dl_{entry.id}_{i}",
            use_container_width=True,
        )

    # ── Chart preview ──────────────────────────────────────────────────────
    st.markdown(_section_label("Chart Preview"), unsafe_allow_html=True)
    png_outputs = [(n, d) for n, d in result.outputs if n.lower().endswith(".png")]
    rendered_any = False

    if png_outputs:
        img_cols = st.columns(len(png_outputs))
        for col, (name, data) in zip(img_cols, png_outputs):
            with col:
                st.image(data, caption=name, use_container_width=True)
        rendered_any = True

    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            chart_df = _chartable(df)
            if chart_df is not None and len(chart_df.columns) <= 6:
                st.markdown(
                    f'<div style="font-size:.72rem;color:#555;margin-bottom:.3rem;">'
                    f'{fname} &rarr; {sheet_name}</div>',
                    unsafe_allow_html=True,
                )
                st.bar_chart(chart_df, stack=True, use_container_width=True)
                rendered_any = True

    if not rendered_any:
        st.markdown(
            '<div style="color:#555;font-size:.82rem;padding:.5rem 0;">No chartable data detected for this report.</div>',
            unsafe_allow_html=True,
        )
    elif not png_outputs:
        st.markdown(
            '<div style="color:#555;font-size:.75rem;margin-top:.3rem;">'
            'Re-rendered from output data &mdash; the downloaded .xlsx contains the original styled chart.</div>',
            unsafe_allow_html=True,
        )

    # ── Tables ─────────────────────────────────────────────────────────────
    st.markdown(_section_label("Output Tables"), unsafe_allow_html=True)
    for fname, sheets in result.tables.items():
        for sheet_name, df in sheets.items():
            st.markdown(
                f'<div style="display:flex;align-items:baseline;gap:.6rem;margin-bottom:.4rem;">'
                f'<span style="font-size:.88rem;font-weight:700;color:#F0F0F0;">{sheet_name}</span>'
                f'<span style="font-size:.72rem;color:#555;font-family:monospace;">{fname}</span>'
                f'<span style="font-size:.7rem;color:#444;margin-left:auto;">'
                f'{len(df)} rows &times; {len(df.columns)} cols</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            st.dataframe(df, use_container_width=True, hide_index=True)
            with st.expander(f"Copy  '{sheet_name}'  as text"):
                if fmt == "Markdown":
                    try:
                        text = df.to_markdown(index=False)
                    except Exception:
                        text = df.to_string(index=False)
                else:
                    text = df.to_string(index=False)
                st.code(text, language="text")


# ---------------------------------------------------------------------------
# Design system — black & gold
# ---------------------------------------------------------------------------

_CSS = """
<style>
:root {
  --gold:        #D4AF37;
  --gold-bright: #F5C842;
  --gold-dim:    #8B7520;
  --gold-glow:   rgba(212,175,55,0.18);
  --bg:          #060606;
  --surface:     #0D0D0D;
  --card:        #131313;
  --card2:       #1A1A1A;
  --border:      #232323;
  --text:        #F0F0F0;
  --text2:       #B0B0B0;
  --muted:       #666666;
  --radius:      8px;
  --radius-lg:   12px;
}

/* ── Page ── */
.stApp { background: var(--bg) !important; }
.stApp > header { background: var(--bg) !important; border-bottom: 1px solid var(--border) !important; }
#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] { display:none !important; }
.main .block-container { padding-top: 1.5rem !important; max-width: 1340px !important; }

/* ── Headings ── */
h1,h2,h3,h4 { color: var(--text) !important; }

/* ── Subheader override: gold label style ── */
[data-testid="stHeading"] h2 {
  font-size: 0.72rem !important;
  text-transform: uppercase !important;
  letter-spacing: 0.14em !important;
  color: var(--gold) !important;
  font-weight: 700 !important;
  border-bottom: 1px solid var(--border) !important;
  padding-bottom: 0.45rem !important;
  margin-bottom: 0.75rem !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
  background: var(--card) !important;
  border: 1.5px dashed var(--gold-dim) !important;
  border-radius: var(--radius-lg) !important;
  transition: border-color .2s, box-shadow .2s;
}
[data-testid="stFileUploader"]:hover {
  border-color: var(--gold) !important;
  box-shadow: 0 0 18px var(--gold-glow) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] div,
[data-testid="stFileUploaderDropzoneInstructions"] span { color: var(--muted) !important; }
[data-testid="stFileUploader"] label { color: var(--muted) !important; font-size:.72rem !important; text-transform:uppercase !important; letter-spacing:.1em !important; }
[data-testid="stFileUploader"] small { color: var(--muted) !important; }

/* ── Selectbox ── */
[data-testid="stSelectbox"] label {
  color: var(--muted) !important; font-size:.72rem !important;
  text-transform:uppercase !important; letter-spacing:.1em !important;
}
[data-testid="stSelectbox"] > div > div {
  background: var(--card) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius) !important;
  color: var(--text) !important;
}
[data-testid="stSelectbox"] > div > div:focus-within {
  border-color: var(--gold) !important;
  box-shadow: 0 0 0 1px var(--gold) !important;
}

/* ── Primary button ── */
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, #C9A227 0%, #F5C842 50%, #C9A227 100%) !important;
  background-size: 200% !important;
  color: #000 !important;
  border: none !important;
  border-radius: var(--radius) !important;
  font-weight: 800 !important;
  font-size: .85rem !important;
  letter-spacing: .1em !important;
  text-transform: uppercase !important;
  padding: .55rem 2.2rem !important;
  box-shadow: 0 0 22px rgba(212,175,55,.22) !important;
  transition: box-shadow .25s, transform .15s !important;
}
.stButton > button[kind="primary"]:hover {
  box-shadow: 0 0 36px rgba(212,175,55,.45) !important;
  transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"]:disabled {
  background: var(--card2) !important;
  color: var(--muted) !important;
  box-shadow: none !important;
  transform: none !important;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
  background: var(--card) !important;
  border: 1px solid var(--gold-dim) !important;
  color: var(--gold) !important;
  border-radius: var(--radius) !important;
  font-weight: 600 !important;
  font-size: .82rem !important;
  letter-spacing: .04em !important;
  transition: all .2s !important;
}
[data-testid="stDownloadButton"] > button:hover {
  background: rgba(212,175,55,.08) !important;
  border-color: var(--gold) !important;
  box-shadow: 0 0 14px var(--gold-glow) !important;
}

/* ── Radio ── */
[data-testid="stRadio"] label { color: var(--muted) !important; font-size:.72rem !important; text-transform:uppercase !important; letter-spacing:.1em !important; }
[data-testid="stRadio"] [data-testid="stMarkdownContainer"] p { color: var(--text2) !important; font-size:.85rem !important; }

/* ── Expander ── */
[data-testid="stExpander"] {
  background: var(--card) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius) !important;
}
[data-testid="stExpander"] summary { color: var(--muted) !important; font-size:.82rem !important; }
[data-testid="stExpander"] summary:hover { color: var(--gold) !important; }
[data-testid="stExpander"] summary svg { fill: var(--muted) !important; }

/* ── Alerts ── */
[data-testid="stAlert"] { background: var(--card) !important; border-radius: var(--radius) !important; }
[data-testid="stAlert"][data-baseweb="notification"] { border-left: 3px solid var(--gold) !important; }
.stSuccess[data-testid="stAlert"] { border-left: 3px solid #3ecf6e !important; }
.stError[data-testid="stAlert"]   { border-left: 3px solid #f55 !important; }
[data-testid="stAlert"] p { color: var(--text2) !important; }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border: 1px solid var(--border) !important; border-radius: var(--radius) !important; overflow: hidden !important; }

/* ── Code block ── */
[data-testid="stCode"] { background: var(--surface) !important; border: 1px solid var(--border) !important; border-radius: var(--radius) !important; }
.stCodeBlock pre { background: var(--surface) !important; color: var(--text2) !important; }

/* ── Image ── */
[data-testid="stImage"] img { border-radius: var(--radius) !important; border: 1px solid var(--border) !important; }

/* ── Spinner ── */
[data-testid="stSpinner"] > div { border-top-color: var(--gold) !important; }

/* ── Caption / text ── */
.stCaption p, [data-testid="stCaptionContainer"] p { color: var(--muted) !important; font-size:.78rem !important; }
.stMarkdown p { color: var(--text2) !important; }

/* ── Divider ── */
hr { border-color: var(--border) !important; margin: 1.2rem 0 !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background: var(--surface); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background: var(--gold-dim); }

/* ── Bar chart (Vega) — force dark ── */
[data-testid="stArrowVegaLiteChart"] {
  background: var(--card) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius) !important;
  padding: .75rem !important;
}
canvas { border-radius: var(--radius) !important; }
</style>
"""

# Inline HTML helpers
def _header_html() -> str:
    return """
<div style="
  background:linear-gradient(135deg,#0D0D0D 0%,#111008 100%);
  border:1px solid #232323;
  border-bottom:2px solid #D4AF37;
  border-radius:14px;
  padding:1.8rem 2.2rem 1.4rem;
  margin-bottom:1.6rem;
  position:relative;
  overflow:hidden;
">
  <div style="position:absolute;top:0;right:0;width:380px;height:100%;
    background:radial-gradient(ellipse at 90% 50%,rgba(212,175,55,.09) 0%,transparent 70%);
    pointer-events:none;"></div>
  <div style="display:flex;align-items:center;gap:1.1rem;">
    <div style="
      width:48px;height:48px;border-radius:11px;flex-shrink:0;
      background:linear-gradient(135deg,#D4AF37,#7A6010);
      display:flex;align-items:center;justify-content:center;
      font-size:1.5rem;box-shadow:0 0 20px rgba(212,175,55,.3);
    ">&#128202;</div>
    <div>
      <div style="font-size:1.55rem;font-weight:800;color:#F0F0F0;letter-spacing:-.02em;line-height:1.1;">
        Excel Automation <span style="color:#D4AF37;">Studio</span>
      </div>
      <div style="font-size:.72rem;color:#666;letter-spacing:.12em;text-transform:uppercase;margin-top:.25rem;">
        AB InBev &nbsp;&bull;&nbsp; TPRM / Risk / Assessment Automation &nbsp;&bull;&nbsp; 12 Reports
      </div>
    </div>
  </div>
</div>
"""

def _info_card_html(sheet: str, notes: str, group: str) -> str:
    return f"""
<div style="
  background:#131313;
  border:1px solid #232323;
  border-left:3px solid #D4AF37;
  border-radius:8px;
  padding:.9rem 1.2rem;
  margin:.6rem 0 1rem;
">
  <div style="display:flex;gap:2rem;flex-wrap:wrap;">
    <div>
      <div style="font-size:.65rem;color:#555;text-transform:uppercase;letter-spacing:.12em;margin-bottom:.2rem;">Category</div>
      <div style="font-size:.88rem;color:#D4AF37;font-weight:600;">{group}</div>
    </div>
    <div>
      <div style="font-size:.65rem;color:#555;text-transform:uppercase;letter-spacing:.12em;margin-bottom:.2rem;">Required Sheet</div>
      <div style="font-size:.88rem;color:#F0F0F0;font-family:monospace;">{sheet}</div>
    </div>
    <div style="flex:1;min-width:200px;">
      <div style="font-size:.65rem;color:#555;text-transform:uppercase;letter-spacing:.12em;margin-bottom:.2rem;">Notes</div>
      <div style="font-size:.82rem;color:#888;">{notes}</div>
    </div>
  </div>
</div>
"""

def _section_label(text: str) -> str:
    return f"""
<div style="
  font-size:.68rem;color:#D4AF37;text-transform:uppercase;
  letter-spacing:.16em;font-weight:700;
  border-bottom:1px solid #232323;padding-bottom:.35rem;
  margin:1.4rem 0 .7rem;
">{text}</div>
"""


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="Excel Automation Studio",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Inject design system CSS
    st.markdown(_CSS, unsafe_allow_html=True)

    # Custom header
    st.markdown(_header_html(), unsafe_allow_html=True)

    # ── Control panel ──────────────────────────────────────────────────────
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown(_section_label("1 · Input Workbook"), unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Drop your Excel workbook here or click to browse",
            type=["xlsx"],
            label_visibility="collapsed",
        )
        if uploaded:
            st.markdown(
                f'<div style="font-size:.78rem;color:#3ecf6e;margin-top:.3rem;">'
                f'&#10003; &nbsp;<strong style="color:#F0F0F0;">{uploaded.name}</strong>'
                f'&nbsp;<span style="color:#555;">({uploaded.size/1024:.1f} KB)</span></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div style="font-size:.75rem;color:#555;margin-top:.3rem;">'
                'Accepts .xlsx files · same workbook the scripts expect</div>',
                unsafe_allow_html=True,
            )

    with col_right:
        st.markdown(_section_label("2 · Select Report"), unsafe_allow_html=True)

        groups: dict[str, list[ScriptEntry]] = {}
        for e in REGISTRY:
            groups.setdefault(e.group, []).append(e)

        options: list[str] = []
        option_to_id: dict[str, str] = {}
        for group, entries in groups.items():
            for e in entries:
                display = f"{group}  ·  {e.label}"
                options.append(display)
                option_to_id[display] = e.id

        choice = st.selectbox(
            "Report",
            options,
            index=0,
            label_visibility="collapsed",
        )
        entry = REGISTRY_BY_ID[option_to_id[choice]]
        st.markdown(_info_card_html(entry.sheet, entry.notes, entry.group), unsafe_allow_html=True)

    # ── Options row ────────────────────────────────────────────────────────
    st.markdown(_section_label("3 · Options & Run"), unsafe_allow_html=True)
    opt_col, run_col = st.columns([2, 1], gap="large")

    with opt_col:
        fmt = st.radio(
            "Copy format",
            ["TSV (Excel / Sheets)", "Markdown"],
            horizontal=True,
        )
        fmt = "Markdown" if "Markdown" in fmt else "TSV"

    with run_col:
        run = st.button(
            "&#9654;  Run Report",
            type="primary",
            disabled=uploaded is None,
            use_container_width=True,
        )
        if uploaded is None:
            st.markdown(
                '<div style="font-size:.72rem;color:#555;text-align:center;margin-top:.3rem;">'
                'Upload a workbook to enable</div>',
                unsafe_allow_html=True,
            )

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── How it works (collapsed by default) ───────────────────────────────
    with st.expander("How it works"):
        st.markdown(
            "1. **Upload** your `.xlsx` workbook — same file the scripts expect "
            "(*PRP Sample Jun (2).xlsx* or equivalent).\n"
            "2. **Select** one of the 12 automation reports from the dropdown.\n"
            "3. **Run** — the original script executes in an isolated temp folder. "
            "You get the log, a chart preview, every output table with a copy button, "
            "and the exact styled `.xlsx` the script generates.\n\n"
            "> Each report requires a specific sheet — shown in the info card above the selector."
        )

    # ── Execute ────────────────────────────────────────────────────────────
    if run and uploaded is not None:
        with st.spinner(f"Running  ·  {entry.label}"):
            try:
                result = run_script(entry, uploaded.getvalue())
            except subprocess.TimeoutExpired:
                st.error("Script exceeded 3-minute timeout and was stopped.")
                return
            except Exception as exc:
                st.error(f"Could not launch script: {exc}")
                return
        # Persist results so download-button reruns don't wipe them.
        st.session_state["last_result"] = result
        st.session_state["last_entry_id"] = entry.id
        st.session_state["last_fmt"] = fmt

    # Render from session state — survives every rerun (download clicks, etc.)
    if "last_result" in st.session_state:
        _entry = REGISTRY_BY_ID[st.session_state["last_entry_id"]]
        render_results(_entry, st.session_state["last_result"], st.session_state["last_fmt"])


if __name__ == "__main__":
    main()
