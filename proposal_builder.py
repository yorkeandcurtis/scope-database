#!/usr/bin/env python3
"""
Y&C Proposal Builder — Estimate & Clarifications Letter Generator

Builds proposal letters from the Y&C template using a searchable database
of historical scope items. Cross-platform (Mac + Windows).
"""

import os
import re
import sys
import copy
import sqlite3
import datetime
import textwrap
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import Counter

from docx import Document
from docx.shared import Pt, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Paths ────────────────────────────────────────────────────────────────────

if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys._MEIPASS)
else:
    SCRIPT_DIR = Path(__file__).resolve().parent

DB_PATH = SCRIPT_DIR / "scope_database.db"

# Template location varies by platform
if sys.platform == "win32":
    _ONEDRIVE_TEMPLATE = os.path.expanduser(
        "~/OneDrive - Yorke and Curtis, Inc/"
        "01 - PRECONSTRUCTION/04 - TEMPLATES/"
        "00 - New Project Template/05 - Quantities & Estimates/"
        "Estimate & Clarifications Template.docx"
    )
else:
    _ONEDRIVE_TEMPLATE = os.path.expanduser(
        "~/Library/CloudStorage/OneDrive-YorkeandCurtis,Inc/"
        "01 - PRECONSTRUCTION/04 - TEMPLATES/"
        "00 - New Project Template/05 - Quantities & Estimates/"
        "Estimate & Clarifications Template.docx"
    )

_BUNDLED_TEMPLATE = SCRIPT_DIR / "Estimate & Clarifications Template.docx"

def get_template_path():
    if _BUNDLED_TEMPLATE.exists():
        return str(_BUNDLED_TEMPLATE)
    if os.path.exists(_ONEDRIVE_TEMPLATE):
        return _ONEDRIVE_TEMPLATE
    return None


# ── Division Definitions ─────────────────────────────────────────────────────

DIVISIONS = [
    ("1", "Division 1 - General", "Division 1 - General"),
    ("2", "Division 2 – Sitework", "Division 2 - Sitework"),
    ("3", "Division 3 – Concrete", "Division 3 - Concrete"),
    ("4", "Division 4 – Masonry", "Division 4 - Masonry"),
    ("5", "Division 5 – Metals", "Division 5 - Metals/Steel"),
    ("6", "Division 6 – Woods & Plastics", "Division 6 - Woods & Plastics"),
    ("7", "Division 7 – Thermal & Moisture Protection", "Division 7 - Thermal & Moisture Protection"),
    ("8", "Division 8 – Doors & Windows", "Division 8 - Doors & Windows"),
    ("9", "Division 9 – Finishes", "Division 9 - Finishes"),
    ("10", "Division 10/11/12 – Specialties/Equipment/Furnishings", "Division 10 - Specialties"),
    ("15", "Division 15 – Mechanical", "Division 15 - Mechanical/Plumbing"),
    ("16", "Division 16 – Electrical", "Division 16 - Electrical"),
]

STANDARD_EXCLUSIONS = [
    "Building Permit, SDC's, Water Meter or associated fees",
    "Architect or consultant fees",
    "Any and all permits except MEP Permits",
    "Special inspections, testing and bonds",
    "Utility company fees, for example (NW Natural, water department, power utility, etc.)",
    "Franchise utility work/Fees including removal or relocation of overhead lines",
    "Work to Existing Power poles, transformers, relocation of these services, and/or underground transformers, vaults, etc.",
    "Testing and/or removal of any contaminated soils",
    "Over Excavation of Soils",
    "Testing and/or removal of hazardous materials",
    "Adjacent property access costs and any costs associated with renting adjacent property for use during construction/staging",
    "Tree grates",
    "CCTV system & Access Controls",
    "ROW/Street improvements",
    "LEED Provisions until final determination",
    "Prevailing Wages",
    "Public Works Bonds",
    "Deferred Submittals & associated costs",
]

ESTIMATE_TYPES = [
    "Preliminary Estimate",
    "ROM Budget Estimate",
    "Budget Estimate",
    "Final Estimate",
    "Bid",
]

PROJECT_TYPES = [
    "All Types",
    "Retail/Commercial",
    "Office/TI",
    "Apartments/Residential",
    "Senior Living",
    "Religious/Non-Profit",
    "Education",
    "Medical/Healthcare",
    "Industrial/Manufacturing",
    "Recreation/Amenity",
    "Mixed Use",
    "Hospitality",
    "Other",
]


# ── Database Layer ───────────────────────────────────────────────────────────

class ScopeDatabase:
    """Interface to the scope items SQLite database."""

    def __init__(self, db_path):
        self.conn = sqlite3.connect(str(db_path))
        self.conn.row_factory = sqlite3.Row

    def get_division_items(self, division_db_name, limit=200,
                           scope_type=None, project_type=None,
                           year_from=None, year_to=None,
                           project_search=None):
        """Get scope items for a division, ranked by frequency, with filters."""
        c = self.conn.cursor()
        div_num = re.search(r'\d+', division_db_name)
        if not div_num:
            return []

        # Use " -" suffix to prevent Division 1 matching Division 10-17
        num = div_num.group(0)
        pattern = f"Division {num} -%"
        # Also match exact "Division {num}" with no suffix (e.g. "Division 17")
        conditions = ["(division LIKE ? OR division = ?)"]
        params = [pattern, f"Division {num}"]

        if scope_type and scope_type != "All":
            conditions.append("scope_type = ?")
            params.append(scope_type)

        if project_type and project_type != "All Types":
            conditions.append("project_type = ?")
            params.append(project_type)

        if year_from:
            conditions.append("document_date >= ?")
            params.append(f"{year_from}-01-01")

        if year_to:
            conditions.append("document_date <= ?")
            params.append(f"{year_to}-12-31")

        if project_search:
            conditions.append("project_name LIKE ?")
            params.append(f"%{project_search}%")

        where = " AND ".join(conditions)
        params.append(limit)

        c.execute(f"""
            SELECT scope_text, scope_type, COUNT(*) as freq,
                   GROUP_CONCAT(DISTINCT project_name) as projects
            FROM scope_items
            WHERE {where}
            GROUP BY scope_text
            ORDER BY freq DESC
            LIMIT ?
        """, params)
        return [dict(row) for row in c.fetchall()]

    def search_items(self, query, division_filter=None,
                     scope_type=None, project_type=None,
                     year_from=None, year_to=None,
                     project_search=None):
        """Full-text search across scope items with filters."""
        c = self.conn.cursor()
        # FTS table doesn't have all columns, so join with main table
        conditions = ["fts.scope_items_fts MATCH ?"]
        params = [query]

        if division_filter:
            div_num = re.search(r'\d+', division_filter)
            if div_num:
                num = div_num.group(0)
                conditions.append("(si.division LIKE ? OR si.division = ?)")
                params.extend([f"Division {num} -%", f"Division {num}"])

        if scope_type and scope_type != "All":
            conditions.append("si.scope_type = ?")
            params.append(scope_type)

        if project_type and project_type != "All Types":
            conditions.append("si.project_type = ?")
            params.append(project_type)

        if year_from:
            conditions.append("si.document_date >= ?")
            params.append(f"{year_from}-01-01")

        if year_to:
            conditions.append("si.document_date <= ?")
            params.append(f"{year_to}-12-31")

        if project_search:
            conditions.append("si.project_name LIKE ?")
            params.append(f"%{project_search}%")

        where = " AND ".join(conditions)

        c.execute(f"""
            SELECT si.scope_text, si.scope_type, si.division,
                   COUNT(*) as freq
            FROM scope_items_fts fts
            JOIN scope_items si ON fts.rowid = si.id
            WHERE {where}
            GROUP BY si.scope_text
            ORDER BY freq DESC
            LIMIT 100
        """, params)
        return [dict(row) for row in c.fetchall()]

    def get_project_names(self, division_db_name=None):
        """Get distinct project names, optionally filtered by division."""
        c = self.conn.cursor()
        if division_db_name:
            div_num = re.search(r'\d+', division_db_name)
            if div_num:
                c.execute("""
                    SELECT DISTINCT project_name FROM scope_items
                    WHERE division LIKE ?
                    ORDER BY project_name
                """, (f"Division {div_num.group(0)}%",))
            else:
                c.execute("SELECT DISTINCT project_name FROM scope_items ORDER BY project_name")
        else:
            c.execute("SELECT DISTINCT project_name FROM scope_items ORDER BY project_name")
        return [r[0] for r in c.fetchall()]

    def close(self):
        self.conn.close()


# ── Document Generator ───────────────────────────────────────────────────────

def _make_paragraph_element(text="", style_name=None):
    """Create a w:p element with optional text and style."""
    p = OxmlElement("w:p")
    if style_name:
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), style_name)
        pPr.append(pStyle)
        p.append(pPr)
    if text:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        p.append(r)
    return p


def _make_run(text, bold=False, size=None):
    """Create a w:r element."""
    r = OxmlElement("w:r")
    if bold or size:
        rPr = OxmlElement("w:rPr")
        if bold:
            b = OxmlElement("w:b")
            rPr.append(b)
        if size:
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(size))  # half-points
            rPr.append(sz)
            szCs = OxmlElement("w:szCs")
            szCs.set(qn("w:val"), str(size))
            rPr.append(szCs)
        r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def generate_proposal(template_path, output_path, project_info, division_items, exclusions):
    """
    Generate a proposal letter from the template, preserving headers/footers/logo.

    Strategy: load the template, remove only the body paragraphs between the first
    and last section properties, then insert new content paragraphs using the same
    XML namespace. This preserves all header/footer/image relationships.
    """
    doc = Document(template_path)
    body = doc.element.body

    # Collect all paragraph and table elements (not sectPr)
    to_remove = []
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("p", "tbl"):
            # Don't remove paragraphs that contain section properties
            has_sect = child.findall(qn("w:pPr") + "/" + qn("w:sectPr"))
            if not has_sect:
                to_remove.append(child)

    for elem in to_remove:
        body.remove(elem)

    # Find insertion point: insert before the first remaining element
    # (which will be a sectPr or a paragraph containing sectPr)
    insert_before = None
    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("p", "sectPr"):
            insert_before = child
            break

    def insert_p(text="", style=None, runs=None):
        """Insert a paragraph before the section properties."""
        p = _make_paragraph_element("" if runs else text, style)
        if runs:
            for run_text, bold, size in runs:
                p.append(_make_run(run_text, bold, size))
        if insert_before is not None:
            body.insert(list(body).index(insert_before), p)
        else:
            body.append(p)
        return p

    def insert_list_item(text):
        """Insert a List Paragraph bullet item."""
        return insert_p(text, style="ListParagraph")

    # ── Build the document content ──

    fmt_date = project_info.get("date", datetime.date.today().strftime("%m-%d-%y"))

    # Date
    insert_p(fmt_date)

    # Blank + Client name
    insert_p("\n" + project_info.get("client_name", ""))

    # Client address
    insert_p(project_info.get("client_address", ""))

    # Blank + Re: subject line
    subject = (project_info.get("project_name", "Project") + " " +
               project_info.get("estimate_type", "Estimate") + " & Clarifications" +
               " Dated " + fmt_date)
    insert_p(runs=[("\n", False, None), ("Re: " + subject, True, None)])

    # Blank
    insert_p()

    # Dear line
    insert_p(project_info.get("dear_line", "Dear Client:"))

    # Opening paragraph
    architect = project_info.get("architect", "Architect")
    drawing_date = project_info.get("drawing_date", "")
    est_type_lower = project_info.get("estimate_type", "Estimate").lower()
    proj_name = project_info.get("project_name", "the")
    opening = f"Below is the {est_type_lower} & clarifications for the {proj_name} Project based upon {architect} drawings"
    if drawing_date:
        opening += f" Dated {drawing_date}"
    opening += ":"
    insert_p(opening)

    # Blank
    insert_p()

    # Estimate amount (bold, 14pt = 28 half-points)
    est_type = project_info.get("estimate_type", "Preliminary Estimate")
    est_amount = project_info.get("estimate_amount", "")
    if est_amount:
        try:
            amt_float = float(str(est_amount).replace(",", "").replace("$", ""))
            est_amount_fmt = f"${amt_float:,.0f}"
        except ValueError:
            est_amount_fmt = est_amount
    else:
        est_amount_fmt = "$X,XXX,XXX"
    insert_p(runs=[(f"{est_type}:  {est_amount_fmt}", True, 28)])

    # Blank
    insert_p()

    # Qualifications header
    insert_p(runs=[
        ("Please note the following ", False, None),
        ("Specific Qualifications & Clarifications", True, None),
        (":", False, None),
    ])

    # Blank
    insert_p()

    # Division scope items
    for div_num, div_label, div_db_name in DIVISIONS:
        items = division_items.get(div_label, [])
        if not items:
            continue

        # Division header
        insert_p(div_label)

        # Bullet items
        for item_text in items:
            if item_text.strip():
                insert_list_item(item_text.strip())

    # Blank lines
    insert_p()
    insert_p()
    insert_p()
    insert_p()

    # Standard exclusions
    if exclusions:
        insert_p(runs=[
            ("Please note the following ", False, None),
            ("Standard ", True, None),
            ("Exclusions", True, None),
            (":", False, None),
        ])
        insert_p()
        for excl in exclusions:
            if excl.strip():
                insert_p(excl.strip())

    # Closing
    insert_p()
    insert_p()
    insert_p()
    insert_p()
    insert_p("Thank you for giving Yorke & Curtis the opportunity to work with you on this project. Please let me know if you have any questions.")
    insert_p()

    # Signature
    insert_p("Sincerely,\nYorke & Curtis, Inc.")
    insert_p()

    pm_name = project_info.get("pm_name", "")
    if pm_name:
        insert_p(pm_name)

    insert_p()
    insert_p()

    # CC line
    cc_line = project_info.get("cc_line", "Jeremiah Dodson & Erik Timmons")
    insert_p(runs=[
        ("C.c. ", False, None),
        (cc_line, False, None),
        (" – Yorke & Curtis, Inc.", False, None),
    ])

    doc.save(output_path)
    return output_path


# ── GUI ──────────────────────────────────────────────────────────────────────

class ScopeItemPicker(tk.Toplevel):
    """Dialog to pick scope items from database for a specific division.
    Uses a Treeview for fast rendering of large lists."""

    CHECK_ON = "\u2611"   # ☑
    CHECK_OFF = "\u2610"  # ☐

    def __init__(self, parent, db, division_label, division_db_name, existing_items=None):
        super().__init__(parent)
        self.title(f"Select Scope Items — {division_label}")
        self.geometry("1100x750")
        self.transient(parent)
        self.grab_set()

        self.db = db
        self.division_label = division_label
        self.division_db_name = division_db_name
        self.result = None

        # selected: scope_text -> True
        self.selected = set()
        self.custom_items = []
        # Map treeview iid -> scope_text
        self._iid_to_text = {}
        # Debounce timer for text search
        self._search_timer = None

        if existing_items:
            self.selected.update(existing_items)

        self._build_ui()
        self._load_items()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        # ── Filter bar (row 1): Search + Scope Type ──
        filter1 = ttk.Frame(self, padding=(5, 5, 5, 2))
        filter1.pack(fill=tk.X)

        ttk.Label(filter1, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_text_filter_change)
        search_entry = ttk.Entry(filter1, textvariable=self.search_var, width=35)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.focus_set()

        ttk.Label(filter1, text="Scope Type:").pack(side=tk.LEFT, padx=(15, 0))
        self.type_var = tk.StringVar(value="All")
        type_combo = ttk.Combobox(
            filter1, textvariable=self.type_var, width=12,
            values=["All", "includes", "excludes", "assumes", "allowance", "recommends", "note"],
            state="readonly"
        )
        type_combo.pack(side=tk.LEFT, padx=5)
        type_combo.bind("<<ComboboxSelected>>", self._on_combo_filter_change)

        self.show_selected_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filter1, text="Selected only",
                        variable=self.show_selected_var,
                        command=self._on_combo_filter_change).pack(side=tk.LEFT, padx=(15, 0))

        # ── Filter bar (row 2): Project Type + Year Range + Project Search ──
        filter2 = ttk.Frame(self, padding=(5, 2, 5, 5))
        filter2.pack(fill=tk.X)

        ttk.Label(filter2, text="Project Type:").pack(side=tk.LEFT)
        self.project_type_var = tk.StringVar(value="All Types")
        pt_combo = ttk.Combobox(
            filter2, textvariable=self.project_type_var, width=22,
            values=PROJECT_TYPES, state="readonly"
        )
        pt_combo.pack(side=tk.LEFT, padx=5)
        pt_combo.bind("<<ComboboxSelected>>", self._on_combo_filter_change)

        ttk.Label(filter2, text="Year From:").pack(side=tk.LEFT, padx=(15, 0))
        self.year_from_var = tk.StringVar(value="")
        year_from = ttk.Combobox(
            filter2, textvariable=self.year_from_var, width=6,
            values=[""] + [str(y) for y in range(2026, 2010, -1)],
            state="readonly"
        )
        year_from.pack(side=tk.LEFT, padx=3)
        year_from.bind("<<ComboboxSelected>>", self._on_combo_filter_change)

        ttk.Label(filter2, text="To:").pack(side=tk.LEFT, padx=(5, 0))
        self.year_to_var = tk.StringVar(value="")
        year_to = ttk.Combobox(
            filter2, textvariable=self.year_to_var, width=6,
            values=[""] + [str(y) for y in range(2026, 2010, -1)],
            state="readonly"
        )
        year_to.pack(side=tk.LEFT, padx=3)
        year_to.bind("<<ComboboxSelected>>", self._on_combo_filter_change)

        ttk.Label(filter2, text="Project:").pack(side=tk.LEFT, padx=(15, 0))
        self.project_search_var = tk.StringVar()
        self.project_search_var.trace_add("write", self._on_text_filter_change)
        ttk.Entry(filter2, textvariable=self.project_search_var, width=20).pack(
            side=tk.LEFT, padx=5)

        ttk.Button(filter2, text="Reset Filters", command=self._reset_filters).pack(
            side=tk.LEFT, padx=(10, 0))

        # ── Results count ──
        self.results_label = ttk.Label(self, text="", padding=(5, 0))
        self.results_label.pack(anchor="w")

        # ── Treeview for items ──
        tree_frame = ttk.Frame(self, padding=5)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("freq", "type", "text")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings",
                                 selectmode="none")

        self.tree.heading("#0", text="", anchor="w")
        self.tree.heading("freq", text="Freq", anchor="e")
        self.tree.heading("type", text="Type", anchor="w")
        self.tree.heading("text", text="Scope Item", anchor="w")

        self.tree.column("#0", width=30, minwidth=30, stretch=False)
        self.tree.column("freq", width=50, minwidth=40, stretch=False, anchor="e")
        self.tree.column("type", width=75, minwidth=60, stretch=False)
        self.tree.column("text", width=900, minwidth=200, stretch=True)

        # Color tags
        self.tree.tag_configure("includes", foreground="#2e7d32")
        self.tree.tag_configure("excludes", foreground="#c62828")
        self.tree.tag_configure("assumes", foreground="#e65100")
        self.tree.tag_configure("allowance", foreground="#1565c0")
        self.tree.tag_configure("recommends", foreground="#6a1b9a")
        self.tree.tag_configure("note", foreground="#616161")
        self.tree.tag_configure("custom", foreground="#00897b")
        self.tree.tag_configure("checked", background="#e8f5e9")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Toggle check on click
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        # Also toggle on space/return
        self.tree.bind("<space>", self._on_tree_toggle_key)
        self.tree.bind("<Return>", self._on_tree_toggle_key)

        # ── Custom item entry ──
        custom_frame = ttk.LabelFrame(self, text="Add Custom Item", padding=5)
        custom_frame.pack(fill=tk.X, padx=5, pady=5)
        self.custom_var = tk.StringVar()
        custom_entry = ttk.Entry(custom_frame, textvariable=self.custom_var)
        custom_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        custom_entry.bind("<Return>", self._add_custom)
        ttk.Button(custom_frame, text="Add", command=self._add_custom).pack(side=tk.LEFT)

        # ── Bottom buttons ──
        btn_frame = ttk.Frame(self, padding=5)
        btn_frame.pack(fill=tk.X)
        self.count_label = ttk.Label(btn_frame, text="0 items selected",
                                     font=("TkDefaultFont", 10, "bold"))
        self.count_label.pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="OK", command=self._on_ok).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Select All Shown",
                   command=self._select_all).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Deselect All",
                   command=self._deselect_all).pack(side=tk.RIGHT, padx=5)

    def _get_filters(self):
        return {
            "scope_type": self.type_var.get(),
            "project_type": self.project_type_var.get(),
            "year_from": self.year_from_var.get() or None,
            "year_to": self.year_to_var.get() or None,
            "project_search": self.project_search_var.get().strip() or None,
        }

    def _reset_filters(self):
        self.search_var.set("")
        self.type_var.set("All")
        self.project_type_var.set("All Types")
        self.year_from_var.set("")
        self.year_to_var.set("")
        self.project_search_var.set("")
        self.show_selected_var.set(False)
        self._load_items()

    def _load_items(self):
        """Load scope items into the Treeview."""
        self.tree.delete(*self.tree.get_children())
        self._iid_to_text.clear()

        search_text = self.search_var.get().strip()
        filters = self._get_filters()

        if search_text and len(search_text) >= 2:
            items = self.db.search_items(search_text, self.division_db_name, **filters)
        else:
            items = self.db.get_division_items(self.division_db_name, **filters)

        if self.show_selected_var.get():
            items = [i for i in items if i["scope_text"] in self.selected]

        for item in items:
            text = item["scope_text"]
            freq = item["freq"]
            stype = item["scope_type"]
            is_checked = text in self.selected
            check = self.CHECK_ON if is_checked else self.CHECK_OFF
            tags = (stype,)
            if is_checked:
                tags = (stype, "checked")

            iid = self.tree.insert("", "end", text=check,
                                   values=(f"{freq}x", stype, text),
                                   tags=tags)
            self._iid_to_text[iid] = text

        # Custom items
        for custom in self.custom_items:
            is_checked = custom in self.selected
            check = self.CHECK_ON if is_checked else self.CHECK_OFF
            tags = ("custom",)
            if is_checked:
                tags = ("custom", "checked")
            iid = self.tree.insert("", "end", text=check,
                                   values=("new", "custom", custom),
                                   tags=tags)
            self._iid_to_text[iid] = custom

        shown = len(self._iid_to_text)
        self.results_label.config(text=f"Showing {shown} items")
        self._update_count()

    def _toggle_item(self, iid):
        """Toggle the checked state of a treeview item."""
        text = self._iid_to_text.get(iid)
        if not text:
            return
        if text in self.selected:
            self.selected.discard(text)
            self.tree.item(iid, text=self.CHECK_OFF)
            # Remove 'checked' from tags
            tags = [t for t in self.tree.item(iid, "tags") if t != "checked"]
            self.tree.item(iid, tags=tuple(tags))
        else:
            self.selected.add(text)
            self.tree.item(iid, text=self.CHECK_ON)
            tags = list(self.tree.item(iid, "tags")) + ["checked"]
            self.tree.item(iid, tags=tuple(tags))
        self._update_count()

    def _on_tree_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self._toggle_item(iid)

    def _on_tree_toggle_key(self, event):
        for iid in self.tree.selection():
            self._toggle_item(iid)

    def _on_text_filter_change(self, *args):
        """Debounced handler for text entry changes (300ms delay)."""
        if self._search_timer is not None:
            self.after_cancel(self._search_timer)
        self._search_timer = self.after(300, self._load_items)

    def _on_combo_filter_change(self, *args):
        """Immediate handler for dropdown/checkbox changes."""
        self._load_items()

    def _update_count(self):
        count = len(self.selected)
        self.count_label.config(text=f"{count} items selected")

    def _select_all(self):
        for iid, text in self._iid_to_text.items():
            self.selected.add(text)
            self.tree.item(iid, text=self.CHECK_ON)
            tags = list(self.tree.item(iid, "tags"))
            if "checked" not in tags:
                tags.append("checked")
            self.tree.item(iid, tags=tuple(tags))
        self._update_count()

    def _deselect_all(self):
        self.selected.clear()
        for iid in self._iid_to_text:
            self.tree.item(iid, text=self.CHECK_OFF)
            tags = [t for t in self.tree.item(iid, "tags") if t != "checked"]
            self.tree.item(iid, tags=tuple(tags))
        self._update_count()

    def _add_custom(self, event=None):
        text = self.custom_var.get().strip()
        if text:
            self.custom_items.append(text)
            self.selected.add(text)
            self.custom_var.set("")
            self._load_items()

    def _on_ok(self):
        self.result = list(self.selected)
        self.destroy()


class ProposalBuilderApp:
    """Main application window."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Y&C Proposal Builder — Estimate & Clarifications")
        self.root.geometry("1000x800")

        if not DB_PATH.exists():
            messagebox.showerror("Database Not Found",
                f"Scope database not found at:\n{DB_PATH}\n\n"
                "Run build_scope_db.py first to create it.")
            sys.exit(1)

        self.db = ScopeDatabase(DB_PATH)
        self.division_items = {}
        self.exclusion_vars = {}

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        style.configure("TNotebook.Tab", padding=[12, 4])

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ── Tab 1: Project Info ──
        info_frame = ttk.Frame(notebook, padding=15)
        notebook.add(info_frame, text="  Project Info  ")

        fields = [
            ("Date:", "date", datetime.date.today().strftime("%m-%d-%y")),
            ("Client Name:", "client_name", ""),
            ("Client Address:", "client_address", ""),
            ("Project Name:", "project_name", ""),
            ("Architect/Drawings By:", "architect", ""),
            ("Drawing Date:", "drawing_date", ""),
            ("Estimate Amount ($):", "estimate_amount", ""),
            ("Dear Line:", "dear_line", "Dear Client:"),
            ("PM Name & Title:", "pm_name", ""),
            ("CC Line:", "cc_line", "Jeremiah Dodson & Erik Timmons"),
        ]

        self.info_vars = {}
        for i, (label, key, default) in enumerate(fields):
            ttk.Label(info_frame, text=label, font=("TkDefaultFont", 11)).grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=4)
            var = tk.StringVar(value=default)
            self.info_vars[key] = var
            entry = ttk.Entry(info_frame, textvariable=var, width=60, font=("TkDefaultFont", 11))
            entry.grid(row=i, column=1, sticky="ew", pady=4)

        row = len(fields)
        ttk.Label(info_frame, text="Estimate Type:", font=("TkDefaultFont", 11)).grid(
            row=row, column=0, sticky="e", padx=(0, 10), pady=4)
        self.est_type_var = tk.StringVar(value="Preliminary Estimate")
        est_combo = ttk.Combobox(
            info_frame, textvariable=self.est_type_var, values=ESTIMATE_TYPES,
            width=30, font=("TkDefaultFont", 11), state="readonly"
        )
        est_combo.grid(row=row, column=1, sticky="w", pady=4)
        info_frame.columnconfigure(1, weight=1)

        # ── Tab 2: Scope Items by Division ──
        scope_frame = ttk.Frame(notebook, padding=10)
        notebook.add(scope_frame, text="  Scope Items  ")

        div_list_frame = ttk.Frame(scope_frame)
        div_list_frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(div_list_frame)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="Select scope items for each division:",
                  font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT)

        self.div_buttons = {}
        self.div_count_labels = {}

        canvas = tk.Canvas(div_list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(div_list_frame, orient=tk.VERTICAL, command=canvas.yview)
        div_inner = ttk.Frame(canvas)

        div_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=div_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for div_num, div_label, div_db_name in DIVISIONS:
            row_frame = ttk.Frame(div_inner, padding=5)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=div_label, font=("TkDefaultFont", 11),
                      width=45, anchor="w").pack(side=tk.LEFT)

            count_label = ttk.Label(row_frame, text="0 items", width=10,
                                    foreground="#888888")
            count_label.pack(side=tk.LEFT, padx=10)
            self.div_count_labels[div_label] = count_label

            btn = ttk.Button(row_frame, text="Select Items...",
                command=lambda dl=div_label, ddn=div_db_name: self._open_picker(dl, ddn))
            btn.pack(side=tk.LEFT, padx=5)

            ttk.Button(row_frame, text="Clear",
                command=lambda dl=div_label: self._clear_division(dl)).pack(side=tk.LEFT)

        # ── Tab 3: Standard Exclusions ──
        excl_frame = ttk.Frame(notebook, padding=15)
        notebook.add(excl_frame, text="  Standard Exclusions  ")

        ttk.Label(excl_frame, text="Select standard exclusions to include:",
                  font=("TkDefaultFont", 11, "bold")).pack(anchor="w", pady=(0, 10))

        btn_row = ttk.Frame(excl_frame)
        btn_row.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btn_row, text="Select All",
            command=lambda: self._toggle_all_exclusions(True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Deselect All",
            command=lambda: self._toggle_all_exclusions(False)).pack(side=tk.LEFT)

        for excl in STANDARD_EXCLUSIONS:
            var = tk.BooleanVar(value=True)
            self.exclusion_vars[excl] = var
            cb_frame = ttk.Frame(excl_frame)
            cb_frame.pack(anchor="w", fill=tk.X, pady=2)
            ttk.Checkbutton(cb_frame, text="", variable=var).pack(side=tk.LEFT)
            ttk.Label(cb_frame, text=excl, wraplength=800).pack(side=tk.LEFT)

        custom_excl_frame = ttk.LabelFrame(excl_frame, text="Add Custom Exclusion", padding=5)
        custom_excl_frame.pack(fill=tk.X, pady=10)
        self.custom_excl_var = tk.StringVar()
        ttk.Entry(custom_excl_frame, textvariable=self.custom_excl_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(custom_excl_frame, text="Add",
            command=self._add_custom_exclusion).pack(side=tk.LEFT)

        # ── Bottom: Generate button ──
        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill=tk.X)

        self.summary_label = ttk.Label(bottom, text="Total: 0 scope items across 0 divisions",
                                       font=("TkDefaultFont", 10))
        self.summary_label.pack(side=tk.LEFT)

        ttk.Button(bottom, text="Generate Proposal",
            command=self._generate, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)

        try:
            style = ttk.Style()
            style.configure("Accent.TButton", font=("TkDefaultFont", 12, "bold"))
        except Exception:
            pass

    def _open_picker(self, div_label, div_db_name):
        existing = self.division_items.get(div_label, [])
        picker = ScopeItemPicker(self.root, self.db, div_label, div_db_name, existing)
        self.root.wait_window(picker)

        if picker.result is not None:
            self.division_items[div_label] = picker.result
            count = len(picker.result)
            self.div_count_labels[div_label].config(
                text=f"{count} items",
                foreground="#2e7d32" if count > 0 else "#888888"
            )
            self._update_summary()

    def _clear_division(self, div_label):
        self.division_items[div_label] = []
        self.div_count_labels[div_label].config(text="0 items", foreground="#888888")
        self._update_summary()

    def _toggle_all_exclusions(self, state):
        for var in self.exclusion_vars.values():
            var.set(state)

    def _add_custom_exclusion(self):
        text = self.custom_excl_var.get().strip()
        if text:
            var = tk.BooleanVar(value=True)
            self.exclusion_vars[text] = var
            self.custom_excl_var.set("")
            messagebox.showinfo("Added", f"Custom exclusion added:\n{text}")

    def _update_summary(self):
        total_items = sum(len(items) for items in self.division_items.values())
        div_count = sum(1 for items in self.division_items.values() if items)
        self.summary_label.config(
            text=f"Total: {total_items} scope items across {div_count} divisions")

    def _generate(self):
        template_path = get_template_path()
        if not template_path:
            messagebox.showerror("Template Not Found",
                "Could not find the Estimate & Clarifications Template.\n\n"
                "Expected locations:\n"
                f"  {_BUNDLED_TEMPLATE}\n"
                f"  {_ONEDRIVE_TEMPLATE}")
            return

        project_name = self.info_vars["project_name"].get().strip()
        if not project_name:
            messagebox.showwarning("Missing Info", "Please enter a Project Name.")
            return

        total_items = sum(len(items) for items in self.division_items.values())
        if total_items == 0:
            if not messagebox.askyesno("No Scope Items",
                    "No scope items selected. Generate anyway?"):
                return

        project_info = {
            "date": self.info_vars["date"].get(),
            "client_name": self.info_vars["client_name"].get(),
            "client_address": self.info_vars["client_address"].get(),
            "project_name": project_name,
            "architect": self.info_vars["architect"].get(),
            "drawing_date": self.info_vars["drawing_date"].get(),
            "estimate_type": self.est_type_var.get(),
            "estimate_amount": self.info_vars["estimate_amount"].get(),
            "dear_line": self.info_vars["dear_line"].get(),
            "pm_name": self.info_vars["pm_name"].get(),
            "cc_line": self.info_vars["cc_line"].get(),
        }

        exclusions = [text for text, var in self.exclusion_vars.items() if var.get()]

        default_name = f"{project_name} Estimate & Clarifications {project_info['date']}.docx"
        output_path = filedialog.asksaveasfilename(
            title="Save Proposal As",
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Desktop"),
        )

        if not output_path:
            return

        try:
            generate_proposal(
                template_path, output_path, project_info,
                self.division_items, exclusions
            )
            messagebox.showinfo("Success",
                f"Proposal generated!\n\n{output_path}\n\n"
                f"{total_items} scope items across "
                f"{sum(1 for i in self.division_items.values() if i)} divisions")

            if sys.platform == "darwin":
                os.system(f'open "{output_path}"')
            elif sys.platform == "win32":
                os.startfile(output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate proposal:\n\n{e}")
            import traceback
            traceback.print_exc()

    def run(self):
        self.root.mainloop()
        self.db.close()


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = ProposalBuilderApp()
    app.run()
