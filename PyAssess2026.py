#!/usr/bin/env python3

# PyAssess2026.py - generates processed exam grids for Physics@Manchester
# Author: Clive Dickinson
# Date: 2026-05-30
# Version: 0.0.1

import argparse
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===========================================================================
# Constants — update as required
# ===========================================================================
AY    = 2025                  # Academic year (2025 = AY 2024-25 etc.)
INDIR = "data2025_newcodes"   # Directory containing input exam grid spreadsheets
OUTDIR = "./"                # Output directory for generated spreadsheets

# ===========================================================================
# Per-class-year file mappings
# ===========================================================================
# classyear keys:
#   1 / 1m  = Year 1  Physics / Maths+Physics
#   2 / 2m  = Year 2  Physics / Maths+Physics
#  31 / 31m = Year 3  MPhys progressing / MMath progressing
#  32 / 32m = Year 3  BSc Physics graduating / BSc MathsPhysics graduating
#   4 / 4m  = Year 4  MPhys (final year) / MMath (final year)
#
# Values: (input_filename, output_filename)

CLASSYEAR_FILES = {
    '1':   ('PHYS_1241_S2_Y1_Exam_Grids.xlsx',         f'1styear_Physics.AY{AY}.xlsx'),
    '1m':  ('PHYS_1241_S2_Y1_MP_Exam_Grids.xlsx',      f'1styear_MathsPhysics.AY{AY}.xlsx'),
    '2':   ('PHYS_1241_S2_Y2_Exam_Grids.xlsx',         f'2ndyear_Physics.AY{AY}.xlsx'),
    '2m':  ('PHYS_1241_S2_Y2_MP_Exam_Grids.xlsx',      f'2ndyear_MathsPhysics.AY{AY}.xlsx'),
    '31':  ('PHYS_1241_S2_Y3_PROG_Exam_Grids.xlsx',    f'3rdyear_MPhys.AY{AY}.xlsx'),
    '31m': ('PHYS_1241_S2_Y3_MP_PROG_Exam_Grids.xlsx', f'3rdyear_MMath.AY{AY}.xlsx'),
    '32':  ('PHYS_1241_S2_Y3_GRAD_Exam_Grids.xlsx',    f'FinalYear_BSc_Physics.AY{AY}.xlsx'),
    '32m': ('PHYS_1241_S2_Y3_MP_GRAD_Exam_Grids.xlsx', f'FinalYear_BSc_MathsPhysics.AY{AY}.xlsx'),
    '4':   ('PHYS_1241_S2_Y4_Exam_Grids.xlsx',         f'FinalYear_MPhys.AY{AY}.xlsx'),
    '4m':  ('PHYS_1241_S2_Y4_MP_Exam_Grids.xlsx',      f'FinalYear_MMath.AY{AY}.xlsx'),
}

ALL_CLASSYEARS = list(CLASSYEAR_FILES.keys())

# ===========================================================================
# Excel sheet layout (same across all input files)
# ===========================================================================
SHEET_NAME     = 'Style A Plus With Gradebook'
UNIT_LABEL_ROW = 4   # 0-indexed: row 5 — contains 'Unit 1', 'Unit 2', ...
COL_HEADER_ROW = 5   # 0-indexed: row 6 — contains 'Module', 'Mark', 'EN', ...
DATA_START_ROW = 6   # 0-indexed: row 7 — first student row

# Within each unit block the four columns of interest sit at these offsets
# from the unit's start column (Module, Link1, Link2, Mark, EN, Mit Circs, ...)
_UNIT_COL_OFFSETS = {'module': 0, 'mark': 3, 'en': 4, 'mit_circs': 5}

# ===========================================================================
# Data classes
# ===========================================================================

_MODULE_RE = re.compile(r'^(.*?)\s*\((\d+)\)\s*$')


def _append_code(existing, new_code):
    """Append *new_code* to *existing* output code, joining with '_' if needed."""
    return f"{existing}_{new_code}" if existing else new_code


_CODE_SPLIT_RE = re.compile(r'[\s,;]+')

def _split_codes(value):
    """Return a frozenset of individual codes from a field that may hold several.

    Handles space-, comma-, and semicolon-separated values, e.g. 'A, EA' or 'EA AA'.
    """
    if not value:
        return frozenset()
    return frozenset(c for c in _CODE_SPLIT_RE.split(str(value).strip()) if c)


# Mit_circs codes that are specifically actioned in calc_credits().
# Any code NOT in this set is copied to the output code column as-is.
_PROCESSED_MIT_CODES = frozenset({'AA', 'EA'})


def _parse_module(module):
    """Split 'PHYS10180 (20)' into ('PHYS10180', 20).

    Returns (coursename, credits) where credits is an int, or
    (module, None) if the string doesn't match the expected pattern.
    """
    if module is None:
        return None, None
    m = _MODULE_RE.match(str(module))
    if m:
        return m.group(1).strip(), int(m.group(2))
    return module, None


class UnitInfo:
    """Data for a single unit (module) for one student."""
    __slots__ = ('unit_name', 'module', 'coursename', 'credits', 'mark', 'en', 'mit_circs',
                 'passed', 'excluded', 'output_code')

    def __init__(self, unit_name, module, mark, en, mit_circs):
        self.unit_name   = unit_name              # 'Unit 1', 'Unit 2', etc.
        self.module      = module                 # raw cell value, e.g. 'PHYS10180 (20)'
        self.coursename, self.credits = _parse_module(module)
        self.mark        = mark                   # unit mark (float) or None
        self.en          = en                     # EN flag (str) or None
        self.mit_circs   = mit_circs              # mitigating circumstances code (str) or None
        self.passed      = None                   # True/False once calc_credits() runs; None = no mark
        self.excluded    = False                  # True if excluded from year mark
        self.output_code = None                   # code(s) written to the output code column

    def __repr__(self):
        return (f'UnitInfo({self.unit_name}, coursename={self.coursename!r}, '
                f'credits={self.credits}, mark={self.mark}, passed={self.passed}, '
                f'excluded={self.excluded}, output_code={self.output_code!r})')


class StudentInfo:
    """All data for a single student, read from one row of the exam grid."""
    __slots__ = (
        'emplid', 'name', 'id_no', 'uf', 'mc', 'bz',
        'admit_term', 'entry_type', 'psi', 'plan',
        'units_passed', 'award', 'classification',
        'units', 'trailing',
        'yearmark',
        'credits_taken', 'credits_passed', 'credits_excluded', 'creds_passed_taken',
        'resits',
    )

    def __init__(self):
        self.emplid         = None
        self.name           = None
        self.id_no          = None   # sequential row number in the grid
        self.uf             = None
        self.mc             = None
        self.bz             = None
        self.admit_term     = None
        self.entry_type     = None
        self.psi            = None
        self.plan           = None   # degree programme, e.g. 'MPhys(Hons) Physics'
        self.units_passed   = None
        self.award          = None
        self.classification = None
        self.units          = []     # list of UnitInfo, one per unit column block
        self.trailing       = {}     # trailing columns: normalised name -> value
        self.yearmark           = None   # credit-weighted average of unit marks
        self.credits_taken      = None   # total credits with a mark
        self.credits_passed     = None   # credits where mark > 39.95
        self.credits_excluded   = 0      # credits excluded from calculation (populated later)
        self.creds_passed_taken = None   # formatted string for output, e.g. '120 / 120'
        self.resits             = None   # deferred/resit courses for output, e.g. 'PHYS10071[1] / PHYS10101[1]'

    def calc_credits(self):
        """Set credits_taken/passed/excluded, creds_passed_taken, and unit flags.

        Exclusion rules applied here:
          AA in mit_circs → unit is excluded from year mark, treated as passed,
                            output_code set to 'X'.
        Excluded credits still count towards credits_passed.
        """
        taken      = 0
        passed     = 0
        excluded   = 0
        resit_list = []

        for unit in self.units:
            if unit.module is None or unit.credits is None:
                continue                             # empty slot

            taken     += unit.credits
            mit_codes  = _split_codes(unit.mit_circs)
            en_codes   = _split_codes(unit.en)
            used_mit   = set()                       # mit codes consumed by a rule
            used_en    = set()                       # EN codes consumed by a rule

            # --- EA deferral (highest priority; trumps AA and all other mit codes) ---
            if 'EA' in mit_codes:
                unit.excluded = True
                unit.passed   = True
                if 'XN' in en_codes:                 # missed exam → XL_R1
                    unit.output_code = _append_code(unit.output_code, 'XL_R1')
                    used_en.add('XN')
                else:
                    unit.output_code = _append_code(unit.output_code, 'R1')
                excluded += unit.credits
                passed   += unit.credits
                resit_list.append(f"{unit.coursename or unit.module}[1]")
                used_mit.update(mit_codes)           # EA trumps: mark all mit codes used

            # --- AA exclusion ---
            elif 'AA' in mit_codes:
                unit.excluded    = True
                unit.passed      = True
                unit.output_code = _append_code(unit.output_code, 'X')
                excluded += unit.credits
                passed   += unit.credits
                used_mit.add('AA')

            # --- no mark: exclude (treat as passed, omit from year mark) ---
            elif unit.mark is None:
                unit.excluded = True
                unit.passed   = True
                excluded += unit.credits
                passed   += unit.credits

            # --- normal pass/fail ---
            else:
                try:
                    unit.passed = float(unit.mark) > 39.95
                except (TypeError, ValueError):
                    unit.passed = False              # non-numeric mark counts as fail
                if unit.passed:
                    passed += unit.credits

            # --- copy through any unprocessed EN and mit_circs codes ---
            for code in sorted(en_codes - used_en):
                unit.output_code = _append_code(unit.output_code, code)
            for code in sorted(mit_codes - used_mit - _PROCESSED_MIT_CODES):
                unit.output_code = _append_code(unit.output_code, code)

        self.credits_taken      = taken
        self.credits_passed     = passed
        self.credits_excluded   = excluded
        self.creds_passed_taken = f"{passed} / {taken}"
        self.resits             = ' / '.join(resit_list) if resit_list else None

    def calc_yearmark(self):
        """Set self.yearmark to the credit-weighted mean of all unit marks.

        Units with a missing mark or missing credits are skipped.
        Result is rounded to 1 decimal place; None if no valid units found.
        """
        total_credits  = 0
        weighted_marks = 0.0
        for unit in self.units:
            if unit.excluded:
                continue                          # excluded units don't count toward year mark
            if unit.mark is None or unit.credits is None:
                continue
            try:
                mark = float(unit.mark)
            except (TypeError, ValueError):
                continue
            weighted_marks += mark * unit.credits
            total_credits  += unit.credits
        self.yearmark = (round(weighted_marks / total_credits, 1)
                         if total_credits > 0 else None)

    def __repr__(self):
        return (f'StudentInfo(emplid={self.emplid!r}, name={self.name!r}, '
                f'units={len(self.units)})')


# ===========================================================================
# Excel reading
# ===========================================================================

# Student-info columns are fixed (left-hand side of the sheet)
_STUDENT_COLS = [
    ('emplid',          0),
    ('name',            1),
    ('id_no',           2),
    ('uf',              3),
    ('mc',              4),
    ('bz',              5),
    ('admit_term',      6),
    ('entry_type',      7),
    ('psi',             8),
    ('plan',            9),
    ('units_passed',   10),
    ('award',          11),
    ('classification', 12),
]

# Sub-header names that belong to a unit block (used to find where trailing
# columns begin after the last unit)
_UNIT_SUBHEADERS = {'Module', 'Link1', 'Link2', 'Mark', 'EN', 'GBN'}


def _cell(row, col):
    """Return row[col] as a Python value, or None if missing/NaN."""
    val = row.iloc[col]
    return None if pd.isna(val) else val


def _norm_header(val):
    """Normalise a column header to a clean string key."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    return str(val).replace('\n', ' ').strip()


def read_students(filepath):
    """Read 'Style A Plus With Gradebook' from *filepath*.

    Returns a list of StudentInfo objects, one per student row.
    """
    df = pd.read_excel(filepath, sheet_name=SHEET_NAME,
                       header=None, dtype=object)

    unit_row   = df.iloc[UNIT_LABEL_ROW]
    header_row = df.iloc[COL_HEADER_ROW]
    n_cols     = len(header_row)

    # ------------------------------------------------------------------
    # Locate unit blocks: columns where the unit-label row has 'Unit N'
    # ------------------------------------------------------------------
    unit_starts = sorted(
        c for c, val in enumerate(unit_row)
        if isinstance(val, str) and val.startswith('Unit ')
    )
    if not unit_starts:
        raise ValueError(f"No unit columns found in {filepath}")

    # Trailing section starts at the first column after the last unit block
    # whose header is not a recognised unit sub-header.
    last_unit_start = unit_starts[-1]
    trailing_start  = n_cols   # default: no trailing
    for c in range(last_unit_start, n_cols):
        h = _norm_header(header_row.iloc[c])
        is_unit_sub = (
            h in _UNIT_SUBHEADERS
            or h.startswith('AM')
            or 'Mit' in h
            or h == ''
        )
        if not is_unit_sub:
            trailing_start = c
            break

    unit_ends = unit_starts[1:] + [trailing_start]

    # For each unit, record the absolute column index for the 4 fields
    # using fixed offsets (Module+0, Mark+3, EN+4, Mit Circs+5).
    unit_col_map = []   # list of (unit_name, module_c, mark_c, en_c, mit_c)
    for start, end in zip(unit_starts, unit_ends):
        unit_name = unit_row.iloc[start]
        cols = {field: start + offset
                for field, offset in _UNIT_COL_OFFSETS.items()
                if start + offset < end}
        unit_col_map.append((
            unit_name,
            cols.get('module'),
            cols.get('mark'),
            cols.get('en'),
            cols.get('mit_circs'),
        ))

    # Trailing column names (normalised)
    trailing_cols = [
        (c, _norm_header(header_row.iloc[c]) or f'col_{c}')
        for c in range(trailing_start, n_cols)
    ]

    # ------------------------------------------------------------------
    # Parse student rows: skip any row where emplid (col 0) is blank
    # ------------------------------------------------------------------
    students = []
    for _, row in df.iloc[DATA_START_ROW:].iterrows():
        if pd.isna(row.iloc[0]):
            continue

        s = StudentInfo()
        for attr, col in _STUDENT_COLS:
            setattr(s, attr, _cell(row, col))

        s.units = [
            UnitInfo(
                unit_name,
                _cell(row, mc),
                _cell(row, mk),
                _cell(row, en),
                _cell(row, mit),
            )
            for unit_name, mc, mk, en, mit in unit_col_map
        ]

        s.trailing = {name: _cell(row, c) for c, name in trailing_cols}

        students.append(s)

    return students


# ===========================================================================
# Output configuration
# ===========================================================================

# Trailing columns written after the unit pairs, keyed by classyear.
# Values are read directly from the "2 Line format" sheet of each input file.
TRAILING_COLS = {
    '1':   ['Creds Passed/Taken', 'Year Mark', 'Status', 'Fail reason',
            'Resits', 'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '1m':  ['Creds Passed/Taken', 'Phys Year Mark', 'Math Year Mark', 'Year Mark',
            'Status', 'Fail reason', 'Resits', 'Notes',
            'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '2':   ['Creds Passed/Taken', 'Phys 1', 'Year Mark', 'Status', 'Fail reason',
            'Resits', 'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '2m':  ['Creds Passed/Taken', 'Phys 1', 'Phys Year Mark', 'Math Year Mark',
            'Year Mark', 'Status', 'Fail reason', 'Resits', 'Notes',
            'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '31':  ['Creds Passed/Taken', 'L3/L4 creds passed', 'Phys 1', 'Phys 2',
            'BZ', 'Year Mark', 'Overall', 'Status', 'Award', 'Notes',
            'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '31m': ['Creds Passed/Taken', 'L3/L4 creds passed', 'Phys 1', 'Phys 2',
            'Phys Year Mark', 'Math Year Mark', 'BZ', 'Year Mark', 'Overall',
            'Status', 'Award', 'Fail reason', 'Notes',
            'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '32':  ['Creds Passed/Taken', 'L3/L4 creds passed', 'Phys 1', 'Phys 2',
            'BZ', 'Year Mark', 'Overall', 'Deg Class Alg', 'Deg Class Rev',
            'Deg Class Actual', 'Fail reason', 'Award', 'Classification',
            'Award Alg', 'Award Actual', 'Classification Alg',
            'Classification Actual', 'Award Change', 'Classification Change',
            'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '32m': ['Creds Passed/Taken', 'L3/L4 creds passed', 'Phys 1', 'Phys 2',
            'Phys Year Mark', 'Math Year Mark', 'BZ', 'Year Mark', 'Overall',
            'Deg Class Alg', 'Deg Class Rev', 'Deg Class Actual', 'Fail reason',
            'Award', 'Classification', 'Award Alg', 'Award Actual',
            'Classification Alg', 'Classification Actual',
            'Award Change', 'Classification Change',
            'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '4':   ['Creds Passed/Taken', 'Y3 creds failed w/wo MCs',
            'L4 creds passed Y3+Y4', 'Phys 1', 'Phys 2', 'Phys 3',
            'BZ', 'Year Mark', 'Overall', 'Deg Class Alg', 'Deg Class Rev',
            'Deg Class Actual', 'Fail reason', 'Award', 'Classification',
            'Award Alg', 'Award Actual', 'Classification Alg',
            'Classification Actual', 'Award Change', 'Classification Change',
            'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
    '4m':  ['Creds Passed/Taken', 'Y3 creds failed w/wo MCs',
            'L4 creds passed Y3+Y4', 'Phys 1', 'Phys 2', 'Phys 3',
            'Phys Year Mark', 'Math Year Mark', 'BZ', 'Year Mark', 'Overall',
            'Deg Class Alg', 'Deg Class Rev', 'Deg Class Actual', 'Fail reason',
            'Award', 'Classification', 'Award Alg', 'Award Actual',
            'Classification Alg', 'Classification Actual',
            'Award Change', 'Classification Change',
            'Notes', 'Pre-Exam Board Minutes', 'Exam Board Minutes'],
}

# Column widths (Excel character units), measured from the "2 Line format" sheet.
# '_unit' and '_code' are the two columns of each unit pair.
_COL_WIDTHS = {
    'ID No.':                    4.66,
    'Emplid':                    9.66,
    'Name':                     11.66,
    'Plan':                     16.66,
    '_unit':                     7.66,
    '_code':                     7.66,
    'Creds Passed/Taken':       16.66,
    'Year Mark':                10.66,
    'Phys Year Mark':           10.66,
    'Math Year Mark':           10.66,
    'Status':                   10.66,
    'Fail reason':              16.66,
    'Resits':                   50.16,
    'Notes':                    36.66,
    'Pre-Exam Board Minutes':   26.66,
    'Exam Board Minutes':       26.66,
    'Phys 1':                   10.66,
    'Phys 2':                   10.66,
    'Phys 3':                   10.66,
    'BZ':                        8.00,
    'Overall':                  10.66,
    'L3/L4 creds passed':       16.66,
    'L4 creds passed Y3+Y4':    18.66,
    'Y3 creds failed w/wo MCs': 18.66,
    'Award':                    10.66,
    'Classification':           14.66,
    'Deg Class Alg':            14.66,
    'Deg Class Rev':            14.66,
    'Deg Class Actual':         14.66,
    'Award Alg':                10.66,
    'Award Actual':             10.66,
    'Classification Alg':       16.66,
    'Classification Actual':    16.66,
    'Award Change':             10.66,
    'Classification Change':    16.66,
}

# Formatting objects (created once, reused for every cell)
_FONT        = Font(name='Aptos Narrow', size=11)
_FONT_BOLD   = Font(name='Aptos Narrow', size=11, bold=True)
_FILL_GREY   = PatternFill(fill_type='solid', fgColor='FFE0E0E0')
_ALIGN_CTR   = Alignment(horizontal='center')
_INFO_ROW_H  = 17.0   # height of student info rows

_SIDE_THIN    = Side(border_style='thin', color='FF808080')   # vertical column separators
_SIDE_THIN_H  = Side(border_style='thin', color='FFD0D0D0')   # horizontal row lines within a student pair
_SIDE_THIN_BK = Side(border_style='thin')                     # horizontal dividers between student pairs (black)
_SIDE_NONE    = Side()

# Fixed left-hand output columns: (header label, StudentInfo attribute name)
_FIXED_COLS = [
    ('ID No.', 'id_no'),
    ('Emplid',  'emplid'),
    ('Name',    'name'),
    ('Plan',    'plan'),
]


# Maps trailing column header names to StudentInfo attribute names.
# Populated as computed attributes are added; write_students uses this to
# fill in values rather than leaving those cells blank.
_TRAILING_ATTR = {
    'Creds Passed/Taken': 'creds_passed_taken',
    'Year Mark':          'yearmark',
    'Resits':             'resits',
}

# Excel number formats applied to trailing columns that hold computed floats.
_TRAILING_FORMAT = {
    'Year Mark':      '0.0',
    'Phys Year Mark': '0.0',
    'Math Year Mark': '0.0',
    'Overall':        '0.0',
    'Phys 1':         '0.0',
    'Phys 2':         '0.0',
    'Phys 3':         '0.0',
}


# ===========================================================================
# Excel writing
# ===========================================================================

def write_students(students, outpath, classyear):
    """Write *students* to *outpath* in 2-row-per-student format.

    Row 1    : bold header — fixed labels, 'Unit N' merged over each pair, trailing labels
    Row 2n   : student info row  — fixed fields, module codes (merged), blank trailing
    Row 2n+1 : student marks row — marks and output codes per unit, blank trailing

    Formatting:
    - Alternating grey fill (FFE0E0E0) on every other student pair
    - Medium (thick) bottom border below the header and after each student pair
    - Thin left border between every column

    Borders are applied to all cells before merging because openpyxl converts
    non-top-left merged cells to read-only MergedCell objects on merge_cells().
    """
    trailer_names = TRAILING_COLS[classyear]
    n_units  = len(students[0].units) if students else 0
    n_fixed  = len(_FIXED_COLS)
    u_start  = n_fixed + 1               # 1-based col of first unit pair
    t_start  = u_start + 2 * n_units     # 1-based col of first trailing col
    last_col = t_start - 1 + len(trailer_names)
    n_rows   = 1 + 2 * len(students)

    # Rows that get a thick bottom border: header + every marks row
    thick_rows = frozenset([1] + [2 + 2*i + 1 for i in range(len(students))])

    wb = Workbook()
    ws = wb.active
    ws.title = 'Assessment'

    # ------------------------------------------------------------------ widths
    for i, (label, _) in enumerate(_FIXED_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = _COL_WIDTHS[label]
    for i in range(n_units):
        ws.column_dimensions[get_column_letter(u_start + 2*i    )].width = _COL_WIDTHS['_unit']
        ws.column_dimensions[get_column_letter(u_start + 2*i + 1)].width = _COL_WIDTHS['_code']
    for j, name in enumerate(trailer_names):
        ws.column_dimensions[get_column_letter(t_start + j)].width = \
            _COL_WIDTHS.get(name, 12.0)

    # ------------------------------------------------------------------ header
    for i, (label, _) in enumerate(_FIXED_COLS, start=1):
        c = ws.cell(row=1, column=i, value=label)
        c.font = _FONT_BOLD
        c.alignment = _ALIGN_CTR
    for i in range(n_units):
        col = u_start + 2 * i
        c = ws.cell(row=1, column=col, value=f'Unit {i + 1}')
        c.font = _FONT_BOLD
        c.alignment = _ALIGN_CTR
        # merge deferred until after border pass
    for j, name in enumerate(trailer_names):
        c = ws.cell(row=1, column=t_start + j, value=name)
        c.font = _FONT_BOLD
        c.alignment = _ALIGN_CTR

    # ------------------------------------------------------------------ rows
    pending_merges = []   # collected here; applied after the border pass

    for idx, s in enumerate(students):
        info_row  = 2 + 2 * idx
        marks_row = info_row + 1
        fill = _FILL_GREY if idx % 2 == 1 else None

        ws.row_dimensions[info_row].height = _INFO_ROW_H

        def _c(row, col, value=None, center=False):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = _FONT
            if fill:
                cell.fill = fill
            if center:
                cell.alignment = _ALIGN_CTR
            return cell

        # info row — fixed columns
        for i, (_, attr) in enumerate(_FIXED_COLS, start=1):
            _c(info_row, i, getattr(s, attr))

        # info row — unit module codes (merge pending)
        for i, unit in enumerate(s.units):
            col = u_start + 2 * i
            _c(info_row, col,     unit.module, center=True)
            _c(info_row, col + 1, None)
            pending_merges.append((info_row, col, col + 1))

        # info row — trailing columns (blank unless already computed)
        for j, tname in enumerate(trailer_names):
            attr  = _TRAILING_ATTR.get(tname)
            value = getattr(s, attr) if attr else None
            cell  = _c(info_row, t_start + j, value)
            fmt   = _TRAILING_FORMAT.get(tname)
            if fmt:
                cell.number_format = fmt

        # marks row — fixed columns (blank; needed for fill and borders)
        for i in range(1, n_fixed + 1):
            _c(marks_row, i)

        # marks row — unit marks and output codes
        for i, unit in enumerate(s.units):
            col = u_start + 2 * i
            _c(marks_row, col,     unit.mark)
            _c(marks_row, col + 1, unit.output_code)

        # marks row — trailing
        for j in range(len(trailer_names)):
            _c(marks_row, t_start + j)

    # ---- border pass (must run before merge_cells) -------------------
    # thick_rows (header + every marks row) get a thin black bottom border.
    # Info rows get a lighter thin bottom border.
    # The output-code column of each unit pair has no left border so it reads
    # as one visual unit with the mark column beside it.
    code_cols = frozenset(u_start + 2*i + 1 for i in range(n_units))
    for r in range(1, n_rows + 1):
        bottom = _SIDE_THIN_BK if r in thick_rows else _SIDE_THIN_H
        for c in range(1, last_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left   = _SIDE_THIN if (c > 1 and c not in code_cols) else _SIDE_NONE,
                bottom = bottom,
            )

    # ---- merges (after borders) --------------------------------------
    for i in range(n_units):
        col = u_start + 2 * i
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
    for row, c1, c2 in pending_merges:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

    wb.save(outpath)
    print(f"  Written {len(students)} students "
          f"({n_units} units, {len(trailer_names)} trailing cols) → {outpath}")


# ===========================================================================
# Argument parsing
# ===========================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description=f'PyAssess2026: Physics undergraduate assessment (AY{AY})'
    )
    parser.add_argument(
        '--classyear',
        default='all',
        metavar='CY',
        help=(
            "Class year to process: 1, 2, 31, 32, 4 "
            "(append m/M for Maths+Physics equivalent). "
            "Use 'all' or '*' to run all 10 (default)."
        )
    )
    return parser.parse_args()


def resolve_classyears(raw):
    """Return the list of classyear keys to process."""
    val = raw.strip().lower().rstrip('/')
    if val in ('all', '*', ''):
        return ALL_CLASSYEARS
    val = val.replace('M', 'm')  # normalise upper-M
    if val not in CLASSYEAR_FILES:
        sys.exit(
            f"Error: unrecognised --classyear '{raw}'.\n"
            f"Valid values: {', '.join(ALL_CLASSYEARS)}, all, *"
        )
    return [val]


# ===========================================================================
# Main
# ===========================================================================

def main():
    args = parse_args()
    classyears = resolve_classyears(args.classyear)

    for cy in classyears:
        infile, outfile = CLASSYEAR_FILES[cy]
        inpath  = os.path.join(INDIR,  infile)
        outpath = os.path.join(OUTDIR, outfile)
        print(f"classyear {cy:>3s}: reading {inpath} ...")
        students = read_students(inpath)
        print(f"             {len(students)} students, "
              f"{len(students[0].units)} units each")
        for s in students:
            s.calc_credits()
            s.calc_yearmark()
        write_students(students, outpath, cy)



    # ***For testing/debugging keep this here (comment out when doing actual runs)
    #from IPython import embed
    #embed()

if __name__ == '__main__':
    main()
