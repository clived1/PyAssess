#!/usr/bin/env python3

# PyAssess2026.py - generates processed exam grids for Physics@Manchester
# Author: Clive Dickinson
# Date: 2026-05-30
# Version: 0.0.1

import argparse
import math
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

# Pass mark for any individual unit
PASS_MARK = 39.95

# Minimum mark in any unit to avoid outright fail (30%)
MIN_MARK = 29.95

# Degree classification boundaries
BOUNDARY_FIRST  = 69.95
BOUNDARY_UPPER2 = 59.95
BOUNDARY_LOWER2 = 49.95
BOUNDARY_THIRD  = 39.95

# Credit thresholds for Y1/Y2 progression
MIN_CREDITS_TO_PROGRESS  = 80   # credits at >= PASS_MARK needed to progress without resits
MIN_PASS_CREDITS         = 60   # credits at >= PASS_MARK needed at first attempt to avoid FAIL

FINAL_CLASSYEARS = ['32', '32m', '4', '4m']   # graduating / final-year students

# Units that must be passed if taken (lab, BSc project, MPhys project)
MUST_PASS = frozenset({'PHYS10180', 'PHYS10280', 'PHYS20180', 'PHYS20280',
                       'PHYS30180', 'PHYS30280', 'PHYS30880', 'PHYS30881',
                       'PHYS30882', 'PHYS40181', 'PHYS40182'})

# Maths units that Y1 M+P students must pass — cannot be compensated
MUST_PASS_MATHS = frozenset({'MATH11121', 'MATH11022'})

# Core Physics units (required for resit decisions)
CORE_PHYSICS = frozenset({
    'PHYS10071',
    'PHYS10101',
    'PHYS10121',
    'PHYS10191',
    'PHYS10302',
    'PHYS10342',
    'PHYS10372',
    'PHYS20101',
    'PHYS20141',
    'PHYS20151',   # Y2 changed in 2025 (replaces PHYS20161 - Judith email 30-Jun-2025)
    'PHYS20171',
    'PHYS20302',   # Y2 changed in 2025 (replaces PHYS20252 - Judith email 06-May-2025)
    'PHYS20342',   # Y2 changed in 2025 (replaces PHYS20312 - Judith email 06-May-2025)
    'PHYS20352',
})

# Additional core maths units for Maths+Physics students
CORE_MATHS_PHYSICS = frozenset({
    'MATH10111', 'MATH10121', 'MATH10212', 'MATH11222',
    'MATH11121', 'MATH11022', 'MATH29141', 'MATH24420',
})

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


# Mit_circs codes that are specifically actioned in exclude_units().
# Any code NOT in this set is copied to the output code column as-is.
_PROCESSED_MIT_CODES = frozenset({'AA', 'EA'})

# Classyears where deferrals (EA) and resits apply.
_DEFERRAL_CLASSYEARS = frozenset({'1', '1m', '2', '2m'})

# EN codes indicating a mark carried forward from a previous attempt.
_CARRIED_EN_CODES = frozenset({'L1C', 'L2C', 'L3C'})



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


_MARK_NUM_RE = re.compile(r'^\s*([\d.]+)\s*[A-Za-z]*\s*$')

def _numeric_mark(value):
    """Return the numeric part of a mark value, stripping any trailing letter suffix.

    Handles plain floats and values like '35C' (compensated) or '30R' (resit).
    Returns None if no numeric value can be extracted.
    """
    if value is None:
        return None
    m = _MARK_NUM_RE.match(str(value).strip())
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


class UnitInfo:
    """Data for a single unit (module) for one student."""
    __slots__ = ('unit_name', 'module', 'coursename', 'credits', 'mark', 'en', 'mit_circs',
                 'passed', 'excluded', 'output_code', 'capped')

    def __init__(self, unit_name, module, mark, en, mit_circs):
        self.unit_name   = unit_name              # 'Unit 1', 'Unit 2', etc.
        self.module      = module                 # raw cell value, e.g. 'PHYS10180 (20)'
        self.coursename, self.credits = _parse_module(module)
        self.mark        = mark                   # unit mark (float) or None
        self.en          = en                     # EN flag (str) or None
        self.mit_circs   = mit_circs              # mitigating circumstances code (str) or None
        self.passed      = None                   # True/False once exclude_units() runs; None = no mark
        self.excluded    = False                  # True if excluded from year mark
        self.output_code = None                   # code(s) written to the output code column
        self.capped      = False                  # True if mark is capped at 30.0 for year mark (R2 attempt)

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
        'credits_taken', 'credits_passed', 'credits_excluded', 'credits_deferred', 'creds_passed_taken',
        'excluded_idx', 'excluded_courses',
        'failed_idx', 'failed_courses',
        'deferred_idx', 'deferred_courses',
        'some_unit_under_30',
        'zone_idx', 'zone_courses',
        'compensated_idx', 'compensated_courses', 'credits_compensated',
        'referred_idx', 'referred_courses',
        'resits',
        'fail', 'fail_reason',
        'status',
        'phys_yearmark', 'math_yearmark',
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
        self.credits_passed     = None   # credits where mark > PASS_MARK
        self.credits_excluded   = 0      # credits excluded from calculation (populated later)
        self.credits_deferred   = 0      # credits from deferred (EA) units (assumed passed)
        self.creds_passed_taken = None   # formatted string for output, e.g. '120 / 120'
        self.excluded_idx       = []     # indices into self.units of excluded units
        self.excluded_courses   = []     # coursenames of excluded units
        self.failed_idx         = []     # indices into self.units of failed units
        self.failed_courses     = []     # coursenames of failed units
        self.deferred_idx       = []     # indices into self.units of deferred (EA) units
        self.deferred_courses   = []     # coursenames of deferred units
        self.some_unit_under_30  = False  # True if any failed unit is below MIN_MARK (30%)
        self.zone_idx            = []     # indices of units in compensation zone (MIN_MARK < mark <= PASS_MARK)
        self.zone_courses        = []     # coursenames of zone units
        self.compensated_idx     = []     # indices of units compensated (failed but allowed to count)
        self.compensated_courses = []     # coursenames of compensated units
        self.credits_compensated = 0      # total credits compensated
        self.referred_idx        = []     # indices of units referred (R2 resit)
        self.referred_courses    = []     # coursenames of referred units
        self.resits              = None   # deferred/resit courses for output, e.g. 'PHYS10071[1] / PHYS10101[1]'
        self.fail                = False  # True if student cannot progress
        self.fail_reason         = ''     # short description of why student failed
        self.status              = None   # set once by calc_status(): 'ACTV', 'A/D', 'REVW', 'FAIL'
        self.phys_yearmark       = None   # M+P only: credit-weighted average of non-MATH units
        self.math_yearmark       = None   # M+P only: credit-weighted average of MATH units

    def exclude_units(self, classyear=None):
        """Set credits_taken/passed/excluded, creds_passed_taken, and unit flags.

        Exclusion rules applied here:
          EA in mit_circs → deferral; only actioned for years 1/2
                            (classyear in _DEFERRAL_CLASSYEARS). Excluded from year
                            mark, treated as passed for outcome checks, but NOT
                            counted in credits_passed.
          AA in mit_circs → excluded from year mark, treated as passed,
                            output_code set to 'X'; credits counted in credits_passed.
        """
        taken            = 0
        passed           = 0
        excluded         = 0
        deferred_creds   = 0
        excluded_idx     = []
        excluded_courses = []
        failed_idx       = []
        failed_courses   = []
        deferred_idx     = []
        deferred_courses = []

        for idx, unit in enumerate(self.units):
            if unit.module is None or unit.credits is None:
                continue                             # empty slot

            taken     += unit.credits
            mit_codes  = _split_codes(unit.mit_circs)
            en_codes   = _split_codes(unit.en)
            used_mit   = set()                       # mit codes consumed by a rule
            used_en    = set()                       # EN codes consumed by a rule
            coursename = unit.coursename or unit.module

            # --- EA deferral (years 1/2 only; highest priority, trumps AA and all other mit codes) ---
            if 'EA' in mit_codes and classyear in _DEFERRAL_CLASSYEARS:
                unit.excluded = True
                unit.passed   = True           # treated as passed for outcome checks
                if 'XN' in en_codes:           # missed exam → XL_R1
                    unit.output_code = _append_code(unit.output_code, 'XL_R1')
                    used_en.add('XN')
                else:
                    unit.output_code = _append_code(unit.output_code, 'R1')
                excluded       += unit.credits  # excluded from year mark
                deferred_creds += unit.credits  # assumed passed for credit-count purposes
                # deferral credits deliberately NOT added to passed/credits_passed
                excluded_idx.append(idx);  excluded_courses.append(coursename)
                deferred_idx.append(idx);  deferred_courses.append(coursename)
                used_mit.update(mit_codes)           # EA trumps: mark all mit codes used

            # --- AA exclusion ---
            elif 'AA' in mit_codes:
                unit.excluded    = True
                unit.passed      = True
                unit.output_code = _append_code(unit.output_code, 'X')
                excluded += unit.credits
                passed   += unit.credits
                excluded_idx.append(idx);  excluded_courses.append(coursename)
                used_mit.add('AA')

            # --- L1C/L2C/L3C (carried mark): exclude from year average, treat as passed ---
            elif en_codes & _CARRIED_EN_CODES:
                unit.excluded = True
                unit.passed   = True
                excluded += unit.credits
                passed   += unit.credits
                excluded_idx.append(idx);  excluded_courses.append(coursename)
                # carried code not added to used_en, so it copies through to the output code column

            # --- XN (missed exam): counts as failed, mark still used in year mark average ---
            elif 'XN' in en_codes:
                unit.passed = False
                failed_idx.append(idx);  failed_courses.append(coursename)
                # XN not added to used_en, so it copies through to the output code column

            # --- no mark: exclude (treat as passed, omit from year mark) ---
            elif unit.mark is None:
                unit.excluded = True
                unit.passed   = True
                excluded += unit.credits
                passed   += unit.credits
                excluded_idx.append(idx);  excluded_courses.append(coursename)

            # --- normal pass/fail ---
            else:
                num = _numeric_mark(unit.mark)
                unit.passed = num is not None and num > PASS_MARK
                if unit.passed:
                    passed += unit.credits
                else:
                    failed_idx.append(idx);  failed_courses.append(coursename)

            # --- copy through any unprocessed EN and mit_circs codes ---
            for code in sorted(en_codes - used_en):
                if code == 'R2':
                    unit.output_code = _append_code(unit.output_code, 'cap')
                    unit.capped = True
                elif code == 'R1':
                    pass  # R1 = 1st-attempt resit; treat as normal, no output annotation
                else:
                    unit.output_code = _append_code(unit.output_code, code)
            for code in sorted(mit_codes - used_mit - _PROCESSED_MIT_CODES):
                unit.output_code = _append_code(unit.output_code, code)

        self.credits_taken      = taken
        self.credits_passed     = passed
        self.credits_excluded   = excluded
        self.credits_deferred   = deferred_creds
        self.creds_passed_taken = f"{passed} / {taken}"
        self.excluded_idx       = excluded_idx
        self.excluded_courses   = excluded_courses
        self.failed_idx         = failed_idx
        self.failed_courses     = failed_courses
        self.deferred_idx       = deferred_idx
        self.deferred_courses   = deferred_courses
        self.resits             = ' / '.join(f"{c}[1]" for c in deferred_courses if c) or None

    def calc_yearmark(self, classyear=None):
        """Set self.yearmark to the credit-weighted mean of all unit marks.

        For M+P classyears also sets self.phys_yearmark (non-MATH units) and
        self.math_yearmark (units whose coursename contains 'MATH').
        Units with a missing mark or missing credits are skipped.
        Results are rounded to 1 decimal place; None if no valid units found.
        """
        def _round1dp(weighted, credits):
            return math.floor(weighted / credits * 10 + 0.5) / 10 if credits > 0 else None

        total_credits  = 0
        weighted_marks = 0.0
        phys_credits   = 0
        phys_weighted  = 0.0
        math_credits   = 0
        math_weighted  = 0.0
        mp = classyear is not None and classyear.endswith('m')

        for unit in self.units:
            if unit.excluded:
                continue
            if unit.mark is None or unit.credits is None:
                continue
            mark = _numeric_mark(unit.mark)
            if mark is None:
                continue
            calc_mark = min(mark, 30.0) if unit.capped else mark
            weighted_marks += calc_mark * unit.credits
            total_credits  += unit.credits
            if mp:
                if (unit.coursename or '').startswith('MATH'):
                    math_weighted += calc_mark * unit.credits
                    math_credits  += unit.credits
                else:
                    phys_weighted += calc_mark * unit.credits
                    phys_credits  += unit.credits

        self.yearmark = _round1dp(weighted_marks, total_credits)
        if mp:
            self.phys_yearmark = _round1dp(phys_weighted, phys_credits)
            self.math_yearmark = _round1dp(math_weighted, math_credits)

    def calc_status(self, classyear):
        """Set self.status once, reading flags set by earlier methods.

        Priority order:
          FAIL  — self.fail is True
          REVW  — one or more referrals (self.referred_idx non-empty), with or without deferrals
          A/D   — one or more deferrals, no referrals (self.deferred_idx non-empty)
          ACTV  — default: active, fine to progress
        """
        if classyear in FINAL_CLASSYEARS:
            return
        if self.fail:
            self.status = 'FAIL'
            self.resits = None
        elif self.referred_idx:
            self.status = 'REVW'
        elif self.deferred_idx:
            self.status = 'A/D'
        else:
            self.status = 'ACTV'

    def calc_referrals(self, classyear):
        """Determine compensation and referrals for non-final-year students.

        A MUST_PASS unit with mark == 39 is treated as a normal zone failure (30-39%)
        and flows through the same compensation/referral logic as any other unit.
        A MUST_PASS unit with mark < 39 is an outright fail ('Failed lab').

        Units with EN code 'R2' were already taken as a 2nd attempt; no further
        resit can be offered.  If such a unit has mark < 30% the student fails
        outright.  In the zone (30-39%) compensation rules apply as normal, but
        if the rules would assign R2 (core/must-pass unit, or over credit cap)
        the student fails outright instead.

        Y1/Y2 paths:
          Full compensation (failed <= 40 credits, no unit under 30%):
            Must-pass units → R2 (or FAIL if R2-in-EN); all others → C.
          Referral (some_unit_under_30):
            Units < 30% → R2 (or FAIL if R2-in-EN); zone must-pass/core → R2
            (or FAIL if R2-in-EN); zone non-core within 40 credits → C;
            zone non-core over cap → R2 (or FAIL if R2-in-EN).
          >40 credits, no unit under 30%:
            Must-pass/core zone units → R2 (or FAIL if R2-in-EN); non-core → C.
        """
        if classyear in FINAL_CLASSYEARS:
            return

        # MUST_PASS units with mark == 39 (and not already at R2 attempt) are
        # zone failures, not outright fails.  R2-in-EN labs lose that exception
        # because no further resit can be offered.
        lab_near_pass_idx = set()
        r2_en_idx         = set()   # failed units whose EN column contains 'R2'
        for idx in self.failed_idx:
            unit       = self.units[idx]
            coursename = unit.coursename or unit.module
            en_codes   = _split_codes(unit.en)
            is_r2_en   = 'R2' in en_codes
            if is_r2_en:
                r2_en_idx.add(idx)
            if coursename in MUST_PASS and not is_r2_en:
                if _numeric_mark(unit.mark) == 39.0:
                    lab_near_pass_idx.add(idx)

        if classyear not in _DEFERRAL_CLASSYEARS:
            # Y3 progressing.
            if (self.credits_passed or 0) + self.credits_deferred < MIN_PASS_CREDITS:
                self.fail        = True
                self.fail_reason = '<60 credits'
                return
            # Refer lab=39 as R2; no other compensation logic yet.
            if lab_near_pass_idx:
                referred_idx     = []
                referred_courses = []
                for idx in self.failed_idx:
                    if idx in lab_near_pass_idx:
                        unit = self.units[idx]
                        coursename = unit.coursename or unit.module
                        unit.output_code = _append_code(unit.output_code, 'R2')
                        referred_idx.append(idx)
                        referred_courses.append(coursename)
                self.referred_idx     = referred_idx
                self.referred_courses = referred_courses
                self.fail_reason      = 'Resit failed lab'
                resit_parts  = [f"{c}[1]" for c in self.deferred_courses]
                resit_parts += referred_courses
                self.resits  = ' / '.join(p for p in resit_parts if p) or None
            return

        # --- Y1/Y2 FAIL checks (all collected so multiple reasons can be combined) ---
        fail_reasons = []

        if any(c in MUST_PASS
               for idx, c in zip(self.failed_idx, self.failed_courses)
               if idx not in lab_near_pass_idx):
            fail_reasons.append('Failed lab')

        if (self.credits_passed or 0) + self.credits_deferred < MIN_PASS_CREDITS:
            fail_reasons.append('<60 credits')

        # R2-in-EN units with mark < 30%: no resit available.
        for idx in r2_en_idx:
            unit = self.units[idx]
            mark = _numeric_mark(unit.mark) or 0.0
            if not (mark > MIN_MARK):
                fail_reasons.append('Failed (<30%) 2nd attempts')
                break

        if fail_reasons:
            self.fail        = True
            self.fail_reason = ' / '.join(fail_reasons)
            return

        # All failed units flow through the same logic (lab=39 is just a zone unit).
        other_failed_idx = list(self.failed_idx)

        # --- classify all failed units ---
        failed_credits     = 0
        some_unit_under_30 = False
        for idx in other_failed_idx:
            unit            = self.units[idx]
            failed_credits += unit.credits or 0
            num = _numeric_mark(unit.mark)
            if num is None or not (num > MIN_MARK):
                some_unit_under_30 = True

        self.some_unit_under_30 = some_unit_under_30

        must_pass_for_cy = MUST_PASS | (MUST_PASS_MATHS if classyear == '1m' else frozenset())
        core_for_cy      = CORE_PHYSICS | (CORE_MATHS_PHYSICS if classyear in ('1m', '2m')
                                           else frozenset())

        zone_idx            = []
        zone_courses        = []
        compensated_idx     = []
        compensated_courses = []
        referred_idx        = []
        referred_courses    = []

        if failed_credits <= 40 and not some_unit_under_30:
            # --- full compensation path ---
            for idx in other_failed_idx:
                unit       = self.units[idx]
                coursename = unit.coursename or unit.module
                if coursename in must_pass_for_cy:
                    if idx in r2_en_idx:
                        self.fail        = True
                        self.fail_reason = 'Failed (<30%) 2nd attempts'
                        return
                    unit.output_code = _append_code(unit.output_code, 'R2')
                    referred_idx.append(idx)
                    referred_courses.append(coursename)
                else:
                    unit.output_code = _append_code(unit.output_code, 'C')
                    compensated_idx.append(idx)
                    compensated_courses.append(coursename)

        elif some_unit_under_30:
            # --- referral path ---
            compensation_used = 0
            for idx in other_failed_idx:
                unit       = self.units[idx]
                coursename = unit.coursename or unit.module
                mark = _numeric_mark(unit.mark) or 0.0
                if not (mark > MIN_MARK):
                    # already caught by pre-check if R2-in-EN, but defensive
                    if idx in r2_en_idx:
                        self.fail        = True
                        self.fail_reason = 'Failed (<30%) 2nd attempts'
                        return
                    unit.output_code = _append_code(unit.output_code, 'R2')
                    referred_idx.append(idx)
                    referred_courses.append(coursename)
                else:
                    zone_idx.append(idx)
                    zone_courses.append(coursename)
                    if coursename in core_for_cy or coursename in must_pass_for_cy:
                        if idx in r2_en_idx:
                            self.fail        = True
                            self.fail_reason = 'Failed (<30%) 2nd attempts'
                            return
                        unit.output_code = _append_code(unit.output_code, 'R2')
                        referred_idx.append(idx)
                        referred_courses.append(coursename)
                    elif compensation_used + (unit.credits or 0) <= 40:
                        unit.output_code = _append_code(unit.output_code, 'C')
                        compensation_used += unit.credits or 0
                        compensated_idx.append(idx)
                        compensated_courses.append(coursename)
                    else:
                        if idx in r2_en_idx:
                            self.fail        = True
                            self.fail_reason = 'Failed (<30%) 2nd attempts'
                            return
                        unit.output_code = _append_code(unit.output_code, 'R2')
                        referred_idx.append(idx)
                        referred_courses.append(coursename)

        else:
            # --- >40 credits, all in 30–39% zone (no unit below 30%) ---
            # Must-pass/core units → R2 (or FAIL if R2-in-EN); non-core → C.
            for idx in other_failed_idx:
                unit       = self.units[idx]
                coursename = unit.coursename or unit.module
                zone_idx.append(idx)
                zone_courses.append(coursename)
                if coursename in core_for_cy or coursename in must_pass_for_cy:
                    if idx in r2_en_idx:
                        self.fail        = True
                        self.fail_reason = 'Failed (<30%) 2nd attempts'
                        return
                    unit.output_code = _append_code(unit.output_code, 'R2')
                    referred_idx.append(idx)
                    referred_courses.append(coursename)
                else:
                    unit.output_code = _append_code(unit.output_code, 'C')
                    compensated_idx.append(idx)
                    compensated_courses.append(coursename)

        # --- Assign all results ---
        self.zone_idx            = zone_idx
        self.zone_courses        = zone_courses
        self.compensated_idx     = compensated_idx
        self.compensated_courses = compensated_courses
        self.credits_compensated = sum(self.units[i].credits or 0 for i in compensated_idx)
        self.referred_idx        = referred_idx
        self.referred_courses    = referred_courses

        if referred_idx:
            if any((self.units[i].coursename or self.units[i].module) in MUST_PASS
                   for i in referred_idx):
                self.fail_reason = 'Resit failed lab'
            resit_parts = [f"{c}[1]" for c in self.deferred_courses]
            resit_parts += referred_courses
            self.resits = ' / '.join(p for p in resit_parts if p) or None

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
    'ID No.':                    5.00,
    'Emplid':                    9.00,
    'Name':                     12.00,
    'Plan':                     17.00,
    '_unit':                     7.50,
    '_code':                     7.50,
    'Creds Passed/Taken':       16.00,
    'Year Mark':                9.00,
    'Phys Year Mark':           12.00,
    'Math Year Mark':           12.00,
    'Status':                   9.00,
    'Fail reason':              22.00,
    'Resits':                   50.00,
    'Notes':                    50.00,
    'Pre-Exam Board Minutes':   40.00,
    'Exam Board Minutes':       40.00,
    'Phys 1':                   10.00,
    'Phys 2':                   10.00,
    'Phys 3':                   10.00,
    'BZ':                        8.00,
    'Overall':                  10.00,
    'L3/L4 creds passed':       16.00,
    'L4 creds passed Y3+Y4':    18.00,
    'Y3 creds failed w/wo MCs': 18.00,
    'Award':                    10.00,
    'Classification':           14.00,
    'Deg Class Alg':            14.00,
    'Deg Class Rev':            14.00,
    'Deg Class Actual':         14.00,
    'Award Alg':                10.00,
    'Award Actual':             10.00,
    'Classification Alg':       16.00,
    'Classification Actual':    16.00,
    'Award Change':             10.00,
    'Classification Change':    16.00,
}

# Formatting objects (created once, reused for every cell)
_FONT        = Font(name='Aptos Narrow', size=11)
_FONT_BOLD   = Font(name='Aptos Narrow', size=11, bold=True)
_FILL_GREY   = PatternFill(fill_type='solid', fgColor='FFE0E0E0')
_FILL_YELLOW = PatternFill(fill_type='solid', fgColor='FFFFFF00')
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
    'Phys Year Mark':     'phys_yearmark',
    'Math Year Mark':     'math_yearmark',
    'Status':             'status',
    'Fail reason':        'fail_reason',
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
    ws.title = '2 Line Format'

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

        # info row — trailing columns (computed attrs take priority; fall back to input value)
        for j, tname in enumerate(trailer_names):
            attr  = _TRAILING_ATTR.get(tname)
            value = getattr(s, attr) if attr else s.trailing.get(tname)
            cell  = _c(info_row, t_start + j, value)
            fmt   = _TRAILING_FORMAT.get(tname)
            if fmt:
                cell.number_format = fmt

        # marks row — fixed columns (blank; needed for fill and borders)
        for i in range(1, n_fixed + 1):
            _c(marks_row, i)

        # marks row — unit marks and output codes
        yellow_set = set(s.failed_idx) | set(s.deferred_idx)
        for i, unit in enumerate(s.units):
            col = u_start + 2 * i
            mark_cell = _c(marks_row, col, unit.mark)
            if i in yellow_set:
                mark_cell.fill = _FILL_YELLOW
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
            s.exclude_units(cy)
            s.calc_yearmark(cy)
            s.calc_referrals(cy)
            s.calc_status(cy)

            # ***For testing/debugging keep this here (comment out when doing actual runs)
            '''if (s.emplid == '11391048'):
                from IPython import embed
                embed()'''
        write_students(students, outpath, cy)



    # ***For testing/debugging keep this here (comment out when doing actual runs)
    #from IPython import embed
    #embed()

if __name__ == '__main__':
    main()
