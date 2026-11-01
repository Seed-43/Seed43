# -*- coding: utf-8 -*-
"""
pylink_numfmt.py -- render an xlsx number the way Excel displays it.

An xlsx stores 97.15 as 97.15000000000001 and the " m²" of a custom
format only in styles.xml, so a reader that copies the raw <v> text
puts float noise and no units on the drawing. format_value() applies
the cell's number format: General, fixed decimals, thousands, percent,
scientific, quoted/escaped literals, colour and condition sections,
and dates and times.

Pure Python with no Revit or .NET imports, so it runs under IronPython 2
in Revit and under CPython for testing.
"""

import datetime as _dt
import re

# Excel's built-in formats, by numFmtId. Ids 14 and 22 follow the system
# short date in Excel; d/mm/yyyy is the NZ/AU form.
BUILTIN_FORMATS = {
    0: u'General',
    1: u'0',
    2: u'0.00',
    3: u'#,##0',
    4: u'#,##0.00',
    9: u'0%',
    10: u'0.00%',
    11: u'0.00E+00',
    12: u'General',
    13: u'General',
    14: u'd/mm/yyyy',
    15: u'd-mmm-yy',
    16: u'd-mmm',
    17: u'mmm-yy',
    18: u'h:mm AM/PM',
    19: u'h:mm:ss AM/PM',
    20: u'h:mm',
    21: u'h:mm:ss',
    22: u'd/mm/yyyy h:mm',
    37: u'#,##0 ;(#,##0)',
    38: u'#,##0 ;(#,##0)',
    39: u'#,##0.00;(#,##0.00)',
    40: u'#,##0.00;(#,##0.00)',
    45: u'mm:ss',
    46: u'[h]:mm:ss',
    47: u'mm:ss.0',
    48: u'##0.0E+0',
    49: u'@',
}

_MONTHS = [u'January', u'February', u'March', u'April', u'May', u'June',
           u'July', u'August', u'September', u'October', u'November',
           u'December']
_DAYS = [u'Monday', u'Tuesday', u'Wednesday', u'Thursday', u'Friday',
         u'Saturday', u'Sunday']

_NUM_CHARS = u'0#?.,%'
_DATE_LETTERS = u'yYdDhHsS'


def format_value(raw, code):
    """The text Excel shows for a numeric cell's raw value under format
    code. Anything this can't read comes back as General, never raised."""
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return raw
    try:
        return _format(value, code)
    except Exception:
        return _general(value)


# ── Sections ──

def _format(value, code):
    if not code or code.strip().lower() == u'general':
        return _general(value)
    sections = _split_sections(code)
    section, absolute = _pick_section(value, sections)
    if section is None:
        return _general(value)
    tokens = _tokenise(section)
    if _is_date(tokens):
        return _format_date(value, tokens)
    if not any(kind == 'num' for kind, _ in tokens):
        # Literal-only section, e.g. "-" for zero, or General inside one.
        text = u''.join(t for kind, t in tokens if kind == 'lit')
        if u'general' in section.lower():
            return text.replace(u'General', _general(value))
        return text.replace(u'@', _general(value))
    shown = abs(value) if absolute else value
    return _format_number(shown, tokens)


def _split_sections(code):
    sections, buf, quoted, i = [], [], False, 0
    while i < len(code):
        ch = code[i]
        if ch == u'"':
            quoted = not quoted
        elif ch == u'\\' and not quoted and i + 1 < len(code):
            buf.append(code[i:i + 2])
            i += 2
            continue
        elif ch == u';' and not quoted:
            sections.append(u''.join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    sections.append(u''.join(buf))
    return sections


def _condition(section):
    m = re.match(u'\\s*\\[(<=|>=|<>|<|>|=)\\s*(-?[0-9.]+)\\]', section)
    if not m:
        return None
    return m.group(1), float(m.group(2))


def _test(value, cond):
    op, limit = cond
    return {u'<': value < limit, u'>': value > limit,
            u'<=': value <= limit, u'>=': value >= limit,
            u'=': value == limit, u'<>': value != limit}[op]


def _pick_section(value, sections):
    """(section, use_absolute). A negative value in a second section is
    shown unsigned: the section supplies its own minus or brackets."""
    sections = [s for s in sections]
    if any(_condition(s) for s in sections[:2]):
        for i, s in enumerate(sections[:2]):
            cond = _condition(s)
            if cond and _test(value, cond):
                return s, i == 1 and value < 0
        return (sections[2] if len(sections) > 2 else sections[-1]), False
    if value > 0 or len(sections) == 1:
        return sections[0], False
    if value < 0:
        return sections[1], True
    return (sections[2] if len(sections) > 2 else sections[0]), False


# ── Tokens ──

def _tokenise(section):
    """[(kind, text)] with kind 'lit' (shown as is), 'num' (a digit
    placeholder or . , %), 'exp' (E+ / E-) or 'date' (a date/time code)."""
    tokens, i, n = [], 0, len(section)
    while i < n:
        ch = section[i]
        if ch == u'"':
            end = section.find(u'"', i + 1)
            end = n if end < 0 else end
            tokens.append(('lit', section[i + 1:end]))
            i = end + 1
        elif ch == u'\\' and i + 1 < n:
            tokens.append(('lit', section[i + 1]))
            i += 2
        elif ch == u'_' and i + 1 < n:
            tokens.append(('lit', u' '))      # pads by a character's width
            i += 2
        elif ch == u'*' and i + 1 < n:
            i += 2                            # fill character: no width here
        elif ch == u'[':
            end = section.find(u']', i)
            end = n if end < 0 else end
            inner = section[i + 1:end]
            if inner.startswith(u'$'):
                # [$€-1809] currency/locale tag: show the symbol part.
                tokens.append(('lit', inner[1:].split(u'-')[0]))
            elif re.match(u'^[hHmMsS]+$', inner):
                tokens.append(('date', u'[' + inner.lower() + u']'))
            i = end + 1                       # colours and conditions drop
        elif section[i:i + 5].upper() == u'AM/PM':
            tokens.append(('date', u'AM/PM'))
            i += 5
        elif section[i:i + 3].upper() == u'A/P':
            tokens.append(('date', u'A/P'))
            i += 3
        elif ch in u'eE' and i + 1 < n and section[i + 1] in u'+-':
            tokens.append(('exp', section[i:i + 2]))
            i += 2
        elif ch in _NUM_CHARS:
            tokens.append(('num', ch))
            i += 1
        elif ch.lower() in u'ymdhs':
            j = i
            while j < n and section[j].lower() == ch.lower():
                j += 1
            tokens.append(('date', section[i:j].lower()))
            i = j
        elif ch == u'@':
            tokens.append(('lit', u'@'))      # text placeholder
            i += 1
        else:
            tokens.append(('lit', ch))
            i += 1
    return tokens


def _is_date(tokens):
    return any(kind == 'date' for kind, _ in tokens) and not any(
        kind == 'num' and t in u'0#?' for kind, t in tokens)


# ── Numbers ──

def _general(value):
    """Excel's General: no float noise, up to 11 significant digits,
    scientific only for very large or very small values."""
    if value == int(value) and abs(value) < 1e11:
        return u'{0:d}'.format(int(value))
    a = abs(value)
    if a >= 1e11 or a < 1e-9:
        text = u'{0:.5E}'.format(value)
        mant, exp = text.split(u'E')
        mant = mant.rstrip(u'0').rstrip(u'.')
        return u'{0}E{1}{2:02d}'.format(mant, exp[0], abs(int(exp)))
    text = u'{0:.11g}'.format(value)
    if u'e' in text:
        text = u'{0:.10f}'.format(value).rstrip(u'0').rstrip(u'.')
    return text


def _format_number(value, tokens):
    # % multiplies by 100 and is shown where it stands, like a literal.
    percent = sum(1 for k, t in tokens if k == 'num' and t == u'%')
    value *= 100 ** percent
    tokens = [('lit', t) if (k == 'num' and t == u'%') else (k, t)
              for k, t in tokens]
    first = next(i for i, (k, _) in enumerate(tokens) if k in ('num', 'exp'))
    last = max(i for i, (k, _) in enumerate(tokens) if k in ('num', 'exp'))
    prefix = u''.join(t for k, t in tokens[:first] if k == 'lit')
    suffix = u''.join(t for k, t in tokens[last + 1:] if k == 'lit')
    body = tokens[first:last + 1]
    # Literals between placeholders (000-000) are kept in place below.
    pattern = u''.join(t if k in ('num', 'exp') else u'\x00' for k, t in body)
    mids = [t for k, t in body if k not in ('num', 'exp')]

    exp_pos = -1
    for marker in (u'E+', u'E-', u'e+', u'e-'):
        if marker in pattern:
            exp_pos = pattern.index(marker)
    mantissa_pat = pattern if exp_pos < 0 else pattern[:exp_pos]

    # Trailing commas after the last placeholder scale by 1000 each.
    stripped = mantissa_pat.rstrip(u',\x00')
    trailing = mantissa_pat[len(stripped):].count(u',')
    value /= 1000.0 ** trailing
    mantissa_pat = stripped

    int_pat, _, dec_pat = mantissa_pat.partition(u'.')
    grouping = u',' in int_pat
    int_digits = [c for c in int_pat if c in u'0#?']
    dec_digits = [c for c in dec_pat if c in u'0#?']

    exponent = None
    if exp_pos >= 0:
        exponent = 0
        if value:
            exponent = int(_floor_log10(abs(value)))
            step = max(1, len(int_digits)) if len(int_digits) > 1 else 1
            exponent -= exponent % step
            value /= 10.0 ** exponent

    sign = u'-' if value < 0 else u''
    text = u'{0:.{1}f}'.format(abs(value), len(dec_digits))
    whole, _, frac = text.partition(u'.')

    # Optional decimals (#) drop trailing zeros; ? leaves a space.
    min_dec = len(dec_digits)
    while min_dec and dec_digits[min_dec - 1] != u'0':
        min_dec -= 1
    frac_out = frac
    while len(frac_out) > min_dec and frac_out.endswith(u'0'):
        frac_out = frac_out[:-1]
    frac_out += u' ' * sum(1 for c in dec_digits[len(frac_out):] if c == u'?')

    min_int = sum(1 for c in int_digits if c == u'0')
    if whole == u'0' and min_int == 0:
        whole = u''
    whole = whole.zfill(min_int) if whole else u'0' * min_int
    if grouping and whole:
        whole = u'{0:,}'.format(int(whole))
        whole = whole.zfill(min_int) if len(whole) < min_int else whole

    if mids and not grouping and not dec_pat and exp_pos < 0:
        whole = _fill_mids(whole, int_pat, mids)

    out = whole
    if u'.' in mantissa_pat:
        # Excel keeps the point even when # drops every decimal: 0.## -> "5."
        out += u'.' + frac_out
    if exponent is not None:
        exp_marker = pattern[exp_pos:exp_pos + 2]
        exp_digits = sum(1 for c in pattern[exp_pos + 2:] if c == u'0')
        e_sign = u'-' if exponent < 0 else (u'+' if exp_marker[1] == u'+'
                                           else u'')
        out += u'E' + e_sign + str(abs(exponent)).zfill(max(1, exp_digits))
    return prefix + sign + out + suffix


def _floor_log10(x):
    import math
    return math.floor(math.log10(x))


def _fill_mids(digits, int_pat, mids):
    """Lay digits into a pattern like 000-000, right to left."""
    out, di, mi = [], len(digits) - 1, len(mids) - 1
    for c in reversed(int_pat):
        if c == u'\x00':
            out.append(mids[mi] if mi >= 0 else u'')
            mi -= 1
        elif c in u'0#?':
            out.append(digits[di] if di >= 0 else u'')
            di -= 1
    lead = digits[:di + 1] if di >= 0 else u''
    return lead + u''.join(reversed(out))


# ── Dates and times ──

def _serial_to_datetime(value):
    # Excel's 1900 system, with its phantom 29 Feb 1900 already counted
    # in serials from 61 up: day 0 is 30 Dec 1899 for those.
    base = _dt.datetime(1899, 12, 30) if value >= 61 else \
        _dt.datetime(1899, 12, 31)
    return base + _dt.timedelta(days=value)


def _format_date(value, tokens):
    when = _serial_to_datetime(value)
    when = when + _dt.timedelta(microseconds=500)   # round to the ms
    has_ampm = any(t in (u'AM/PM', u'A/P') for _, t in tokens)
    out = []
    dates = [i for i, (k, _) in enumerate(tokens) if k == 'date']
    for idx, (kind, t) in enumerate(tokens):
        if kind != 'date':
            if kind == 'num' and t == u'.':
                out.append(u'.')
            elif kind == 'num' and t == u'0':
                out.append(str(when.microsecond // 100000))
            else:
                out.append(t)
            continue
        if t in (u'AM/PM', u'A/P'):
            pm = when.hour >= 12
            out.append((u'PM' if pm else u'AM') if t == u'AM/PM'
                       else (u'P' if pm else u'A'))
        elif t.startswith(u'['):
            total = value * 24 if u'h' in t else (
                value * 1440 if u'm' in t else value * 86400)
            out.append(str(int(total)).zfill(len(t) - 2))
        elif t[0] == u'y':
            out.append(str(when.year) if len(t) > 2
                       else str(when.year)[-2:])
        elif t[0] == u'd':
            if len(t) >= 4:
                out.append(_DAYS[when.weekday()])
            elif len(t) == 3:
                out.append(_DAYS[when.weekday()][:3])
            else:
                out.append(str(when.day).zfill(len(t)))
        elif t[0] == u'h':
            hour = when.hour
            if has_ampm:
                hour = hour % 12 or 12
            out.append(str(hour).zfill(min(len(t), 2)))
        elif t[0] == u's':
            out.append(str(when.second).zfill(min(len(t), 2)))
        elif t[0] == u'm':
            if _is_minutes(tokens, dates, idx) and len(t) <= 2:
                out.append(str(when.minute).zfill(len(t)))
            elif len(t) >= 5:
                out.append(_MONTHS[when.month - 1][0])
            elif len(t) == 4:
                out.append(_MONTHS[when.month - 1])
            elif len(t) == 3:
                out.append(_MONTHS[when.month - 1][:3])
            else:
                out.append(str(when.month).zfill(len(t)))
    return u''.join(out)


def _is_minutes(tokens, dates, idx):
    """m means minutes straight after an hour or straight before seconds."""
    pos = dates.index(idx)
    if pos > 0 and tokens[dates[pos - 1]][1][:1] in (u'h', u'['):
        return True
    if pos + 1 < len(dates) and tokens[dates[pos + 1]][1][:1] == u's':
        return True
    return False
