# -*- coding: utf-8 -*-
# generate_bundle_yaml.py
#
# Scans all .pushbutton folders (recursively) from the directory this script
# lives in. For each folder it:
#
#   1. Reads metadata from the .py file(s)  OR  the existing bundle.yaml
#   2. Always rewrites bundle.yaml with:
#        - title  formatted as bold italic unicode
#        - author always "𝐒𝐄𝐄𝐃𝟒𝟑"  (bold caps unicode)
#        - tooltip with VERSION line converted to bold unicode
#   3. Strips the metadata block from the .py file and adds a filename comment
#
# Run from anywhere inside your pyRevit extension folder:
#     python generate_bundle_yaml.py

import os
import re
import textwrap


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

_CODING_LINE_RE = re.compile(r'^[ \t]*#.*coding[:=].*$', re.MULTILINE)

_DUNDER_SINGLE_RE = re.compile(
    r'^[ \t]*__(?P<name>title|author|doc)__[ \t]*=[ \t]*(?P<q>["\'])(?P<val>.+?)(?P=q)[ \t]*$',
    re.MULTILINE,
)

_DUNDER_TRIPLE_RE = re.compile(
    r'^[ \t]*__(?P<name>title|author|doc)__[ \t]*=[ \t]*'
    r'(?P<q>"{3}|\'{3})(?P<val>.*?)(?P=q)',
    re.MULTILINE | re.DOTALL,
)

_MODULE_DOC_RE = re.compile(
    r'^(?P<q>"{3}|\'{3})(?P<val>.*?)(?P=q)',
    re.MULTILINE | re.DOTALL,
)

# Matches   title: "some value"   in an existing bundle.yaml
_YAML_TITLE_RE = re.compile(r'^title:\s*["\']?(.+?)["\']?\s*$', re.MULTILINE)


# ---------------------------------------------------------------------------
# Unicode formatting helpers
# ---------------------------------------------------------------------------

# Fixed author string — bold caps unicode
AUTHOR = '𝐒𝐄𝐄𝐃𝟒𝟑'



def _to_bold(text):
    """Convert uppercase letters and digits to Unicode mathematical bold."""
    result = []
    for ch in text:
        if 'A' <= ch <= 'Z':
            result.append(chr(0x1D400 + ord(ch) - ord('A')))
        elif '0' <= ch <= '9':
            result.append(chr(0x1D7CE + ord(ch) - ord('0')))
        else:
            result.append(ch)
    return ''.join(result)


def _format_title(title):
    """Return title as plain text."""
    return title


def _bold_version_line(tooltip):
    """Convert the first VERSION XXXXXX line to bold unicode.
    Only the text is reformatted — the version number itself is unchanged.
    """
    lines = tooltip.splitlines()
    pattern = re.compile(r'^(VERSION\s+\S+)(.*)', re.IGNORECASE)
    for i, line in enumerate(lines):
        m = pattern.match(line.strip())
        if m:
            lines[i] = _to_bold(m.group(1).upper()) + m.group(2)
            break
    return '\n'.join(lines)


# ---------------------------------------------------------------------------
# Extraction from .py source
# ---------------------------------------------------------------------------

def _extract_from_py(source):
    """Extract title, author, tooltip from .py source. Returns dict."""
    found = {'title': '', 'author': '', 'tooltip': ''}

    for m in _DUNDER_TRIPLE_RE.finditer(source):
        name = m.group('name')
        val  = textwrap.dedent(m.group('val')).strip()
        key  = 'tooltip' if name == 'doc' else name
        if not found[key]:
            found[key] = val

    for m in _DUNDER_SINGLE_RE.finditer(source):
        name = m.group('name')
        val  = m.group('val').strip()
        key  = 'tooltip' if name == 'doc' else name
        if not found[key]:
            found[key] = val

    if not found['tooltip']:
        stripped = _CODING_LINE_RE.sub('', source, count=1).lstrip()
        m = _MODULE_DOC_RE.match(stripped)
        if m:
            found['tooltip'] = textwrap.dedent(m.group('val')).strip()

    return found


def _extract_from_yaml(yaml_path):
    """Read title (and tooltip) from an existing bundle.yaml. Returns dict."""
    found = {'title': '', 'tooltip': ''}
    try:
        with open(yaml_path, encoding='utf-8', errors='replace') as fh:
            content = fh.read()
    except OSError:
        return found

    # Pull title from   title: "value"
    m = _YAML_TITLE_RE.search(content)
    if m:
        # Strip any existing unicode bold italic back to plain text is complex,
        # so we just grab whatever is there as the display title.
        raw = m.group(1).strip().strip('"').strip("'")
        # Strip any unicode bold/italic formatting back to plain ASCII
        plain = []
        for ch in raw:
            cp = ord(ch)
            if 0x1D400 <= cp <= 0x1D433:   # bold A-Z a-z
                base = cp - 0x1D400
                plain.append(chr(ord('A') + base) if base < 26 else chr(ord('a') + base - 26))
            elif 0x1D434 <= cp <= 0x1D467:  # italic A-Z a-z
                base = cp - 0x1D434
                plain.append(chr(ord('A') + base) if base < 26 else chr(ord('a') + base - 26))
            elif 0x1D468 <= cp <= 0x1D49B:  # bold italic A-Z a-z
                base = cp - 0x1D468
                plain.append(chr(ord('A') + base) if base < 26 else chr(ord('a') + base - 26))
            elif 0x1D7CE <= cp <= 0x1D7D7:  # bold digits 0-9
                plain.append(chr(ord('0') + cp - 0x1D7CE))
            else:
                plain.append(ch)
        found['title'] = ''.join(plain)

    # Pull tooltip from block scalar or quoted scalar
    tip_block = re.search(r'^tooltip:\s*\|\n((?:[ \t]+.*\n?)*)', content, re.MULTILINE)
    if tip_block:
        lines = tip_block.group(1).splitlines()
        found['tooltip'] = '\n'.join(l[2:] if l.startswith('  ') else l for l in lines).strip()
    else:
        tip_inline = re.search(r'^tooltip:\s*["\'](.+?)["\']', content, re.MULTILINE)
        if tip_inline:
            found['tooltip'] = tip_inline.group(1).strip()

    return found


# ---------------------------------------------------------------------------
# Stripping metadata from .py source
# ---------------------------------------------------------------------------

def _strip_metadata(source):
    """Remove metadata blocks from source, keep coding comment, collapse blanks."""
    result = _DUNDER_TRIPLE_RE.sub('', source)
    result = _DUNDER_SINGLE_RE.sub('', result)

    # Remove implicit module docstring if it is the first non-comment code
    lines = result.splitlines(keepends=True)
    first_code_idx = None
    for i, line in enumerate(lines):
        s = line.strip()
        if s and not s.startswith('#'):
            first_code_idx = i
            break

    if first_code_idx is not None:
        first_line = lines[first_code_idx].strip()
        if first_line.startswith('"""') or first_line.startswith("'''"):
            q = first_line[:3]
            if q in first_line[3:]:
                lines[first_code_idx] = ''
            else:
                end = first_code_idx + 1
                while end < len(lines):
                    if q in lines[end]:
                        break
                    end += 1
                for j in range(first_code_idx, min(end + 1, len(lines))):
                    lines[j] = ''
            result = ''.join(lines)

    # Collapse blank lines: none before first real code, max 1 elsewhere
    out = []
    prev_blank    = False
    coding_seen   = False
    real_code_seen = False
    for line in result.splitlines(keepends=True):
        is_blank  = not line.strip()
        is_coding = bool(_CODING_LINE_RE.match(line.rstrip('\n')))
        if is_coding and not coding_seen:
            out.append(line)
            coding_seen = True
            prev_blank  = False
            continue
        if is_blank:
            if not real_code_seen:
                continue
            if not prev_blank:
                out.append(line)
            prev_blank = True
        else:
            real_code_seen = True
            prev_blank     = False
            out.append(line)

    return ''.join(out)


# ---------------------------------------------------------------------------
# YAML writing
# ---------------------------------------------------------------------------

def _yaml_scalar(text, indent=2):
    if not text:
        return '""'
    lines = text.splitlines()
    if len(lines) == 1:
        return '"{}"'.format(text.replace('\\', '\\\\').replace('"', '\\"'))
    pad  = ' ' * indent
    body = '\n'.join(pad + ln for ln in lines)
    return '|\n' + body


def _write_bundle_yaml(folder, meta):
    out_path = os.path.join(folder, 'bundle.yaml')
    content = (
        'title: {}\n'
        'author: {}\n'
        'tooltip: {}\n'
    ).format(
        _yaml_scalar(meta['title']),
        _yaml_scalar(meta['author']),
        _yaml_scalar(meta['tooltip']),
    )
    with open(out_path, 'w', encoding='utf-8') as fh:
        fh.write(content)
    print('  [YAML]   -> {}'.format(out_path))


# ---------------------------------------------------------------------------
# Per-folder processing
# ---------------------------------------------------------------------------

def _process_py_file(filepath):
    """Extract metadata and strip it from the .py file in-place."""
    try:
        with open(filepath, encoding='utf-8', errors='replace') as fh:
            source = fh.read()
    except OSError as exc:
        print('  [WARN] Cannot read {}: {}'.format(filepath, exc))
        return {}

    meta = _extract_from_py(source)
    if not any(meta.values()):
        return {}

    cleaned = _strip_metadata(source)

    # Add filename comment after coding line (or at top if no coding line)
    fname_comment = '# {}\n'.format(os.path.basename(filepath))
    lines = cleaned.splitlines(keepends=True)
    if lines and _CODING_LINE_RE.match(lines[0].rstrip('\n')):
        cleaned = lines[0] + fname_comment + ''.join(lines[1:])
    else:
        cleaned = fname_comment + cleaned

    if cleaned != source:
        try:
            with open(filepath, 'w', encoding='utf-8') as fh:
                fh.write(cleaned)
            print('  [STRIP]  {}'.format(os.path.basename(filepath)))
        except OSError as exc:
            print('  [WARN] Cannot write {}: {}'.format(filepath, exc))

    return meta


def _process_pushbutton(folder):
    bundle_path = os.path.join(folder, 'bundle.yaml')

    py_files = sorted(
        f for f in os.listdir(folder)
        if f.endswith('.py') and os.path.isfile(os.path.join(folder, f))
    )

    # Collect metadata — py files take priority, yaml fills gaps
    merged = {'title': '', 'author': '', 'tooltip': ''}

    for fname in py_files:
        meta = _process_py_file(os.path.join(folder, fname))
        for key in ('title', 'author', 'tooltip'):
            if not merged[key] and meta.get(key):
                merged[key] = meta[key]

    # Fill any still-missing fields from existing bundle.yaml
    if not all([merged['title'], merged['tooltip']]) and os.path.isfile(bundle_path):
        yaml_meta = _extract_from_yaml(bundle_path)
        if not merged['title']   and yaml_meta.get('title'):
            merged['title']   = yaml_meta['title']
        if not merged['tooltip'] and yaml_meta.get('tooltip'):
            merged['tooltip'] = yaml_meta['tooltip']

    if not any(merged.values()):
        print('  [SKIP] No usable metadata found')
        return

    # Apply formatting
    merged['title']   = _format_title(merged['title'])
    merged['author']  = AUTHOR
    if merged['tooltip']:
        merged['tooltip'] = _bold_version_line(merged['tooltip'])

    _write_bundle_yaml(folder, merged)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _find_pushbutton_folders(root):
    for dirpath, dirnames, _ in os.walk(root):
        dirnames.sort()
        for d in sorted(dirnames):
            if d.endswith('.pushbutton'):
                yield os.path.join(dirpath, d)


def main():
    root = os.path.dirname(os.path.abspath(__file__))
    print('Scanning from: {}\n'.format(root))

    count = 0
    for pb_folder in _find_pushbutton_folders(root):
        print('-> {}'.format(pb_folder))
        _process_pushbutton(pb_folder)
        count += 1

    if count == 0:
        print('No .pushbutton folders found.')
    else:
        print('\nDone - processed {} .pushbutton folder(s).'.format(count))


if __name__ == '__main__':
    main()
