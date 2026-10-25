"""
check_names.py - finds project-wide name clashes that VB6 reports one at a
time as "Ambiguous name detected" (VB identifiers are case-insensitive).

  python tools/check_names.py [project_dir]

Reports any two Public declarations (Sub/Function/Const/variable/Enum member/
Type) with the same case-insensitive name in different places, and module
names that equal a public identifier.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DECL = [
    re.compile(r"^\s*Public\s+(?:Declare\s+)?(?:Sub|Function)\s+(\w+)", re.I),
    re.compile(r"^\s*(?:Sub|Function)\s+(\w+)", re.I),                 # default Public
    re.compile(r"^\s*Public\s+Const\s+(\w+)", re.I),
    re.compile(r"^\s*Public\s+(?:Enum|Type)\s+(\w+)", re.I),
    re.compile(r"^\s*Public\s+(?!Sub|Function|Const|Enum|Type|Declare|Property)(\w+)", re.I),
]


def scan(projdir):
    vbp = os.path.join(projdir, "DoomVB.vbp")
    files = []
    for line in open(vbp, encoding="latin-1"):
        m = re.match(r"(?:Module|Form)=(?:[^;]*;\s*)?(.+)", line.strip())
        if m:
            files.append(os.path.join(projdir, m.group(1).strip()))
    names = {}
    modules = {}
    for path in files:
        enum = False
        modname = None
        for no, line in enumerate(open(path, encoding="latin-1"), 1):
            m = re.match(r'Attribute VB_Name = "(\w+)"', line)
            if m:
                modname = m.group(1)
                modules[modname.lower()] = path
                continue
            code = line.split("'")[0]
            if re.match(r"^\s*(Public\s+)?Enum\s+\w+", code, re.I):
                if not re.match(r"^\s*Private", code, re.I):
                    enum = True
            if enum:
                if re.match(r"^\s*End\s+Enum", code, re.I):
                    enum = False
                    continue
                mm = re.match(r"^\s*(\w+)\s*(=|$)", code)
                if mm and not re.match(r"^\s*(Public\s+)?Enum", code, re.I):
                    names.setdefault(mm.group(1).lower(), []).append((path, no, mm.group(1)))
            for rx in DECL:
                mm = rx.match(code)
                if mm:
                    if path.lower().endswith(".frm") and not code.strip().lower().startswith("public"):
                        break
                    # skip "Private Sub" (rx 1 would not match it anyway)
                    names.setdefault(mm.group(1).lower(), []).append((path, no, mm.group(1)))
                    # multiple variables on one Public line
                    if rx is DECL[4]:
                        flat = code
                        while re.search(r"\([^()]*\)", flat):
                            flat = re.sub(r"\([^()]*\)", "", flat)
                        for extra in re.findall(r",\s*(\w+)", flat):
                            names.setdefault(extra.lower(), []).append((path, no, extra))
                    break
    bad = 0
    for k, v in sorted(names.items()):
        if len(v) > 1:
            bad += 1
            print("CLASH %s:" % v[0][2])
            for p, no, n in v:
                print("    %s:%d  %s" % (os.path.relpath(p, projdir), no, n))
        if k in modules:
            bad += 1
            print("CLASH module name %s with %s:%d" % (k, os.path.relpath(v[0][0], projdir), v[0][1]))
    print("%d clash(es)" % bad)
    return bad


if __name__ == "__main__":
    d = sys.argv[1] if len(sys.argv) > 1 else ROOT
    sys.exit(1 if scan(os.path.abspath(d)) else 0)
