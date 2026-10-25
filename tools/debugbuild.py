"""
debugbuild.py - builds an instrumented copy of DoomVB for locating run time
errors in the compiled program.

  python tools/debugbuild.py [outdir]

The copy (default: build_debug/) gets line numbers on executable lines and
an error handler in every procedure that reports "module.procedure line N"
to doomvb.log and re-raises, so a crash logs a call trace.  The original
sources are not modified.
"""
import os
import re
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402

SKIP_MODULES = {"info_tables", "sounds_tables", "dstrings", "d_items", "dbg"}
PROC_START = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)", re.I)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function)\b", re.I)
NO_NUMBER = re.compile(r"^\s*(Case\b|Else\b|ElseIf\b|End\b|Loop\b|Next\b|Wend\b|Dim\b|Static\b|Const\b|"
                       r"Exit\b|#|'|Rem\b|\w+:\s*$|\w+:\s*'.*$)", re.I)

DBG_MODULE = '''Attribute VB_Name = "dbg"
Option Explicit
Public Sub DbgRaise(ByVal where As String, ByVal ln As Long)
    Dim n As Long, d As String, s As String
    n = Err.Number: d = Err.Description: s = Err.Source
    I_Printf "  at " & where & " line " & ln & " (error " & n & ": " & d & ")"
    Err.Raise n, s, d
End Sub
'''


def instrument(path, modname):
    with open(path, encoding="latin-1") as f:
        lines = f.read().split("\n")
    out = []
    inproc = None
    has_onerror = False
    body = []
    lineno = 0
    pending_header = False

    def flush_proc(endline):
        nonlocal lineno
        kind, name, header = inproc
        if has_onerror:
            out.append(header)
            out.extend(body)
            out.append(endline)
            return
        out.append(header)
        out.append("On Error GoTo DbgErr_")
        prev_cont = False
        for b in body:
            if prev_cont or NO_NUMBER.match(b) or b.strip() == "":
                out.append(b)
            else:
                lineno += 1
                out.append("%d %s" % (lineno, b.lstrip()))
            prev_cont = b.split("'")[0].rstrip().endswith(" _")
        out.append("Exit " + kind)
        out.append("DbgErr_:")
        out.append('DbgRaise "%s.%s", Erl' % (modname, name))
        out.append(endline)

    for ln in lines:
        raw = ln.rstrip("\r")
        if inproc is None:
            m = PROC_START.match(raw)
            if m and not raw.strip().lower().startswith(("declare", "public declare", "private declare")) \
                    and " Declare " not in raw:
                # one-line procedure (stub style "Sub X(): End Sub")
                if re.search(r":\s*End\s+(Sub|Function)\s*$", raw, re.I):
                    out.append(raw)
                    continue
                inproc = (m.group(1), m.group(2), raw)
                has_onerror = False
                body = []
                pending_header = raw.split("'")[0].rstrip().endswith(" _")
                # header may continue on following lines
                continue
            out.append(raw)
        else:
            if pending_header:
                kind, name, header = inproc
                inproc = (kind, name, header + "\n" + raw)
                pending_header = raw.split("'")[0].rstrip().endswith(" _")
                continue
            if PROC_END.match(raw):
                flush_proc(raw)
                inproc = None
                continue
            if re.match(r"^\s*On\s+Error\b", raw, re.I):
                has_onerror = True
            body.append(raw)
    with open(path, "w", encoding="latin-1", newline="\r\n") as f:
        f.write("\n".join(out))


def main():
    outdir = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "build_debug"))
    if os.path.exists(outdir):
        shutil.rmtree(outdir)
    for sub in ("src", "stubs", "res"):
        shutil.copytree(os.path.join(ROOT, sub), os.path.join(outdir, sub))
    with open(os.path.join(outdir, "src", "dbg.bas"), "w", newline="\r\n") as f:
        f.write(DBG_MODULE)
    for name in os.listdir(os.path.join(outdir, "src")):
        if name.lower().endswith(".bas"):
            modname = os.path.splitext(name)[0]
            if modname in SKIP_MODULES:
                continue
            instrument(os.path.join(outdir, "src", name), modname)
    vbp = os.path.join(outdir, "DoomVB.vbp")
    sys.argv = [sys.argv[0], vbp]
    rc = build.main()
    # run from the project folder so the WAD and config are found
    exe = os.path.join(outdir, "DoomVB.exe")
    if os.path.exists(exe):
        shutil.copy(exe, os.path.join(ROOT, "DoomVB_debug.exe"))
        print("copied to DoomVB_debug.exe")
    return rc


if __name__ == "__main__":
    sys.exit(main())
