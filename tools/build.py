"""
build.py - builds DoomVB.exe with the Visual Basic 6 command line compiler.

  python tools/build.py [project.vbp]

* normalises every source file to CRLF line endings (VB6 requires them)
* compiles res/doomvb.rc -> res/doomvb.res with the VB6 resource compiler
* runs VB6.EXE /make and prints the compile log (first error, if any)
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
RC = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE"


def normalise(folder):
    for name in os.listdir(folder):
        if not name.lower().endswith((".bas", ".frm", ".cls", ".vbp", ".ctl")):
            continue
        path = os.path.join(folder, name)
        with open(path, "rb") as f:
            data = f.read()
        fixed = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        if fixed != data:
            with open(path, "wb") as f:
                f.write(fixed)


# cluster -> (stub module, modules that replace it)
CLUSTERS = {
    "A": ("stubs_r", ["r_main", "r_bsp", "r_segs", "r_plane", "r_things", "r_draw", "r_data", "r_sky"]),
    "B": ("stubs_p", ["p_setup", "p_tick", "p_mobj", "p_maputl", "p_map", "p_sight", "p_user", "p_inter"]),
    "C": ("stubs_ai", ["p_enemy", "p_pspr"]),
    "D": ("stubs_spec", ["p_spec", "p_lights", "p_switch", "p_plats", "p_doors", "p_ceilng", "p_floor",
                         "p_telept", "p_saveg"]),
    "E": ("stubs_g", ["g_game", "m_menu", "hu_lib", "hu_stuff"]),
    "F": ("stubs_ui", ["st_lib", "st_stuff", "am_map", "wi_stuff", "f_finale", "f_wipe"]),
    "G": ("stubs_snd", ["s_sound", "i_sound"]),
}

VBP_TAIL = """ResFile32="res\\doomvb.res"
Startup="Sub Main"
HelpFile=""
Title="DoomVB"
ExeName32="DoomVB.exe"
Command32=""
Name="DoomVB"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionComments="Visual Basic 6 port of the DOOM source release (GNU GPL v2)"
VersionCompanyName="DoomVB"
VersionFileDescription="DoomVB"
VersionProductName="DoomVB"
CompilationType=0
OptimizationType=2
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck={bounds}
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
"""


def module_name(path):
    with open(path, "r", encoding="latin-1") as f:
        for line in f:
            if line.startswith("Attribute VB_Name"):
                return line.split('"')[1]
    return os.path.splitext(os.path.basename(path))[0]


def write_vbp(projdir, vbp, release):
    """Project = every module in src/ plus the stub of each cluster not yet present."""
    src = os.path.join(projdir, "src")
    present = {os.path.splitext(n)[0].lower() for n in os.listdir(src)}
    lines = ["Type=Exe",
             "Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
             "C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation"]
    for n in sorted(os.listdir(src)):
        if n.lower().endswith(".frm"):
            lines.append("Form=src\\" + n)
    for n in sorted(os.listdir(src)):
        if n.lower().endswith(".bas"):
            lines.append("Module=%s; src\\%s" % (module_name(os.path.join(src, n)), n))
    stubdir = os.path.join(projdir, "stubs")
    for key, (stub, mods) in sorted(CLUSTERS.items()):
        have = [m for m in mods if m in present]
        if have and len(have) != len(mods):
            print("WARNING: cluster %s is incomplete: missing %s" %
                  (key, ", ".join(m for m in mods if m not in present)))
        if not have:
            if not os.path.exists(os.path.join(stubdir, stub + ".bas")):
                print("ERROR: missing stub", stub)
                sys.exit(1)
            lines.append("Module=%s; stubs\\%s.bas" % (stub, stub))
    lines.append(VBP_TAIL.format(bounds="-1" if release else "0"))
    with open(vbp, "w", newline="\r\n") as f:
        f.write("\n".join(lines))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    release = "--release" in sys.argv
    vbp = args[0] if args else os.path.join(ROOT, "DoomVB.vbp")
    vbp = os.path.abspath(vbp)
    projdir = os.path.dirname(vbp)
    write_vbp(projdir, vbp, release)
    normalise(projdir)
    for sub in ("src",):
        if os.path.isdir(os.path.join(projdir, sub)):
            normalise(os.path.join(projdir, sub))

    resdir = os.path.join(projdir, "res")
    rcfile = os.path.join(resdir, "doomvb.rc")
    if os.path.exists(rcfile):
        r = subprocess.run([RC, "/r", "/fo", os.path.join(resdir, "doomvb.res"), rcfile],
                           cwd=resdir, capture_output=True, text=True)
        if r.returncode != 0:
            print("RC failed:", r.stdout, r.stderr)
            return 1

    log = os.path.join(projdir, "build.log")
    if os.path.exists(log):
        os.remove(log)
    subprocess.run([VB6, "/make", vbp, "/out", log], cwd=projdir)
    with open(log, "r", encoding="latin-1") as f:
        text = f.read().strip()
    print(text)
    return 0 if "succeeded" in text else 1


if __name__ == "__main__":
    sys.exit(main())
