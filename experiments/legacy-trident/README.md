# Trident + DSOFramer feasibility probe

Development experiment, outside the production VSIX. G0 only observes HTML
and empty controls; explicit G1 modes open/create and save owned synthetic DOCX. No Office software is included. See
[the analysis](../../docs/LEGACY_TRIDENT_PROBE.md) for the two separate hosting
boundaries and the evidence limits.

## Build

From the repository root in Windows PowerShell:

```powershell
./experiments/legacy-trident/build.ps1
# Optional: download pinned source, verify Git blob hashes, then compile an x86 OCX.
./experiments/legacy-trident/build.ps1 -BuildDsoSample
```

The executables require Windows with .NET Framework 4.x and MSHTML. Building the
optional OCX also requires Visual C++ x86 Build Tools, MIDL and the Windows SDK.
The output is ignored under `.local/build/legacy-trident/`; no installation,
registry writes, `regsvr32`, or browser policy edits are performed by the build.
The original upstream project is deliberately not executed: it contains a
registration post-build action.

## Run only synthetic G0 checks

Run from an interactive desktop, in the isolated experiment directory. The
probe creates its own small window, records a single JSON report, and closes
normally. G0 has a ten-second UI deadline; G1 uses an external supervisor deadline. The optional output file can
be supplied as `--output=<absolute path within the experiment directory>`.

```powershell
./.local/build/legacy-trident/TridentProbe.exe
# These need the source-built oie-legacy-framer.ocx alongside TridentDsoProbe.exe.
./.local/build/legacy-trident/TridentDsoProbe.exe --dso-object
./.local/build/legacy-trident/TridentDsoProbe.exe --dso-native
```

`html` checks the self-authored HTML/Trident engine only. `dso-object` attempts
an empty OBJECT element; `dso-native` is an empty AxHost control comparison.
G0 does not open documents or script the Office control.
Return code 0 means the observation completed, **not** that the control loaded:
inspect `framerNativeWindowPresent`. Code 1 means a caught exception; code 2
means bad arguments; code 3 means a blocking kill bit or incomplete policy
preflight; code 4 means the ten-second deadline expired.

The report also compares a fixed HTML overlay's DOM hit with the OS native
window hit, without moving or clicking the mouse. It records an actual hit
class only for windows owned by this probe; unrelated coverage is inconclusive.
The own-child geometry route is supplementary. Neither result verifies pixel
composition, editing, or IDE menus.

Before either DSO mode, a read-only preflight checks the original and private
CLSID against known IE/Office COM kill-bit locations in HKLM/HKCU and both
registry views. A match or unreadable check stops activation. The private GUID
only avoids collisions; it must not bypass a system block on the original.
Additional host policy still applies. If a host or operating-system security
block is observed, record it and stop that path. Do not change IObjectSafety,
IE zones, feature controls, kill bits, system time, or other policies to pass.

## Source and redistribution

The optional source is from the fixed mirror commit
[`shines77/DsoFramer@224163f5`](https://github.com/shines77/DsoFramer/tree/224163f5a5acb692118a5be95e5e49fe750bfd50).
That mirror is not a Microsoft-owned repository; the claimed original installer
has not been independently authenticated. The build preserves source notices
and downloads the mirror LICENSE and original Microsoft sample EULA alongside
the artifacts. Do not describe all Microsoft sample code as MIT solely because
the mirror contains an MIT license. Distribution review remains a separate
gate before production packaging.

The experimental copy changes COM identities, removes registration routines
and automatic type-library registration, uses `LoadTypeLibEx(REGKIND_NONE)`,
fixes the icon-destruction API, and replaces an obsolete resource header. It
keeps IObjectSafety unchanged. The legacy x86 code is not an audited or
supported Office container and is not claimed to support modern Office yet.

## G1 synthetic Word checks

Build the compiled Word probe and this probe/OCX before testing. Then run
`./experiments/office-probe-common/deploy.ps1 -Remote <ssh-alias> -Mode G1Open`
or `-Mode G1Create`. Use `-Mode G1HtmlOpen` for the HTML OBJECT variant.
Each uses a separate fresh owned fixture, never an
unverified user path. The x86 control and x64 Office compatibility is measured,
not assumed. Kill bits and IObjectSafety remain enforced.

G1 calls the pinned private native interface with PreserveSig, records actual
HRESULTs and a 48-entry memory-only native trace. Trace stages: 10=Open,
11=CreateNew, 12=CreateFromFile returned, 13=IPActivateView returned,
20=GetClassFile, 21=moniker BindToObject, 22=IPersistFile.Load,
30=ActivateMe, 31=CreateView, 32=UIActivate, 33=SetActiveObject,
34=SetRect, 35=Show, 40=DoVerb. A missing active
document is a failure, not a reason to silently retry. Save uses the original
source for Open or the owned target for CreateNew. Shared application instances
are never quit. Failed saves retain their container for recovery.

Success requires passed/documentVerified, normal closure and exit 0. The outer
G1 also requires a real Word child HWND belonging to the bound Word process;
a successful save alone cannot pass embedding. Host DPI is PerMonitorV2 and
host/document DPI awareness values are recorded. The HTML variant obtains
the private control through the native host's IOleContainer, so the page has
no script automation access. IObjectSafety and browser policies are unchanged.
Both variants perform two bounded periodic saves without closing/reopening
the document and require the same native document HWND throughout.

The outer
supervisor reports retained PIDs without terminating Office. A UI timer cannot
interrupt blocking COM. Old G0 observations do not prove editable Word or IDE
composition. UseCachedSource rebuilds only the previously verified private
cache; the default build fetches and verifies the pinned source again.
