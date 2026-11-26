# OpenVBS-HotPatch

**Keep your VBScript code alive.**

OpenVBS-HotPatch is an open-source, VBScript-compatible scripting engine for Windows applications that use **Active Scripting (`IActiveScript`)**.

Microsoft is retiring VBScript, including [a planned removal from Windows Server 2028](https://learn.microsoft.com/en-us/windows-server/get-started/removed-deprecated-features-windows-server?tabs=ws28-preview). If you maintain an application that embeds VBScript, OpenVBS-HotPatch is an alternative engine you can **evaluate without immediately rewriting all of your scripts**.

> **Compatibility notice:** OpenVBS-HotPatch is an independent implementation, **not** Microsoft's VBScript engine. It is not a guaranteed drop-in replacement for every script, COM host, or Classic ASP application. Windows Server 2028 compatibility has **not been verified** here. Test your application before relying on it in production.

## Who is this for?

| Your situation | Where to start |
| --- | --- |
| You maintain a Windows application hosting scripts through `IActiveScript` | [Embed the engine](#2-embed-openvbs-in-a-windows-application) |
| You run `.vbs` files with Windows Script Host | [Try CScript with an explicitly selected engine](#1-run-a-vbs-file-with-cscript) |
| You want to host the engine without COM registration | [Try registration-free hosting](#3-try-registration-free-hosting) |
| You maintain a Classic ASP application | Read [Classic ASP and IIS](#classic-asp-and-iis) before testing |
| You need the cross-platform language engine | See [OpenVBS](https://github.com/yrm006/openvbs) |

## What this project provides

- A Windows COM scripting engine implementing the `IActiveScript` hosting interface, built on the [OpenVBS](https://github.com/yrm006/openvbs) / Jujube language core.
- A `cscript` invocation example using an explicitly selected `OpenVBS` engine.
- A C# Active Scripting host example (`sample.cs`).
- A C# example of loading the engine without COM registration (`sample-regfree.cs`).

These are integration paths demonstrated in the repository, **not a claim of complete VBScript API or host compatibility**. See [COMPATIBILITY.md](COMPATIBILITY.md) for the scope of validation still needed.

## Quick start

### Download

Download a [Release package](./openvbs-hotpatch/releases).

### Install

Unpack the package to your directory and regsvr32 it to your registry.(see README.txt inside)

## Quick start (from source code)

### Prerequisites

- Windows and Visual Studio 2022 C++ build tools.
- **x64 Native Tools Command Prompt for VS 2022**, as used in the project's existing build instructions.
- Administrator privileges **only if** registering the DLL system-wide.

### Build

In the x64 Native Tools Command Prompt, from the repository directory:

```bat
nmake
```

The build is expected to produce `openvbs.dll`. Before shipping any binary, verify its architecture and dependencies on your target system.

### 1. Run a `.vbs` file with CScript

From an elevated command prompt, **if registration is needed**:

```bat
regsvr32 openvbs.dll
```

Then explicitly select OpenVBS:

```bat
cscript //e:openvbs your_script.vbs
```

`//e:openvbs` selects this engine for the command; it does **not** replace Windows' default `VBScript` engine or `vbscript.dll`.

### 2. Embed OpenVBS in a Windows application

The repository contains `sample.cs`, which creates the engine through the `OpenVBS` ProgID and uses Active Scripting interfaces:

```csharp
var engine = (IActiveScript)Activator.CreateInstance(
    Type.GetTypeFromProgID("OpenVBS"));

((IActiveScriptParse)engine).InitNew();

var site = new Site();
engine.SetScriptSite(site);
engine.AddNamedItem("bridge", (uint)ScriptItemFlags.IsVisible);

const string script =
    "bridge.Log \"Hello from OpenVBS!\"\r\n" +
    "Dim total\r\n" +
    "total = bridge.Add(40, 2)\r\n" +
    "bridge.Log \"Add returned: \" & total";

object result = null;
((IActiveScriptParse)engine).ParseScriptText(
    script, null, null, null, 0, 0, 0, out result, IntPtr.Zero);
engine.SetScriptState(SCRIPTSTATE.STARTED);
```

This is an excerpt, **not a standalone compilable program**. See the complete `sample.cs` for the interop definitions, script site and bridge implementation. COM registration is required for this specific ProgID-based example.

### 3. Try registration-free hosting

The repository also contains `sample-regfree.cs`, which loads a DLL directly:

```csharp
var engine = (IActiveScript)RegFree.CreateInstance(
    "openvbs.dll", new Guid("{23ADC41D-068C-4D0B-B3F6-0792F675E1B6}"));
```

See the complete sample for the `RegFree` implementation. This approach avoids `regsvr32` in that host example; it does **not** guarantee compatibility with arbitrary existing applications that expect Microsoft's VBScript CLSID or ProgID.

## Classic ASP and IIS

Classic ASP is a separate compatibility target. Active Scripting support alone does not establish full compatibility with `asp.dll`, ASP's intrinsic objects, IIS configuration, ADO, or an existing website.

If you maintain a Classic ASP site, first evaluate a small, isolated test page. Do **not** assume that merely installing this DLL will redirect existing `Language="VBScript"` pages to OpenVBS. Please report the precise IIS/ASP configuration and test results.

For the broader, cross-platform engine and ongoing Classic ASP work, see [OpenVBS](https://github.com/yrm006/openvbs).

## Compatibility and limitations

| Area | Current documentation status |
| --- | --- |
| COM / `IActiveScript` integration | Engine and C# sample are provided; host-by-host validation required |
| Windows Script Host / `cscript` | Explicit `//e:openvbs` usage is documented; verify your own scripts and host objects |
| COM registration-free hosting | C# sample is provided; limited to compatible hosting approaches |
| VBScript language and built-in objects | Compatibility is a project goal; no complete conformance claim |
| Classic ASP / IIS | Evaluate separately; full compatibility is not claimed |
| Windows Server 2028 Preview | Not verified in this documentation |
| Transparent replacement of `vbscript.dll` | Not provided or claimed |

See [COMPATIBILITY.md](COMPATIBILITY.md) for a reproducible checklist.

## Report a compatibility issue

**Real-world compatibility reports are especially valuable.** If you depend on VBScript or `IActiveScript`, please [open an issue](https://github.com/yrm006/openvbs-hotpatch/issues) with:

1. Windows version, architecture (x64/x86), and host (`cscript`, an in-house COM host, IIS, etc.).
2. OpenVBS-HotPatch commit or build version and the exact way you loaded it.
3. A **minimal, sanitized** VBScript example or host call sequence.
4. Expected result with Microsoft's VBScript engine, actual result with OpenVBS, and any error message.

Remove secrets, credentials, personal data and proprietary business logic before posting. Do not test on a production system without a rollback plan.

## Why this exists

Some applications cannot be migrated to PowerShell or JavaScript just by rewriting a script file: they use the Active Scripting COM interfaces to embed VBScript in existing software. OpenVBS-HotPatch explores preserving that *hosting model* with a separately maintained language engine.

This project is based on [OpenVBS](https://github.com/yrm006/openvbs), a modern BASIC engine powered by the original Jujube core.

## License and attribution

MIT License — see [LICENSE](LICENSE).

Author: [NaturalStyle PREMIUM](https://p.na-s.jp/).

OpenVBS-HotPatch is an independent open-source project and is not affiliated with or endorsed by Microsoft.
