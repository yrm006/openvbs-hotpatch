# Compatibility and validation

This page intentionally distinguishes **documented integration examples** from **verified compatibility**. A sample in the repository is not a certification of a particular host, Windows version or VBScript language feature.

## Status definitions

- **Example available:** the repository includes a sample or instructions; environment-specific behavior still needs verification.
- **Verified:** reproducible test results are published with Windows version, architecture, build and expected/actual output.
- **Known limitation:** a reproducible difference has been confirmed.
- **Not yet verified:** no current test result is published here.

## Coverage to validate

| Area | Documentation | What to verify |
| --- | --- | --- |
| `IActiveScript` / `IActiveScriptParse` | Example available (`sample.cs`) | Engine lifecycle, site callbacks, named items, errors, script state |
| `cscript //e:openvbs` | Invocation documented | `WScript` host object; argument handling; output and exit codes |
| Registration-free COM | Example available (`sample-regfree.cs`) | Host architecture, DLL location, loader failure behavior |
| Core VBScript syntax | Not yet verified here | Procedures, scoping, variants, errors, arrays, classes, `With` |
| Built-in functions and objects | Not yet verified here | Conversion, date/time, string, filesystem, `RegExp`, `Dictionary` |
| COM automation | Not yet verified here | `CreateObject`, `GetObject`, IDispatch, eventing, threading |
| Classic ASP and IIS | Not yet verified here | ASP engine selection, intrinsics, sessions, ADO, errors |
| Windows Server 2028 Preview | Not yet verified here | Build, load, execution, sandbox and security policy differences |
| Transparent `vbscript.dll` replacement | Not claimed | **Do not assume this is supported** |

## Minimal smoke tests

These are *suggested tests to run*, **not already-passing results**. Save as separate `.vbs` files and execute with `cscript //e:openvbs file.vbs` after registering the engine.

### Language basics

```vbscript
Option Explicit
Dim x
x = 40 + 2
WScript.Echo x
```

Expected output: `42`.

### Functions and arrays

```vbscript
Option Explicit
Function Twice(n)
    Twice = n * 2
End Function
Dim a
a = Array(1, 2, 3)
WScript.Echo Twice(a(1))
```

Expected output: `4`.

### COM automation (depends on system registration)

```vbscript
Option Explicit
Dim d
Set d = CreateObject("Scripting.Dictionary")
d.Add "answer", 42
WScript.Echo d("answer")
```

Expected output: `42` if the host and relevant COM object are available.

## Reporting results

Please include: exact Windows version/build; x64 or x86; OpenVBS commit; build environment; command line or host API used; expected versus actual output; sanitized minimal sample; and whether Microsoft's VBScript engine behaves differently.

For a new result or failure, [open an issue](https://github.com/yrm006/openvbs-hotpatch/issues). Avoid including credentials and proprietary scripts.
