OpenVBS - HotPatch Edition
====

OpenVBS is a replacement engine for VBScript. You can continue to use your existing VBScript code simply.

## Description
This project was forked from [OpenVBS](https://github.com/yrm006/openvbs). Please refer to their documentation for more details.

## How to build

[Windows]

x64 Native Tools Command Prompt for VS 2022

    > nmake

regist the dll as Administrator (if you need):

    > regsvr32 openvbs.dll

Test for CScript (need regsvr32):

    > cscript //e:openvbs your_script.vbs

Test for .NET (need regsvr32):

    > sample.exe

Test for .NET (no-need regsvr32):

    > sample-regfree.exe

## Usage

[C#]

```C#
    var engine = (IActiveScript)Activator.CreateInstance( Type.GetTypeFromProgID("OpenVBS") );

    ((IActiveScriptParse)engine).InitNew();

    var site = new Site();
    engine.SetScriptSite(site);

    engine.AddNamedItem("bridge", (uint)ScriptItemFlags.IsVisible);

    const string script = 
        "bridge.Log \"Hello from OBScript!\"\r\n" +
        "Dim total\r\n" +
        "total = bridge.Add(40, 2)\r\n" +
        "bridge.Log \"Add returned: \" & total";

    object result = null;
    ((IActiveScriptParse)engine).ParseScriptText(script, null, null, null, 0, 0, 0, out result, IntPtr.Zero);

    engine.SetScriptState(SCRIPTSTATE.STARTED);
```

[C# (RegSvr32-Free)]

```C#
    var engine = (IActiveScript)RegFree.CreateInstance("openvbs.dll", new Guid("{23ADC41D-068C-4D0B-B3F6-0792F675E1B6}"));

    // ...
```

## Author
[NaturalStyle PREMIUM](https://p.na-s.jp)

---
yrm.20251009
