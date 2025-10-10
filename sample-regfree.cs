using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;



public class RegFree{
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern IntPtr LoadLibrary(string lpFileName);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetModuleHandle(string lpModuleName);
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern IntPtr GetProcAddress(IntPtr hModule, string procName);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DllGetClassObjectDelegate(
        [In] ref Guid clsid, 
        [In] ref Guid iid, 
        [MarshalAs(UnmanagedType.Interface)] out IClassFactory obj);

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000001-0000-0000-C000-000000000046")]
    interface IClassFactory{
        void CreateInstance(object outer, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object obj);
        void LockServer(bool fLock);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000000-0000-0000-C000-000000000046")]
    interface IUnknown{
        int QueryInterface(ref Guid guid, [MarshalAs(UnmanagedType.Interface)] out object obj);
        int AddRef();
        int Release();
    }

    public static object CreateInstance(string dllpath, Guid clsid){
        IntPtr hLib = GetModuleHandle(dllpath);
        if(hLib == IntPtr.Zero){
            hLib = LoadLibrary(dllpath);
        }
        IntPtr pGetClassObject = GetProcAddress(hLib, "DllGetClassObject");

        var dGetClassObject = (DllGetClassObjectDelegate)Marshal.GetDelegateForFunctionPointer(pGetClassObject, typeof(DllGetClassObjectDelegate));

        IClassFactory factory;
        Guid iid_factory = typeof(IClassFactory).GUID;
        dGetClassObject(ref clsid, ref iid_factory, out factory);

        object obj;
        Guid iid = typeof(IUnknown).GUID;
        factory.CreateInstance(IntPtr.Zero, ref iid, out obj);

        return obj;
    }
}



[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("EAE1BA61-A4ED-11cf-8F20-00805F2CD064")]
public interface IActiveScriptError
{
    void GetExceptionInfo(out System.Runtime.InteropServices.ComTypes.EXCEPINFO info);
    void GetSourcePosition(out uint line, out int charPosition, [MarshalAs(UnmanagedType.LPWStr)] out string sourceLine);
    void GetSourceLineText([MarshalAs(UnmanagedType.LPWStr)] out string sourceLine);
}

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("DB01A1E3-A42B-11cf-8F20-00805F2CD064")]
public interface IActiveScriptSite
{
    void GetLCID(out int lcid);
    void GetItemInfo([MarshalAs(UnmanagedType.LPWStr)] string name, uint mask, [MarshalAs(UnmanagedType.Interface)] out object item, out ITypeInfo typeInfo);
    void GetDocVersionString([MarshalAs(UnmanagedType.LPWStr)] out string version);
    void OnScriptTerminate([MarshalAs(UnmanagedType.Interface)] object result, IntPtr exceptionInfo);
    void OnStateChange(SCRIPTSTATE state);
    void OnScriptError(IActiveScriptError err);
    void OnEnterScript();
    void OnLeaveScript();
}

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("BB1A2AE1-A4F9-11CF-8F20-00805F2CD064")]
public interface IActiveScript
{
    void SetScriptSite([MarshalAs(UnmanagedType.Interface)] IActiveScriptSite site);
    void GetScriptSite(ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object site);
    void SetScriptState(SCRIPTSTATE state);
    void GetScriptState(out SCRIPTSTATE state);
    void Close();
    void AddNamedItem([MarshalAs(UnmanagedType.LPWStr)] string name, uint flags);
    void AddTypeLib([In] ref Guid libid, uint dwMajor, uint dwMinor, uint flags);
    void GetScriptDispatch([MarshalAs(UnmanagedType.LPWStr)] string itemName, [MarshalAs(UnmanagedType.Interface)] out object disp);
    void InvokeScript();
}

[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17")]
public interface IActiveScriptParse
{
    void InitNew();
    void AddScriptlet(
        [MarshalAs(UnmanagedType.LPWStr)] string defaultName,
        [MarshalAs(UnmanagedType.LPWStr)] string code,
        [MarshalAs(UnmanagedType.LPWStr)] string itemName,
        [MarshalAs(UnmanagedType.LPWStr)] string subItemName,
        [MarshalAs(UnmanagedType.LPWStr)] string eventName,
        [MarshalAs(UnmanagedType.LPWStr)] string delimiter,
        uint sourceContextCookie,
        uint startingLineNumber,
        uint flags,
        [MarshalAs(UnmanagedType.LPWStr)] out string name,
        IntPtr pexcepinfo);
    void ParseScriptText(
        [MarshalAs(UnmanagedType.LPWStr)] string code,
        [MarshalAs(UnmanagedType.LPWStr)] string itemName,
        [MarshalAs(UnmanagedType.Interface)] object punkContext,
        [MarshalAs(UnmanagedType.LPWStr)] string delimiter,
        uint sourceContextCookie,
        uint startingLineNumber,
        uint flags,
        [MarshalAs(UnmanagedType.Interface)] out object result,
        IntPtr pexcepinfo);
}

public enum SCRIPTSTATE
{
    UNINITIALIZED = 0,
    STARTED = 1,
    CONNECTED = 2,
    DISCONNECTED = 3,
    CLOSED = 4,
    INITIALIZED = 5
}

public enum ScriptItemFlags
{
    None = 0,
    IsVisible = 0x00000002,
    IsSource = 0x00000004,
    GlobalMembers = 0x00000008,
    IsPersistent = 0x00000040,
    CodeOnly = 0x00000200,
    NoCode = 0x00000400
}



public class Site : IActiveScriptSite{
    Bridge m_bridge;

    public Site(){
        m_bridge = new Bridge();
    }

    // IActiveScriptSize
    public void GetLCID(out int lcid){
        lcid = 0;
    }
    public void GetItemInfo(string name, uint mask, out object item, out ITypeInfo typeInfo){
        item = null;
        typeInfo = null;
        if(name == "bridge"){
            item = m_bridge;
        }
    }
    public void GetDocVersionString(out string version){
        version = null;
    }
    public void OnScriptTerminate(object result, IntPtr exceptionInfo){

    }
    public void OnStateChange(SCRIPTSTATE state){

    }
    public void OnScriptError(IActiveScriptError err){

    }
    public void OnEnterScript(){

    }
    public void OnLeaveScript(){

    }
}

public class Bridge{
    public void Log(string message){
        Console.WriteLine($"[Bridge] {message}");
    }

    public double Add(double a, double b){
        var result = a + b;
        Console.WriteLine($"[Bridge] {a} + {b} = {result}");
        return result;
    }
}

public static class Sample{
    public static void Main(){
        var engine = (IActiveScript)RegFree.CreateInstance("openvbs.dll", new Guid("{23ADC41D-068C-4D0B-B3F6-0792F675E1B6}"));

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
    }
}
