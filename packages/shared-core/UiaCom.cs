using System.Runtime.InteropServices;

namespace OfficeInEditor.Core;

internal static class UiaCom
{
    internal const int Name = 30005;
    internal const int AutomationId = 30011;
    internal const int NativeWindowHandle = 30020;
    internal const int BoundingRectangle = 30001;
    internal const int IsOffscreen = 30022;
    internal const int TreeDescendants = 4;
    [ThreadStatic] internal static int LastCreateHr;

    [ComImport, Guid("352FFBA8-0973-437C-A61F-F64CAFD81DF9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ICondition { }

    [ComImport, Guid("D22108AA-8AC5-49A5-837B-37BBB3D7591E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IElement
    {
        void _01();
        void _02();
        [PreserveSig] int FindFirst(int scope, ICondition condition, out IElement found);
        void _04();
        void _05();
        void _06();
        void _07();
        [PreserveSig] int GetCurrentPropertyValue(int propertyId, [MarshalAs(UnmanagedType.Struct)] out object value);
    }

    [ComImport, Guid("30CBE57D-D9D0-452A-AB13-7AC5AC4825EE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAutomation
    {
        void _01();
        void _02();
        [PreserveSig] int GetRootElement(out IElement root);
        [PreserveSig] int ElementFromHandle(IntPtr hwnd, out IElement element);
        void _05();
        [PreserveSig] int GetFocusedElement(out IElement element);
        void _07();
        void _08();
        void _09();
        void _10();
        void _11();
        void _12();
        void _13();
        void _14();
        void _15();
        void _16();
        void _17();
        void _18();
        void _19();
        void _20();
        [PreserveSig] int CreatePropertyCondition(int propertyId, [In, MarshalAs(UnmanagedType.Struct)] object value, out ICondition condition);
    }

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, uint context, ref Guid iid, out IntPtr created);

    [ThreadStatic] private static string? lastCreateHow;
    internal static string LastCreateHow => lastCreateHow ?? "";

    internal static IAutomation? Create()
    {
        LastCreateHr = 0;
        lastCreateHow = "";
        try
        {
            var clsid = new Guid("FF48DBA4-60EF-4201-AA87-54103EEF594E");
            var iid = new Guid("30CBE57D-D9D0-452A-AB13-7AC5AC4825EE");
            LastCreateHr = CoCreateInstance(ref clsid, IntPtr.Zero, 0x1, ref iid, out var unk);
            lastCreateHow = "cocreate";
            if (LastCreateHr != 0 || unk == IntPtr.Zero)
            {
                LastCreateHr = CoCreateInstance(ref clsid, IntPtr.Zero, 0x17, ref iid, out unk);
                lastCreateHow = "cocreate-all";
            }
            if (LastCreateHr != 0 || unk == IntPtr.Zero) return null;
            try
            {
                var obj = Marshal.GetObjectForIUnknown(unk);
                lastCreateHow = "getobj";
                return (IAutomation)obj;
            }
            finally { Marshal.Release(unk); }
        }
        catch (Exception error)
        {
            LastCreateHr = error.HResult;
            lastCreateHow = error.GetType().Name;
            return null;
        }
    }
}
