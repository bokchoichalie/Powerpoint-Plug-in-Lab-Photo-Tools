using System;
using System.Runtime.InteropServices;
using System.Xml;
using LabPhotoTools;

internal static class ComBoundary
{
    [DllImport("oleaut32.dll")] private static extern IntPtr SafeArrayCreateVector(ushort vt, int lowerBound, uint count);
    [DllImport("oleaut32.dll")] private static extern int SafeArrayDestroy(IntPtr array);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Connection(IntPtr self, IntPtr app, int mode, IntPtr instance, ref IntPtr array);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Disconnection(IntPtr self, int mode, ref IntPtr array);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Lifecycle(IntPtr self, ref IntPtr array);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int Ribbon(IntPtr self, [MarshalAs(UnmanagedType.BStr)] string id, out IntPtr xml);
    private static T Function<T>(IntPtr instance, int index) where T : class
    {
        return Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(instance),index*IntPtr.Size),typeof(T)) as T;
    }
    [STAThread]
    private static int Main()
    {
        IntPtr unknown=IntPtr.Zero, lifecycle=IntPtr.Zero, ribbon=IntPtr.Zero;
        try
        {
            unknown=Marshal.GetIUnknownForObject(new Connect());
            Guid iid=new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744");
            Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown,ref iid,out lifecycle));
            for(int kind=0;kind<3;kind++)
            {
                IntPtr array=kind==0?IntPtr.Zero:SafeArrayCreateVector(12,kind==1?0:3,kind==1?0u:2u);
                try
                {
                    Marshal.ThrowExceptionForHR(Function<Connection>(lifecycle,7)(lifecycle,IntPtr.Zero,0,IntPtr.Zero,ref array));
                    Marshal.ThrowExceptionForHR(Function<Disconnection>(lifecycle,8)(lifecycle,0,ref array));
                    for(int slot=9;slot<=11;slot++) Marshal.ThrowExceptionForHR(Function<Lifecycle>(lifecycle,slot)(lifecycle,ref array));
                }
                finally { if(array!=IntPtr.Zero) SafeArrayDestroy(array); }
            }
            iid=new Guid("000C0396-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown,ref iid,out ribbon));
            IntPtr text;
            Marshal.ThrowExceptionForHR(Function<Ribbon>(ribbon,7)(ribbon,"Microsoft.PowerPoint.Presentation",out text));
            try { XmlDocument xml=new XmlDocument();xml.LoadXml(Marshal.PtrToStringBSTR(text)); }
            finally { Marshal.FreeBSTR(text); }
            Console.WriteLine("PASS: "+(IntPtr.Size*8)+"-bit native COM lifecycle and ribbon XML."); return 0;
        }
        catch(Exception ex) { Console.Error.WriteLine(ex);return 1; }
        finally { if(ribbon!=IntPtr.Zero)Marshal.Release(ribbon);if(lifecycle!=IntPtr.Zero)Marshal.Release(lifecycle);if(unknown!=IntPtr.Zero)Marshal.Release(unknown); }
    }
}

