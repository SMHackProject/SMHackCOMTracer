namespace SMHackCOMTracer {
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;
    using Microsoft.Win32;
    using SMHackCore;

    public class PluginMain : IPlugin {
        [SuppressMessage("ReSharper", "UnusedParameter.Local")]
        public PluginMain(ServerInterfaceProxy proxy) {
        }

        public void Init() { }

        public ApiHook[] GetApiHooks() {
            return new[] {
                new ApiHook(
                    "ole32.dll",
                    "CoGetClassObject",
                    proxy => new CoGetClassObjectDelegate(
                        delegate(Guid rclsid, int context, IntPtr info, Guid riid, IntPtr ppv) {
                            proxy.DoLog(new LogData(rclsid));
                            return CoGetClassObject(rclsid, context, info, riid, ppv);
                        }
                    )),
                new ApiHook(
                    "ole32.dll",
                    "CoCreateInstance",
                    proxy => new CoCreateInstanceDelegate(
                        delegate(Guid rclsid, IntPtr ptr, int context, Guid riid, IntPtr ppv) {
                            proxy.DoLog(new LogData(rclsid));
                            return CoCreateInstance(rclsid, ptr, context, riid, ppv);
                        }
                    )),
                new ApiHook(
                    "ole32.dll",
                    "CoCreateInstanceEx",
                    proxy => new CoCreateInstanceExDelegate(
                        delegate(Guid rclsid, IntPtr ptr, int context, IntPtr info, uint cmq, IntPtr results) {
                            proxy.DoLog(new LogData(rclsid));
                            return CoCreateInstanceEx(rclsid, ptr, context, info, cmq, results);
                        }
                    ))
            };
        }

        [Serializable]
        public class LogData {
            public Guid Guid;
            public string Name = string.Empty;

            public LogData(Guid clsid) {
                Guid = clsid;
                var basekey = Registry.ClassesRoot.OpenSubKey($"CLSID\\{{{clsid}}}");
                if (basekey != null)
                    Name = (basekey.GetValue("") ??
                            basekey.OpenSubKey("VersionIndependentProgID")?.GetValue("") ??
                            basekey.OpenSubKey("ProgID")?.GetValue("") ??
                            basekey.OpenSubKey("InProcServer32")?.GetValue("Class") ??
                            basekey.OpenSubKey("InProcServer32")?.GetValue("") ??
                            basekey.OpenSubKey("LocalServer32")?.GetValue("")
                    )?.ToString() ?? string.Empty;
            }
        }

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CoGetClassObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            int dwClsContext,
            IntPtr pServerInfo,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out] IntPtr ppv
        );

        private delegate int CoGetClassObjectDelegate(
            Guid rclsid,
            int dwClsContext,
            IntPtr pServerInfo,
            Guid riid,
            IntPtr ppv
        );

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CoCreateInstance(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr ptr,
            int dwClsContext,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out] IntPtr ppv
        );

        private delegate int CoCreateInstanceDelegate(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr ptr,
            int dwClsContext,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [Out] IntPtr ppv
        );

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CoCreateInstanceEx(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr ptr,
            int dwClsContext,
            IntPtr pServerInfo,
            uint cmq,
            [In, Out] IntPtr pResults
        );

        private delegate int CoCreateInstanceExDelegate(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr ptr,
            int dwClsContext,
            IntPtr pServerInfo,
            uint cmq,
            [In, Out] IntPtr pResults
        );
    }
}