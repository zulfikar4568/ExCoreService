using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ExCoreService
{
    public class NetUtil
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct NETRESOURCE
        {
            public uint dwScope;
            public uint dwType;
            public uint dwDisplayType;
            public uint dwUsage;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpLocalName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpRemoteName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpComment;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpProvider;
        }

        public const uint NO_ERROR = 0;
        public const uint ERROR = 1;
        public const uint RESOURCETYPE_DISK = 1;
        public const uint CONNECT_UPDATE_PROFILE = 1;

        [DllImport("mpr.dll", EntryPoint = "WNetAddConnection2W")]
        public static extern uint WNetAddConnection2(ref NETRESOURCE lpNetResource, [In(), MarshalAs(UnmanagedType.LPWStr)] string lpPassword, [In(), MarshalAs(UnmanagedType.LPWStr)] string lpUsername, uint dwFlags);

        [DllImport("mpr.dll", EntryPoint = "WNetCancelConnection2W")]
        public static extern uint WNetCancelConnection2([In(), MarshalAs(UnmanagedType.LPWStr)] string lpName, uint dwFlags, [MarshalAs(UnmanagedType.LPWStr)] bool fForce);

        public static uint Connect(string UNCPath, string Username, string Password, bool IgnoreException = true)
        {
            uint statusConnect = NetUtil.NO_ERROR;
            try
            {
                if (UNCPath != "")
                {
                    if (!System.IO.Directory.Exists(UNCPath))
                    {
                        NetUtil.NETRESOURCE nrDrive = new NetUtil.NETRESOURCE() { dwType = NetUtil.RESOURCETYPE_DISK, lpRemoteName = UNCPath };
                        uint uiRes = NetUtil.WNetAddConnection2(ref nrDrive, Password, Username, 0);
                        statusConnect = uiRes;
                        if (uiRes != NetUtil.NO_ERROR)
                        {
                            Exception oException = new System.ComponentModel.Win32Exception((int)uiRes);
                            throw new Exception($"Unabled to connect to {UNCPath} with Username {Username} ({oException.Message})");
                        }
                    }
                }
                return statusConnect;
            } catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return NetUtil.ERROR;
            }
        }

        public static uint Disconnect(string UNCPath, bool IgnoreException = true)
        {
            uint statusDisconnect = NetUtil.NO_ERROR;
            try
            {
                if (UNCPath != "")
                {
                    if (System.IO.Directory.Exists(UNCPath))
                    {
                        uint uiRes = NetUtil.WNetCancelConnection2(UNCPath, CONNECT_UPDATE_PROFILE, true);
                        statusDisconnect = uiRes;
                        if (uiRes != NetUtil.NO_ERROR)
                        {
                            Exception oException = new System.ComponentModel.Win32Exception((int)uiRes);
                            throw new Exception($"Unabled to disconnect to {UNCPath} ({oException.Message})");
                        }
                    }
                }
                return statusDisconnect;
            } catch (Exception ex)
            {
                ex.Source = typeof(Program).Assembly.GetName().Name == ex.Source ? MethodBase.GetCurrentMethod().Name : MethodBase.GetCurrentMethod().Name + "." + ex.Source;
                EventLogUtil.LogErrorEvent(ex.Source, ex);
                if (!IgnoreException) throw ex;
                return NetUtil.ERROR;
            }
        }
    }
}
