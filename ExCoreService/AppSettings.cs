using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExCoreService.Simple3Des;
using System.Configuration;
using Camstar.WCF.ObjectStack;

namespace ExCoreService
{
    class AppSettings
    {
        public static string DefaultOrderStatus
        {
            get
            {
                return ConfigurationManager.AppSettings["DefaultOrderStatus"];
            }
        }
        public static string DefaultUOM
        {
            get
            {
                return ConfigurationManager.AppSettings["DefaultUOM"];
            }
        }
        public static string DefaultInventoryLocation
        {
            get
            {
                return ConfigurationManager.AppSettings["DefaultInventoryLocation"];
            }
        }
        public static string ServicesMode
        {
            get
            {
                return ConfigurationManager.AppSettings["ServicesMode"];
            }
        }
        public static string Workflow
        {
            get
            {
                return ConfigurationManager.AppSettings["Workflow"];
            }
        }
        public static TimeSpan UTCOffset
        {
            get
            {
                string sUTCOffset = ConfigurationManager.AppSettings["UTCOffset"];
                string[] aUTCOffset = sUTCOffset.Split(':');
                return new TimeSpan(Int32.Parse(aUTCOffset[0]), Int32.Parse(aUTCOffset[1]), Int32.Parse(aUTCOffset[2]));
            }
        }
        public static int TimerPollingInterval
        {
            get
            {
                return Convert.ToInt16(ConfigurationManager.AppSettings["TimerPollingInterval"]);
            }
        }
        public static string ExCoreHost
        {
            get
            {
                return ConfigurationManager.AppSettings["ExCoreHost"];
            }
        }
        public static string ExCorePort
        {
            get
            {
                return ConfigurationManager.AppSettings["ExCorePort"];
            }
        }
        public static string ExCoreUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["ExCoreUsername"];
            }
        }
        public static string ExCorePassword
        {
            get
            {

                Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["ExCorePassword"]);
            }
        }
        private static UserProfile _ExCoreUserProfile = null;
        public static UserProfile ExCoreUserProfile
        {
            get
            {
                if (_ExCoreUserProfile == null)
                {
                    _ExCoreUserProfile = new UserProfile(ExCoreUsername, ExCorePassword, UTCOffset);
                }
                if (_ExCoreUserProfile.Name != ExCoreUsername || _ExCoreUserProfile.Password.Value != ExCorePassword)
                {
                    _ExCoreUserProfile = new UserProfile(ExCoreUsername, ExCorePassword, UTCOffset);
                }
                return _ExCoreUserProfile;
            }
        }
        public static string SourceUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["SourceUNCPath"];
            }
        }
        public static string SourceUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["SourceUNCPathUsername"];
            }
        }
        public static string SourceUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["SourceUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["SourceUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string SourceFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["SourceFolder"];
            }
        }
        public static string CompletedUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["CompletedUNCPath"];
            }
        }
        public static string CompletedUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["CompletedUNCPathUsername"];
            }
        }
        public static string CompletedUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["CompletedUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["CompletedUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string CompletedFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["CompletedFolder"];
            }
        }
        public static string ErrorUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["ErrorUNCPath"];
            }
        }
        public static string ErrorUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["ErrorUNCPathUsername"];
            }
        }
        public static string ErrorUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["ErrorUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["ErrorUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string ErrorFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["ErrorFolder"];
            }
        }
    }
}
