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

        #region SOURCE FOLDER
        public static string OrderSourceUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderSourceUNCPath"];
            }
        }
        public static string OrderSourceUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderSourceUNCPathUsername"];
            }
        }
        public static string OrderSourceUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderSourceUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderSourceUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderSourceFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderSourceFolder"];
            }
        }

        public static string OrderBOMSourceUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMSourceUNCPath"];
            }
        }
        public static string OrderBOMSourceUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMSourceUNCPathUsername"];
            }
        }
        public static string OrderBOMSourceUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderBOMSourceUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderBOMSourceUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderBOMSourceFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMSourceFolder"];
            }
        }

        #endregion

        #region COMPLETED FOLDER
        public static string OrderCompletedUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderCompletedUNCPath"];
            }
        }
        public static string OrderCompletedUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderCompletedUNCPathUsername"];
            }
        }
        public static string OrderCompletedUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderCompletedUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderCompletedUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderCompletedFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderCompletedFolder"];
            }
        }

        public static string OrderBOMCompletedUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMCompletedUNCPath"];
            }
        }
        public static string OrderBOMCompletedUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMCompletedUNCPathUsername"];
            }
        }
        public static string OrderBOMCompletedUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderBOMCompletedUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderBOMCompletedUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderBOMCompletedFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMCompletedFolder"];
            }
        }
        #endregion

        #region ERROR FOLDER
        public static string OrderErrorUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderErrorUNCPath"];
            }
        }
        public static string OrderErrorUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderErrorUNCPathUsername"];
            }
        }
        public static string OrderErrorUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderErrorUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderErrorUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderErrorFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderErrorFolder"];
            }
        }

        public static string OrderBOMErrorUNCPath
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMErrorUNCPath"];
            }
        }
        public static string OrderBOMErrorUNCPathUsername
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMErrorUNCPathUsername"];
            }
        }
        public static string OrderBOMErrorUNCPathPassword
        {
            get
            {
                if (ConfigurationManager.AppSettings["OrderBOMErrorUNCPathPassword"] != "")
                {
                    Simple3Des oSimple3Des = new Simple3Des(ConfigurationManager.AppSettings["ExCorePasswordKey"]);
                    return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["OrderBOMErrorUNCPathPassword"]);
                }
                else
                {
                    return "";
                }
            }
        }
        public static string OrderBOMErrorFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["OrderBOMErrorFolder"];
            }
        }
        #endregion
    }
}
