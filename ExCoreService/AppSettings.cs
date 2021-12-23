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
        public static TimeSpan UTCOffset 
        {
            get
            {
                string sUTCOffset = ConfigurationManager.AppSettings["UTCOffset"];
                string[] aUTCOffset = sUTCOffset.Split(':');
                return new TimeSpan(Int32.Parse(aUTCOffset[0]), Int32.Parse(aUTCOffset[1]), Int32.Parse(aUTCOffset[2]));
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
                //Console.WriteLine(oSimple3Des.EncryptData(ConfigurationManager.AppSettings["ExCorePasswordReal"]));
                //Console.WriteLine(oSimple3Des.DecryptData(ConfigurationManager.AppSettings["ExCorePassword"]));
                return oSimple3Des.DecryptData(ConfigurationManager.AppSettings["ExCorePassword"]);

                //return ConfigurationManager.AppSettings["ExCorePasswordKey"];
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

        public static string SourceFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["SourceFolder"];
            }
        }

        public static string CompletedFolder
        {
            get
            {
                return ConfigurationManager.AppSettings["CompletedFolder"];
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
