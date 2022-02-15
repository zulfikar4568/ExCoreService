using OpcenterWikLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace ExCoreServiceProductMaster
{
    class Program
    {
        static void Main(string[] args)
        {
            var exitCode = HostFactory.Run(x =>
            {
                AppSettings.AssemblyName = "3" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                x.Service<StreamFile>(s =>
                {
                    s.ConstructUsing(streamfile => new StreamFile());
                    s.WhenStarted(streamfile => {
                        streamfile.Start();
                        EventLogUtil.LogEvent("ExCoreServiceProductMaster started successfully", System.Diagnostics.EventLogEntryType.Information, 3);
                    });
                    s.WhenStopped(streamfile => {
                        streamfile.Stop();
                        EventLogUtil.LogEvent("ExCoreServiceProductMaster stopped successfully", System.Diagnostics.EventLogEntryType.Information, 3);
                    });
                });
                x.RunAsLocalSystem();
                x.SetServiceName("ExCoreServiceProductMaster");
                x.SetDisplayName("ExCoreServiceProductMaster");
                x.SetDescription("This is service for MES Opcenter Execution Core");
            });

            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
            Environment.ExitCode = exitCodeValue;
        }
    }
}
