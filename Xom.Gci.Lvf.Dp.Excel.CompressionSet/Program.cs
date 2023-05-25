using log4net.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Xom.Gci.Lvf.Dp.Helper.ErrorHandling.ErrorLog;

namespace Xom.Gci.Lvf.Dp.Excel.CompressionSet
{
    class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            int exitCode = 0;
            XmlConfigurator.ConfigureAndWatch(new FileInfo(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\Config\\LogSetting.xml"));
            if (args.Length != 1)
            {
                Log(IncorrectParserParameters(), log4net.Core.Level.Error);
                return exitCode;
            }
            exitCode = FileProcessor.ParseFile(args);
            return exitCode;
        }
    }
}
