using System;
using System.Security.Policy;
using System.Security;
using System.Security.Permissions;
using System.Reflection;
using System.Configuration;

namespace SalesUtil
{
    public enum Areas { SE, NO }

    public class Program
    {
        public static string TempatePath;
        public static string InputDataPath;
        public static string InputDataSheetName;
        public static string ProducDataPath;
        public static string ProductDataSheetName;
        public static string DuDataPath;
        public static string DuDataSheetName;
        public static string OutputRootDirectory;
        public static string InputRootDirectory;
        public static string EmlSubject;
        public static string GoogleKey;
        public static Formats Format;
        public static Areas Localization;

        static void Main(string[] args)
        {
            if (AppDomain.CurrentDomain.IsDefaultAppDomain())
            {
                // RazorEngine cannot clean up from the default appdomain...
                Console.WriteLine("Switching to secound AppDomain, for RazorEngine...");
                AppDomainSetup adSetup = new AppDomainSetup();
                adSetup.ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
                var current = AppDomain.CurrentDomain;
                // You only need to add strongnames when your appdomain is not a full trust environment.
                var strongNames = new StrongName[0];

                var domain = AppDomain.CreateDomain(
                    "MyMainDomain", null,
                    current.SetupInformation, new PermissionSet(PermissionState.Unrestricted),
                    strongNames);
                try
                {
                    var exitCode = domain.ExecuteAssembly(Assembly.GetExecutingAssembly().Location);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e);
                }
                // RazorEngine will cleanup. 
                AppDomain.Unload(domain);
                return;
            }

            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Sales Operation Utility");

            AppConfig();

            Console.WriteLine("Application Configured");

            var app = new Presentation()
            {
                Format = Format
            };
            app.GenerateLetters();
        }

        private static void AppConfig()
        {
            OutputRootDirectory = ConfigurationManager.AppSettings["OutputRootDirectory"];
            InputRootDirectory = ConfigurationManager.AppSettings["InputRootDirectory"];
            TempatePath = string.Format(@"{0}\{1}", InputRootDirectory, ConfigurationManager.AppSettings["tempatePath"]);

#if (DEBUG)
            TempatePath = ConfigurationManager.AppSettings["tempatePath"];
#endif

            InputDataPath = string.Format(@"{0}\{1}", InputRootDirectory, ConfigurationManager.AppSettings["ExcelInputFilePath"]);
            InputDataSheetName = ConfigurationManager.AppSettings["ExcelInputDataSheet"];

            ProducDataPath = string.Format(@"{0}\{1}", InputRootDirectory, ConfigurationManager.AppSettings["ExcelProducDataPath"]);
            ProductDataSheetName = ConfigurationManager.AppSettings["ExcelProductDataSheetName"];

            DuDataPath = string.Format(@"{0}\{1}", InputRootDirectory, ConfigurationManager.AppSettings["ExcelDuDataPath"]);
            DuDataSheetName = ConfigurationManager.AppSettings["ExcelDuDataSheetName"];

            GoogleKey = ConfigurationManager.AppSettings["GoogleKey"];

            EmlSubject = ConfigurationManager.AppSettings["Subject"];
            var s = ConfigurationManager.AppSettings["Format"];
            if (s.Equals("word")) Format = Formats.word;
            else
                if (s.Equals("eml")) Format = Formats.eml;
            else
                if (s.Equals("none")) Format = Formats.none;
            else
            {
                Console.WriteLine("Format option is not configured. Default value 'word' is being used");
                Format = Formats.word;
            }

            var l = ConfigurationManager.AppSettings["BusinessAreaApplication"];
            if (l.Equals("SE")) Localization = Areas.SE;
            else
                if (l.Equals("NO")) Localization = Areas.NO;
            else
            {
                Console.WriteLine("Localization option is not configured. Default value 'SE' is being used.");
                Localization = Areas.SE;
            }

        }
    }
}
