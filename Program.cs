namespace Fix
{
    using System;
    using System.Linq;
    using Microsoft.Win32;

    // based on // https://www.howtogeek.com/howto/windows-vista/make-windows-vista-explorer-preview-pane-work-for-more-file-types/
    public class Program
    {
        private const string KeyDefaultValueName = "";
        private const string KeyShellex = "shellex";
        private const string KeyPreviewHandler = "{8895b1c6-b41f-4c1c-a562-0d564250836f}";
        private const string KeyWindowsTxtPreviewHandler = "{1531d583-8375-4d3f-b5fb-d23bbd169f22}";
        private const string KeyPreviewHandlersRegistrationPath = "Software\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers";

        public static void Main(string[] args)
        {
            if (!CheckPreviewHandlerInstalled(KeyWindowsTxtPreviewHandler))
            {
                Console.WriteLine("Default handler is not installed.");
                ListInstalledPreviewHandlers();
                return;
            }

            Console.WriteLine("Enter file extension (starting with '.') to register default Txt preview handler.");
            string ext = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(ext))
            {
                Console.WriteLine("You should enter something");
                return;
            }

            ext = ext.Trim();
            if (!ext.StartsWith("."))
            {
                Console.WriteLine("Extension should start with a dot like '.cs', or '.txt'.");
                return;
            }

            RegistryKey extKey;

            try
            {
                extKey = Registry.ClassesRoot.OpenSubKey(ext, true);
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred. You should probably run this app with administator rights.");
                Console.WriteLine(e.Message);
                return;
            }

            extKey ??= Registry.ClassesRoot.CreateSubKey(ext);

            if (extKey == null)
            {
                return;
            }

            // check default for redirection
            var keyDefaultValue = extKey.GetValue(KeyDefaultValueName);
            if (keyDefaultValue != null)
            {
                // redirect
                extKey.Close();
                extKey.Dispose();

                Console.WriteLine("No support for redirect yet.");
            }
            else
            {
                // direct
                var key1 = extKey.CreateSubKey(KeyShellex);

                extKey.Close();
                extKey.Dispose();

                if (key1 == null)
                {
                    return;
                }

                var key2 = key1.CreateSubKey(KeyPreviewHandler);

                key1.Close();
                key1.Dispose();

                if (key2 == null)
                {
                    return;
                }

                var defaultValue = key2.GetValue(KeyDefaultValueName);
                if (defaultValue == null)
                {
                    key2.SetValue(KeyDefaultValueName, KeyWindowsTxtPreviewHandler, RegistryValueKind.String);
                }

                key2.Close();
                key2.Dispose();
            }

            Console.WriteLine("Done. Press enter to exit.");
            Console.ReadLine();
        }

        private static void ListInstalledPreviewHandlers()
        {
            var handlersKey = Registry.LocalMachine.OpenSubKey(KeyPreviewHandlersRegistrationPath, false);
            if (handlersKey == null)
            {
                return;
            }

            Console.WriteLine("Installed PreviewHandlers in the registry.");
            foreach (var key in handlersKey.GetValueNames())
            {
                var value = handlersKey.GetValue(key);
                Console.WriteLine($"- {key ?? ""} = {value ?? ""}");
            }
        }

        private static bool CheckPreviewHandlerInstalled(string key)
        {
            var handlersKey = Registry.LocalMachine.OpenSubKey(KeyPreviewHandlersRegistrationPath, false);
            if (handlersKey == null)
            {
                return false;
            }

            if (!handlersKey.GetValueNames().Contains(key))
            {
                return false;
            }

            var value = handlersKey.GetValue(key);
            if (value == null)
            {
                return false;
            }

            return true;
        }
    }
}
