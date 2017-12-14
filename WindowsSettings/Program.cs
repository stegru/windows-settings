using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace WindowsSettings
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                switch (args[0])
                {
                    case "-list-all":
                    case "-list":
                        bool listAll = args[0] == "-list-all";
                        string path = SettingItem.RegistryPath.Replace("HKEY_LOCAL_MACHINE\\", "");

                        using (RegistryKey regKey = Registry.LocalMachine.OpenSubKey(path, false))
                        {
                            foreach (string key in regKey.GetSubKeyNames())
                            {
                                string settingPath = Path.Combine(SettingItem.RegistryPath, key);
                                string typeName = Registry.GetValue(settingPath, "Type", null) as string;
                                SettingType type;
                                if (Enum.TryParse(typeName, true, out type))
                                {
                                    if (listAll || (type != SettingType.Custom && type != SettingType.SettingCollection))
                                    {
                                        Console.WriteLine("{0}: {1}", key, type);
                                    }
                                }
                            }
                        }

                        break;

                    case "-methods":
                        MethodInfo[] methods = typeof(SettingItem).GetMethods(
                            BindingFlags.Instance | BindingFlags.Public | BindingFlags.FlattenHierarchy);
                        foreach (MethodInfo method in methods)
                        {
                            if (method.IsExposed())
                            {
                                IEnumerable<string> paras = method.GetParameters().Select(p => {
                                    return string.Format(CultureInfo.InvariantCulture, "{1}: {0}",
                                        p.ParameterType.Name, p.Name);
                                });
                                Console.WriteLine("{1}({2}): {0}", method.ReturnType.Name, method.Name, String.Join(", ", paras));
                            }
                        }

                        break;

                    default:
                        break;
                }
                return;
            }

            using (Stream input = Console.OpenStandardInput())
            {
                IEnumerable<Payload> payloads = null;

                try
                {
                    payloads = Payload.FromStream(input);
                }
                catch (SerializationException e)
                {
                    Console.Error.Write("Invalid JSON: ");
                    Console.Error.WriteLine((e.InnerException ?? e).Message);
                    Environment.ExitCode = 1;
                    return;
                }

                foreach (Payload payload in payloads)
                {
                    Result result = SettingHandler.Apply(payload);
                    Console.Write(result.ToString());
                    Console.WriteLine(",");
                }
            }
        }
    }
}
