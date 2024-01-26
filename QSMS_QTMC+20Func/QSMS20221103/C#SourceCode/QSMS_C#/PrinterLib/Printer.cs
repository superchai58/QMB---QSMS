using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
//using System.Threading.Tasks;

namespace PrinterLib
{
    public class Printer
    {
        private static readonly string AssemblyName = "PrinterLib";
        public static PrintBase GenPrinter(PrinterSetting setting)
        {
            string className = "PrinterLib.Print" + setting.PrinterType;
            string[] args = setting.Setting.Split(';');

            return (PrintBase)Assembly.Load(AssemblyName).CreateInstance(className, true, BindingFlags.Default, null, args, null, null);
        }
    }
}
