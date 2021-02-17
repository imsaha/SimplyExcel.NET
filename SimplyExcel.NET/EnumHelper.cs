using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET
{
    internal static class EnumHelper
    {
        public static IList<string> GetNames(Type type)
        {
            return type.GetFields(BindingFlags.Static | BindingFlags.Public).Select(fi => fi.Name).ToList();
        }
    }
}
