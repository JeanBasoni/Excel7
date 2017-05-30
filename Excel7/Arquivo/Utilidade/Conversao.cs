using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Arquivo.Utilidade
{
    public static class Conversao
    {
        public static object FormatarParametro(this object obj)
        {

            if (obj.GetType() == typeof(string) && !IsNumeric(obj.ToString()))
                obj.ToString();
            else if (obj.GetType() == typeof(int))
                obj = Convert.ToInt32(obj);
            else if (obj.GetType() == typeof(decimal))
                obj = Convert.ToDecimal(obj);
            else if (obj.GetType() == typeof(bool))
                obj = Convert.ToBoolean(obj);
            else if (obj.GetType() == typeof(float))
                obj = float.Parse(obj.ToString());
            else if (obj.GetType() == typeof(DateTime))
                obj = Convert.ToDateTime(obj);
            else if (obj.GetType() == typeof(System.DBNull))
                obj = null;


            return obj;
        }

        private static bool IsNumeric(string data)
        {
            bool isnumeric = false;
            char[] datachars = data.ToCharArray();

            foreach (var datachar in datachars)
                isnumeric = isnumeric ? char.IsDigit(datachar) : isnumeric;


            return isnumeric;
        }
    }
}
