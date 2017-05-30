using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Arquivo.Enum;

namespace Arquivo.Utilidade
{
    public static class TipoDados
    {
        public static string TipoFormatoCelula(this object obj)
        {

            if (obj.GetType() == typeof(DateTime))
                obj = "dd/mm/yyyy hh:mm:ss";
            else if (obj.GetType() == typeof(decimal))
                obj = "#,##0.00;[Red]-#,##0.00";

            return obj.ToString();
        }

    }
}
