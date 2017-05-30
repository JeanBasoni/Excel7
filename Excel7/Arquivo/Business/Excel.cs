using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Arquivo.Enum;
using Arquivo.Interface;
using Arquivo.Entidade;

namespace Arquivo.Business
{
    public class Excel : IMetodo
    {
        public void GerarExcel(ConfiguracaoExcel configExcel)
        {
            switch (configExcel.TipoExcel)
            {
                case TipoExcel.Aba:
                    new Repositorio.Excel().GerarExcelDataSetUmaAba(configExcel);
                    break;
                case TipoExcel.NAbas:
                    new Repositorio.Excel().GerarExcelDataSetNAbas(configExcel);
                    break;
                default:
                    break;
            }
        }

    }
}
