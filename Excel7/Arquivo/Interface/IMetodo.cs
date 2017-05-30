using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Arquivo.Enum;
using Arquivo.Entidade;

namespace Arquivo.Interface
{
    public interface IMetodo
    {
        void GerarExcel(ConfiguracaoExcel configExcel);
    }
}
