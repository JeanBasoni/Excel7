using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Arquivo.Entidade
{
    public class Propriedade
    {
        public string[] ListaGenerica(object objeto)
        {
            var val = new List<string>();
            foreach (var item in objeto.GetType().GetProperties())
                val.Add((item.GetValue(objeto) ?? "").ToString());

            return val.ToArray();
        }
    }
}
