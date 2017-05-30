using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB
{
    public class Pessoa
    {
        public int CdPessoa { get; set; }
        public string Nome { get; set; }
        public decimal Dinheiro { get; set; }
        public DateTime? DhNascimento { get; set; }
        public bool BlAtivo { get; set; }   
    }
}
