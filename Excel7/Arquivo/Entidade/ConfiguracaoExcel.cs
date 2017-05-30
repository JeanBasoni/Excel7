using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Arquivo.Enum;
using System.Data;
using System.Drawing;
using OfficeOpenXml;

namespace Arquivo.Entidade
{
    public class ConfiguracaoExcel
    {
        public Estilo Estilo { get; set; }
        public ConfiguracaoFile ConfiguracaoFile { get; set; }
        public TipoExcel TipoExcel { get; set; }
        public ExtensaoExcel ExtensaoExcel { get; set; }
        public Dados Dados { get; set; }
        public ExcelPackage ExcelPackage { get; set; }
    }

    public class Estilo
    {
        public Cabecalho Cabecalho { get; set; }
    }

    public class Cabecalho
    {
        public Color Background { get; set; }
        public Color CorFont { get; set; }
        public Font Fonte { get; set; }
        public bool isBold { get; set; }
        public bool isItalic { get; set; }
    }

    public class ConfiguracaoFile
    {
        public string NomeArquivo { get; set; }
        public string CaminhoArquivo { get; set; }
        public string NomeArquivoPath { get; set; }
        public string Extensao { get; set; }

        /// <summary>
        /// Retorna o caminho completo do arquivo com a extensão
        /// </summary>
        public string FormatarPath(object value, ExtensaoExcel tipoExt)
        {
            FormatarNomeFile(tipoExt);
            var pathFinal = CaminhoArquivo + value + Extensao;
            return pathFinal;
        }

        /// <summary>
        /// Retorna o nome do arquivo com a extensão
        /// </summary>
        private void FormatarNomeFile(ExtensaoExcel tipoExt)
        {
            //var extensao = new ExtensaoExcel();
            switch (tipoExt)
            {
                case ExtensaoExcel.Xls:
                    NomeArquivoPath = NomeArquivo + ".xls";
                    Extensao = ".xls";
                    break;
                case ExtensaoExcel.Xlsx:
                    NomeArquivoPath = NomeArquivo + ".xlsx";
                    Extensao = ".xlsx";
                    break;
                default:
                    break;
            }
        }

    }

    public class Dados
    {
        public TipoDados TipoDados { get; set; }
        public List<object> Data { get; set; }
        public DataSet Dataset { get; set; }
    }
}
