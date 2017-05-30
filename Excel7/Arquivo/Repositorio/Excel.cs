using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Arquivo.Utilidade;
using OfficeOpenXml.Drawing.Chart;
using Arquivo.Entidade;
using Arquivo.Enum;

namespace Arquivo.Repositorio
{
    public class Excel : ManagerExcel
    {
        public ConfiguracaoExcel _config { get; set; }

        public Excel()
        {
            _config = new ConfiguracaoExcel();
        }

        public void GerarExcelDataSetNAbas(ConfiguracaoExcel configExcel)
        {
            try
            {
                _config = configExcel;

                //var t = new Propriedade();
                //t.ListaGenerica(Config.Dados.Data);

                CarregarConfig(_config);

                var p = new ExcelPackage();

                var numAba = 1;
                foreach (DataTable dt in configExcel.Dados.Dataset.Tables)
                {
                    //Definir as propriedades do workbook e adicionar uma folha predefinida no mesmo
                    //SetWorkbookProperties(p);
                    //Cria Aba
                    ExcelWorksheet ws = CreateSheet(p, "Aba " + numAba, numAba);

                    int rowIndex = 1;

                    CreateHeader(ws, ref rowIndex, dt);
                    CreateData(ws, ref rowIndex, dt);
                    //CreateFooter(ws, ref rowIndex, dt);

                    //AddComment(ws, 5, 10, "Analista de Sistemas\n", "\n\n\n Jean Basoni");

                    //string path = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Application.StartupPath)), "Zeeshan Umar.jpg");
                    //string path = Path.Combine(@"C:\Projetos", "59780.jpg");
                    //AddImage(ws, 10, 0, path);

                    //AddCustomShape(ws, 10, 7, eShapeStyle.Ellipse, "Text inside Ellipse.");

                    //Generate A File with Random name
                    //Byte[] bin = p.GetAsByteArray();
                    //string file = Guid.NewGuid().ToString() + ".xlsx";
                    //File.WriteAllBytes(file, bin);
                    //p.Workbook.Worksheets.Add("Tabela " + numAba.ToString(), ws);
                    numAba++;
                }

                //Escreva de volta para o cliente e criando arquivo excel no disco físico
                var p_strPath = ValidarArquivo(p);

                //Download do arquivo
                DownloadArquivo(p_strPath);
            }
            catch (Exception ex)
            {
                new Error().GerarLog(ex);
            }
        }

        public void GerarExcelDataSetUmaAba(ConfiguracaoExcel configExcel)
        {
            try
            {

                _config = configExcel;

                CarregarConfig(_config);

                var p = new ExcelPackage();

                var numAba = 1;
                //Cria abas
                ExcelWorksheet ws = CreateSheet(p, "Aba " + numAba, numAba);
                int rowIndex = 1;

                //set the workbook properties and add a default sheet in it
                //SetWorkbookProperties(p);

                foreach (DataTable dt in configExcel.Dados.Dataset.Tables)
                {
                    CreateHeader(ws, ref rowIndex, dt);
                    CreateData(ws, ref rowIndex, dt);

                    rowIndex += 2;
                    //CriarFooterBranco(ws, ref rowIndex, dt);
                    //CreateFooter(ws, ref rowIndex, dt);

                    //AddComment(ws, 5, 10, "Analista de Sistemas\n", "\n\n\n Jean Basoni");

                    //string path = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Application.StartupPath)), "Zeeshan Umar.jpg");
                    //string path = Path.Combine(@"C:\Projetos", "59780.jpg");
                    //AddImage(ws, 10, 0, path);



                    //Generate A File with Random name
                    //Byte[] bin = p.GetAsByteArray();
                    //string file = Guid.NewGuid().ToString() + ".xlsx";
                    //File.WriteAllBytes(file, bin);
                    //p.Workbook.Worksheets.Add("Tabela " + numAba.ToString(), ws);
                    numAba++;
                }

                //Escreva de volta para o cliente e criando arquivo excel no disco físico
                var p_strPath = ValidarArquivo(p);

                //Download do arquivo
                DownloadArquivo(p_strPath);

                //Estas linhas irão abri-lo no Excel
                //ProcessStartInfo pi = new ProcessStartInfo(p_strPath);
                //Process.Start(pi);
            }
            catch (Exception ex)
            {
                new Error().GerarLog(ex);
            }

        }

        private string ValidarArquivo(ExcelPackage p)
        {
            var arquivoLocal = Guid.NewGuid();
            var p_strPath = _config.ConfiguracaoFile.FormatarPath(arquivoLocal, _config.ExtensaoExcel);
            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            CriarArquivo(p_strPath);

            File.WriteAllBytes(p_strPath, p.GetAsByteArray());

            return p_strPath;
        }

        private void CriarArquivo(string path)
        {
            FileStream objFileStrm = File.Create(path);
            objFileStrm.Close();
            objFileStrm.Dispose();
        }

        private void DownloadArquivo(string path)
        {
            HttpResponse response = HttpContext.Current.Response;

            // Primeiro vamos limpar o objeto response.object
            response.Clear();
            response.Charset = "";

            // Defina o tipo mime de resposta para excel
            response.ContentEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1");
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Disposition", "attachment;filename=" + _config.ConfiguracaoFile.NomeArquivoPath);
            response.WriteFile(path);
            response.Flush();

            //Deleta arquivo
            File.Delete(path);

            //response.End();
        }
    }
}