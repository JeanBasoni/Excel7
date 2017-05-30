using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using Arquivo.Business;

namespace Excel7
{
    public partial class Pagina1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var ds = new DB.Conexao().Conectar();

            Excel ex = new Excel();
            var dt = CreateDataTable();

            Arquivo.Entidade.ConfiguracaoExcel config = new Arquivo.Entidade.ConfiguracaoExcel();
            config.ConfiguracaoFile = new Arquivo.Entidade.ConfiguracaoFile();
            config.ConfiguracaoFile.CaminhoArquivo = @"E:\ProjetosTeste\";
            config.ConfiguracaoFile.NomeArquivo = "ArquivoTeste";

            config.Dados = new Arquivo.Entidade.Dados();
            config.Dados.Dataset = ds;

            config.ExtensaoExcel = Arquivo.Enum.ExtensaoExcel.Xlsx;

            config.TipoExcel = Arquivo.Enum.TipoExcel.Aba;

            config.Estilo = new Arquivo.Entidade.Estilo();
            config.Estilo.Cabecalho = new Arquivo.Entidade.Cabecalho();
            config.Estilo.Cabecalho.Background = Color.FromArgb(25, 25, 112);
            config.Estilo.Cabecalho.CorFont = Color.White;
            config.Estilo.Cabecalho.Fonte = new Font("Calibri", 11);
            config.Estilo.Cabecalho.isBold = true;

            ex.GerarExcel(config);

        }


        /// <summary>
        /// Creates the data table with some dummy data.
        /// </summary>
        /// <returns>DataTable</returns>
        private static DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < 10; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            for (int i = 0; i < 10; i++)
            {
                DataRow dr = dt.NewRow();
                foreach (DataColumn dc in dt.Columns)
                {
                    dr[dc.ToString()] = i;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        /// <summary>
        /// Creates the data table with some dummy data.
        /// </summary>
        /// <returns>DataTable</returns>
        private static DataSet CreateDataSet()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            for (int i = 1; i <= 10; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            for (int j = 1; j <= 10; j++)
            {
                DataRow dr = dt.NewRow();
                foreach (DataColumn dc in dt.Columns)
                {
                    dr[dc] = Convert.ToDecimal(100.25);
                }

                dt.Rows.Add(dr);
            }
            ds.Tables.Add(dt);

            //DataTable dt2 = new DataTable();
            //for (int i = 11; i <= 20; i++)
            //{
            //    dt2.Columns.Add(i.ToString());
            //}

            //for (int i = 11; i <= 20; i++)
            //{
            //    DataRow dr = dt2.NewRow();
            //    foreach (DataColumn dc in dt2.Columns)
            //    {
            //        dr[dc] = i;
            //    }

            //    dt2.Rows.Add(dr);
            //}
            //ds.Tables.Add(dt2);


            return ds;
        }
    }
}