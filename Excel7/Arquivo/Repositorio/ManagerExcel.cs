using Arquivo.Entidade;
using Arquivo.Utilidade;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Arquivo.Repositorio
{
    public class ManagerExcel
    {
        public ConfiguracaoExcel Config = new ConfiguracaoExcel();

        public void CarregarConfig(ConfiguracaoExcel c)
        {
            Config = c;
        }

        public ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName, int number)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[number];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet
            return ws;
        }

        /// <summary>
        /// Sets the workbook properties and adds a default sheet.
        /// </summary>
        /// <param name="p">The p.</param>
        /// <returns></returns>
        public void SetWorkbookProperties(ExcelPackage p)
        {
            //Here setting some document properties
            p.Workbook.Properties.Author = "Zeeshan Umar";
            p.Workbook.Properties.Title = "EPPlus Sample";


        }

        public void CreateHeader(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            int colIndex = 1;

            string limitFilter = "";

            foreach (DataColumn dc in dt.Columns) //Creating Headings
            {
                var cell = ws.Cells[rowIndex, colIndex];
                //cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Font.SetFromFont(Config.Estilo.Cabecalho.Fonte);
                cell.Style.Font.Color.SetColor(Config.Estilo.Cabecalho.CorFont);
                cell.Style.Font.Bold = Config.Estilo.Cabecalho.isBold;
                cell.Style.Font.Italic = Config.Estilo.Cabecalho.isItalic;

                //Setting the background color of header cells to Gray
                var fill = cell.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                //fill.BackgroundColor.SetColor(Color.FromArgb(27, 143, 28));
                fill.BackgroundColor.SetColor(Config.Estilo.Cabecalho.Background);


                //Setting Top/left,right/bottom borders.
                var border = cell.Style.Border;
                border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;


                //Setting Value in cell
                cell.Value = dc.ColumnName;
                //cell.AutoFitColumns();
                ws.Column(colIndex).AutoFit();

                colIndex++;

                //Determinando Filtro
                if ((colIndex - 1) == 1)
                {
                    limitFilter = cell.Address;
                }
                else if ((colIndex - 1) == dt.Columns.Count)
                {
                    limitFilter += ":" + cell.Address;
                }
            }

            //Adicionando filtro
            ws.Cells[limitFilter].AutoFilter = true;
        }

        public void CreateData(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            int colIndex = 0;
            foreach (DataRow dr in dt.Rows) // Adding Data into rows
            {
                colIndex = 1;
                rowIndex++;

                foreach (DataColumn dc in dt.Columns)
                {
                    var cell = ws.Cells[rowIndex, colIndex];

                    //Adicionando valor na célula
                    cell.Value = dr[dc.ColumnName].FormatarParametro();
                    
                    //Formatando a célula
                    cell.Style.Numberformat.Format = dr[dc.ColumnName].TipoFormatoCelula();

                    //Adicionando borda na célula
                    var border = cell.Style.Border;
                    border.Bottom.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;

                    //Tamanho da Celula automático
                    //cell.AutoFitColumns();
                    ws.Column(colIndex).AutoFit();
                    colIndex++;
                }
            }

        }

        public void CreateFooter(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            int colIndex = 0;
            foreach (DataColumn dc in dt.Columns) //Creating Formula in footers
            {
                colIndex++;
                var cell = ws.Cells[rowIndex, colIndex];

                //Setting Sum Formula
                cell.Formula = "Sum(" + ws.Cells[3, colIndex].Address + ":" + ws.Cells[rowIndex - 1, colIndex].Address + ")";

                //Setting Background fill color to Gray
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.Gray);
            }
        }

        public void CriarFooterBranco(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            int colIndex = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                colIndex++;
                var cell = ws.Cells[rowIndex, colIndex, rowIndex, colIndex];
                cell.Value = " ";
                cell.Merge = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.White);

                cell.Style.Border.Right.Color.SetColor(Color.White);
                cell.Style.Border.Left.Color.SetColor(Color.White);
                cell.Style.Border.Bottom.Color.SetColor(Color.White);
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            rowIndex++;
        }

        public void CreatGraphic(ExcelWorksheet ws, ref int rowIndex, DataTable dt)
        {
            ExcelChart chart = ws.Drawings.AddChart("chart", eChartType.ColumnClustered);

            ExcelChartSerie serie = chart.Series.Add(ws.Cells[5, 3, 7, 3], ws.Cells[5, 2, 7, 2]);
            serie.HeaderAddress = ws.Cells[4, 2];

            chart.SetPosition(150, 10);
            chart.SetSize(500, 300);
            chart.Title.Text = "Salário";
            chart.Title.Font.Color = Color.FromArgb(89, 89, 89);
            chart.Title.Font.Size = 15;
            chart.Title.Font.Bold = true;
            chart.Style = eChartStyle.Style15;
            chart.Legend.Border.LineStyle = eLineStyle.Solid;
            chart.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);

        }

        /// <summary>
        /// Adds the custom shape.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="shapeStyle">Shape style</param>
        /// <param name="text">Text for the shape</param>
        public void AddCustomShape(ExcelWorksheet ws, int colIndex, int rowIndex, eShapeStyle shapeStyle, string text)
        {
            ExcelShape shape = ws.Drawings.AddShape("cs" + rowIndex.ToString() + colIndex.ToString(), shapeStyle);
            shape.From.Column = colIndex;
            shape.From.Row = rowIndex;
            shape.From.ColumnOff = Pixel2MTU(5);
            shape.SetSize(100, 100);
            shape.RichText.Add(text);
        }

        /// <summary>
        /// Adds the image in excel sheet.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="filePath">The file path</param>
        public void AddImage(ExcelWorksheet ws, int columnIndex, int rowIndex, string filePath)
        {
            //How to Add a Image using EP Plus
            Bitmap image = new Bitmap(filePath);
            ExcelPicture picture = null;
            if (image != null)
            {
                picture = ws.Drawings.AddPicture("pic" + rowIndex.ToString() + columnIndex.ToString(), image);
                picture.From.Column = columnIndex;
                picture.From.Row = rowIndex;
                picture.From.ColumnOff = Pixel2MTU(2); //Two pixel space for better alignment
                picture.From.RowOff = Pixel2MTU(2);//Two pixel space for better alignment
                picture.SetSize(100, 100);
            }
        }

        /// <summary>
        /// Adds the comment in excel sheet.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="colIndex">Column Index</param>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="comments">Comment text</param>
        /// <param name="author">Author Name</param>
        public void AddComment(ExcelWorksheet ws, int colIndex, int rowIndex, string comment, string author)
        {
            //Adding a comment to a Cell
            var commentCell = ws.Cells[rowIndex, colIndex];
            commentCell.AddComment(comment, author);
        }

        /// <summary>
        /// Pixel2s the MTU.
        /// </summary>
        /// <param name="pixels">The pixels.</param>
        /// <returns></returns>
        public int Pixel2MTU(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }

        /// <summary>
        /// Creates the data table with some dummy data.
        /// </summary>
        /// <returns>DataTable</returns>
        public DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < 10; i++)
            {
                dt.Columns.Add(i.ToString());
            }

            for (int i = 10; i < 21; i++)
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
    }
}
