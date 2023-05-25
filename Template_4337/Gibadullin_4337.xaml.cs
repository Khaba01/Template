using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Gibadullin_4337.xaml
    /// </summary>
    public partial class Gibadullin_4337 : Window
    {
        public Gibadullin_4337()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            using (ZakazyEntities db = new ZakazyEntities())
            {
                for (int i = 1; i < 51; i++)
                {
                    if (list[i, 7] == "")
                    {
                        db.Zakazy.Add(new Zakazy()
                        {
                            id = int.Parse(list[i, 0]),
                            Code = list[i, 1],
                            DateCreate = DateTime.ParseExact(list[i, 2].ToString(), "dd.mm.yyyy", System.Globalization.CultureInfo.InvariantCulture),
                            TimeCreate = TimeSpan.ParseExact(list[i, 3].ToString(), "t", CultureInfo.InvariantCulture),
                            CodClient = list[i, 4],
                            Servic = list[i, 5],
                            Stat = list[i, 6],
                            TimeProcat = list[i, 8]
                        });
                    }
                    else
                    {
                        db.Zakazy.Add(new Zakazy()
                        {
                            id = int.Parse(list[i, 0]),
                            Code = list[i, 1],
                            DateCreate = DateTime.ParseExact(list[i, 2].ToString(), "dd.mm.yyyy", System.Globalization.CultureInfo.InvariantCulture),
                            TimeCreate = TimeSpan.ParseExact(list[i, 3].ToString(), "t", CultureInfo.InvariantCulture),
                            CodClient = list[i, 4],
                            Servic = list[i, 5],
                            Stat = list[i, 6],
                            DateClose = DateTime.ParseExact(list[i, 7].ToString(), "dd.mm.yyyy", System.Globalization.CultureInfo.InvariantCulture),
                            TimeProcat = list[i, 8]
                        });
                    }
                }
                db.SaveChanges();
                MessageBox.Show("Данные добавлены");
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<Zakazy> zakazies;
            List<Statu> status;
            using (ZakazyEntities db = new ZakazyEntities())
            {
                zakazies = db.Zakazy.ToList();
                status = db.Statu.ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = status.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < status.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = status[i].Stat;
                worksheet.Cells[1][startRowIndex + 1] = "ID";
                worksheet.Cells[2][startRowIndex + 1] = "Код заказа";
                worksheet.Cells[3][startRowIndex + 1 ] = "Дата создания";
                worksheet.Cells[4][startRowIndex + 1] = "Код клиента";
                worksheet.Cells[5][startRowIndex + 1] = "Услуги";
                startRowIndex++;
                var categ = zakazies.GroupBy(s => s.Stat).ToList();
                foreach (var c in categ)
                {
                    if (c.Key == status[i].Stat)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = status[i].Stat;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (Zakazy c1 in zakazies)
                        {
                            if (c1.Stat == c.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = c1.id;
                                worksheet.Cells[2][startRowIndex] = c1.Code;
                                worksheet.Cells[3][startRowIndex] = c1.DateCreate;
                                worksheet.Cells[4][startRowIndex] = c1.CodClient;
                                worksheet.Cells[5][startRowIndex] = c1.Servic;
                                startRowIndex++;
                            }
                        }
                        worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                        worksheet.Cells[1][startRowIndex].Font.Bold = true;
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
