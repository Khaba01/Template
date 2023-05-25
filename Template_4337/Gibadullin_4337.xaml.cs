using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using System.Text.Json;
using Newtonsoft.Json.Serialization;
using System.Reflection;
using JsonProperty = Newtonsoft.Json.Serialization.JsonProperty;
using Newtonsoft.Json.Converters;
using Word = Microsoft.Office.Interop.Word;
using JsonSerializer = Newtonsoft.Json.JsonSerializer;

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
        public class CustomDateTimeConverter : IsoDateTimeConverter
        {
            public CustomDateTimeConverter()
            {
                base.DateTimeFormat = "dd.mm.yyyy";
            }
        }
        public class TimeSpanConverter : JsonConverter<TimeSpan>
        {
            public const string TimeSpanFormatString = @"d\.hh\:mm\:ss\:FFF";
            public override void WriteJson(JsonWriter writer, TimeSpan value, JsonSerializer serializer)
            {
                var timespanFormatted = $"{value.ToString(TimeSpanFormatString)}";
                writer.WriteValue(timespanFormatted);
            }
            public override TimeSpan ReadJson(JsonReader reader, Type objectType, TimeSpan existingValue, bool hasExistingValue, JsonSerializer serializer)
            {
                TimeSpan parsedTimeSpan;
                TimeSpan.TryParseExact((string)reader.Value, TimeSpanFormatString, null, out parsedTimeSpan);
                return parsedTimeSpan;
            }
        }
        public class ShouldDeserializeContractResolver : DefaultContractResolver
        {
            public static new readonly ShouldDeserializeContractResolver Instance = new ShouldDeserializeContractResolver();
            protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization)
            {
                JsonProperty property = base.CreateProperty(member, memberSerialization);
                MethodInfo shouldDeserializeMethodInfo = member.DeclaringType.GetMethod("ShouldDeserialize" + member.Name);
                if (shouldDeserializeMethodInfo != null)
                {
                    property.ShouldDeserialize = o => { return (bool)shouldDeserializeMethodInfo.Invoke(o, null); };
                }
                return property;
            }
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
                            //TimeCreate = TimeSpan.ParseExact(list[i, 3].ToString(), "t", CultureInfo.InvariantCulture),
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
                            //TimeCreate = TimeSpan.ParseExact(list[i, 3].ToString(), "t", CultureInfo.InvariantCulture),
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
                worksheet.Cells[3][startRowIndex + 1] = "Дата создания";
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

        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл JSON (Spisok.json)|*.json",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            using (StreamReader reader = new StreamReader(ofd.FileName))
            {
                var settings = new JsonSerializerSettings
                {
                    ContractResolver = ShouldDeserializeContractResolver.Instance
                };
                List<Zakazy> zakazies = JsonConvert.DeserializeObject<List<Zakazy>>(await reader.ReadToEndAsync(), settings);
                using (ZakazyEntities db = new ZakazyEntities())
                {
                    db.Zakazy.RemoveRange(db.Zakazy);
                    foreach (var c in zakazies)
                    {
                        db.Zakazy.Add(c);
                    }
                    db.SaveChanges();
                }
                MessageBox.Show("Объекты добавлены в бд");
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            List<Zakazy> zakazies;
            List<Statu> status;
            using (ZakazyEntities db = new ZakazyEntities())
            {
                zakazies = db.Zakazy.ToList();
                status = db.Statu.ToList();
                var categ = zakazies.GroupBy(s => s.Stat).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                foreach (var c in categ)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = status.Where(g => g.Stat == c.Key).FirstOrDefault().Stat;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table zakazTable = document.Tables.Add(tableRange, c.Count() + 1, 5);
                    zakazTable.Borders.InsideLineStyle = zakazTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    zakazTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = zakazTable.Cell(1, 1).Range;
                    cellRange.Text = "ID";
                    cellRange = zakazTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = zakazTable.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = zakazTable.Cell(1, 4).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = zakazTable.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    zakazTable.Rows[1].Range.Bold = 1;
                    zakazTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int i = 1;
                    foreach (var current in c)
                    {
                        cellRange = zakazTable.Cell(i + 1, 1).Range;
                        cellRange.Text = current.id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = zakazTable.Cell(i + 1, 2).Range;
                        cellRange.Text = current.Code;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = zakazTable.Cell(i + 1, 3).Range;
                        cellRange.Text = current.DateCreate.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = zakazTable.Cell(i + 1, 4).Range;
                        cellRange.Text = current.CodClient;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = zakazTable.Cell(i + 1, 5).Range;
                        cellRange.Text = current.Servic;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                    Word.Paragraph countStudentsParagraph = document.Paragraphs.Add();
                    Word.Range countStudentsRange = countStudentsParagraph.Range;
                    countStudentsRange.Text = $"Количество заказов с данным статусом - {c.Count()}";
                    countStudentsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    countStudentsRange.InsertParagraphAfter();
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;
            }
        }
    }
}
