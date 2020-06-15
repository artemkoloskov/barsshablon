using System.Windows;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using БАРСШаблон.DataTypes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Collections.ObjectModel;
using System;
using System.Configuration;

namespace БАРСШаблон
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            DropRectangle.AllowDrop = true;
        }

        private void DropRectangle_PreviewDrop(object sender, DragEventArgs e)
        {
            object text = e.Data.GetData(DataFormats.FileDrop);
            Rectangle dropRectangle = sender as Rectangle;

            if (dropRectangle != null)
            {
                string filePath = string.Format("{0}", ((string[])text)[0]);

                DropHereTextBlock.Text = filePath;
                DropHereTextBlock.Width = DropRectangle.Width;
                DropHereTextBlock.Height = DropRectangle.Height;

                StartConverting(filePath);
            }
        }

        private void DropRectangle_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void StartConverting(string excelWorkbookFilePath)
        {
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook templateWorkbook = excelApp.Workbooks.Open(excelWorkbookFilePath);

            Excel.Worksheet worksheet = excelApp.ActiveSheet;

            Excel.Range range = worksheet.UsedRange;

            ОписаниеФормы описаниеФормы = new ОписаниеФормы();
            описаниеФормы.Мета = new Мета(); 
            описаниеФормы.Мета.ВерсияМетаописания = range.Rows.Count + "";
            описаниеФормы.Справочники = new Справочник[] { };
            описаниеФормы.Структура = new Структура();
            описаниеФормы.Структура.СвободнаяЯчейка = 
                new СвободнаяЯчейка[] {
                    new СвободнаяЯчейка("Учреждение", typeof(Учреждение).ToString().Split('.')[2]),
                    new СвободнаяЯчейка("Должность", typeof(Строковый).ToString().Split('.')[2]),
                    new СвободнаяЯчейка("Ответственный", typeof(Строковый).ToString().Split('.')[2]),
                    new СвободнаяЯчейка("Телефон", typeof(Строковый).ToString().Split('.')[2]),
                };

            Столбец столбец1 = new Столбец("1", typeof(Целочисленный).ToString().Split('.')[2]);
            Столбец столбец2 = new Столбец("2", typeof(Финансовый).ToString().Split('.')[2]);

            Строка строка1 = new Строка() { Идентификатор = "001", Код = "001", НаименованиеЭлемента = "Охуеть", Тег = "Охт" };
            Строка строка2 = new Строка() { Идентификатор = "002", Код = "002", НаименованиеЭлемента = "Заебись", Тег = "Збс" };

            Таблица таблица1 = new Таблица()
            {
                Идентификатор = "Таблица1",
                Код = "Тбл1",
                Наименование = "Крутая ваще таблица",
                РучноеДобавлениеСтрок = false,
                Тег = "КртВщТабла",
                Столбцы = new Столбец[] { столбец1, столбец2 },
                Строки = new Строка[] { строка1, строка2},
                СвободныеЯчейки = new СвободнаяЯчейка[] { new СвободнаяЯчейка("Суки", typeof(Целочисленный).ToString().Split('.')[2]) },
            };
            
            описаниеФормы.Структура.Таблицы = new Таблица[]
            {
                таблица1,
            };
            
            string outputFilePath = 
                ConfigurationManager.AppSettings.Get("ПутьКПапкеСгенерированныхШаблонов") + 
                описаниеФормы.Мета.Идентификатор + "\\" + 
                описаниеФормы.Мета.ДатаНачалаДействия.Substring(0, 10) + "-" + 
                описаниеФормы.Мета.ДатаОкончанияДействия.Substring(0, 10);

            string outputFileName = описаниеФормы.Мета.Идентификатор + ".xml";

            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(описаниеФормы.GetType());

                System.IO.Directory.CreateDirectory(outputFilePath);

                XDocument xDocument = new XDocument();

                using (XmlWriter xmlWriter = xDocument.CreateWriter())
                {
                    xmlSerializer.Serialize(xmlWriter, описаниеФормы);
                }

                XElement mainXmlStream = xDocument.Root;

                mainXmlStream.Save(outputFilePath + "\\" + outputFileName);
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }           
        }
    }
}
