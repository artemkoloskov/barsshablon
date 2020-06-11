using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using БАРСШаблон.DataTypes;
using БАРСШаблон.Properties;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using БАРСШаблон.Structure;
using System.Collections.ObjectModel;

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

            Мета мета = new Мета();
            мета.ВерсияМетаописания = range.Rows.Count + "";

            Collection<СвободнаяЯчейка> свободныеЯчейки = new Collection<СвободнаяЯчейка>();
            свободныеЯчейки.Add(new СвободнаяЯчейка("Учреждение", new Учреждение()));
            свободныеЯчейки.Add(new СвободнаяЯчейка("Должность", new Строковый()));
            свободныеЯчейки.Add(new СвободнаяЯчейка("Ответственный", new Строковый()));
            свободныеЯчейки.Add(new СвободнаяЯчейка("Телефон", new Строковый()));

            string outputFilePath = "C:\\БАРС\\ШАблоныФормы\\" + мета.Идентификатор + "\\" + мета.ДатаНачалаДействия.Substring(0, 10) + "-" + мета.ДатаОкончанияДействия.Substring(0, 10);

            string outputFileName = мета.Идентификатор + ".xml";

            XmlSerializer xmlSerializer = new XmlSerializer(мета.GetType());
            
            System.IO.Directory.CreateDirectory(outputFilePath);

            XDocument xDocument = new XDocument();

            using (XmlWriter xmlWriter = xDocument.CreateWriter())
            {
                xmlSerializer.Serialize(xmlWriter, мета);
            }

            XElement mainXmlStream = xDocument.Root;
            mainXmlStream =
                new XElement("Описание",
                    mainXmlStream,
                    new XElement("Структура"));


            mainXmlStream.Save(outputFilePath + "\\" + outputFileName);            
        }
    }
}
