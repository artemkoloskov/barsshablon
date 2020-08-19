using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Configuration;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace БАРСШаблон
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : System.Windows.Window
	{
		public MainWindow()
		{
			InitializeComponent();

			DropRectangle.AllowDrop = true;
		}

		private void DropRectangle_PreviewDrop(object sender, DragEventArgs e)
		{
			object text = e.Data.GetData(DataFormats.FileDrop);

			if (sender is System.Windows.Shapes.Rectangle)
			{
				string путьКФайлу = string.Format("{0}", ((string[])text)[0]);

				DropHereTextBlock.Text = путьКФайлу;
				DropHereTextBlock.Width = DropRectangle.Width;
				DropHereTextBlock.Height = DropRectangle.Height;

				КонвертироватьКнигуВШаблон(путьКФайлу);
			}
		}

		private void DropRectangle_PreviewDragOver(object sender, DragEventArgs e)
		{
			e.Effects = DragDropEffects.Copy;
			e.Handled = true;
		}

		private void КонвертироватьКнигуВШаблон(string путьККнигеExcel)
		{
			ОписаниеФормы описаниеФормы = ПолучитьОписаниеФормыИзКнигиExcel(путьККнигеExcel);

			СеарилизоватьВXMLИСохранить(описаниеФормы);
		}

		private ОписаниеФормы ПолучитьОписаниеФормыИзКнигиExcel(string путьККнигеExcel)
		{
			Excel.Application excelApp = new Excel.Application();

			Workbook книгаExcel = excelApp.Workbooks.Open(путьККнигеExcel);

			Мета мета = new Мета(книгаExcel);

			List<Таблица> таблицы = ПолучитьТаблицыФормы(книгаExcel.Sheets);

			книгаExcel.Close();

			List<СвободнаяЯчейка> свободныеЯчейки = ПолучитьСвободныеЯчейкиФормы();

			ОписаниеФормы описаниеФормы = new ОписаниеФормы(мета, таблицы, свободныеЯчейки);

			return описаниеФормы;
		}

		private List<Таблица> ПолучитьТаблицыФормы(Sheets листыКниги)
		{
			List<Таблица> таблицы = new List<Таблица>();

			int n = 1;

			foreach (Worksheet листКниги in листыКниги)
			{
				Таблица таблица = new Таблица(листКниги, n);

				n++;

				if (таблица != null)
				{
					таблицы.Add(таблица);
				}
			}

			return таблицы;
		}

		private List<СвободнаяЯчейка> ПолучитьСвободныеЯчейкиФормы()
		{
			return new List<СвободнаяЯчейка>();
		}

		private void СеарилизоватьВXMLИСохранить(ОписаниеФормы описаниеФормы)
		{
			string путьКПапкеШаблона =
				ConfigurationManager.AppSettings.Get("ПутьКПапкеСгенерированныхШаблонов") +
				описаниеФормы.Мета.Идентификатор + "\\" +
				описаниеФормы.Мета.ДатаНачалаДействия.Substring(0, 10) + "-" +
				описаниеФормы.Мета.ДатаОкончанияДействия.Substring(0, 10);

			System.IO.Directory.CreateDirectory(путьКПапкеШаблона);

			string имяФайла = описаниеФормы.Мета.Идентификатор + ".xml";

			XmlSerializer xmlSerializer = new XmlSerializer(описаниеФормы.GetType());

			XDocument xDocument = new XDocument();

			using (XmlWriter xmlWriter = xDocument.CreateWriter())
			{
				xmlSerializer.Serialize(xmlWriter, описаниеФормы);
			}

			XElement mainXmlStream = xDocument.Root;

			mainXmlStream.Save(путьКПапкеШаблона + "\\" + имяФайла);
		}
	}
}
