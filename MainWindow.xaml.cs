using Microsoft.Office.Interop.Excel;
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
			Excel.Application excelApp = new Excel.Application();

			Workbook книгаExcel = excelApp.Workbooks.Open(путьККнигеExcel);

			try
			{
				ОписаниеФормы описаниеФормы = ОписаниеФормы.ПолучитьОписаниеФормыИзКнигиExcel(книгаExcel);

				СеарилизоватьВXMLИСохранить(описаниеФормы);
			}
			catch (System.Exception e)
			{
				_ = MessageBox.Show($"Возникла ошибка при получении метаструктуры из файла. Текст ошибки:\n\n{e.Message}");

				throw;
			}

			книгаExcel.Close();

			excelApp.Quit();
		}

		private void СеарилизоватьВXMLИСохранить(ОписаниеФормы описаниеФормы)
		{
			string путьКПапкеШаблона =
				ConfigManager.ПутьКПапкеСгенерированныхШаблонов +
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

		private void SettingsButton_Click(object sender, RoutedEventArgs e)
		{
			SettingsWindow settingsWindow = new SettingsWindow();
			settingsWindow.Show();

			Close();
		}
	}
}
