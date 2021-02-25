using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;
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

				Обработать(путьКФайлу);
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

				СеарилизоватьВXMLИСохранить(описаниеФормы, out string путьКПапкеШаблона);

				if (путьКПапкеШаблона != "")
				{
					fileDropLabel.Content = $"Сконвертировано успешно. Путь к сгенерированному файлу:\n\n{путьКПапкеШаблона}";

					_ = Process.Start(fileName: путьКПапкеШаблона);
				}
			}
			catch (System.Exception e)
			{
				_ = MessageBox.Show($"Возникла ошибка при получении метаструктуры из файла. Текст ошибки:\n\n{e.Message}");

				throw;
			}

			книгаExcel.Close();

			excelApp.Quit();
		}

		private void СеарилизоватьВXMLИСохранить(ОписаниеФормы описаниеФормы, out string путьКПапкеШаблона)
		{
			путьКПапкеШаблона =
				МенеджерНастроек.Настройки.ПутьКПапкеСгенерированныхШаблонов.Value +
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

		private string OpenFileDialog()
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			
			if (openFileDialog.ShowDialog() == true)
			{
				return openFileDialog.FileName;
			}

			return "";
		}

		private void RadioButton_Checked(object sender, RoutedEventArgs e)
		{
			fileDropGrid.IsEnabled = true;
			DropRectangle.Stroke = Brushes.Black;

			if (sender == запросRadioButton)
			{
				МенеджерНастроек.Настройки.Мета.ЯвляетсяЗапросом.Value = true;
			}
			else if (sender == мониторингRadioButton)
			{
				МенеджерНастроек.Настройки.Мета.ЯвляетсяЗапросом.Value = false;
			}
		}

		private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
		{
			string путьКФайлу = OpenFileDialog();

			Обработать(путьКФайлу);
		}

		private void Обработать(string путьКФайлу)
		{
			if (путьКФайлу.EndsWith(".xls") || путьКФайлу.EndsWith(".xlsx") || путьКФайлу.EndsWith(".xlsm"))
			{
				fileDropLabel.Content = $"{путьКФайлу}";

				КонвертироватьКнигуВШаблон(путьКФайлу); 
			}
			else
			{
				_ = MessageBox.Show($"Ошибка в пути к файлу, указанном как: {путьКФайлу}");
			}
		}
	}
}
