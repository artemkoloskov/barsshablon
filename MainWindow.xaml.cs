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
				string filePath = string.Format("{0}", ((string[])text)[0]);

				ProcessFile(filePath);
			}
		}

		private void DropRectangle_PreviewDragOver(object sender, DragEventArgs e)
		{
			e.Effects = DragDropEffects.Copy;
			e.Handled = true;
		}

		private void ConvertWorkbookToTemplate(string workbookPath)
		{
			Excel.Application excelApp = new Excel.Application();

			Workbook workbook = excelApp.Workbooks.Open(workbookPath);

			try
			{
				FormDescription fromDescription = FormDescription.GetDescription(workbook);

				SerializeAndSaveTemplate(fromDescription, out string templatePath);

				if (templatePath != "")
				{
					fileDropLabel.Content = $"Successfully converted. Result file path:\n\n{templatePath}";

					_ = Process.Start(fileName: templatePath);
				}
			}
			catch (Exception e)
			{
				_ = MessageBox.Show($"Error getting metastructure:\n\n{e.Message}");

				throw;
			}

			workbook.Close();

			excelApp.Quit();
		}

		private void SerializeAndSaveTemplate(FormDescription formDescription, out string templatePath)
		{
			templatePath =
				SettingsManager.Settings.GeneratedTemplatesPath.Value +
				formDescription.Meta.Id + "\\" +
				formDescription.Meta.DateFrom.Substring(0, 10) + "-" +
				formDescription.Meta.DateTo.Substring(0, 10);

			_ = System.IO.Directory.CreateDirectory(templatePath);

			string fileName = formDescription.Meta.Id + ".xml";

			XmlSerializer xmlSerializer = new XmlSerializer(formDescription.GetType());

			XDocument xDocument = new XDocument();

			using (XmlWriter xmlWriter = xDocument.CreateWriter())
			{
				xmlSerializer.Serialize(xmlWriter, formDescription);
			}

			XElement mainXmlStream = xDocument.Root;

			mainXmlStream.Save(templatePath + "\\" + fileName);
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

			return openFileDialog.ShowDialog() == true ? openFileDialog.FileName : "";
		}

		private void RadioButton_Checked(object sender, RoutedEventArgs e)
		{
			fileDropGrid.IsEnabled = true;
			DropRectangle.Stroke = Brushes.Black;

			if (sender == запросRadioButton)
			{
				SettingsManager.Settings.Meta.IsARequest.Value = true;
			}
			else if (sender == мониторингRadioButton)
			{
				SettingsManager.Settings.Meta.IsARequest.Value = false;
			}
		}

		private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
		{
			string filePath = OpenFileDialog();

			ProcessFile(filePath);
		}

		private void ProcessFile(string filePath)
		{
			if (filePath.EndsWith(".xls") || filePath.EndsWith(".xlsx") || filePath.EndsWith(".xlsm"))
			{
				fileDropLabel.Content = $"{filePath}";

				ConvertWorkbookToTemplate(filePath);
			}
			else
			{
				_ = MessageBox.Show($"Ошибка в пути к файлу, указанном как: {filePath}");
			}
		}
	}
}
