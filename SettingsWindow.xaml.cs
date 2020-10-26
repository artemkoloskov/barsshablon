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
using System.Windows.Shapes;

namespace БАРСШаблон
{
	/// <summary>
	/// Логика взаимодействия для SettingsWindow.xaml
	/// </summary>
	public partial class SettingsWindow : Window
	{
		public SettingsWindow()
		{
			InitializeComponent();

			МетаВерсияTextBox.Text = ConfigManager.МетаВерсияМетаописания;
			МетаИдентификаторTextBox.Text = ConfigManager.МетаИдентификатор;
			МетаГруппаTextBox.Text = ConfigManager.МетаГруппа;
			МетаДатаНачалаTextBox.Text = ConfigManager.МетаДатаНачалаДействия;
			МетаДатаОкончанияTextBox.Text = ConfigManager.МетаДатаОкончанияДействия;
			МетаАвторствоTextBox.Text = ConfigManager.МетаАвторство;
			МетаНомерВерсииTextBox.Text = ConfigManager.МетаНомерВерсии;
			МетаРасположениеШапкиTextBox.Text = ConfigManager.МетаРасположениеШапки;
			МетаВерсияФорматаTextBox.Text = ConfigManager.МетаВерсияФорматаМетаструктуры;
			МетаМеткаНаименованиеTextBox.Text = ConfigManager.МетаМеткаНаименование;

			ПрефиксыТаблицаTextBox.Text = ConfigManager.ТаблицаПрефиксТега;
			ПрефиксыСвободнаяЯчейкаTextBox.Text = ConfigManager.СвободнаяЯчейкаТегПрефикс;
			ПрефиксыСтолбецTextBox.Text = ConfigManager.СтолбецТегПрефикс;
			ПрефиксыСтрокаTextBox.Text = ConfigManager.СтрокаТегПрефикс;
			КоличествоСловТегаTextBox.Text = ConfigManager.КоличествоСловВТеге.ToString();
			КоличествоСимволовТегаTextBox.Text = ConfigManager.КоличествоСимволовВТеге.ToString();

			МеткаДинамическаяTextBox.Text = ConfigManager.ТаблицаМеткаТипТаблицыДинамическая;
			МеткаСтатическаяTextBox.Text = ConfigManager.ТаблицаМеткаТипТаблицыСтатическая;
			МеткаКодыСтрокИСтолбцовTextBox.Text = ConfigManager.ТаблицаМеткаКодыСтрокИСтолбцов;
			МеткаКодыСтолбцовTextBox.Text = ConfigManager.ТаблицаМеткаКодыСтолбцов;
			МеткаКодыСтрокTextBox.Text = ConfigManager.ТаблицаМеткаКодыСтрок;
			МеткаНаименованиеТаблицыTextBox.Text = ConfigManager.ТаблицаМеткаНаименование;
			МеткаТегТаблицыTextBox.Text = ConfigManager.ТаблицаМеткаТег;
			МеткаКодТаблицыTextBox.Text = ConfigManager.ТаблицаМеткаКод;

		}

		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			MainWindow mainWindow = new MainWindow();
			mainWindow.Show();
		}

		private void CancelButton_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}

		private void SaveButton_Click(object sender, RoutedEventArgs e)
		{
			ConfigManager.МетаВерсияМетаописания = МетаВерсияTextBox.Text;
			ConfigManager.МетаИдентификатор = МетаИдентификаторTextBox.Text;
			ConfigManager.МетаГруппа = МетаГруппаTextBox.Text;
			ConfigManager.МетаДатаНачалаДействия = МетаДатаНачалаTextBox.Text;
			ConfigManager.МетаДатаОкончанияДействия = МетаДатаОкончанияTextBox.Text;
			ConfigManager.МетаАвторство = МетаАвторствоTextBox.Text;
			ConfigManager.МетаНомерВерсии = МетаНомерВерсииTextBox.Text;
			ConfigManager.МетаРасположениеШапки = МетаРасположениеШапкиTextBox.Text;
			ConfigManager.МетаВерсияФорматаМетаструктуры = МетаВерсияФорматаTextBox.Text;
			ConfigManager.МетаМеткаНаименование = МетаМеткаНаименованиеTextBox.Text;

			ConfigManager.ТаблицаПрефиксТега = ПрефиксыТаблицаTextBox.Text;
			ConfigManager.СвободнаяЯчейкаТегПрефикс = ПрефиксыСвободнаяЯчейкаTextBox.Text;
			ConfigManager.СтолбецТегПрефикс = ПрефиксыСтолбецTextBox.Text;
			ConfigManager.СтрокаТегПрефикс = ПрефиксыСтрокаTextBox.Text;
			ConfigManager.КоличествоСловВТеге = int.Parse(КоличествоСловТегаTextBox.Text);
			ConfigManager.КоличествоСимволовВТеге = int.Parse(КоличествоСимволовТегаTextBox.Text);

			ConfigManager.ТаблицаМеткаТипТаблицыДинамическая = МеткаДинамическаяTextBox.Text;
			ConfigManager.ТаблицаМеткаТипТаблицыСтатическая = МеткаСтатическаяTextBox.Text;
			ConfigManager.ТаблицаМеткаКодыСтрокИСтолбцов = МеткаКодыСтрокИСтолбцовTextBox.Text;
			ConfigManager.ТаблицаМеткаКодыСтолбцов = МеткаКодыСтолбцовTextBox.Text;
			ConfigManager.ТаблицаМеткаКодыСтрок = МеткаКодыСтрокTextBox.Text;
			ConfigManager.ТаблицаМеткаНаименование = МеткаНаименованиеТаблицыTextBox.Text;
			ConfigManager.ТаблицаМеткаТег = МеткаТегТаблицыTextBox.Text;
			ConfigManager.ТаблицаМеткаКод = МеткаКодТаблицыTextBox.Text;

			ConfigManager.SaveConfig();

			Close();
		}
	}
}
