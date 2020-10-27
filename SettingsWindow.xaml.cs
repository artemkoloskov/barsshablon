using System;
using System.Collections.Generic;
using System.Configuration;
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
		private Настройки настройки;

		private Configuration configApp = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

		public SettingsWindow()
		{
			InitializeComponent();

			настройки = (Настройки)configApp.GetSection("барсШаблонНастройки");

			МетаВерсияTextBox.Text = настройки.Мета.ВерсияМетаописания.Value;
			МетаИдентификаторTextBox.Text = настройки.Мета.Идентификатор.Value;
			МетаГруппаTextBox.Text = настройки.Мета.Группа.Value;
			МетаДатаНачалаTextBox.Text = настройки.Мета.ДатаНачалаДействия.Value;
			МетаДатаОкончанияTextBox.Text = настройки.Мета.ДатаОкончанияДействия.Value;
			МетаАвторствоTextBox.Text = настройки.Мета.Авторство.Value;
			МетаНомерВерсииTextBox.Text = настройки.Мета.НомерВерсии.Value;
			МетаРасположениеШапкиTextBox.Text = настройки.Мета.РасположениеШапки.Value;
			МетаВерсияФорматаTextBox.Text = настройки.Мета.ВерсияФорматаМетаструктуры.Value;
			МетаМеткаНаименованиеTextBox.Text = настройки.Мета.МеткаНаименование.Value;

			ПрефиксыТаблицаTextBox.Text = настройки.Теги.ПрефиксТаблицы.Value;
			ПрефиксыСвободнаяЯчейкаTextBox.Text = настройки.Теги.ПрефиксСвободнойЯчейки.Value;
			ПрефиксыСтолбецTextBox.Text = настройки.Теги.ПрефиксСтолбца.Value;
			ПрефиксыСтрокаTextBox.Text = настройки.Теги.ПрефиксСтроки.Value;
			КоличествоСловТегаTextBox.Text = настройки.Теги.КоличествоСловВТеге.Value;
			КоличествоСимволовТегаTextBox.Text = настройки.Теги.КоличествоСимволовВТеге.Value;

			МеткаДинамическаяTextBox.Text = настройки.Разметка.МеткаТипТаблицыДинамическая.Value;
			МеткаСтатическаяTextBox.Text = настройки.Разметка.МеткаТипТаблицыСтатическая.Value;
			МеткаКодыСтрокИСтолбцовTextBox.Text = настройки.Разметка.МеткаКодыСтрокИСтолбцов.Value;
			МеткаКодыСтолбцовTextBox.Text = настройки.Разметка.МеткаКодыСтолбцов.Value;
			МеткаКодыСтрокTextBox.Text = настройки.Разметка.МеткаКодыСтрок.Value;
			МеткаНаименованиеТаблицыTextBox.Text = настройки.Разметка.МеткаНаименование.Value;
			МеткаТегТаблицыTextBox.Text = настройки.Разметка.МеткаТег.Value;
			МеткаКодТаблицыTextBox.Text = настройки.Разметка.МеткаКод.Value;

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
			настройки.Мета.ВерсияМетаописания.Value = МетаВерсияTextBox.Text;
			настройки.Мета.Идентификатор.Value = МетаИдентификаторTextBox.Text;
			настройки.Мета.Группа.Value = МетаГруппаTextBox.Text;
			настройки.Мета.ДатаНачалаДействия.Value = МетаДатаНачалаTextBox.Text;
			настройки.Мета.ДатаОкончанияДействия.Value = МетаДатаОкончанияTextBox.Text;
			настройки.Мета.Авторство.Value = МетаАвторствоTextBox.Text;
			настройки.Мета.НомерВерсии.Value = МетаНомерВерсииTextBox.Text;
			настройки.Мета.РасположениеШапки.Value = МетаРасположениеШапкиTextBox.Text;
			настройки.Мета.ВерсияФорматаМетаструктуры.Value = МетаВерсияФорматаTextBox.Text;
			настройки.Мета.МеткаНаименование.Value = МетаМеткаНаименованиеTextBox.Text;

			настройки.Теги.ПрефиксТаблицы.Value = ПрефиксыТаблицаTextBox.Text;
			настройки.Теги.ПрефиксСвободнойЯчейки.Value = ПрефиксыСвободнаяЯчейкаTextBox.Text;
			настройки.Теги.ПрефиксСтолбца.Value = ПрефиксыСтолбецTextBox.Text;
			настройки.Теги.ПрефиксСтроки.Value = ПрефиксыСтрокаTextBox.Text;
			настройки.Теги.КоличествоСловВТеге.Value = КоличествоСловТегаTextBox.Text;
			настройки.Теги.КоличествоСимволовВТеге.Value = КоличествоСимволовТегаTextBox.Text;

			настройки.Разметка.МеткаТипТаблицыДинамическая.Value = МеткаДинамическаяTextBox.Text;
			настройки.Разметка.МеткаТипТаблицыСтатическая.Value = МеткаСтатическаяTextBox.Text;
			настройки.Разметка.МеткаКодыСтрокИСтолбцов.Value = МеткаКодыСтрокИСтолбцовTextBox.Text;
			настройки.Разметка.МеткаКодыСтолбцов.Value = МеткаКодыСтолбцовTextBox.Text;
			настройки.Разметка.МеткаКодыСтрок.Value = МеткаКодыСтрокTextBox.Text;
			настройки.Разметка.МеткаНаименование.Value = МеткаНаименованиеТаблицыTextBox.Text;
			настройки.Разметка.МеткаТег.Value = МеткаТегТаблицыTextBox.Text;
			настройки.Разметка.МеткаКод.Value = МеткаКодТаблицыTextBox.Text;

			configApp.Save(ConfigurationSaveMode.Full);

			Close();
		}

		private void ResetSettingsButton_Click(object sender, RoutedEventArgs e)
		{
			МетаВерсияTextBox.Text = настройки.Мета.ВерсияМетаописания.Default;
			МетаИдентификаторTextBox.Text = настройки.Мета.Идентификатор.Default;
			МетаГруппаTextBox.Text = настройки.Мета.Группа.Default;
			МетаДатаНачалаTextBox.Text = настройки.Мета.ДатаНачалаДействия.Default;
			МетаДатаОкончанияTextBox.Text = настройки.Мета.ДатаОкончанияДействия.Default;
			МетаАвторствоTextBox.Text = настройки.Мета.Авторство.Default;
			МетаНомерВерсииTextBox.Text = настройки.Мета.НомерВерсии.Default;
			МетаРасположениеШапкиTextBox.Text = настройки.Мета.РасположениеШапки.Default;
			МетаВерсияФорматаTextBox.Text = настройки.Мета.ВерсияФорматаМетаструктуры.Default;
			МетаМеткаНаименованиеTextBox.Text = настройки.Мета.МеткаНаименование.Default;

			ПрефиксыТаблицаTextBox.Text = настройки.Теги.ПрефиксТаблицы.Default;
			ПрефиксыСвободнаяЯчейкаTextBox.Text = настройки.Теги.ПрефиксСвободнойЯчейки.Default;
			ПрефиксыСтолбецTextBox.Text = настройки.Теги.ПрефиксСтолбца.Default;
			ПрефиксыСтрокаTextBox.Text = настройки.Теги.ПрефиксСтроки.Default;
			КоличествоСловТегаTextBox.Text = настройки.Теги.КоличествоСловВТеге.Default;
			КоличествоСимволовТегаTextBox.Text = настройки.Теги.КоличествоСимволовВТеге.Default;

			МеткаДинамическаяTextBox.Text = настройки.Разметка.МеткаТипТаблицыДинамическая.Default;
			МеткаСтатическаяTextBox.Text = настройки.Разметка.МеткаТипТаблицыСтатическая.Default;
			МеткаКодыСтрокИСтолбцовTextBox.Text = настройки.Разметка.МеткаКодыСтрокИСтолбцов.Default;
			МеткаКодыСтолбцовTextBox.Text = настройки.Разметка.МеткаКодыСтолбцов.Default;
			МеткаКодыСтрокTextBox.Text = настройки.Разметка.МеткаКодыСтрок.Default;
			МеткаНаименованиеТаблицыTextBox.Text = настройки.Разметка.МеткаНаименование.Default;
			МеткаТегТаблицыTextBox.Text = настройки.Разметка.МеткаТег.Default;
			МеткаКодТаблицыTextBox.Text = настройки.Разметка.МеткаКод.Default;
		}
	}
}
