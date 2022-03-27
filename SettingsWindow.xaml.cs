using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
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

			МетаВерсияTextBox.Text = SettingsManager.Settings.Meta.MetaVersion.Value;
			МетаИдентификаторTextBox.Text = SettingsManager.Settings.Meta.Id.Value;
			МетаГруппаTextBox.Text = SettingsManager.Settings.Meta.Group.Value;
			МетаДатаНачалаTextBox.Text = SettingsManager.Settings.Meta.BeginDate.Value;
			МетаДатаОкончанияTextBox.Text = SettingsManager.Settings.Meta.EndDate.Value;
			МетаАвторствоTextBox.Text = SettingsManager.Settings.Meta.Author.Value;
			МетаНомерВерсииTextBox.Text = SettingsManager.Settings.Meta.VersionNumber.Value;
			МетаРасположениеШапкиTextBox.Text = SettingsManager.Settings.Meta.HeaderPlacement.Value;
			МетаВерсияФорматаTextBox.Text = SettingsManager.Settings.Meta.MetaFormatVersion.Value;
			МетаМеткаНаименованиеTextBox.Text = SettingsManager.Settings.Meta.TitleMark.Value;

			ПрефиксыТаблицаTextBox.Text = SettingsManager.Settings.Tags.TablePrefix.Value;
			ПрефиксыСвободнаяЯчейкаTextBox.Text = SettingsManager.Settings.Tags.FreeCellPrefix.Value;
			ПрефиксыСтолбецTextBox.Text = SettingsManager.Settings.Tags.ColumnPrefix.Value;
			ПрефиксыСтрокаTextBox.Text = SettingsManager.Settings.Tags.RowPrefix.Value;
			КоличествоСловТегаTextBox.Text = SettingsManager.Settings.Tags.TagWordCount.Value;
			КоличествоСимволовТегаTextBox.Text = SettingsManager.Settings.Tags.TagCharCount.Value;

			МеткаДинамическаяTextBox.Text = SettingsManager.Settings.Markup.TableIsDynamicMark.Value;
			МеткаСтатическаяTextBox.Text = SettingsManager.Settings.Markup.TableIsStaticMark.Value;
			МеткаКодыСтрокИСтолбцовTextBox.Text = SettingsManager.Settings.Markup.RowAndColumnCodesMark.Value;
			МеткаКодыСтолбцовTextBox.Text = SettingsManager.Settings.Markup.ColumnCodesMark.Value;
			МеткаКодыСтрокTextBox.Text = SettingsManager.Settings.Markup.RowCodesMark.Value;
			МеткаНаименованиеТаблицыTextBox.Text = SettingsManager.Settings.Markup.TitleMark.Value;
			МеткаТегТаблицыTextBox.Text = SettingsManager.Settings.Markup.TagMark.Value;
			МеткаКодТаблицыTextBox.Text = SettingsManager.Settings.Markup.CodeMark.Value;

			ВерсияLabel.Content = Assembly.GetEntryAssembly().GetName().Version;
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
			SettingsManager.Settings.Meta.MetaVersion.Value = МетаВерсияTextBox.Text;
			SettingsManager.Settings.Meta.Id.Value = МетаИдентификаторTextBox.Text;
			SettingsManager.Settings.Meta.Group.Value = МетаГруппаTextBox.Text;
			SettingsManager.Settings.Meta.BeginDate.Value = МетаДатаНачалаTextBox.Text;
			SettingsManager.Settings.Meta.EndDate.Value = МетаДатаОкончанияTextBox.Text;
			SettingsManager.Settings.Meta.Author.Value = МетаАвторствоTextBox.Text;
			SettingsManager.Settings.Meta.VersionNumber.Value = МетаНомерВерсииTextBox.Text;
			SettingsManager.Settings.Meta.HeaderPlacement.Value = МетаРасположениеШапкиTextBox.Text;
			SettingsManager.Settings.Meta.MetaFormatVersion.Value = МетаВерсияФорматаTextBox.Text;
			SettingsManager.Settings.Meta.TitleMark.Value = МетаМеткаНаименованиеTextBox.Text;

			SettingsManager.Settings.Tags.TablePrefix.Value = ПрефиксыТаблицаTextBox.Text;
			SettingsManager.Settings.Tags.FreeCellPrefix.Value = ПрефиксыСвободнаяЯчейкаTextBox.Text;
			SettingsManager.Settings.Tags.ColumnPrefix.Value = ПрефиксыСтолбецTextBox.Text;
			SettingsManager.Settings.Tags.RowPrefix.Value = ПрефиксыСтрокаTextBox.Text;
			SettingsManager.Settings.Tags.TagWordCount.Value = КоличествоСловТегаTextBox.Text;
			SettingsManager.Settings.Tags.TagCharCount.Value = КоличествоСимволовТегаTextBox.Text;

			SettingsManager.Settings.Markup.TableIsDynamicMark.Value = МеткаДинамическаяTextBox.Text;
			SettingsManager.Settings.Markup.TableIsStaticMark.Value = МеткаСтатическаяTextBox.Text;
			SettingsManager.Settings.Markup.RowAndColumnCodesMark.Value = МеткаКодыСтрокИСтолбцовTextBox.Text;
			SettingsManager.Settings.Markup.ColumnCodesMark.Value = МеткаКодыСтолбцовTextBox.Text;
			SettingsManager.Settings.Markup.RowCodesMark.Value = МеткаКодыСтрокTextBox.Text;
			SettingsManager.Settings.Markup.TitleMark.Value = МеткаНаименованиеТаблицыTextBox.Text;
			SettingsManager.Settings.Markup.TagMark.Value = МеткаТегТаблицыTextBox.Text;
			SettingsManager.Settings.Markup.CodeMark.Value = МеткаКодТаблицыTextBox.Text;

			SettingsManager.SaveSettings();

			Close();
		}

		private void ResetSettingsButton_Click(object sender, RoutedEventArgs e)
		{
			МетаВерсияTextBox.Text = SettingsManager.Settings.Meta.MetaVersion.Default;
			МетаИдентификаторTextBox.Text = SettingsManager.Settings.Meta.Id.Default;
			МетаГруппаTextBox.Text = SettingsManager.Settings.Meta.Group.Default;
			МетаДатаНачалаTextBox.Text = SettingsManager.Settings.Meta.BeginDate.Default;
			МетаДатаОкончанияTextBox.Text = SettingsManager.Settings.Meta.EndDate.Default;
			МетаАвторствоTextBox.Text = SettingsManager.Settings.Meta.Author.Default;
			МетаНомерВерсииTextBox.Text = SettingsManager.Settings.Meta.VersionNumber.Default;
			МетаРасположениеШапкиTextBox.Text = SettingsManager.Settings.Meta.HeaderPlacement.Default;
			МетаВерсияФорматаTextBox.Text = SettingsManager.Settings.Meta.MetaFormatVersion.Default;
			МетаМеткаНаименованиеTextBox.Text = SettingsManager.Settings.Meta.TitleMark.Default;

			ПрефиксыТаблицаTextBox.Text = SettingsManager.Settings.Tags.TablePrefix.Default;
			ПрефиксыСвободнаяЯчейкаTextBox.Text = SettingsManager.Settings.Tags.FreeCellPrefix.Default;
			ПрефиксыСтолбецTextBox.Text = SettingsManager.Settings.Tags.ColumnPrefix.Default;
			ПрефиксыСтрокаTextBox.Text = SettingsManager.Settings.Tags.RowPrefix.Default;
			КоличествоСловТегаTextBox.Text = SettingsManager.Settings.Tags.TagWordCount.Default;
			КоличествоСимволовТегаTextBox.Text = SettingsManager.Settings.Tags.TagCharCount.Default;

			МеткаДинамическаяTextBox.Text = SettingsManager.Settings.Markup.TableIsDynamicMark.Default;
			МеткаСтатическаяTextBox.Text = SettingsManager.Settings.Markup.TableIsStaticMark.Default;
			МеткаКодыСтрокИСтолбцовTextBox.Text = SettingsManager.Settings.Markup.RowAndColumnCodesMark.Default;
			МеткаКодыСтолбцовTextBox.Text = SettingsManager.Settings.Markup.ColumnCodesMark.Default;
			МеткаКодыСтрокTextBox.Text = SettingsManager.Settings.Markup.RowCodesMark.Default;
			МеткаНаименованиеТаблицыTextBox.Text = SettingsManager.Settings.Markup.TitleMark.Default;
			МеткаТегТаблицыTextBox.Text = SettingsManager.Settings.Markup.TagMark.Default;
			МеткаКодТаблицыTextBox.Text = SettingsManager.Settings.Markup.CodeMark.Default;
		}
	}
}
