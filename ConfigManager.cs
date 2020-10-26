using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public static class ConfigManager
	{
		public static string МетаВерсияМетаописания
		{
			get => ConfigurationManager.AppSettings["МетаВерсияМетаописания"];

			set => ConfigurationManager.AppSettings["МетаВерсияМетаописания"] = value;
		}

		public static string МетаИдентификатор
		{
			get => ConfigurationManager.AppSettings["МетаИдентификатор"];

			set => ConfigurationManager.AppSettings["МетаИдентификатор"] = value;
		}
		public static string МетаГруппа 
		{
			get => ConfigurationManager.AppSettings["МетаГруппа"]; 
			
			set => ConfigurationManager.AppSettings["МетаГруппа"] = value;
		}
		public static string МетаДатаНачалаДействия 
		{
			get => ConfigurationManager.AppSettings["МетаДатаНачалаДействия"];
			
			set => ConfigurationManager.AppSettings["МетаДатаНачалаДействия"] = value;
		}
		public static string МетаДатаОкончанияДействия 
		{
			get => ConfigurationManager.AppSettings["МетаДатаОкончанияДействия"];
			
			set => ConfigurationManager.AppSettings["МетаДатаОкончанияДействия"] = value;
		}
		public static string МетаАвторство 
		{
			get => ConfigurationManager.AppSettings["МетаАвторство"];
			
			set => ConfigurationManager.AppSettings["МетаАвторство"] = value;
		}
		public static string МетаНомерВерсии 
		{
			get => ConfigurationManager.AppSettings["МетаНомерВерсии"];
			
			set => ConfigurationManager.AppSettings["МетаНомерВерсии"] = value;
		}
		public static string МетаРасположениеШапки 
		{
			get => ConfigurationManager.AppSettings["МетаРасположениеШапки"];
			
			set => ConfigurationManager.AppSettings["МетаРасположениеШапки"] = value;
		}
		public static string МетаВерсияФорматаМетаструктуры 
		{
			get => ConfigurationManager.AppSettings["МетаВерсияФорматаМетаструктуры"]; 

			set => ConfigurationManager.AppSettings["МетаВерсияФорматаМетаструктуры"] = value;
		}

		public static double МетаВесДлиныПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесДлиныПотенциальногоНаименования"]);
		public static double МетаВесНомераСтрокиПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесНомераСтрокиПотенциальногоНаименования"]);
		public static double МетаВесНомераСтолбцаПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесНомераСтолбцаПотенциальногоНаименования"]);
		public static double МетаВесКоличестваЯчеекВОбъединеннойЯчейкеПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесКоличестваЯчеекВОбъединеннойЯчейкеПотенциальногоНаименования"]);
		public static double МетаВесГраницыВнизуПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесГраницыВнизуПотенциальногоНаименования"]);
		public static double МетаВесГраницыВверхуПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесГраницыВверхуПотенциальногоНаименования"]);
		public static double МетаВесГраницыСлеваПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесГраницыСлеваПотенциальногоНаименования"]);
		public static double МетаВесГраницыСправаПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесГраницыСправаПотенциальногоНаименования"]);
		public static double МетаВесВыравниванияПоСерединеПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесВыравниванияПоСерединеПотенциальногоНаименования"]);
		public static double МетаВесВыравниванияСлеваПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесВыравниванияСлеваПотенциальногоНаименования"]);
		public static double МетаВесВыравниванияСправаПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесВыравниванияСправаПотенциальногоНаименования"]);
		public static double МетаВесЖирностиТекстаПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесЖирностиТекстаПотенциальногоНаименования"]);
		public static double МетаВесПустойСтрокиПодЯчейкойПотенциальногоНаименования => double.Parse(ConfigurationManager.AppSettings["МетаВесПустойСтрокиПодЯчейкойПотенциальногоНаименования"]);
		public static double МетаВесЧастоИспользуемогоТермина => double.Parse(ConfigurationManager.AppSettings["МетаВесЧастоИспользуемогоТермина"]);

		public static string МетаМеткаНаименование 
		{
			get => ConfigurationManager.AppSettings["МетаМеткаНаименование"];
			
			set => ConfigurationManager.AppSettings["МетаМеткаНаименование"] = value;
		}
		public static bool МетаЯвляетсяЗапросом
		{
			get => bool.Parse(ConfigurationManager.AppSettings["МетаЯвляетсяЗапросом"]);

			set => ConfigurationManager.AppSettings["МетаЯвляетсяЗапросом"] = value.ToString();
		}


		public static string СвободнаяЯчейкаТегПрефикс 
		{
			get => ConfigurationManager.AppSettings["СвободнаяЯчейкаТегПрефикс"];
			
			set => ConfigurationManager.AppSettings["СвободнаяЯчейкаТегПрефикс"] = value;
		}
		public static string СвободнаяЯчейкаМеткаКодыЯчеек 
		{
			get => ConfigurationManager.AppSettings["СвободнаяЯчейкаМеткаКодыЯчеек"];
			
			set => ConfigurationManager.AppSettings["СвободнаяЯчейкаМеткаКодыЯчеек"] = value;
		}

		public static string СтолбецТегПрефикс 
		{
			get => ConfigurationManager.AppSettings["СтолбецТегПрефикс"];
			
			set => ConfigurationManager.AppSettings["СтолбецТегПрефикс"] = value;
		}

		public static string СтрокаТегПрефикс 
		{
			get => ConfigurationManager.AppSettings["СтрокаТегПрефикс"];
			
			set => ConfigurationManager.AppSettings["СтрокаТегПрефикс"] = value;
		}

		public static string ТаблицаПрефиксТега 
		{
			get => ConfigurationManager.AppSettings["ТаблицаПрефиксТега"];
			
			set => ConfigurationManager.AppSettings["ТаблицаПрефиксТега"] = value;
		}
		public static string ТаблицаМеткаТипТаблицыДинамическая 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаТипТаблицыДинамическая"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаТипТаблицыДинамическая"] = value;
		}
		public static string ТаблицаМеткаТипТаблицыСтатическая 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаТипТаблицыСтатическая"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаТипТаблицыСтатическая"] = value;
		}
		public static string ТаблицаМеткаКодыСтрокИСтолбцов 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтрокИСтолбцов"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтрокИСтолбцов"] = value;
		}
		public static string ТаблицаМеткаКодыСтрок 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтрок"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтрок"] = value;
		}
		public static string ТаблицаМеткаКодыСтолбцов
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтолбцов"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаКодыСтолбцов"] = value;
		}
		public static string ТаблицаМеткаНаименование 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаНаименование"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаНаименование"] = value;
		}
		public static string ТаблицаМеткаТег 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаТег"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаТег"] = value;
		}
		public static string ТаблицаМеткаКод 
		{
			get => ConfigurationManager.AppSettings["ТаблицаМеткаКод"];
			
			set => ConfigurationManager.AppSettings["ТаблицаМеткаКод"] = value;
		}

		public static string ПутьКПапкеСгенерированныхШаблонов 
		{
			get => ConfigurationManager.AppSettings["ПутьКПапкеСгенерированныхШаблонов"];
			
			set => ConfigurationManager.AppSettings["ПутьКПапкеСгенерированныхШаблонов"] = value;
		}

		public static int КоличествоСловВТеге
		{
			get => int.Parse(ConfigurationManager.AppSettings["КоличествоСловВТеге"]);

			set => ConfigurationManager.AppSettings["КоличествоСловВТеге"] = value.ToString();
		}
		public static int КоличествоСимволовВТеге 
		{
			get => int.Parse(ConfigurationManager.AppSettings["КоличествоСимволовВТеге"]); 
			
			set => ConfigurationManager.AppSettings["КоличествоСимволовВТеге"] = value.ToString();
		}

		public static string[] ЧастоИспользуемыеТермины => ConfigurationManager.AppSettings["ЧастоИспользуемыеТермины"].Split(',');

		public static Dictionary<string[], Type> ТипыДанных => new Dictionary<string[], Type>()
		{
			{ МаскаТипаДанныхОбщий, typeof(Финансовый) },
			{ МаскаТипаДанныхЧисловой, typeof(Числовой) },
			{ МаскаТипаДанныхЦелочисленный, typeof(Целочисленный) },
			{ МаскаТипаДанныхФинансовый, typeof(Финансовый) },
			{ МаскаТипаДанныхДатаВремя, typeof(ДатаВремя) },
			{ МаскаТипаДанныхСтроковый, typeof(Строковый) },
		};

		public static string[] МаскаТипаДанныхОбщий => ConfigurationManager.AppSettings["МаскаТипаДанныхОбщий"].Split('|');
		public static string[] МаскаТипаДанныхЧисловой => ConfigurationManager.AppSettings["МаскаТипаДанныхЧисловой"].Split('|');
		public static string[] МаскаТипаДанныхЦелочисленный => ConfigurationManager.AppSettings["МаскаТипаДанныхЦелочисленный"].Split('|');
		public static string[] МаскаТипаДанныхФинансовый => ConfigurationManager.AppSettings["МаскаТипаДанныхФинансовый"].Split('|');
		public static string[] МаскаТипаДанныхДатаВремя => ConfigurationManager.AppSettings["МаскаТипаДанныхДатаВремя"].Split('|');
		public static string[] МаскаТипаДанныхСтроковый => ConfigurationManager.AppSettings["МаскаТипаДанныхСтроковый"].Split('|');

		public static void SaveConfig()
		{
			Configuration configApp = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

			configApp.AppSettings.Settings["МетаВерсияМетаописания"].Value = МетаВерсияМетаописания;
			configApp.AppSettings.Settings["МетаИдентификатор"].Value = МетаИдентификатор;
			configApp.AppSettings.Settings["МетаГруппа"].Value = МетаГруппа;
			configApp.AppSettings.Settings["МетаДатаНачалаДействия"].Value = МетаДатаНачалаДействия;
			configApp.AppSettings.Settings["МетаДатаОкончанияДействия"].Value = МетаДатаОкончанияДействия;
			configApp.AppSettings.Settings["МетаАвторство"].Value = МетаАвторство;
			configApp.AppSettings.Settings["МетаНомерВерсии"].Value = МетаНомерВерсии;
			configApp.AppSettings.Settings["МетаРасположениеШапки"].Value = МетаРасположениеШапки;
			configApp.AppSettings.Settings["МетаВерсияФорматаМетаструктуры"].Value = МетаВерсияФорматаМетаструктуры;
			configApp.AppSettings.Settings["МетаМеткаНаименование"].Value = МетаМеткаНаименование;

			configApp.AppSettings.Settings["ТаблицаПрефиксТега"].Value = ТаблицаПрефиксТега;
			configApp.AppSettings.Settings["СвободнаяЯчейкаТегПрефикс"].Value = СвободнаяЯчейкаТегПрефикс;
			configApp.AppSettings.Settings["СтолбецТегПрефикс"].Value = СтолбецТегПрефикс;
			configApp.AppSettings.Settings["СтрокаТегПрефикс"].Value = СтрокаТегПрефикс;
			configApp.AppSettings.Settings["КоличествоСловВТеге"].Value = КоличествоСловВТеге.ToString();
			configApp.AppSettings.Settings["КоличествоСимволовВТеге"].Value = КоличествоСимволовВТеге.ToString();

			configApp.AppSettings.Settings["ТаблицаМеткаТипТаблицыДинамическая"].Value = ТаблицаМеткаТипТаблицыДинамическая;
			configApp.AppSettings.Settings["ТаблицаМеткаТипТаблицыСтатическая"].Value = ТаблицаМеткаТипТаблицыСтатическая;
			configApp.AppSettings.Settings["ТаблицаМеткаКодыСтрокИСтолбцов"].Value = ТаблицаМеткаКодыСтрокИСтолбцов;
			configApp.AppSettings.Settings["ТаблицаМеткаКодыСтолбцов"].Value = ТаблицаМеткаКодыСтолбцов;
			configApp.AppSettings.Settings["ТаблицаМеткаКодыСтрок"].Value = ТаблицаМеткаКодыСтрок;
			configApp.AppSettings.Settings["ТаблицаМеткаНаименование"].Value = ТаблицаМеткаНаименование;
			configApp.AppSettings.Settings["ТаблицаМеткаТег"].Value = ТаблицаМеткаТег;
			configApp.AppSettings.Settings["ТаблицаМеткаКод"].Value = ТаблицаМеткаКод;

			configApp.Save(ConfigurationSaveMode.Full);
		}
	}
}
