using System;
using System.Collections.Generic;
using System.Configuration;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public static class ConfigManager
	{
		public static string МетаВерсияМетаописания => ConfigurationManager.AppSettings["МетаВерсияМетаописания"];
		public static string МетаИдентификатор => ConfigurationManager.AppSettings["МетаИдентификатор"];
		public static string МетаГруппа => ConfigurationManager.AppSettings["МетаГруппа"];
		public static string МетаДатаНачалаДействия => ConfigurationManager.AppSettings["МетаДатаНачалаДействия"];
		public static string МетаДатаОкончанияДействия => ConfigurationManager.AppSettings["МетаДатаОкончанияДействия"];
		public static string МетаАвторство => ConfigurationManager.AppSettings["МетаАвторство"];
		public static string МетаНомерВерсии => ConfigurationManager.AppSettings["МетаНомерВерсии"];
		public static string МетаРасположениеШапки => ConfigurationManager.AppSettings["МетаРасположениеШапки"];
		public static string МетаВерсияФорматаМетаструктуры => ConfigurationManager.AppSettings["МетаВерсияФорматаМетаструктуры"];

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

		public static string МетаТегНаименование => ConfigurationManager.AppSettings["МетаТегНаименование"];

		public static string СвободнаяЯчейкаТегПрефикс => ConfigurationManager.AppSettings["СвободнаяЯчейкаТегПрефикс"];
		public static string СвободнаяЯчейкаСтрокаТегаКодыЯчеек => ConfigurationManager.AppSettings["СвободнаяЯчейкаСтрокаТегаКодыЯчеек"];

		public static string СтолбецТегПрефикс => ConfigurationManager.AppSettings["СтолбецТегПрефикс"];

		public static string СтрокаТегПрефикс => ConfigurationManager.AppSettings["СтрокаТегПрефикс"];

		public static string ТаблицаПрефиксТега => ConfigurationManager.AppSettings["ТаблицаПрефиксТега"];
		public static string ТаблицаСтрокаТегаТипТаблицыДинамическая => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаТипТаблицыДинамическая"];
		public static string ТаблицаСтрокаТегаТипТаблицыСтатическая => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаТипТаблицыСтатическая"];
		public static string ТаблицаСтрокаТегаКодыСтрокИСтолбцов => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаКодыСтрокИСтолбцов"];
		public static string ТаблицаСтрокаТегаКодыСтрок => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаКодыСтрок"];
		public static string ТаблицаСтрокаТегаКодыСтолбцов => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаКодыСтолбцов"];
		public static string ТаблицаСтрокаТегаНаименование => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаНаименование"];
		public static string ТаблицаСтрокаТегаТег => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаТег"];
		public static string ТаблицаСтрокаТегаКод => ConfigurationManager.AppSettings["ТаблицаСтрокаТегаКод"];

		public static string ПутьКПапкеСгенерированныхШаблонов => ConfigurationManager.AppSettings["ПутьКПапкеСгенерированныхШаблонов"];

		public static int КоличествоСловВТеге => int.Parse(ConfigurationManager.AppSettings["КоличествоСловВТеге"]);
		public static int КоличествоСимволовВТеге => int.Parse(ConfigurationManager.AppSettings["КоличествоСимволовВТеге"]);

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
	}
}
