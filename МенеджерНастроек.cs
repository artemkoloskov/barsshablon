using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace БАРСШаблон
{
	public static class МенеджерНастроек
	{
		private static readonly Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

		public static Настройки Настройки = (Настройки)config.GetSection("барсШаблонНастройки");

		public static void СохранитьНастройки()
		{
			config.Save(ConfigurationSaveMode.Full);
		}
	}

	public class Настройки : ConfigurationSection
	{
		public Настройки()
		{
		}

		[ConfigurationProperty("мета", IsRequired = false)]
		public НастройкиМеты Мета
		{
			get => (НастройкиМеты)this["мета"];
			set => this["мета"] = value;
		}

		[ConfigurationProperty("теги", IsRequired = false)]
		public НастройкиТегов Теги
		{
			get => (НастройкиТегов)this["теги"];
			set => this["теги"] = value;
		}

		[ConfigurationProperty("разметка", IsRequired = false)]
		public НастройкиРазметки Разметка
		{
			get => (НастройкиРазметки)this["разметка"];
			set => this["разметка"] = value;
		}

		[ConfigurationProperty("веса", IsRequired = false)]
		public ВесаПризнаковНаименования Вес
		{
			get => (ВесаПризнаковНаименования)this["веса"];
			set => this["веса"] = value;
		}

		[ConfigurationProperty("путьКПапкеСгенерированныхШаблонов", IsRequired = false)]
		public ЭлементНастроек ПутьКПапкеСгенерированныхШаблонов
		{
			get => (ЭлементНастроек)this["путьКПапкеСгенерированныхШаблонов"];
			set => this["путьКПапкеСгенерированныхШаблонов"] = value;
		}
	}

	public class НастройкиМеты : ConfigurationElement
	{
		[ConfigurationProperty("версияМетаописания", IsRequired = false)]
		public ЭлементНастроек ВерсияМетаописания
		{
			get => (ЭлементНастроек)this["версияМетаописания"];
			set => this["версияМетаописания"] = value;
		}

		[ConfigurationProperty("идентификатор", IsRequired = false)]
		public ЭлементНастроек Идентификатор
		{
			get => (ЭлементНастроек)this["идентификатор"];
			set => this["идентификатор"] = value;
		}

		[ConfigurationProperty("группа", IsRequired = false)]
		public ЭлементНастроек Группа
		{
			get => (ЭлементНастроек)this["группа"];
			set => this["группа"] = value;
		}

		[ConfigurationProperty("датаНачалаДействия", IsRequired = false)]
		public ЭлементНастроек ДатаНачалаДействия
		{
			get => (ЭлементНастроек)this["датаНачалаДействия"];
			set => this["датаНачалаДействия"] = value;
		}

		[ConfigurationProperty("датаОкончанияДействия", IsRequired = false)]
		public ЭлементНастроек ДатаОкончанияДействия
		{
			get => (ЭлементНастроек)this["датаОкончанияДействия"];
			set => this["датаОкончанияДействия"] = value;
		}

		[ConfigurationProperty("авторство", IsRequired = false)]
		public ЭлементНастроек Авторство
		{
			get => (ЭлементНастроек)this["авторство"];
			set => this["авторство"] = value;
		}

		[ConfigurationProperty("номерВерсии", IsRequired = false)]
		public ЭлементНастроек НомерВерсии
		{
			get => (ЭлементНастроек)this["номерВерсии"];
			set => this["номерВерсии"] = value;
		}

		[ConfigurationProperty("расположениеШапки", IsRequired = false)]
		public ЭлементНастроек РасположениеШапки
		{
			get => (ЭлементНастроек)this["расположениеШапки"];
			set => this["расположениеШапки"] = value;
		}

		[ConfigurationProperty("версияФорматаМетаструктуры", IsRequired = false)]
		public ЭлементНастроек ВерсияФорматаМетаструктуры
		{
			get => (ЭлементНастроек)this["версияФорматаМетаструктуры"];
			set => this["версияФорматаМетаструктуры"] = value;
		}

		[ConfigurationProperty("меткаНаименование", IsRequired = false)]
		public ЭлементНастроек МеткаНаименование
		{
			get => (ЭлементНастроек)this["меткаНаименование"];
			set => this["меткаНаименование"] = value;
		}

		[ConfigurationProperty("являетсяЗапросом", IsRequired = false)]
		public ЛогическийЭлементНастроек ЯвляетсяЗапросом
		{
			get => (ЛогическийЭлементНастроек)this["являетсяЗапросом"];
			set => this["являетсяЗапросом"] = value;
		}
	}

	public class НастройкиТегов : ConfigurationElement
	{
		[ConfigurationProperty("префиксСвободнойЯчейки", IsRequired = false)]
		public ЭлементНастроек ПрефиксСвободнойЯчейки
		{
			get => (ЭлементНастроек)this["префиксСвободнойЯчейки"];
			set => this["префиксСвободнойЯчейки"] = value;
		}

		[ConfigurationProperty("префиксСтолбца", IsRequired = false)]
		public ЭлементНастроек ПрефиксСтолбца
		{
			get => (ЭлементНастроек)this["префиксСтолбца"];
			set => this["префиксСтолбца"] = value;
		}

		[ConfigurationProperty("префиксСтроки", IsRequired = false)]
		public ЭлементНастроек ПрефиксСтроки
		{
			get => (ЭлементНастроек)this["префиксСтроки"];
			set => this["префиксСтроки"] = value;
		}

		[ConfigurationProperty("префиксТаблицы", IsRequired = false)]
		public ЭлементНастроек ПрефиксТаблицы
		{
			get => (ЭлементНастроек)this["префиксТаблицы"];
			set => this["префиксТаблицы"] = value;
		}

		[ConfigurationProperty("количествоСловВТеге", IsRequired = false)]
		public ЭлементНастроек КоличествоСловВТеге
		{
			get => (ЭлементНастроек)this["количествоСловВТеге"];
			set => this["количествоСловВТеге"] = value;
		}

		[ConfigurationProperty("количествоСимволовВТеге", IsRequired = false)]
		public ЭлементНастроек КоличествоСимволовВТеге
		{
			get => (ЭлементНастроек)this["количествоСимволовВТеге"];
			set => this["количествоСимволовВТеге"] = value;
		}
	}

	public class НастройкиРазметки : ConfigurationElement
	{
		[ConfigurationProperty("меткаКодыЯчеек", IsRequired = false)]
		public ЭлементНастроек МеткаКодыЯчеек
		{
			get => (ЭлементНастроек)this["меткаКодыЯчеек"];
			set => this["меткаКодыЯчеек"] = value;
		}

		[ConfigurationProperty("меткаТипТаблицыДинамическая", IsRequired = false)]
		public ЭлементНастроек МеткаТипТаблицыДинамическая
		{
			get => (ЭлементНастроек)this["меткаТипТаблицыДинамическая"];
			set => this["меткаТипТаблицыДинамическая"] = value;
		}

		[ConfigurationProperty("меткаТипТаблицыСтатическая", IsRequired = false)]
		public ЭлементНастроек МеткаТипТаблицыСтатическая
		{
			get => (ЭлементНастроек)this["меткаТипТаблицыСтатическая"];
			set => this["меткаТипТаблицыСтатическая"] = value;
		}

		[ConfigurationProperty("меткаКодыСтрокИСтолбцов", IsRequired = false)]
		public ЭлементНастроек МеткаКодыСтрокИСтолбцов
		{
			get => (ЭлементНастроек)this["меткаКодыСтрокИСтолбцов"];
			set => this["меткаКодыСтрокИСтолбцов"] = value;
		}

		[ConfigurationProperty("меткаКодыСтрок", IsRequired = false)]
		public ЭлементНастроек МеткаКодыСтрок
		{
			get => (ЭлементНастроек)this["меткаКодыСтрок"];
			set => this["меткаКодыСтрок"] = value;
		}

		[ConfigurationProperty("меткаКодыСтолбцов", IsRequired = false)]
		public ЭлементНастроек МеткаКодыСтолбцов
		{
			get => (ЭлементНастроек)this["меткаКодыСтолбцов"];
			set => this["меткаКодыСтолбцов"] = value;
		}

		[ConfigurationProperty("меткаНаименование", IsRequired = false)]
		public ЭлементНастроек МеткаНаименование
		{
			get => (ЭлементНастроек)this["меткаНаименование"];
			set => this["меткаНаименование"] = value;
		}

		[ConfigurationProperty("меткаТег", IsRequired = false)]
		public ЭлементНастроек МеткаТег
		{
			get => (ЭлементНастроек)this["меткаТег"];
			set => this["меткаТег"] = value;
		}

		[ConfigurationProperty("меткаКод", IsRequired = false)]
		public ЭлементНастроек МеткаКод
		{
			get => (ЭлементНастроек)this["меткаКод"];
			set => this["меткаКод"] = value;
		}
	}

	public class ВесаПризнаковНаименования : ConfigurationElement
	{
		[ConfigurationProperty("длина", IsRequired = false)]
		public ВесПризнакаНаименования Длина
		{
			get => (ВесПризнакаНаименования)this["длина"];
			set => this["длина"] = value;
		}

		[ConfigurationProperty("номерСтроки", IsRequired = false)]
		public ВесПризнакаНаименования НомерСтроки
		{
			get => (ВесПризнакаНаименования)this["номерСтроки"];
			set => this["номерСтроки"] = value;
		}

		[ConfigurationProperty("номерСтолбца", IsRequired = false)]
		public ВесПризнакаНаименования НомерСтолбца
		{
			get => (ВесПризнакаНаименования)this["номерСтолбца"];
			set => this["номерСтолбца"] = value;
		}

		[ConfigurationProperty("количествоЯчеекВОбъединеннойЯчейке", IsRequired = false)]
		public ВесПризнакаНаименования КоличествоЯчеекВОбъединеннойЯчейке
		{
			get => (ВесПризнакаНаименования)this["количествоЯчеекВОбъединеннойЯчейке"];
			set => this["количествоЯчеекВОбъединеннойЯчейке"] = value;
		}

		[ConfigurationProperty("границаВнизу", IsRequired = false)]
		public ВесПризнакаНаименования ГраницаВнизу
		{
			get => (ВесПризнакаНаименования)this["границаВнизу"];
			set => this["границаВнизу"] = value;
		}

		[ConfigurationProperty("границаВверху", IsRequired = false)]
		public ВесПризнакаНаименования ГраницаВверху
		{
			get => (ВесПризнакаНаименования)this["границаВверху"];
			set => this["границаВверху"] = value;
		}

		[ConfigurationProperty("границаСлева", IsRequired = false)]
		public ВесПризнакаНаименования ГраницаСлева
		{
			get => (ВесПризнакаНаименования)this["границаСлева"];
			set => this["границаСлева"] = value;
		}

		[ConfigurationProperty("границаСправа", IsRequired = false)]
		public ВесПризнакаНаименования ГраницаСправа
		{
			get => (ВесПризнакаНаименования)this["границаСправа"];
			set => this["границаСправа"] = value;
		}

		[ConfigurationProperty("выравниваниеПоСередине", IsRequired = false)]
		public ВесПризнакаНаименования ВыравниваниеПоСередине
		{
			get => (ВесПризнакаНаименования)this["выравниваниеПоСередине"];
			set => this["выравниваниеПоСередине"] = value;
		}

		[ConfigurationProperty("выравниваниеСлева", IsRequired = false)]
		public ВесПризнакаНаименования ВыравниваниеСлева
		{
			get => (ВесПризнакаНаименования)this["выравниваниеСлева"];
			set => this["выравниваниеСлева"] = value;
		}

		[ConfigurationProperty("выравниваниеСправа", IsRequired = false)]
		public ВесПризнакаНаименования ВыравниваниеСправа
		{
			get => (ВесПризнакаНаименования)this["выравниваниеСправа"];
			set => this["выравниваниеСправа"] = value;
		}

		[ConfigurationProperty("жирностьТекста", IsRequired = false)]
		public ВесПризнакаНаименования ЖирностьТекста
		{
			get => (ВесПризнакаНаименования)this["жирностьТекста"];
			set => this["жирностьТекста"] = value;
		}

		[ConfigurationProperty("пустаяСтрокаПодЯчейкой", IsRequired = false)]
		public ВесПризнакаНаименования ПустаяСтрокаПодЯчейкой
		{
			get => (ВесПризнакаНаименования)this["пустаяСтрокаПодЯчейкой"];
			set => this["пустаяСтрокаПодЯчейкой"] = value;
		}

		[ConfigurationProperty("частоИспользуемыйТермин", IsRequired = false)]
		public ВесПризнакаНаименования ЧастоИспользуемыйТермин
		{
			get => (ВесПризнакаНаименования)this["частоИспользуемыйТермин"];
			set => this["частоИспользуемыйТермин"] = value;
		}
	}

	public class ЭлементНастроек : ConfigurationElement
	{
		[ConfigurationProperty("value", IsRequired = false)]
		public string Value
		{
			get => (string)this["value"];
			set => this["value"] = value;
		}

		[ConfigurationProperty("default", IsRequired = false)]
		public string Default => (string)this["default"];
	}

	public class ВесПризнакаНаименования : ConfigurationElement
	{
		[ConfigurationProperty("value", IsRequired = false)]
		public double Value
		{
			get => double.Parse(this["value"].ToString());
			set => this["value"] = value;
		}

		[ConfigurationProperty("default", IsRequired = false)]
		public double Default => double.Parse(this["default"].ToString());
	}

	public class ЛогическийЭлементНастроек : ConfigurationElement
	{
		[ConfigurationProperty("value", IsRequired = false)]
		public bool Value
		{
			get => (bool)this["value"];
			set => this["value"] = value;
		}

		[ConfigurationProperty("default", IsRequired = false)]
		public bool Default => (bool)this["default"];
	}
}
