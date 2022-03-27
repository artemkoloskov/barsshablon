using System.Configuration;

namespace БАРСШаблон
{
	public static class SettingsManager
	{
		private static readonly Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

		public static Settings Settings = (Settings)config.GetSection("барсШаблонНастройки");

		public static void SaveSettings()
		{
			config.Save(ConfigurationSaveMode.Full);
		}
	}

	public class Settings : ConfigurationSection
	{
		public Settings()
		{
		}

		[ConfigurationProperty("мета", IsRequired = false)]
		public MetaSettings Meta
		{
			get => (MetaSettings)this["мета"];
			set => this["мета"] = value;
		}

		[ConfigurationProperty("теги", IsRequired = false)]
		public TagsSettings Tags
		{
			get => (TagsSettings)this["теги"];
			set => this["теги"] = value;
		}

		[ConfigurationProperty("разметка", IsRequired = false)]
		public MarkupSettingsНастройки Markup
		{
			get => (MarkupSettingsНастройки)this["разметка"];
			set => this["разметка"] = value;
		}

		[ConfigurationProperty("веса", IsRequired = false)]
		public TitleMarkerWeights Weight
		{
			get => (TitleMarkerWeights)this["веса"];
			set => this["веса"] = value;
		}

		[ConfigurationProperty("путьКПапкеСгенерированныхШаблонов", IsRequired = false)]
		public SettingsElement GeneratedTemplatesPath
		{
			get => (SettingsElement)this["путьКПапкеСгенерированныхШаблонов"];
			set => this["путьКПапкеСгенерированныхШаблонов"] = value;
		}
	}

	public class MetaSettings : ConfigurationElement
	{
		[ConfigurationProperty("версияМетаописания", IsRequired = false)]
		public SettingsElement MetaVersion
		{
			get => (SettingsElement)this["версияМетаописания"];
			set => this["версияМетаописания"] = value;
		}

		[ConfigurationProperty("идентификатор", IsRequired = false)]
		public SettingsElement Id
		{
			get => (SettingsElement)this["идентификатор"];
			set => this["идентификатор"] = value;
		}

		[ConfigurationProperty("группа", IsRequired = false)]
		public SettingsElement Group
		{
			get => (SettingsElement)this["группа"];
			set => this["группа"] = value;
		}

		[ConfigurationProperty("датаНачалаДействия", IsRequired = false)]
		public SettingsElement BeginDate
		{
			get => (SettingsElement)this["датаНачалаДействия"];
			set => this["датаНачалаДействия"] = value;
		}

		[ConfigurationProperty("датаОкончанияДействия", IsRequired = false)]
		public SettingsElement EndDate
		{
			get => (SettingsElement)this["датаОкончанияДействия"];
			set => this["датаОкончанияДействия"] = value;
		}

		[ConfigurationProperty("авторство", IsRequired = false)]
		public SettingsElement Author
		{
			get => (SettingsElement)this["авторство"];
			set => this["авторство"] = value;
		}

		[ConfigurationProperty("номерВерсии", IsRequired = false)]
		public SettingsElement VersionNumber
		{
			get => (SettingsElement)this["номерВерсии"];
			set => this["номерВерсии"] = value;
		}

		[ConfigurationProperty("расположениеШапки", IsRequired = false)]
		public SettingsElement HeaderPlacement
		{
			get => (SettingsElement)this["расположениеШапки"];
			set => this["расположениеШапки"] = value;
		}

		[ConfigurationProperty("версияФорматаМетаструктуры", IsRequired = false)]
		public SettingsElement MetaFormatVersion
		{
			get => (SettingsElement)this["версияФорматаМетаструктуры"];
			set => this["версияФорматаМетаструктуры"] = value;
		}

		[ConfigurationProperty("меткаНаименование", IsRequired = false)]
		public SettingsElement TitleMark
		{
			get => (SettingsElement)this["меткаНаименование"];
			set => this["меткаНаименование"] = value;
		}

		[ConfigurationProperty("являетсяЗапросом", IsRequired = false)]
		public LogicalSettingsElement IsARequest
		{
			get => (LogicalSettingsElement)this["являетсяЗапросом"];
			set => this["являетсяЗапросом"] = value;
		}
	}

	public class TagsSettings : ConfigurationElement
	{
		[ConfigurationProperty("префиксСвободнойЯчейки", IsRequired = false)]
		public SettingsElement FreeCellPrefix
		{
			get => (SettingsElement)this["префиксСвободнойЯчейки"];
			set => this["префиксСвободнойЯчейки"] = value;
		}

		[ConfigurationProperty("префиксСтолбца", IsRequired = false)]
		public SettingsElement ColumnPrefix
		{
			get => (SettingsElement)this["префиксСтолбца"];
			set => this["префиксСтолбца"] = value;
		}

		[ConfigurationProperty("префиксСтроки", IsRequired = false)]
		public SettingsElement RowPrefix
		{
			get => (SettingsElement)this["префиксСтроки"];
			set => this["префиксСтроки"] = value;
		}

		[ConfigurationProperty("префиксТаблицы", IsRequired = false)]
		public SettingsElement TablePrefix
		{
			get => (SettingsElement)this["префиксТаблицы"];
			set => this["префиксТаблицы"] = value;
		}

		[ConfigurationProperty("количествоСловВТеге", IsRequired = false)]
		public SettingsElement TagWordCount
		{
			get => (SettingsElement)this["количествоСловВТеге"];
			set => this["количествоСловВТеге"] = value;
		}

		[ConfigurationProperty("количествоСимволовВТеге", IsRequired = false)]
		public SettingsElement TagCharCount
		{
			get => (SettingsElement)this["количествоСимволовВТеге"];
			set => this["количествоСимволовВТеге"] = value;
		}
	}

	public class MarkupSettingsНастройки : ConfigurationElement
	{
		[ConfigurationProperty("меткаКодыЯчеек", IsRequired = false)]
		public SettingsElement CellCodesMark
		{
			get => (SettingsElement)this["меткаКодыЯчеек"];
			set => this["меткаКодыЯчеек"] = value;
		}

		[ConfigurationProperty("меткаТипТаблицыДинамическая", IsRequired = false)]
		public SettingsElement TableIsDynamicMark
		{
			get => (SettingsElement)this["меткаТипТаблицыДинамическая"];
			set => this["меткаТипТаблицыДинамическая"] = value;
		}

		[ConfigurationProperty("меткаТипТаблицыСтатическая", IsRequired = false)]
		public SettingsElement TableIsStaticMark
		{
			get => (SettingsElement)this["меткаТипТаблицыСтатическая"];
			set => this["меткаТипТаблицыСтатическая"] = value;
		}

		[ConfigurationProperty("меткаКодыСтрокИСтолбцов", IsRequired = false)]
		public SettingsElement RowAndColumnCodesMark
		{
			get => (SettingsElement)this["меткаКодыСтрокИСтолбцов"];
			set => this["меткаКодыСтрокИСтолбцов"] = value;
		}

		[ConfigurationProperty("меткаКодыСтрок", IsRequired = false)]
		public SettingsElement RowCodesMark
		{
			get => (SettingsElement)this["меткаКодыСтрок"];
			set => this["меткаКодыСтрок"] = value;
		}

		[ConfigurationProperty("меткаКодыСтолбцов", IsRequired = false)]
		public SettingsElement ColumnCodesMark
		{
			get => (SettingsElement)this["меткаКодыСтолбцов"];
			set => this["меткаКодыСтолбцов"] = value;
		}

		[ConfigurationProperty("меткаНаименование", IsRequired = false)]
		public SettingsElement TitleMark
		{
			get => (SettingsElement)this["меткаНаименование"];
			set => this["меткаНаименование"] = value;
		}

		[ConfigurationProperty("меткаТег", IsRequired = false)]
		public SettingsElement TagMark
		{
			get => (SettingsElement)this["меткаТег"];
			set => this["меткаТег"] = value;
		}

		[ConfigurationProperty("меткаКод", IsRequired = false)]
		public SettingsElement CodeMark
		{
			get => (SettingsElement)this["меткаКод"];
			set => this["меткаКод"] = value;
		}
	}

	public class TitleMarkerWeights : ConfigurationElement
	{
		[ConfigurationProperty("длина", IsRequired = false)]
		public TitleMarkerWeight Length
		{
			get => (TitleMarkerWeight)this["длина"];
			set => this["длина"] = value;
		}

		[ConfigurationProperty("номерСтроки", IsRequired = false)]
		public TitleMarkerWeight RowNumber
		{
			get => (TitleMarkerWeight)this["номерСтроки"];
			set => this["номерСтроки"] = value;
		}

		[ConfigurationProperty("номерСтолбца", IsRequired = false)]
		public TitleMarkerWeight ColumnNumber
		{
			get => (TitleMarkerWeight)this["номерСтолбца"];
			set => this["номерСтолбца"] = value;
		}

		[ConfigurationProperty("количествоЯчеекВОбъединеннойЯчейке", IsRequired = false)]
		public TitleMarkerWeight NumberOfCellsInMergedCell
		{
			get => (TitleMarkerWeight)this["количествоЯчеекВОбъединеннойЯчейке"];
			set => this["количествоЯчеекВОбъединеннойЯчейке"] = value;
		}

		[ConfigurationProperty("границаВнизу", IsRequired = false)]
		public TitleMarkerWeight BottomBorder
		{
			get => (TitleMarkerWeight)this["границаВнизу"];
			set => this["границаВнизу"] = value;
		}

		[ConfigurationProperty("границаВверху", IsRequired = false)]
		public TitleMarkerWeight TopBorder
		{
			get => (TitleMarkerWeight)this["границаВверху"];
			set => this["границаВверху"] = value;
		}

		[ConfigurationProperty("границаСлева", IsRequired = false)]
		public TitleMarkerWeight LeftBorder
		{
			get => (TitleMarkerWeight)this["границаСлева"];
			set => this["границаСлева"] = value;
		}

		[ConfigurationProperty("границаСправа", IsRequired = false)]
		public TitleMarkerWeight RightBorder
		{
			get => (TitleMarkerWeight)this["границаСправа"];
			set => this["границаСправа"] = value;
		}

		[ConfigurationProperty("выравниваниеПоСередине", IsRequired = false)]
		public TitleMarkerWeight CenterAligned
		{
			get => (TitleMarkerWeight)this["выравниваниеПоСередине"];
			set => this["выравниваниеПоСередине"] = value;
		}

		[ConfigurationProperty("выравниваниеСлева", IsRequired = false)]
		public TitleMarkerWeight LeftAligned
		{
			get => (TitleMarkerWeight)this["выравниваниеСлева"];
			set => this["выравниваниеСлева"] = value;
		}

		[ConfigurationProperty("выравниваниеСправа", IsRequired = false)]
		public TitleMarkerWeight RightAligned
		{
			get => (TitleMarkerWeight)this["выравниваниеСправа"];
			set => this["выравниваниеСправа"] = value;
		}

		[ConfigurationProperty("жирностьТекста", IsRequired = false)]
		public TitleMarkerWeight BoldText
		{
			get => (TitleMarkerWeight)this["жирностьТекста"];
			set => this["жирностьТекста"] = value;
		}

		[ConfigurationProperty("пустаяСтрокаПодЯчейкой", IsRequired = false)]
		public TitleMarkerWeight EmptyRowBelowCell
		{
			get => (TitleMarkerWeight)this["пустаяСтрокаПодЯчейкой"];
			set => this["пустаяСтрокаПодЯчейкой"] = value;
		}

		[ConfigurationProperty("частоИспользуемыйТермин", IsRequired = false)]
		public TitleMarkerWeight CommonlyUsedWord
		{
			get => (TitleMarkerWeight)this["частоИспользуемыйТермин"];
			set => this["частоИспользуемыйТермин"] = value;
		}
	}

	public class SettingsElement : ConfigurationElement
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

	public class TitleMarkerWeight : ConfigurationElement
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

	public class LogicalSettingsElement : ConfigurationElement
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
