using Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Столбец
	{
		public Столбец()
		{
		}

		public Столбец(string кодСтолбца, string типСтолбца)
		{
			идентификатор = кодСтолбца;
			код = кодСтолбца;
			тип = типСтолбца;
			тег = ConfigurationManager.AppSettings.Get("СтолбецТегПрефикс") + ДопМетоды.ПолучитьТег(идентификатор);
			описание = ДопМетоды.ПолучитьСриализованныйТип(тип);
		}

		public Столбец(Range клеткаСтрокиСКодамиСтолбцов)
		{
			идентификатор = клеткаСтрокиСКодамиСтолбцов.Value.ToString();
			код = клеткаСтрокиСКодамиСтолбцов.Value.ToString();
			наименованиеЭлемента = ДопМетоды.ПолучитьНаименованиеСтрокиИлиСтолбца(клеткаСтрокиСКодамиСтолбцов, false);
			тег = ConfigurationManager.AppSettings.Get("СтолбецТегПрефикс") + ДопМетоды.ПолучитьТег(идентификатор);
			тип = ПолучитьТипСтолбца(клеткаСтрокиСКодамиСтолбцов);
			описание = ДопМетоды.ПолучитьСриализованныйТип(тип);
		}

		private string ПолучитьТипСтолбца(Range клеткаСтрокиСКодамиСтолбцов)
		{
			return "Строковый";
		}

		private string идентификатор;
		private string код;
		private string наименованиеЭлемента;
		private string тег;
		private string тип;
		private string описание;

		private Range клеткаСтрокиСКодамиСтолбцов;

		[XmlAttribute()]
		public string Идентификатор { get => идентификатор; set => идентификатор = value; }

		[XmlAttribute()]
		public string Код { get => код; set => код = value; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get => наименованиеЭлемента; set => наименованиеЭлемента = value; }

		[XmlAttribute()]
		public string Тег { get => тег; set => тег = value; }

		[XmlAttribute()]
		public string Тип { get => тип; set => тип = value; }

		[XmlAttribute()]
		public string Описание { get => описание; set => описание = value; }

		[XmlIgnore]
		public Range КлеткаСтрокиСКодамиСтолбцов { get => клеткаСтрокиСКодамиСтолбцов; set => клеткаСтрокиСКодамиСтолбцов = value; }
	}
}