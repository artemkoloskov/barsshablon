using Microsoft.Office.Interop.Excel;
using System.Xml.Serialization;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Столбец
	{
		public Столбец()
		{
		}

		public Столбец(string кодСтолбца, object типСтолбца)
		{
			идентификатор = кодСтолбца;
			код = кодСтолбца;
			объектТипаСтолбца = типСтолбца;
			тип = объектТипаСтолбца.GetType().Name;
			тег = ConfigManager.СтолбецТегПрефикс + ДопМетоды.ПолучитьТег(идентификатор);
			описание = ДопМетоды.ПолучитьСриализованныйТип(объектТипаСтолбца);
		}

		public Столбец(Range клеткаСтрокиСКодамиСтолбцов, bool являетсяКлючевым)
		{
			идентификатор = клеткаСтрокиСКодамиСтолбцов.Value.ToString();
			код = клеткаСтрокиСКодамиСтолбцов.Value.ToString();
			наименованиеЭлемента = ДопМетоды.ПолучитьНаименованиеСтрокиИлиСтолбца(клеткаОбластиСКодами: клеткаСтрокиСКодамиСтолбцов, ищемДляСтроки: false);
			тег = ConfigManager.СтолбецТегПрефикс + ДопМетоды.ПолучитьТег(идентификатор);
			ключевой = являетсяКлючевым;
			объектТипаСтолбца = ПолучитьТипСтолбца(клеткаСтрокиСКодамиСтолбцов);
			тип = объектТипаСтолбца.GetType().Name;
			описание = ДопМетоды.ПолучитьСриализованныйТип(объектТипаСтолбца);
		}

		private object ПолучитьТипСтолбца(Range клеткаСтрокиСКодамиСтолбцов)
		{
			return ДопМетоды.ПолучитьТип(клеткаСтрокиСКодамиСтолбцов.Offset[1, 0].NumberFormat, ключевой);
		}

		private string идентификатор;
		private string код;
		private string наименованиеЭлемента;
		private string тег;
		private string тип;
		private string описание;

		private object объектТипаСтолбца;
		private bool ключевой;

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