using Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Строка
	{
		public Строка()
		{
		}

		public Строка(string кодСтроки)
		{
			идентификатор = кодСтроки;
			код = кодСтроки;
			тег = ConfigManager.СтрокаТегПрефикс + ДопМетоды.ПолучитьТег(идентификатор);
		}

		public Строка(Range клеткаСтолбцаСКодамиСтрок)
		{
			идентификатор = клеткаСтолбцаСКодамиСтрок.Value.ToString();
			код = клеткаСтолбцаСКодамиСтрок.Value.ToString();
			наименованиеЭлемента = ДопМетоды.ПолучитьНаименованиеСтрокиИлиСтолбца(клеткаСтолбцаСКодамиСтрок, true);
			тег = ConfigManager.СтрокаТегПрефикс + ДопМетоды.ПолучитьТег(Идентификатор);
		}

		private string идентификатор;
		private string код;
		private string наименованиеЭлемента;
		private string тег;

		[XmlAttribute()]
		public string Идентификатор { get => идентификатор; set => идентификатор = value; }

		[XmlAttribute()]
		public string Код { get => код; set => код = value; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get => наименованиеЭлемента; set => наименованиеЭлемента = value; }

		[XmlAttribute()]
		public string Тег { get => тег; set => тег = value; }
	}
}