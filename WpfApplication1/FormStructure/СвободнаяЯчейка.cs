using System;
using System.Configuration;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class СвободнаяЯчейка
	{
		public СвободнаяЯчейка()
		{
		}

		public СвободнаяЯчейка(string кодЯчейки, string типЯчейки)
		{
			идентификатор = кодЯчейки;
			код = кодЯчейки;
			тип = типЯчейки;
			тег = ConfigurationManager.AppSettings.Get("СвободнаяЯчейкаТегПрефикс") + ДопМетоды.ПолучитьТег(идентификатор);
			описание = ДопМетоды.ПолучитьСриализованныйТип(тип);
		}

		private string идентификатор;
		private string код;
		private string наименованиеЭлемента;
		private string тип;
		private string описание;
		private string тег;

		[XmlAttribute()]
		public string Идентификатор
		{
			get => идентификатор;
			set => идентификатор = value;
		}

		[XmlAttribute()]
		public string Код
		{
			get => код;
			set => код = value;
		}

		[XmlAttribute()]
		public string НаименованиеЭлемента
		{
			get => наименованиеЭлемента;
			set => наименованиеЭлемента = value;
		}

		[XmlAttribute()]
		public string Тип
		{
			get => тип;
			set => тип = value;
		}

		[XmlAttribute()]
		public string Описание
		{
			get => описание;
			set => описание = value;
		}

		[XmlAttribute()]
		public string Тег
		{
			get => тег;
			set => тег = value;
		}
	}
}
