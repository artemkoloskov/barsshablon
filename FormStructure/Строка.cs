using Microsoft.Office.Interop.Excel;
using System;
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

		public Строка(Range клеткаСтолбцаСКодамиСтрок)
		{
			Console.WriteLine(DateTime.Now + ": строка " + клеткаСтолбцаСКодамиСтрок.Value.ToString() + ", начата");

			Идентификатор = клеткаСтолбцаСКодамиСтрок.Value.ToString();

			Код = клеткаСтолбцаСКодамиСтрок.Value.ToString();

			НаименованиеЭлемента = ДопМетоды.ПолучитьНаименованиеСтрокиИлиСтолбца(клеткаСтолбцаСКодамиСтрок, true);

			Тег = МенеджерНастроек.Настройки.Теги.ПрефиксСтроки.Value +
				ДопМетоды.ПолучитьТегПоМетке(клеткаСтолбцаСКодамиСтрок, true) == "" ?
				ДопМетоды.ПолучитьТег(Идентификатор) :
				ДопМетоды.ПолучитьТегПоМетке(клеткаСтолбцаСКодамиСтрок, true);

			Console.WriteLine(DateTime.Now + ": строка " + клеткаСтолбцаСКодамиСтрок.Value.ToString() + ", закончена");
		}

		[XmlAttribute()]
		public string Идентификатор { get; set; }

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get; set; }

		[XmlAttribute()]
		public string Тег { get; set; }
	}
}