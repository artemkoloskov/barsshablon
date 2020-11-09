using Microsoft.Office.Interop.Excel;
using System;
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

		public Столбец(Range клеткаСтрокиСКодамиСтолбцов, bool являетсяКлючевым)
		{
			Console.WriteLine(DateTime.Now + ": столбец " + клеткаСтрокиСКодамиСтолбцов.Value.ToString() + ", начат");

			Идентификатор = клеткаСтрокиСКодамиСтолбцов.Value.ToString();

			Код = клеткаСтрокиСКодамиСтолбцов.Value.ToString();

			НаименованиеЭлемента = ДопМетоды.ПолучитьНаименованиеСтрокиИлиСтолбца(клеткаОбластиСКодами: клеткаСтрокиСКодамиСтолбцов, поискДляСтроки: false);

			Тег = МенеджерНастроек.Настройки.Теги.ПрефиксСтолбца.Value + 
				ДопМетоды.ПолучитьТегПоМетке(клеткаСтрокиСКодамиСтолбцов, false) == "" ? 
				ДопМетоды.ПолучитьТег(Идентификатор) : 
				ДопМетоды.ПолучитьТегПоМетке(клеткаСтрокиСКодамиСтолбцов, false);

			ЯвлетсяКлючевым = являетсяКлючевым;

			ТипСтолбца = ПолучитьТипСтолбца(клеткаСтрокиСКодамиСтолбцов);

			тип = ТипСтолбца.GetType().Name;

			Описание = ДопМетоды.ПолучитьСриализованныйТип(ТипСтолбца);

			Console.WriteLine(DateTime.Now + ": столбец " + клеткаСтрокиСКодамиСтолбцов.Value.ToString() + ", закончен");
		}

		private object ПолучитьТипСтолбца(Range клеткаСтрокиСКодамиСтолбцов)
		{
			return ДопМетоды.ПолучитьТип(клеткаСтрокиСКодамиСтолбцов.Offset[1, 0].NumberFormat, ЯвлетсяКлючевым);
		}

		private string тип;

		[XmlAttribute()]
		public string Идентификатор { get; set; }

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get; set; }

		[XmlAttribute()]
		public string Тег { get; set; }

		[XmlAttribute()]
		public string Тип { get => тип; set => тип = value; }

		[XmlAttribute()]
		public string Описание { get; set; }

		[XmlIgnore]
		public object ТипСтолбца { get; }

		[XmlIgnore]
		public bool ЯвлетсяКлючевым { get; }
	}
}