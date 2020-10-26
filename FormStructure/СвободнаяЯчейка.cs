using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
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

		public СвободнаяЯчейка(string кодЯчейки, object типЯчейки)
		{
			Идентификатор = кодЯчейки;

			Код = кодЯчейки;

			ТипЯчейки = типЯчейки;

			тип = ТипЯчейки.GetType().Name;

			Тег = ConfigManager.СвободнаяЯчейкаТегПрефикс + ДопМетоды.ПолучитьТег(Идентификатор);

			Описание = ДопМетоды.ПолучитьСриализованныйТип(ТипЯчейки);
		}

		public static List<СвободнаяЯчейка> ПолучитьСвободныеЯчейки(Worksheet лист)
		{
			List<СвободнаяЯчейка> свободныеЯчейки = new List<СвободнаяЯчейка>();

			Range меткаКодыЯчеек = ПолучитьМеткуИзЛиста(лист);

			if (меткаКодыЯчеек != null)
			{
				Range столбецСКодамиЯчеек = лист.Application.Intersect(меткаКодыЯчеек.EntireColumn, лист.UsedRange);

				if (меткаКодыЯчеек != null)
				{
					foreach (Range клетка in столбецСКодамиЯчеек)
					{
						if (клетка.Row > меткаКодыЯчеек.Row)
						{
							if (!ДопМетоды.КлеткаПустаИлиСодержитМетку(клетка))
							{
								свободныеЯчейки.Add(new СвободнаяЯчейка(клетка.Value.ToString(), ПолучитьТипЯчейки(клетка)));
							}

							if (ДопМетоды.КлеткаПустаИлиСодержитМетку(клетка.Offset[1, 0]))
							{
								break;
							}
						}
					}
				}
			}

			return свободныеЯчейки;
		}

		private static object ПолучитьТипЯчейки(Range клеткаСтолбцаСКодамиЯчеек)
		{
			return ДопМетоды.ПолучитьТип(клеткаСтолбцаСКодамиЯчеек.Offset[0, 1].NumberFormat, false);
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и возвращает ячейку с меткой
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private static Range ПолучитьМеткуИзЛиста(Worksheet лист)
		{
			string меткаКодыЯчеек = ConfigManager.СвободнаяЯчейкаМеткаКодыЯчеек;

			foreach (Range клеткаТаблицы in лист.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == меткаКодыЯчеек)
					{
						return клеткаТаблицы;
					}
				}
			}

			return null;
		}

		private string тип;

		[XmlAttribute()]
		public string Идентификатор { get; set; }

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get; set; }

		[XmlAttribute()]
		public string Тип { get => тип; set => тип = value; }

		[XmlAttribute()]
		public string Описание { get; set; }

		[XmlAttribute()]
		public string Тег { get; set; }

		[XmlIgnore()]
		public object ТипЯчейки { get; }
	}
}
