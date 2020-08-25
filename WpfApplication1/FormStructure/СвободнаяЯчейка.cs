using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows.Documents;
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
			тег = ConfigurationManager.AppSettings["СвободнаяЯчейкаТегПрефикс"] + ДопМетоды.ПолучитьТег(идентификатор);
			описание = ДопМетоды.ПолучитьСриализованныйТип(тип);
		}

		public static List<СвободнаяЯчейка> ПолучитьСвободныеЯчейки(Worksheet лист)
		{
			List<СвободнаяЯчейка> свободныеЯчейки = new List<СвободнаяЯчейка>();

			Range тегСКодамиЯчеек = ПолучитьТегИзЛиста(лист);

			if (тегСКодамиЯчеек != null)
			{
				Range столбецСКодамиЯчеек = лист.Application.Intersect(тегСКодамиЯчеек.EntireColumn, лист.UsedRange);

				if (тегСКодамиЯчеек != null)
				{
					foreach (Range клетка in столбецСКодамиЯчеек)
					{
						if (клетка.Row > тегСКодамиЯчеек.Row)
						{
							if (!ДопМетоды.КлеткаПустаИлиСодержитТег(клетка))
							{
								свободныеЯчейки.Add(new СвободнаяЯчейка(клетка.Value.ToString(), ПолучитьТипЯчейки(клетка)));
							}

							if (ДопМетоды.КлеткаПустаИлиСодержитТег(клетка.Offset[1, 0]))
							{
								break;
							}
						}
					}
				} 
			}

			return свободныеЯчейки;
		}

		private static string ПолучитьТипЯчейки(Range клеткаСтолбцаСКодамиЯчеек)
		{
			return "Строковый"; //TODO
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и возвращает ячейку с тегом
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private static Range ПолучитьТегИзЛиста(Worksheet лист)
		{
			string строкаТегаКодыЯчеек = ConfigurationManager.AppSettings["СвободнаяЯчейкаСтрокаТегаКодыЯчеек"];

			foreach (Range клеткаТаблицы in лист.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == строкаТегаКодыЯчеек)
					{
						return клеткаТаблицы;
					}
				}
			}

			return null;
		}

		private string идентификатор;
		private string код;
		private string наименованиеЭлемента;
		private string тип;
		private string описание;
		private string тег;

		[XmlAttribute()]
		public string Идентификатор { get => идентификатор; set => идентификатор = value; }

		[XmlAttribute()]
		public string Код { get => код; set => код = value; }

		[XmlAttribute()]
		public string НаименованиеЭлемента { get => наименованиеЭлемента; set => наименованиеЭлемента = value; }

		[XmlAttribute()]
		public string Тип { get => тип; set => тип = value; }

		[XmlAttribute()]
		public string Описание { get => описание; set => описание = value; }

		[XmlAttribute()]
		public string Тег { get => тег; set => тег = value; }
	}
}
