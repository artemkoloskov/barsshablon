using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Таблица
	{
		public Таблица()
		{
		}

		public static List<Таблица> ПолучитьТаблицыФормы(Sheets листыКниги)
		{
			List<Таблица> таблицы = new List<Таблица>();

			int n = 1;

			foreach (Worksheet листКниги in листыКниги)
			{
				Таблица таблица = new Таблица(листКниги, n);

				n++;

				if (таблица != null)
				{
					таблицы.Add(таблица);
				}
			}

			return таблицы;
		}

		public Таблица(Worksheet листКниги, int n)
		{
			ЛистКниги = листКниги;

			НайтиТегиНаЛисте();

			Строки = ПолучитьСтрокиТаблицы();

			Столбцы = ПолучитьСтолбцыТаблицы();

			Наименование = ПолучитьНаименованиеТаблицы(n);

			Идентификатор = "Таблица" + n;

			Код = "Таблица" + n;

			Тег = ConfigManager.ТаблицаПрефиксТега + ДопМетоды.ПолучитьТег(Наименование);
		}

		private string ПолучитьНаименованиеТаблицы(int n)
		{
			if (тегНаименование != null)
			{
				if (ДопМетоды.ПолучитьНаименованиеПоТегу(тегНаименование, out string наименование))
				{
					return наименование;
				}
			}

			return "Таблица" + n;
		}

		private Строка[] ПолучитьСтрокиТаблицы()
		{
			if (тегКодыСтрок != null)
			{
				List<Строка> списокСтрок = new List<Строка>();

				foreach (Range клеткаСтолбцаСКодамиСтрок in ЛистКниги.Application.Intersect(тегКодыСтрок.EntireColumn, ЛистКниги.UsedRange).Cells)
				{
					if (!ДопМетоды.КлеткаПустаИлиСодержитТег(клеткаСтолбцаСКодамиСтрок) &&
						клеткаСтолбцаСКодамиСтрок.Row > тегКодыСтрок.Row)
					{
						списокСтрок.Add(new Строка(клеткаСтолбцаСКодамиСтрок));
					}
				}

				Строка[] строки = new Строка[списокСтрок.Count];

				int i = 0;

				foreach (Строка строка in списокСтрок)
				{
					строки[i] = строка;

					i++;
				}

				return строки;
			}

			return null;
		}

		private Столбец[] ПолучитьСтолбцыТаблицы()
		{
			if (тегКодыСтолбцов != null)
			{
				List<Столбец> списокСтолбцов = new List<Столбец>();

				foreach (Range клеткаСтрокиСКодамиСтолбцов in ЛистКниги.Application.Intersect(тегКодыСтолбцов.EntireRow, ЛистКниги.UsedRange).Cells)
				{
					if (!ДопМетоды.КлеткаПустаИлиСодержитТег(клеткаСтрокиСКодамиСтолбцов) &&
						клеткаСтрокиСКодамиСтолбцов.Column > тегКодыСтолбцов.Column)
					{
						списокСтолбцов.Add(new Столбец(клеткаСтрокиСКодамиСтолбцов, Динамическая));
					}
				}

				Столбец[] столбцы = new Столбец[списокСтолбцов.Count];

				int i = 0;

				foreach (Столбец столбец in списокСтолбцов)
				{
					столбцы[i] = столбец;

					i++;
				}

				return столбцы;
			}

			return null;
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и запоминает ячейки с тегами
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private void НайтиТегиНаЛисте()
		{
			foreach (Range клеткаТаблицы in ЛистКниги.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыДинамическая ||
								клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыСтатическая)
					{
						тегТипТаблицы = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаКодыСтрок)
					{
						тегКодыСтрок = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаКодыСтолбцов)
					{
						тегКодыСтолбцов = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаНаименование)
					{
						тегНаименование = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаТег)
					{
						тегТег = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаКод)
					{
						тегКод = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == ConfigManager.ТаблицаСтрокаТегаКодыСтрокИСтолбцов)
					{
						тегКодыСтрок = клеткаТаблицы;

						тегКодыСтолбцов = клеткаТаблицы;
					}
				}
			}
		}

		private Range тегТипТаблицы;
		private Range тегКодыСтолбцов;
		private Range тегКодыСтрок;
		private Range тегНаименование;
		private Range тегКод;
		private Range тегТег;

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("СвободнаяЯчейка", typeof(СвободнаяЯчейка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public СвободнаяЯчейка[] СвободныеЯчейки { get; set; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Строка", typeof(Строка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Строка[] Строки { get; set; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Столбец", typeof(Столбец), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Столбец[] Столбцы { get; set; }

		[XmlAttribute()]
		public string Идентификатор { get; set; }

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string Наименование { get; set; }

		[XmlAttribute()]
		public string Тег { get; set; }

		[XmlAttribute()]
		public string СсылкаНаМетодическийСправочник { get; set; }

		[XmlAttribute()]
		public bool РучноеДобавлениеСтрок { get; set; } = false;

		[XmlIgnore]
		public Worksheet ЛистКниги { get; set; }

		[XmlIgnore]
		public bool Динамическая =>
			!((Строки != null && Строки.Length > 0) ||
			(тегТипТаблицы != null && тегТипТаблицы.Value.toString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыСтатическая)) ||
			(тегТипТаблицы != null && тегТипТаблицы.Value.toString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыДинамическая);

		[XmlIgnore]
		public Range ТегТипТаблицы { get => тегТипТаблицы; set => тегТипТаблицы = value; }
		[XmlIgnore]
		public Range ТегКодыСтолбцов { get => тегКодыСтолбцов; set => тегКодыСтолбцов = value; }
		[XmlIgnore]
		public Range ТегКодыСтрок { get => тегКодыСтрок; set => тегКодыСтрок = value; }
		[XmlIgnore]
		public Range ТегНаименование { get => тегНаименование; set => тегНаименование = value; }
		[XmlIgnore]
		public Range ТегКод { get => тегКод; set => тегКод = value; }
		[XmlIgnore]
		public Range ТегТег { get => тегТег; set => тегТег = value; }
	}
}