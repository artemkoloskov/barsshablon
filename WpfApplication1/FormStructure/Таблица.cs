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

		public Таблица(string кодТаблицы)
		{
			идентификатор = кодТаблицы;
			код = кодТаблицы;
			тег = ДопМетоды.ПолучитьТег(идентификатор);
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

			НайтиТэгиНаЛисте();

			строки = ПолучитьСтрокиТаблицы();

			столбцы = ПолучитьСтолбцыТаблицы();

			идентификатор = "Таблица" + n;
			код = "Таблица" + n;
			наименование = ПолучитьНаименованиеТаблицы(n);
			ручноеДобавлениеСтрок = false;
			тег = "Таблица" + n;
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
		private void НайтиТэгиНаЛисте()
		{
			string строкаТегаТипТаблицыДинамическая = ConfigManager.ТаблицаСтрокаТегаТипТаблицыДинамическая;
			string строкаТегаТипТаблицыСтатическая = ConfigManager.ТаблицаСтрокаТегаТипТаблицыСтатическая;
			string строкаТегаКодыСтрок = ConfigManager.ТаблицаСтрокаТегаКодыСтрок;
			string строкаТегаКодыСтрокИСтолбцов = ConfigManager.ТаблицаСтрокаТегаКодыСтрокИСтолбцов;
			string строкаТегаКодыСтолбцов = ConfigManager.ТаблицаСтрокаТегаКодыСтолбцов;
			string строкаТегаНаименование = ConfigManager.ТаблицаСтрокаТегаНаименование;
			string строкаТегаТег = ConfigManager.ТаблицаСтрокаТегаТег;
			string строкаТегаКод = ConfigManager.ТаблицаСтрокаТегаКод;

			foreach (Range клеткаТаблицы in ЛистКниги.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == строкаТегаТипТаблицыДинамическая ||
								клеткаТаблицы.Value.ToString() == строкаТегаТипТаблицыСтатическая)
					{
						тегТипТаблицы = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаКодыСтрок)
					{
						тегКодыСтрок = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаКодыСтолбцов)
					{
						тегКодыСтолбцов = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаНаименование)
					{
						тегНаименование = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаТег)
					{
						тегТег = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаКод)
					{
						тегКод = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == строкаТегаКодыСтрокИСтолбцов)
					{
						тегКодыСтрок = клеткаТаблицы;

						тегКодыСтолбцов = клеткаТаблицы;
					}
				}
			}
		}

		private СвободнаяЯчейка[] свободныеЯчейки;
		private Строка[] строки;
		private Столбец[] столбцы;
		private string идентификатор;
		private string код;
		private string наименование;
		private string тег;
		private string ссылкаНаМетодическийСправочник;
		private bool ручноеДобавлениеСтрок = false;

		private Range тегТипТаблицы;
		private Range тегКодыСтолбцов;
		private Range тегКодыСтрок;
		private Range тегНаименование;
		private Range тегКод;
		private Range тегТег;
		private Worksheet листКниги;

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("СвободнаяЯчейка", typeof(СвободнаяЯчейка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public СвободнаяЯчейка[] СвободныеЯчейки { get => свободныеЯчейки; set => свободныеЯчейки = value; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Строка", typeof(Строка), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Строка[] Строки { get => строки; set => строки = value; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Столбец", typeof(Столбец), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Столбец[] Столбцы { get => столбцы; set => столбцы = value; }

		[XmlAttribute()]
		public string Идентификатор { get => идентификатор; set => идентификатор = value; }

		[XmlAttribute()]
		public string Код { get => код; set => код = value; }

		[XmlAttribute()]
		public string Наименование { get => наименование; set => наименование = value; }

		[XmlAttribute()]
		public string Тег { get => тег; set => тег = value; }

		[XmlAttribute()]
		public string СсылкаНаМетодическийСправочник { get => ссылкаНаМетодическийСправочник; set => ссылкаНаМетодическийСправочник = value; }

		[XmlAttribute()]
		public bool РучноеДобавлениеСтрок { get => ручноеДобавлениеСтрок; set => ручноеДобавлениеСтрок = value; }

		[XmlIgnore]
		public Worksheet ЛистКниги { get => листКниги; set => листКниги = value; }

		[XmlIgnore]
		public bool Динамическая =>
			!((Строки != null && Строки.Length > 0) ||
			(тегТипТаблицы != null && тегТипТаблицы.Value.toString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыСтатическая)) ||
			(тегТипТаблицы != null && тегТипТаблицы.Value.toString() == ConfigManager.ТаблицаСтрокаТегаТипТаблицыДинамическая);
	}
}