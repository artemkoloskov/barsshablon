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

			int порядковыйНомерЛиста = 1;

			foreach (Worksheet листКниги in листыКниги)
			{
				Таблица таблица = new Таблица(листКниги, порядковыйНомерЛиста);

				порядковыйНомерЛиста++;

				if (таблица != null)
				{
					таблицы.Add(таблица);
				}
			}

			return таблицы;
		}

		public Таблица(Worksheet листКниги, int порядковыйНомерЛиста)
		{
			ЛистКниги = листКниги;

			НайтиМеткиНаЛисте();

			Строки = ПолучитьСтрокиТаблицы();

			Столбцы = ПолучитьСтолбцыТаблицы();

			Наименование = ПолучитьНаименованиеТаблицы(порядковыйНомерЛиста);

			Идентификатор = "Таблица" + порядковыйНомерЛиста;

			Код = "Таблица" + порядковыйНомерЛиста;

			Тег = МенеджерНастроек.Настройки.Теги.ПрефиксТаблицы.Value + ДопМетоды.ПолучитьТег(Наименование);
		}

		private string ПолучитьНаименованиеТаблицы(int порядковыйНомерЛиста)
		{
			if (меткаНаименование != null)
			{
				if (ДопМетоды.ПолучитьНаименованиеПоМетке(меткаНаименование, out string наименование))
				{
					return наименование;
				}
			}

			return "Таблица" + порядковыйНомерЛиста;
		}

		private Строка[] ПолучитьСтрокиТаблицы()
		{
			if (меткаКодыСтрок != null)
			{
				List<Строка> списокСтрок = new List<Строка>();

				foreach (Range клеткаСтолбцаСКодамиСтрок in ЛистКниги.Application.Intersect(меткаКодыСтрок.EntireColumn, ЛистКниги.UsedRange).Cells)
				{
					if (!ДопМетоды.КлеткаПустаИлиСодержитМетку(клеткаСтолбцаСКодамиСтрок) &&
						клеткаСтолбцаСКодамиСтрок.Row > меткаКодыСтрок.Row)
					{
						списокСтрок.Add(new Строка(клеткаСтолбцаСКодамиСтрок));
					}
				}

				return списокСтрок.ToArray();
			}

			return null;
		}

		private Столбец[] ПолучитьСтолбцыТаблицы()
		{
			if (меткаСтолбцов != null)
			{
				List<Столбец> списокСтолбцов = new List<Столбец>();

				foreach (Range клеткаСтрокиСКодамиСтолбцов in ЛистКниги.Application.Intersect(меткаСтолбцов.EntireRow, ЛистКниги.UsedRange).Cells)
				{
					if (!ДопМетоды.КлеткаПустаИлиСодержитМетку(клеткаСтрокиСКодамиСтолбцов) &&
						клеткаСтрокиСКодамиСтолбцов.Column > меткаСтолбцов.Column)
					{
						списокСтолбцов.Add(new Столбец(клеткаСтрокиСКодамиСтолбцов, Динамическая));
					}
				}

				return списокСтолбцов.ToArray();
			}

			return null;
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и запоминает ячейки с метками
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private void НайтиМеткиНаЛисте()
		{
			foreach (Range клеткаТаблицы in ЛистКниги.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыДинамическая.Value ||
								клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыСтатическая.Value)
					{
						меткаТипТаблицы = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаКодыСтрок.Value)
					{
						меткаКодыСтрок = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаКодыСтолбцов.Value)
					{
						меткаСтолбцов = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаНаименование.Value)
					{
						меткаНаименование = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаТег.Value)
					{
						меткаТег = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаКод.Value)
					{
						меткаКод = клеткаТаблицы;
					}

					if (клеткаТаблицы.Value.ToString() == МенеджерНастроек.Настройки.Разметка.МеткаКодыСтрокИСтолбцов.Value)
					{
						меткаКодыСтрок = клеткаТаблицы;

						меткаСтолбцов = клеткаТаблицы;
					}
				}
			}
		}

		private Range меткаТипТаблицы;
		private Range меткаСтолбцов;
		private Range меткаКодыСтрок;
		private Range меткаНаименование;
		private Range меткаКод;
		private Range меткаТег;

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
			(меткаТипТаблицы != null && 
			меткаТипТаблицы.Value.toString() == МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыСтатическая.Value)) ||
			(меткаТипТаблицы != null && 
			меткаТипТаблицы.Value.toString() == МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыДинамическая.Value);

		[XmlIgnore]
		public Range МеткаТипТаблицы { get => меткаТипТаблицы; set => меткаТипТаблицы = value; }
		[XmlIgnore]
		public Range МеткаКодыСтолбцов { get => меткаСтолбцов; set => меткаСтолбцов = value; }
		[XmlIgnore]
		public Range МеткаКодыСтрок { get => меткаКодыСтрок; set => меткаКодыСтрок = value; }
		[XmlIgnore]
		public Range МеткаНаименование { get => меткаНаименование; set => меткаНаименование = value; }
		[XmlIgnore]
		public Range МеткаКод { get => меткаКод; set => меткаКод = value; }
		[XmlIgnore]
		public Range МеткаТег { get => меткаТег; set => меткаТег = value; }
	}
}