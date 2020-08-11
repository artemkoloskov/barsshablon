using System.Xml.Serialization;
using System.Xml.Schema;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using System;
using БАРСШаблон.DataTypes;
using System.Collections.Generic;

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

		public Таблица (Worksheet листКниги)
		{
			ЛистКниги = листКниги;

			НайтиТэгиНаЛисте();

			строки = ПолучитьСтрокиТаблицы();

			столбцы = ПолучитьСтолбцыТаблицы();
			
			идентификатор = "Таблица1";
			код = "Тбл1";
			наименование = "Крутая ваще таблица";
			ручноеДобавлениеСтрок = false;
			тег = "КртВщТабла";
			свободныеЯчейки = new СвободнаяЯчейка[] { new СвободнаяЯчейка("Суки", typeof(Целочисленный).ToString().Split('.')[2]) };
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
						списокСтолбцов.Add(new Столбец(клеткаСтрокиСКодамиСтолбцов));
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
			string строкаТегаТипТаблицыДинамическая = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыДинамическая");
			string строкаТегаТипТаблицыСтатическая = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыСтатическая");
			string строкаТегаКодыСтрок = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрок");
			string строкаТегаКодыСтрокИСтолбцов = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрокИСтолбцов");
			string строкаТегаКодыСтолбцов = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтолбцов");
			string строкаТегаНазвание = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаНазвание");
			string строкаТегаТег = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТег");
			string строкаТегаКод = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКод");

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

					if (клеткаТаблицы.Value.ToString() == строкаТегаНазвание)
					{
						тегНазвание = клеткаТаблицы;
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
		private Range тегНазвание;
		private Range тегКод;
		private Range тегТег;

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
		public Worksheet ЛистКниги { get; set; }
	}
}