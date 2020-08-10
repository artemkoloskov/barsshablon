using System.Xml.Serialization;
using System.Xml.Schema;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using System;
using БАРСШаблон.DataTypes;

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
			тег = CommonMethods.ПолучитьТег(идентификатор);
		}

		public Таблица (Worksheet листКниги)
		{
			НайтиТэгиНаЛисте(листКниги);

			Столбец столбец1 = new Столбец("1", typeof(Целочисленный).ToString().Split('.')[2]);
			Столбец столбец2 = new Столбец("2", typeof(Финансовый).ToString().Split('.')[2]);

			Строка строка1 = new Строка() { Идентификатор = "001", Код = "001", НаименованиеЭлемента = "Охуеть", Тег = "Охт" };
			Строка строка2 = new Строка() { Идентификатор = "002", Код = "002", НаименованиеЭлемента = "Заебись", Тег = "Збс" };

			идентификатор = "Таблица1";
			код = "Тбл1";
			наименование = "Крутая ваще таблица";
			ручноеДобавлениеСтрок = false;
			тег = "КртВщТабла";
			столбцы = new Столбец[] { столбец1, столбец2 };
			строки = new Строка[] { строка1, строка2 };
			свободныеЯчейки = new СвободнаяЯчейка[] { new СвободнаяЯчейка("Суки", typeof(Целочисленный).ToString().Split('.')[2]) };
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и запоминает ячейки с тегами
		/// </summary>
		/// <param name="листКниги"></param>
		private void НайтиТэгиНаЛисте(Worksheet листКниги)
		{
			string строкаТегаТипТаблицыДинамическая = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыДинамическая");
			string строкаТегаТипТаблицыСтатическая = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыСтатическая");
			string строкаТегаКодыСтрок = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрок");
			string строкаТегаКодыСтрокИСтолбцов = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрокИСтолбцов");
			string строкаТегаКодыСтолбцов = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтолбцов");
			string строкаТегаНазвание = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаНазвание");
			string строкаТегаТег = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТег");
			string строкаТегаКод = ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКод");

			foreach (Range клеткаТаблицы in листКниги.UsedRange.Cells)
			{
				if (клеткаТаблицы.Value == строкаТегаТипТаблицыДинамическая ||
					клеткаТаблицы.Value == строкаТегаТипТаблицыСтатическая)
				{
					тегТипТаблицы = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаКодыСтрок)
				{
					тегКодыСтрок = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаКодыСтолбцов)
				{
					тегКодыСтолбцов = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаНазвание)
				{
					тегНазвание = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаТег)
				{
					тегТег = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаКод)
				{
					тегКод = клеткаТаблицы;
				}

				if (клеткаТаблицы.Value == строкаТегаКодыСтрокИСтолбцов)
				{
					тегКодыСтрокИСтолбцов = клеткаТаблицы;
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
		private Range тегКодыСтрокИСтолбцов;
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
	}
}