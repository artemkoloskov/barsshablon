using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public class ДопМетоды
	{
		/// <summary>
		/// Сокращает строку то приемлемого полю тег вида
		/// </summary>
		/// <param name="наименование"></param>
		/// <returns></returns>
		public static string ПолучитьТег(string наименование)
		{
			string тег = "";

			string[] словаНаименования = наименование.Split(' ');

			int количествоСловВТеге = int.Parse(ConfigurationManager.AppSettings.Get("КоличествоСловВТеге"));
			int количествоСимволовВТеге = int.Parse(ConfigurationManager.AppSettings.Get("КоличествоСимволовВТеге"));

			int i = 1;

			foreach (string слово in словаНаименования)
			{
				if (i > количествоСловВТеге || тег.Length >= количествоСимволовВТеге)
				{
					break;
				}

				тег += слово.Substring(0, 1).ToUpper() + слово.Substring(1).ToLower();

				i++;
			}

			if(тег.Length > количествоСимволовВТеге)
			{
				тег = тег.Substring(0, количествоСимволовВТеге);
			}

			return тег; //TODO
		}

		/// <summary>
		/// Возвращает сериализованный в XML тип ячейки или столбца, соответствующий
		/// строке переданной методу аргументом
		/// </summary>
		/// <param name="тип"></param>
		/// <returns></returns>
		public static string ПолучитьСриализованныйТип(string тип)
		{
			switch (тип)
			{
				case "ДатаВремя":
					return new ДатаВремя().ToXML();
				case "Логический":
					return new Логический().ToXML();
				case "Строковый":
					return new Строковый().ToXML();
				case "Учреждение":
					return new Учреждение().ToXML();
				case "Финансовый":
					return new Финансовый().ToXML();
				case "Целочисленный":
					return new Целочисленный().ToXML();
				default:
					return "";
			}
		}

		/// <summary>
		/// Использует расстояние Дамерау-Левенштейна для приблизительного сравнения двух строк.
		/// Результат проверки так же зависит от длины строки.
		/// </summary>
		/// <param name="строка1"></param>
		/// <param name="строка2"></param>
		/// <returns></returns>
		public static bool СтрокиПриблизительноСовпадают(string строка1, string строка2)
		{
			if (Math.Min(строка1.Length, строка2.Length) <= 2)
			{
				return DamerauLevenshteinDistance.РасстояниеМеждуСтроками(строка1, строка2) == 0;
			}

			if (Math.Min(строка1.Length, строка2.Length) <= 4)
			{
				return DamerauLevenshteinDistance.РасстояниеМеждуСтроками(строка1, строка2) == 1;
			}

			if (Math.Min(строка1.Length, строка2.Length) > 20)
			{
				return DamerauLevenshteinDistance.РасстояниеМеждуСтроками(строка1, строка2) < 5;
			}

			return DamerauLevenshteinDistance.РасстояниеМеждуСтроками(строка1, строка2) < 3;
		}

		/// <summary>
		/// Прверяет, не попадает ли строка в список часто используемых
		/// терминов.
		/// </summary>
		/// <param name="строка"></param>
		/// <returns></returns>
		public static bool СтрокаЯвлетсяЧастоИспользуемой(string строка)
		{
			string[] частоИсспользуемыТермины = ConfigurationManager.AppSettings.Get("ЧастоИспользуемыеТермины").Split(',');

			foreach (string термин in частоИсспользуемыТермины)
			{
				if (СтрокиПриблизительноСовпадают(строка, термин))
				{
					return true;
				}
			}

			return false;
		}

		/// <summary>
		/// Определяет наличие содержимого в клетке таблицы Excel.
		/// Пробел и два пробела не считаются содержимым.
		/// </summary>
		/// <param name="клетка"></param>
		/// <returns></returns>
		public static bool КлеткаПуста(Range клетка)
		{
			return
				клетка.Value == null ||
				string.IsNullOrEmpty(клетка.Value.ToString()) ||
				клетка.Value.ToString() == " " ||
				клетка.Value.ToString() == "  ";
		}

		public static bool КлеткаПустаИлиСодержитТег(Range клетка)
		{
			List<string> теги = new List<string>()
			{
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыДинамическая"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТипТаблицыСтатическая"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрок"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтрокИСтолбцов"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКодыСтолбцов"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаНаименование"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТег"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКод"),
				ConfigurationManager.AppSettings.Get("МетаТегНаименование"),
			};

			return
				КлеткаПуста(клетка) ||
				теги.Contains(клетка.Value.ToString());

		}

		public static string ПолучитьНаименованиеСтрокиИлиСтолбца(Range клеткаОбластиСКодами, bool ищемДляСтроки)
		{
			return КлеткаПустаИлиСодержитТег(клеткаОбластиСКодами.Offset[ищемДляСтроки ? 0 : -1, ищемДляСтроки ? -1 : 0]) ?
			"" :
			клеткаОбластиСКодами.Offset[ищемДляСтроки ? 0 : -1, ищемДляСтроки ? -1 : 0].Value.ToString();
		}

		public static bool ПолучитьНаименованиеПоТегу(Range тег, out string наименвание)
		{
			if (тег != null)
			{
				if (!КлеткаПустаИлиСодержитТег(тег.Offset[1, 0]))
				{
					наименвание = тег.Offset[1, 0].Value.ToString();

					return true;
				}

				if (!КлеткаПустаИлиСодержитТег(тег.Offset[0, 1]))
				{
					наименвание = тег.Offset[0, 1].Value.ToString();

					return true;
				} 
			}

			наименвание = "";

			return false;
		}
	}
}