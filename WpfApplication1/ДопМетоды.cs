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
		/// <param name="идентификатор"></param>
		/// <returns></returns>
		public static string ПолучитьТег(string идентификатор)
		{
			return идентификатор; //TODO
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
		public static bool КлеткаПуста (Range клетка)
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
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаНазвание"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаТег"),
				ConfigurationManager.AppSettings.Get("ТаблицаСтрокаТегаКод"),
			};

			return
				КлеткаПуста(клетка) ||
				теги.Contains(клетка.Value.ToString());

		}
	}
}
