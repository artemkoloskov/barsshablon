using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public static class ДопМетоды
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

			int количествоСловВТеге = ConfigManager.КоличествоСловВТеге;
			int количествоСимволовВТеге = ConfigManager.КоличествоСимволовВТеге;

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

			if (тег.Length > количествоСимволовВТеге)
			{
				тег = тег.Substring(0, количествоСимволовВТеге);
			}

			return тег; //TODO
		}

		public static dynamic ПолучитьТип(string форматКлетки, bool являетсяКлючевым)
		{
			foreach (var типДанных in ConfigManager.ТипыДанных)
			{
				if (типДанных.Key.Contains(форматКлетки))
				{
					dynamic ass = Activator.CreateInstance(типДанных.Value);

					ass.ЯвляетсяКлючевым = являетсяКлючевым;

					if (типДанных.Value == typeof(Числовой))
					{
						ass.Точность = форматКлетки.Split('.').Length;
					}

					return ass;
				}
			}

			Финансовый финансовыйТип = new Финансовый
			{
				ЯвляетсяКлючевым = являетсяКлючевым
			};

			return финансовыйТип;
		}

		/// <summary>
		/// Возвращает сериализованный в XML тип ячейки или столбца, соответствующий
		/// строке переданной методу аргументом
		/// </summary>
		/// <param name="тип"></param>
		/// <returns></returns>
		public static string ПолучитьСриализованныйТип(object тип)
		{
			switch (тип)
			{
				case ДатаВремя датаВремя:
					return датаВремя.ToXML();
				case Логический логический:
					return логический.ToXML();
				case Строковый строковый:
					return строковый.ToXML();
				case Учреждение учреждение:
					return учреждение.ToXML();
				case Финансовый финансовый:
					return финансовый.ToXML();
				case Целочисленный целочисленный:
					return целочисленный.ToXML();
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
			foreach (string термин in ConfigManager.ЧастоИспользуемыеТермины)
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
				ConfigManager.ТаблицаСтрокаТегаТипТаблицыДинамическая,
				ConfigManager.ТаблицаСтрокаТегаТипТаблицыСтатическая,
				ConfigManager.ТаблицаСтрокаТегаКодыСтрок,
				ConfigManager.ТаблицаСтрокаТегаКодыСтрокИСтолбцов,
				ConfigManager.ТаблицаСтрокаТегаКодыСтолбцов,
				ConfigManager.ТаблицаСтрокаТегаНаименование,
				ConfigManager.ТаблицаСтрокаТегаТег,
				ConfigManager.ТаблицаСтрокаТегаКод,
				ConfigManager.МетаТегНаименование,
				ConfigManager.СвободнаяЯчейкаСтрокаТегаКодыЯчеек,
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