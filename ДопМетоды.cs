using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public static class ДопМетоды
	{
		/// <summary>
		/// Сокращает строку до приемлемого полю тег вида
		/// </summary>
		/// <param name="наименование"></param>
		/// <returns></returns>
		public static string ПолучитьТег(string наименование)
		{
			string тег = "";

			string[] словаНаименования = наименование.Split(' ');

			int количествоСловВТеге = int.Parse(МенеджерНастроек.Настройки.Теги.КоличествоСловВТеге.Value);

			int количествоСимволовВТеге = int.Parse(МенеджерНастроек.Настройки.Теги.КоличествоСимволовВТеге.Value);

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
			string[] маскаТипаДанныхОбщий = ConfigurationManager.AppSettings["МаскаТипаДанныхОбщий"].Split('|');
			string[] маскаТипаДанныхЧисловой = ConfigurationManager.AppSettings["МаскаТипаДанныхЧисловой"].Split('|');
			string[] маскаТипаДанныхЦелочисленный = ConfigurationManager.AppSettings["МаскаТипаДанныхЦелочисленный"].Split('|');
			string[] маскаТипаДанныхФинансовый = ConfigurationManager.AppSettings["МаскаТипаДанныхФинансовый"].Split('|');
			string[] маскаТипаДанныхДатаВремя = ConfigurationManager.AppSettings["МаскаТипаДанныхДатаВремя"].Split('|');
			string[] маскаТипаДанныхСтроковый = ConfigurationManager.AppSettings["МаскаТипаДанныхСтроковый"].Split('|');

			Dictionary<string[], Type> типыДанных = new Dictionary<string[], Type>()
			{
				{ маскаТипаДанныхОбщий, typeof(Финансовый) },
				{ маскаТипаДанныхЧисловой, typeof(Числовой) },
				{ маскаТипаДанныхЦелочисленный, typeof(Целочисленный) },
				{ маскаТипаДанныхФинансовый, typeof(Финансовый) },
				{ маскаТипаДанныхДатаВремя, typeof(ДатаВремя) },
				{ маскаТипаДанныхСтроковый, typeof(Строковый) },
			};

			foreach (var типДанных in типыДанных)
			{
				if (типДанных.Key.Contains(форматКлетки))
				{
					dynamic результирующийТип = Activator.CreateInstance(типДанных.Value);

					if (типДанных.Value != typeof(Учреждение) && типДанных.Value != typeof(Логический) && типДанных.Value != typeof(ДатаВремя))
					{
						результирующийТип.ЯвляетсяКлючевым = являетсяКлючевым;
					}

					if (типДанных.Value == typeof(Числовой))
					{
						результирующийТип.Точность = форматКлетки.Split('.')[1].Length;
					}

					return результирующийТип;
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
				case Числовой числовой:
					return числовой.ToXML();
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
			string[] частоИспользуемыеТермины = ConfigurationManager.AppSettings["ЧастоИспользуемыеТермины"].Split(',');

			foreach (string термин in частоИспользуемыеТермины)
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

		public static bool КлеткаПустаИлиСодержитМетку(Range клетка)
		{
			List<string> метки = new List<string>()
			{
				МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыДинамическая.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаТипТаблицыСтатическая.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаКодыСтрок.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаКодыСтрокИСтолбцов.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаКодыСтолбцов.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаНаименование.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаТег.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаКод.Value,
				МенеджерНастроек.Настройки.Мета.МеткаНаименование.Value,
				МенеджерНастроек.Настройки.Разметка.МеткаКодыЯчеек.Value,
		};

			return
				КлеткаПуста(клетка) ||
				метки.Contains(клетка.Value.ToString());

		}

		public static string ПолучитьНаименованиеСтрокиИлиСтолбца(Range клеткаОбластиСКодами, bool ищемДляСтроки)
		{
			return КлеткаПустаИлиСодержитМетку(клеткаОбластиСКодами.Offset[ищемДляСтроки ? 0 : -1, ищемДляСтроки ? -1 : 0]) ?
			"" :
			клеткаОбластиСКодами.Offset[ищемДляСтроки ? 0 : -1, ищемДляСтроки ? -1 : 0].Value.ToString();
		}

		public static bool ПолучитьНаименованиеПоМетке(Range метка, out string наименвание)
		{
			if (метка != null)
			{
				if (!КлеткаПустаИлиСодержитМетку(метка.Offset[1, 0]))
				{
					наименвание = метка.Offset[1, 0].Value.ToString();

					return true;
				}

				if (!КлеткаПустаИлиСодержитМетку(метка.Offset[0, 1]))
				{
					наименвание = метка.Offset[0, 1].Value.ToString();

					return true;
				}
			}

			наименвание = "";

			return false;
		}
	}
}