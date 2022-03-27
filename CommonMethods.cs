using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	public static class CommonMethods
	{
		/// <summary>
		/// Сокращает строку до приемлемого полю тег вида
		/// </summary>
		/// <param name="title"></param>
		/// <returns></returns>
		public static string GetTag(string title)
		{
			string тег = "";

			string[] titleWords = title.Trim().Split(' ');

			int tagWordCount = int.Parse(SettingsManager.Settings.Tags.TagWordCount.Value);

			int tagCharCount = int.Parse(SettingsManager.Settings.Tags.TagCharCount.Value);

			int i = 1;

			foreach (string word in titleWords)
			{
				if (i > tagWordCount || тег.Length >= tagCharCount)
				{
					break;
				}

				тег += word.Substring(0, 1).ToUpper() + word.Substring(1).ToLower();

				i++;
			}

			if (тег.Length > tagCharCount)
			{
				тег = тег.Substring(0, tagCharCount);
			}

			return тег; //TODO
		}

		/// <summary>
		/// Сокращает строку до приемлемого полю тег вида
		/// </summary>
		/// <param name="наименование"></param>
		/// <returns></returns>
		public static string GetTagFromMarkup(Range cellWithCode, bool searchingForRow)
		{
			string tag = "";

			int tagCharCount = int.Parse(SettingsManager.Settings.Tags.TagCharCount.Value);

			if (!CellIsEmptyOrContainsMark(cellWithCode.Offset[searchingForRow ? 0 : 1, searchingForRow ? 1 : 0]))
			{
				tag = cellWithCode.Offset[searchingForRow ? 0 : 1, searchingForRow ? 1 : 0].Value.ToString().Replace(" ", "_");
			}

			if (tag.Length > tagCharCount)
			{
				tag = tag.Substring(0, tagCharCount);
			}

			return tag; //TODO
		}

		public static dynamic GetCellType(string cellType, bool isKey)
		{
			string[] defaultDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхОбщий"].Split('|');
			string[] nuericDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхЧисловой"].Split('|');
			string[] integerDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхЦелочисленный"].Split('|');
			string[] moneyDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхФинансовый"].Split('|');
			string[] dateTimeDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхДатаВремя"].Split('|');
			string[] textDataTypeMask = ConfigurationManager.AppSettings["МаскаТипаДанныхСтроковый"].Split('|');

			Dictionary<string[], Type> dataTypes = new Dictionary<string[], Type>()
			{
				{ defaultDataTypeMask, typeof(MoneyType) },
				{ nuericDataTypeMask, typeof(NumericType) },
				{ integerDataTypeMask, typeof(IntegerType) },
				{ moneyDataTypeMask, typeof(MoneyType) },
				{ dateTimeDataTypeMask, typeof(DateTimeType) },
				{ textDataTypeMask, typeof(TextType) },
			};

			foreach (KeyValuePair<string[], Type> dataType in dataTypes)
			{
				if (dataType.Key.Contains(cellType))
				{
					dynamic resultingType = Activator.CreateInstance(dataType.Value);

					if (dataType.Value != typeof(OrganisationType) && dataType.Value != typeof(LogicalType) && dataType.Value != typeof(DateTimeType))
					{
						resultingType.IsKey = isKey;
					}

					if (dataType.Value == typeof(NumericType))
					{
						resultingType.Precision = cellType.Split('.')[1].Length;
					}

					return resultingType;
				}
			}

			MoneyType moneyType = new MoneyType
			{
				IsKey = isKey
			};

			return moneyType;
		}

		/// <summary>
		/// Возвращает сериализованный в XML тип ячейки или столбца, соответствующий
		/// строке переданной методу аргументом
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		public static string GetSerializedType(object type)
		{
			switch (type)
			{
				case DateTimeType dateTimeType:
					return dateTimeType.ToXML();
				case LogicalType logicalType:
					return logicalType.ToXML();
				case TextType textType:
					return textType.ToXML();
				case OrganisationType organisationType:
					return organisationType.ToXML();
				case MoneyType MoneyType:
					return MoneyType.ToXML();
				case IntegerType integerType:
					return integerType.ToXML();
				case NumericType numericType:
					return numericType.ToXML();
				default:
					return "";
			}
		}

		/// <summary>
		/// Использует расстояние Дамерау-Левенштейна для приблизительного сравнения двух строк.
		/// Результат проверки так же зависит от длины строки.
		/// </summary>
		/// <param name="str1"></param>
		/// <param name="str2"></param>
		/// <returns></returns>
		public static bool StringsAreRoughlyComparable(string str1, string str2)
		{
			if (Math.Min(str1.Length, str2.Length) <= 2)
			{
				return DamerauLevenshteinDistance.DistanceBetweenStrings(str1, str2) == 0;
			}

			if (Math.Min(str1.Length, str2.Length) <= 4)
			{
				return DamerauLevenshteinDistance.DistanceBetweenStrings(str1, str2) == 1;
			}

			if (Math.Min(str1.Length, str2.Length) > 20)
			{
				return DamerauLevenshteinDistance.DistanceBetweenStrings(str1, str2) < 5;
			}

			return DamerauLevenshteinDistance.DistanceBetweenStrings(str1, str2) < 3;
		}

		/// <summary>
		/// Прверяет, не попадает ли строка в список часто используемых
		/// терминов.
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static bool StringIsCommonlyUsed(string str)
		{
			string[] commonlyUsedWords = ConfigurationManager.AppSettings["ЧастоИспользуемыеТермины"].Split(',');

			foreach (string термин in commonlyUsedWords)
			{
				if (StringsAreRoughlyComparable(str, термин))
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
		/// <param name="cell"></param>
		/// <returns></returns>
		public static bool CellIsEmpty(Range cell)
		{
			return
				cell.Value == null ||
				string.IsNullOrEmpty(cell.Value.ToString()) ||
				cell.Value.ToString() == " " ||
				cell.Value.ToString() == "  ";
		}

		public static bool CellIsEmptyOrContainsMark(Range cell)
		{
			List<string> marks = new List<string>()
			{
				SettingsManager.Settings.Markup.TableIsDynamicMark.Value,
				SettingsManager.Settings.Markup.TableIsStaticMark.Value,
				SettingsManager.Settings.Markup.RowCodesMark.Value,
				SettingsManager.Settings.Markup.RowAndColumnCodesMark.Value,
				SettingsManager.Settings.Markup.ColumnCodesMark.Value,
				SettingsManager.Settings.Markup.TitleMark.Value,
				SettingsManager.Settings.Markup.TagMark.Value,
				SettingsManager.Settings.Markup.CodeMark.Value,
				SettingsManager.Settings.Meta.TitleMark.Value,
				SettingsManager.Settings.Markup.CellCodesMark.Value,
		};

			return
				CellIsEmpty(cell) ||
				marks.Contains(cell.Value.ToString());

		}

		public static string GetRowOrColumnTitle(Range codesRangeCell, bool searchingForRow)
		{
			if (searchingForRow)
			{
				if (codesRangeCell.Column == 1)
				{
					return CellIsEmptyOrContainsMark(codesRangeCell.Offset[0, 1]) ?
						"" :
						codesRangeCell.Offset[0, 1].Value.ToString();
				}
				else
				{
					return CellIsEmptyOrContainsMark(codesRangeCell.Offset[0, -1]) ?
						"" :
						codesRangeCell.Offset[0, -1].Value.ToString();
				}
			}
			else
			{
				if (codesRangeCell.Row == 1)
				{
					return CellIsEmptyOrContainsMark(codesRangeCell.Offset[1, 0]) ?
						"" :
						codesRangeCell.Offset[1, 0].Value.ToString();
				}
				else
				{
					return CellIsEmptyOrContainsMark(codesRangeCell.Offset[-1, 0]) ?
						"" :
						codesRangeCell.Offset[-1, 0].Value.ToString();
				}
			}
		}

		public static bool GetTitleFromMarkedCell(Range markedCell, out string title)
		{
			if (markedCell != null)
			{
				if (!CellIsEmptyOrContainsMark(markedCell.Offset[1, 0]))
				{
					title = markedCell.Offset[1, 0].Value.ToString();

					return true;
				}

				if (!CellIsEmptyOrContainsMark(markedCell.Offset[0, 1]))
				{
					title = markedCell.Offset[0, 1].Value.ToString();

					return true;
				}
			}

			title = "";

			return false;
		}

		/// <summary>
		/// Заменяет все запрещенные символы в строке на указанную строку
		/// </summary>
		/// <param name="path"></param>
		/// <param name="replacementString"></param>
		/// <returns></returns>
		public static string RemoveForbiddenSymbols(string path, string replacementString, bool removePunctuation = false)
		{
			string forbiddenSymbols = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

			if (removePunctuation)
			{
				forbiddenSymbols += ",.-;:\"'?!";
			}

			foreach (char запрещенныйСимвол in forbiddenSymbols)
			{
				path = path.Replace(запрещенныйСимвол.ToString(), replacementString);
			}

			return path;
		}
	}
}