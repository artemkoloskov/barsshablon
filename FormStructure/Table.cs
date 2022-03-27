using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Table
	{
		public Table()
		{
		}

		public static List<Table> GetFormTables(Sheets sheets)
		{
			List<Table> tables = new List<Table>();

			int sheetIndex = 1;

			foreach (Worksheet worksheet in sheets)
			{
				Table table = new Table(worksheet, sheetIndex);

				sheetIndex++;

				if (table != null)
				{
					tables.Add(table);
				}
			}

			return tables;
		}

		public Table(Worksheet worksheet, int sheetIndex)
		{
			Worksheet = worksheet;

			ScanSheetForMarkUp();

			Rows = GetTableRows();

			Columns = GetTableColumns();

			Title = GetTableTitle(sheetIndex);

			Id = "Таблица" + sheetIndex;

			Code = "Таблица" + sheetIndex;

			Tag = SettingsManager.Settings.Tags.TablePrefix.Value + CommonMethods.GetTag(Title);
		}

		private string GetTableTitle(int sheetIndex)
		{
			if (TitleMark != null)
			{
				if (CommonMethods.GetTitleFromMarkedCell(TitleMark, out string title))
				{
					return title;
				}
			}

			return "Таблица" + sheetIndex;
		}

		private Row[] GetTableRows()
		{
			if (RowCodesMark != null)
			{
				List<Row> rows = new List<Row>();

				foreach (Range rowCodeCell in Worksheet.Application.Intersect(RowCodesMark.EntireColumn, Worksheet.UsedRange).Cells)
				{
					if (!CommonMethods.CellIsEmptyOrContainsMark(rowCodeCell) &&
						rowCodeCell.Row > RowCodesMark.Row)
					{
						rows.Add(new Row(rowCodeCell));
					}
				}

				return rows.ToArray();
			}

			return null;
		}

		private Column[] GetTableColumns()
		{
			if (ColumnCodesMark != null)
			{
				List<Column> columns = new List<Column>();

				foreach (Range columnCodeCell in Worksheet.Application.Intersect(ColumnCodesMark.EntireRow, Worksheet.UsedRange).Cells)
				{
					if (!CommonMethods.CellIsEmptyOrContainsMark(columnCodeCell) &&
						columnCodeCell.Column > ColumnCodesMark.Column)
					{
						columns.Add(new Column(columnCodeCell, Dynamic));
					}
				}

				return columns.ToArray();
			}

			return null;
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и запоминает ячейки с метками
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private void ScanSheetForMarkUp()
		{
			foreach (Range cell in Worksheet.UsedRange.Cells)
			{
				if (cell.Value != null)
				{
					if (cell.Value.ToString() == SettingsManager.Settings.Markup.TableIsDynamicMark.Value ||
								cell.Value.ToString() == SettingsManager.Settings.Markup.TableIsStaticMark.Value)
					{
						TableTypeMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.RowCodesMark.Value)
					{
						RowCodesMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.ColumnCodesMark.Value)
					{
						ColumnCodesMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.TitleMark.Value)
					{
						TitleMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.TagMark.Value)
					{
						TagMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.CodeMark.Value)
					{
						CodeMark = cell;
					}

					if (cell.Value.ToString() == SettingsManager.Settings.Markup.RowAndColumnCodesMark.Value)
					{
						RowCodesMark = cell;

						ColumnCodesMark = cell;
					}
				}
			}
		}

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("СвободнаяЯчейка", typeof(FreeCell), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public FreeCell[] СвободныеЯчейки { get; set; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Строка", typeof(Row), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Row[] Rows { get; set; }

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Столбец", typeof(Column), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Column[] Columns { get; set; }

		[XmlAttribute()]
		public string Id { get; set; }

		[XmlAttribute()]
		public string Code { get; set; }

		[XmlAttribute()]
		public string Title { get; set; }

		[XmlAttribute()]
		public string Tag { get; set; }

		[XmlAttribute()]
		public string СсылкаНаМетодическийСправочник { get; set; }

		[XmlAttribute()]
		public bool РучноеДобавлениеСтрок { get; set; } = false;

		[XmlIgnore]
		public Worksheet Worksheet { get; set; }

		[XmlIgnore]
		public bool Dynamic =>
			!((Rows != null && Rows.Length > 0) ||
			(TableTypeMark != null &&
			TableTypeMark.Value.toString() == SettingsManager.Settings.Markup.TableIsStaticMark.Value)) ||
			(TableTypeMark != null &&
			TableTypeMark.Value.toString() == SettingsManager.Settings.Markup.TableIsDynamicMark.Value);

		[XmlIgnore]
		public Range TableTypeMark { get; set; }
		[XmlIgnore]
		public Range ColumnCodesMark { get; set; }
		[XmlIgnore]
		public Range RowCodesMark { get; set; }
		[XmlIgnore]
		public Range TitleMark { get; set; }
		[XmlIgnore]
		public Range CodeMark { get; set; }
		[XmlIgnore]
		public Range TagMark { get; set; }
	}
}