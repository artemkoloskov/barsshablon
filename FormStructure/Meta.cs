using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Linq;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Meta
	{
		public Meta()
		{
		}

		public Meta(Workbook workbook)
		{
			Workbook = workbook;

			ScanForMarkup();

			Title = GetTitle();

			Id += CommonMethods.RemoveForbiddenSymbols(
				$"{(SettingsManager.Settings.Meta.IsARequest.Value ? "З_" : "М_")}{CommonMethods.GetTag(Title)}",
				"_",
				removePunctuation: true);

			Group += DateTime.Today.Year;

			DateFrom = DateFrom.Replace("0001", DateTime.Today.Year.ToString());

			DateTo = DateTo.Replace("9999", DateTime.Today.Year.ToString());

			LastEditDate = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

			Tag = Id;
		}

		private void ScanForMarkup()
		{
			string titleMark = SettingsManager.Settings.Meta.TitleMark.Value;

			foreach (Range cell in Workbook.Worksheets[1].UsedRange.Cells)
			{
				if (cell.Value != null)
				{
					if (cell.Value.ToString() == titleMark)
					{
						TitleMark = cell;
					}
				}
			}
		}

		private string GetTitle()
		{
			Dictionary<string, double> possibleTitles = new Dictionary<string, double>();

			Range usedRange = Workbook.Worksheets[1].UsedRange;

			if (!TitleIsDefined(out string title))
			{
				foreach (Range column in usedRange.Columns)
				{
					Range topCellInColumn = FindUpperMostNotEmptyCellInColumn(column);

					if (topCellInColumn != null && !possibleTitles.ContainsKey(topCellInColumn.Value.ToString()))
					{
						possibleTitles.Add(topCellInColumn.Value.ToString(), GetProbabilityCellContainsTitle(topCellInColumn));
					}
				}

				KeyValuePair<string, double> mostProbableTitle = new KeyValuePair<string, double>("", 0);

				foreach (KeyValuePair<string, double> possibleTitle in possibleTitles.Where(possibleTitle => possibleTitle.Value > mostProbableTitle.Value))
				{
					mostProbableTitle = possibleTitle;
				}

				title = mostProbableTitle.Key;
			}

			return title.Length > 240 ? title.Substring(0, 239) : title;
		}

		private bool TitleIsDefined(out string title)
		{
			return CommonMethods.GetTitleFromMarkedCell(TitleMark, out title);
		}

		private Range FindUpperMostNotEmptyCellInColumn(Range column)
		{
			foreach (Range cell in column.Cells)
			{
				if (cell.Value != null && cell.Value.ToString() != "" && cell.Value.ToString() != " ")
				{
					return cell;
				}
			}

			return null;
		}

		private double GetProbabilityCellContainsTitle(Range cell)
		{
			double probability = 0;

			probability +=
				cell.Value.ToString().Length * SettingsManager.Settings.Weight.Length.Value;

			probability +=
				cell.Row * SettingsManager.Settings.Weight.RowNumber.Value;

			probability +=
				cell.Column * SettingsManager.Settings.Weight.ColumnNumber.Value;

			probability +=
				GetNumberOfMergedCells(cell) * SettingsManager.Settings.Weight.NumberOfCellsInMergedCell.Value;

			probability +=
				cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				SettingsManager.Settings.Weight.BottomBorder.Value;

			probability +=
				cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				SettingsManager.Settings.Weight.TopBorder.Value;

			probability +=
				cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				SettingsManager.Settings.Weight.LeftBorder.Value;

			probability +=
				cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				SettingsManager.Settings.Weight.RightBorder.Value;

			probability +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignCenter ?
				0 :
				SettingsManager.Settings.Weight.CenterAligned.Value;

			probability +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignLeft ?
				0 :
				SettingsManager.Settings.Weight.LeftAligned.Value;

			probability +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignRight ?
				0 :
				SettingsManager.Settings.Weight.RightAligned.Value;

			probability +=
				cell.Font.Bold ?
				SettingsManager.Settings.Weight.BoldText.Value :
				0;

			probability +=
				GetNumberOfEmptyCellsBelow(cell) * SettingsManager.Settings.Weight.EmptyRowBelowCell.Value;

			probability +=
				CommonMethods.StringIsCommonlyUsed(cell.Value.ToString()) ?
				SettingsManager.Settings.Weight.CommonlyUsedWord.Value :
				0;

			return probability;
		}

		private int GetNumberOfMergedCells(Range cell)
		{
			if (cell.MergeCells)
			{
				return cell.MergeArea.Cells.Count;
			}

			return 0;
		}

		private int GetNumberOfEmptyCellsBelow(Range cell)
		{
			int number = 0;

			do
			{
				number++;
			}
			while
				(number < 10 && (cell.Offset[number, 0].Value == null ||
				cell.Offset[number, 0].Value.ToString() == "" ||
				cell.Offset[number, 0].Value.ToString() == " "));

			return number;
		}

		[XmlElement("ВерсияМетаописания", Form = XmlSchemaForm.Unqualified)]
		public string MetaDescriptionVersion { get; set; } = SettingsManager.Settings.Meta.MetaVersion.Value;

		[XmlElement("Идентификатор", Form = XmlSchemaForm.Unqualified)]
		public string Id { get; set; } = SettingsManager.Settings.Meta.Id.Value;

		[XmlElement("Наименование", Form = XmlSchemaForm.Unqualified)]
		public string Title { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Group { get; set; } = SettingsManager.Settings.Meta.Group.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string DateFrom { get; set; } = SettingsManager.Settings.Meta.BeginDate.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string DateTo { get; set; } = SettingsManager.Settings.Meta.EndDate.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Author { get; set; } = SettingsManager.Settings.Meta.Author.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string LastEditDate { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string VersionNumber { get; set; } = SettingsManager.Settings.Meta.VersionNumber.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string HeaderLocation { get; set; } = SettingsManager.Settings.Meta.HeaderPlacement.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Host { get; set; } = Environment.MachineName;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string LinkToDictionary { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string LinkToExternalDictionary { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string MetaDescriptionFormatVersion { get; set; } = SettingsManager.Settings.Meta.MetaFormatVersion.Value;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Tag { get; set; } = "";

		[XmlIgnore]
		public Workbook Workbook { get; set; }

		[XmlIgnore]
		public Range TitleMark { get; set; }
	}
}