using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class FreeCell
	{
		public FreeCell()
		{
		}

		public FreeCell(string cellCode, object cellType)
		{
			Id = cellCode;

			Code = cellCode;

			CellType = cellType;

			Type = CellType.GetType().Name;

			Tag = SettingsManager.Settings.Tags.FreeCellPrefix.Value + CommonMethods.GetTag(Id);

			Description = CommonMethods.GetSerializedType(CellType);
		}

		public static List<FreeCell> GetFreeCells(Worksheet worksheet)
		{
			List<FreeCell> freeCells = GetDefaultFreeCells();

			Range cellCodesMark = GetMark(worksheet);

			if (cellCodesMark != null)
			{
				Range cellCodesColumn = worksheet.Application.Intersect(cellCodesMark.EntireColumn, worksheet.UsedRange);

				if (cellCodesMark != null)
				{
					foreach (Range cell in cellCodesColumn)
					{
						if (cell.Row > cellCodesMark.Row)
						{
							if (!CommonMethods.CellIsEmptyOrContainsMark(cell))
							{
								freeCells.Add(new FreeCell(cell.Value.ToString(), GetCellType(cell)));
							}

							if (CommonMethods.CellIsEmptyOrContainsMark(cell.Offset[1, 0]))
							{
								break;
							}
						}
					}
				}
			}

			return freeCells;
		}

		private static List<FreeCell> GetDefaultFreeCells()
		{
			List<FreeCell> defaultFreeCells = new List<FreeCell>
			{
				new FreeCell("Учреждение", new OrganisationType()),
				new FreeCell("Должность", new TextType() { IsKey = true }),
				new FreeCell("Ответственный", new TextType() { IsKey = true, Mandatory = true }),
				new FreeCell("Телефон", new TextType() { IsKey = true, Mandatory = true })
			};

			return defaultFreeCells;
		}

		private static object GetCellType(Range cell)
		{
			return CommonMethods.GetCellType(cell.Offset[0, 1].NumberFormat, false);
		}

		/// <summary>
		/// Просматривает все используемые клетки листа и возвращает ячейку с меткой
		/// </summary>
		/// <param name="ЛистКниги"></param>
		private static Range GetMark(Worksheet worksheet)
		{
			string cellCodesMark = SettingsManager.Settings.Markup.CellCodesMark.Value;

			foreach (Range cell in worksheet.UsedRange.Cells)
			{
				if (cell.Value != null)
				{
					if (cell.Value.ToString() == cellCodesMark)
					{
						return cell;
					}
				}
			}

			return null;
		}

		[XmlAttribute("Идентификатор")]
		public string Id { get; set; }

		[XmlAttribute("Код")]
		public string Code { get; set; }

		[XmlAttribute("НаименованиеЭлемента")]
		public string ElementName { get; set; }

		[XmlAttribute("Тип")]
		public string Type { get; set; }

		[XmlAttribute("Описание")]
		public string Description { get; set; }

		[XmlAttribute("Тег")]
		public string Tag { get; set; }

		[XmlIgnore()]
		public object CellType { get; }
	}
}
