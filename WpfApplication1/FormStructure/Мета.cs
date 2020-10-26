using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Мета
	{
		public Мета()
		{
		}

		public Мета(Workbook книгаExcel)
		{
			КнигаExcel = книгаExcel;

			НаитиМеткиВКниге();

			Наименование = ПолучитьНаименование();

			Идентификатор += $"{(ConfigManager.МетаЯвляетсяЗапросом ? "З_" : "М_")}{ДопМетоды.ПолучитьТег(Наименование)}";

			Группа += DateTime.Today.Year;

			ДатаНачалаДействия = ДатаНачалаДействия.Replace("0001", DateTime.Today.Year.ToString());

			ДатаОкончанияДействия = ДатаОкончанияДействия.Replace("9999", DateTime.Today.Year.ToString());

			ДатаПоследнегоИзменения = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

			Тег = Идентификатор;
		}

		private void НаитиМеткиВКниге()
		{
			string меткаНаименование = ConfigManager.МетаМеткаНаименование;

			foreach (Range клеткаТаблицы in КнигаExcel.Worksheets[1].UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == меткаНаименование)
					{
						МеткаНаименование = клеткаТаблицы;
					}
				}
			}
		}

		private string ПолучитьНаименование()
		{
			Dictionary<string, double> возможныеНаименования = new Dictionary<string, double>();

			Range usedRange = КнигаExcel.Worksheets[1].UsedRange;

			if (!НаименованиеУказаноВШаблоне(out string наименование))
			{
				foreach (Range column in usedRange.Columns)
				{
					Range topCellInColumn = НайтиВКолонкеВерхнююНеПустуюЯчейку(column);

					if (topCellInColumn != null && !возможныеНаименования.ContainsKey(topCellInColumn.Value.ToString()))
					{
						возможныеНаименования.Add(topCellInColumn.Value.ToString(), ПолучитьВероятностьНаименованиеВЯчейке(topCellInColumn));
					}
				}

				KeyValuePair<string, double> наиболееВероятноеНаименование = new KeyValuePair<string, double>("", 0);

				foreach (var возможноеНаименование in возможныеНаименования)
				{
					if (возможноеНаименование.Value > наиболееВероятноеНаименование.Value)
					{
						наиболееВероятноеНаименование = возможноеНаименование;
					}
				}

				наименование = наиболееВероятноеНаименование.Key;
			}

			return наименование.Length > 240 ? наименование.Substring(0, 239) : наименование;
		}

		private bool НаименованиеУказаноВШаблоне(out string наименование)
		{
			return ДопМетоды.ПолучитьНаименованиеПоМетке(МеткаНаименование, out наименование);
		}

		private Range НайтиВКолонкеВерхнююНеПустуюЯчейку(Range column)
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

		private double ПолучитьВероятностьНаименованиеВЯчейке(Range cell)
		{
			double вероятность = 0;

			вероятность +=
				cell.Value.ToString().Length * ConfigManager.МетаВесДлиныПотенциальногоНаименования;

			вероятность +=
				cell.Row * ConfigManager.МетаВесНомераСтрокиПотенциальногоНаименования;

			вероятность +=
				cell.Column * ConfigManager.МетаВесНомераСтолбцаПотенциальногоНаименования;

			вероятность +=
				ПолучитьКоличествоЯчеекВОбъединении(cell) * ConfigManager.МетаВесКоличестваЯчеекВОбъединеннойЯчейкеПотенциальногоНаименования;

			вероятность +=
				cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				ConfigManager.МетаВесГраницыВнизуПотенциальногоНаименования;

			вероятность +=
				cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				ConfigManager.МетаВесГраницыВверхуПотенциальногоНаименования;

			вероятность +=
				cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				ConfigManager.МетаВесГраницыСлеваПотенциальногоНаименования;

			вероятность +=
				cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle == (int)XlLineStyle.xlLineStyleNone ?
				0 :
				ConfigManager.МетаВесГраницыСправаПотенциальногоНаименования;

			вероятность +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignCenter ?
				0 :
				ConfigManager.МетаВесВыравниванияПоСерединеПотенциальногоНаименования;

			вероятность +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignLeft ?
				0 :
				ConfigManager.МетаВесВыравниванияСлеваПотенциальногоНаименования;

			вероятность +=
				cell.HorizontalAlignment == (int)XlHAlign.xlHAlignRight ?
				0 :
				ConfigManager.МетаВесВыравниванияСправаПотенциальногоНаименования;

			вероятность +=
				cell.Font.Bold ?
				ConfigManager.МетаВесЖирностиТекстаПотенциальногоНаименования :
				0;

			вероятность +=
				ПолучитьКоличествоПустыхСтрокПодЯчейкой(cell) * ConfigManager.МетаВесПустойСтрокиПодЯчейкойПотенциальногоНаименования;

			вероятность +=
				ДопМетоды.СтрокаЯвлетсяЧастоИспользуемой(cell.Value.ToString()) ?
				ConfigManager.МетаВесЧастоИспользуемогоТермина :
				0;

			return вероятность;
		}

		private static int ПолучитьКоличествоЯчеекВОбъединении(Range cell)
		{
			if (cell.MergeCells)
			{
				return cell.MergeArea.Cells.Count;
			}

			return 0;
		}

		private int ПолучитьКоличествоПустыхСтрокПодЯчейкой(Range cell)
		{
			int количество = 0;

			do
			{
				количество++;
			}
			while
				(количество < 10 && (cell.Offset[количество, 0].Value == null ||
				cell.Offset[количество, 0].Value.ToString() == "" ||
				cell.Offset[количество, 0].Value.ToString() == " "));

			return количество;
		}
		
		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ВерсияМетаописания { get; set; } = ConfigManager.МетаВерсияМетаописания;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Идентификатор { get; set; } = ConfigManager.МетаИдентификатор;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Наименование { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Группа { get; set; } = ConfigManager.МетаГруппа;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаНачалаДействия { get; set; } = ConfigManager.МетаДатаНачалаДействия;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаОкончанияДействия { get; set; } = ConfigManager.МетаДатаОкончанияДействия;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Авторство { get; set; } = ConfigManager.МетаАвторство;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаПоследнегоИзменения { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string НомерВерсии { get; set; } = ConfigManager.МетаНомерВерсии;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string РасположениеШапки { get; set; } = ConfigManager.МетаРасположениеШапки;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Хост { get; set; } = Environment.MachineName;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string СсылкаНаМетодическийСправочник { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string СсылкаНаВнешнююСправку { get; set; } = "";

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ВерсияФорматаМетаструктуры { get; set; } = ConfigManager.МетаВерсияФорматаМетаструктуры;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Тег { get; set; } = "";

		[XmlIgnore]
		public Workbook КнигаExcel { get; set; }

		[XmlIgnore]
		public Range МеткаНаименование { get; set; }
	}
}