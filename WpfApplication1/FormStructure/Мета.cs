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

			НаитиТегиВКниге();

			наименование = ПолучитьНаименование();
			идентификатор += ДопМетоды.ПолучитьТег(наименование);
			группа += DateTime.Today.Year;
			датаНачалаДействия = датаНачалаДействия.Replace("0001", DateTime.Today.Year.ToString());
			датаОкончанияДействия = датаОкончанияДействия.Replace("9999", DateTime.Today.Year.ToString());
			датаПоследнегоИзменения = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
			тег = идентификатор;
		}

		private void НаитиТегиВКниге()
		{
			string строкаТегаНаименование = ConfigManager.МетаТегНаименование;

			foreach (Range клеткаТаблицы in КнигаExcel.Worksheets[1].UsedRange.Cells)
			{
				if (клеткаТаблицы.Value != null)
				{
					if (клеткаТаблицы.Value.ToString() == строкаТегаНаименование)
					{
						тегНаименование = клеткаТаблицы;
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
			return ДопМетоды.ПолучитьНаименованиеПоТегу(тегНаименование, out наименование);
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
				ConfigManager.МетаВесЖирностиТекстаПотенциальногоНаименования:
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
			} while (количество < 10 && (cell.Offset[количество, 0].Value == null ||
					cell.Offset[количество, 0].Value.ToString() == "" ||
					cell.Offset[количество, 0].Value.ToString() == " "));

			return количество;
		}

		private string версияМетаописания = ConfigManager.МетаВерсияМетаописания;
		private string идентификатор = ConfigManager.МетаИдентификатор;
		private string наименование = "";
		private string группа = ConfigManager.МетаГруппа;
		private string датаНачалаДействия = ConfigManager.МетаДатаНачалаДействия;
		private string датаОкончанияДействия = ConfigManager.МетаДатаОкончанияДействия;
		private string авторство = ConfigManager.МетаАвторство;
		private string датаПоследнегоИзменения = "";
		private string номерВерсии = ConfigManager.МетаНомерВерсии;
		private string расположениеШапки = ConfigManager.МетаРасположениеШапки;
		private string хост = Environment.MachineName;
		private string ссылкаНаМетодическийСправочник = "";
		private string ссылкаНаВнешнююСправку = "";
		private string версияФорматаМетаструктуры = ConfigManager.МетаВерсияФорматаМетаструктуры;
		private string тег = "";

		private Workbook книгаExcel;
		private Range тегНаименование;

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ВерсияМетаописания { get => версияМетаописания; set => версияМетаописания = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Идентификатор { get => идентификатор; set => идентификатор = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Наименование { get => наименование; set => наименование = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Группа { get => группа; set => группа = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаНачалаДействия { get => датаНачалаДействия; set => датаНачалаДействия = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаОкончанияДействия { get => датаОкончанияДействия; set => датаОкончанияДействия = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Авторство { get => авторство; set => авторство = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ДатаПоследнегоИзменения { get => датаПоследнегоИзменения; set => датаПоследнегоИзменения = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string НомерВерсии { get => номерВерсии; set => номерВерсии = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string РасположениеШапки { get => расположениеШапки; set => расположениеШапки = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Хост { get => хост; set => хост = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string СсылкаНаМетодическийСправочник { get => ссылкаНаМетодическийСправочник; set => ссылкаНаМетодическийСправочник = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string СсылкаНаВнешнююСправку { get => ссылкаНаВнешнююСправку; set => ссылкаНаВнешнююСправку = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string ВерсияФорматаМетаструктуры { get => версияФорматаМетаструктуры; set => версияФорматаМетаструктуры = value; }

		[XmlElement(Form = XmlSchemaForm.Unqualified)]
		public string Тег { get => тег; set => тег = value; }

		[XmlIgnore]
		public Workbook КнигаExcel { get => книгаExcel; set => книгаExcel = value; }

		[XmlIgnore]
		public Range ТегНаименование { get => тегНаименование; set => тегНаименование = value; }
	}
}