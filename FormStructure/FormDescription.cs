using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	[XmlRoot(Namespace = "", IsNullable = false)]
	public class FormDescription
	{
		public FormDescription()
		{
		}

		public FormDescription(Meta meta, List<Table> tableList, List<FreeCell> freeCellList)
		{
			Meta = meta;

			Dictionaries = new Dictionary[] { };

			Structure = new Structure(tableList, freeCellList);
		}

		public static FormDescription GetDescription(Workbook workbook)
		{
			Meta meta = new Meta(workbook);

			List<Table> tables = Table.GetFormTables(workbook.Sheets);

			List<FreeCell> freeCells = FreeCell.GetFreeCells(workbook.Worksheets[1]);

			FormDescription formDescription = new FormDescription(meta, tables, freeCells);

			return formDescription;
		}

		[XmlElement("Мета", Form = XmlSchemaForm.Unqualified)]
		public Meta Meta { get; set; }

		[XmlElement("Структура", Form = XmlSchemaForm.Unqualified)]
		public Structure Structure { get; set; }

		[XmlElement("Меню", Form = XmlSchemaForm.Unqualified)]
		public string Menu { get; set; } = "";

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("Справочник", typeof(Dictionary), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public Dictionary[] Dictionaries { get; set; }
	}
}