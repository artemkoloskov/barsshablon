using System.Collections.Generic;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Structure
	{
		public Structure()
		{
		}

		public Structure(List<Table> tablesList, List<FreeCell> freeCellsList)
		{
			if (tablesList.Count > 0)
			{
				Tables = tablesList.ToArray();
			}

			if (freeCellsList.Count > 0)
			{
				FreeCells = freeCellsList.ToArray();
			}
		}

		[XmlElement("СвободнаяЯчейка", Form = XmlSchemaForm.Unqualified)]
		public FreeCell[] FreeCells { get; set; }

		[XmlElement("Таблица", Form = XmlSchemaForm.Unqualified)]
		public Table[] Tables { get; set; }
	}
}