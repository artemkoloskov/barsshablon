using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class IntegerType : CellTypeDescription
	{
		public new string Action = "Суммировать";
		public int Precision = 0;
		public string ValueRange = "";
	}
}
