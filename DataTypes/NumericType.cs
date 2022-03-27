using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class NumericType : CellTypeDescription
	{
		public new string Action = "Суммировать";
		public int Precision = 2;
		public string ValueRange = "";
	}
}
