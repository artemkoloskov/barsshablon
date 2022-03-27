using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class MoneyType : CellTypeDescription
	{
		public new string Action = "Суммировать";
		public int Precision = 2;
		public string ValueRange = "";
	}
}
