using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class DateTimeType : CellTypeDescription
	{
		public new bool IsKey = true;
		public string ViewingFormat = "";
		public string DateAttributes = "";
		public string DateRangeBegin = "";
		public string DateRangeEnd = "";
	}
}
