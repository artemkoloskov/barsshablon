using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class LogicalType : CellTypeDescription
	{

		public new bool IsKey = true;
		public new bool DefaultValue;
	}
}
