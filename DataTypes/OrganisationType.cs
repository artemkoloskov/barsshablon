using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class OrganisationType : CellTypeDescription
	{
		public new bool IsKey = true;
	}
}
