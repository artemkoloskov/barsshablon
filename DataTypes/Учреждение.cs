using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class Учреждение : ОписаниеТипаЯчейки
	{
		public new bool ЯвляетсяКлючевым = true;
	}
}
