using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class TextType : CellTypeDescription
	{
		public string Devider = ";";
		public bool MultyRowEditor = false;
		public string Mask = "";
		public string ToolTip = "";
	}
}
