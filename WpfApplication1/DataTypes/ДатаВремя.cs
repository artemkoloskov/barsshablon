using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class ДатаВремя : ОписаниеТипаЯчейки
	{
		public new bool ЯвляетсяКлючевым = true;
		public string ФорматОтображения = "";
		public string DateAttributes = "";
		public string DateRangeBegin = "";
		public string DateRangeEnd = "";
	}
}
