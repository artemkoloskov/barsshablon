using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class Логический : ОписаниеТипаЯчейки
	{

		public new bool ЯвляетсяКлючевым = true;
		public new bool ЗначениеПоУмолчанию;
	}
}
