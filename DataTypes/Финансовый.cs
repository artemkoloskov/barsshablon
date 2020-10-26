using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class Финансовый : ОписаниеТипаЯчейки
	{
		public new string ДействиеСПолем = "Суммировать";
		public int Точность = 2;
		public string ValueRange = "";
	}
}
