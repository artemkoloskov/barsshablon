using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class Целочисленный : ОписаниеТипаЯчейки
	{
		public new string ДействиеСПолем = "Суммировать";
		public int Точность = 0;
		public string ValueRange = "";
	}
}
