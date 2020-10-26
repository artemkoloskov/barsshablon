using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейки")]
	public class Строковый : ОписаниеТипаЯчейки
	{
		public string Разделитель = ";";
		public bool МногострочныйРедактор = false;
		public string МаскаВвода = "";
		public string ВсплывающаяПодсказка = "";
	}
}
