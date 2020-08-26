using System.Xml.Schema;
using System.Xml.Serialization;

namespace БАРСШаблон
{
	[System.Serializable()]
	[XmlType(AnonymousType = true)]
	public partial class Справочник
	{
		public Справочник()
		{
		}

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("ЭлементСправочника", typeof(ЭлементСправочника), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public ЭлементСправочника[] Элементы { get; set; }

		[XmlAttribute()]
		public string Код { get; set; }

		[XmlAttribute()]
		public string Наименование { get; set; }
	}
}