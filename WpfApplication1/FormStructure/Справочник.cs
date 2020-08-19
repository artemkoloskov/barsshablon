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

		private ЭлементСправочника[] элементы;
		private string код;
		private string наименование;

		[XmlArray(Form = XmlSchemaForm.Unqualified)]
		[XmlArrayItem("ЭлементСправочника", typeof(ЭлементСправочника), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
		public ЭлементСправочника[] Элементы
		{
			get => элементы;
			set => элементы = value;
		}

		[XmlAttribute()]
		public string Код
		{
			get => код;
			set => код = value;
		}

		[XmlAttribute()]
		public string Наименование
		{
			get => наименование;
			set => наименование = value;
		}
	}
}