using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейкиАбстракт")]
	public abstract class ОписаниеТипаЯчейки
	{
		public bool ОбязательноДляЗаполнения = false;
		public bool ТолькоЧтение = false;
		public string Комментарий = "";
		public bool ЯвляетсяКлючевым = false;
		public string ЗначениеПоУмолчанию = "";
		public string ДействиеСПолем = "БезИтогов";

		public string ToXML()
		{
			using (StringWriter stringWriter = new StringWriter())
			{
				XmlSerializerNamespaces nameSpaces = new XmlSerializerNamespaces();
				nameSpaces.Add("", "");

				XmlWriterSettings settings = new XmlWriterSettings
				{
					OmitXmlDeclaration = true,
					NamespaceHandling = NamespaceHandling.OmitDuplicates,
					NewLineHandling = NewLineHandling.None,
					Indent = false
				};

				using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, settings))
				{
					XmlSerializer serializer = new XmlSerializer(GetType());

					serializer.Serialize(xmlWriter, this, nameSpaces);
				}

				return stringWriter.ToString();
			}
		}
	}
}
