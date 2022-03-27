using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	[XmlType(TypeName = "ОписаниеТипаЯчейкиАбстракт")]
	public abstract class CellTypeDescription
	{
		public bool Mandatory = false;
		public bool ReadOnly = false;
		public string Comment = "";
		public bool IsKey = false;
		public string DefaultValue = "";
		public string Action = "БезИтогов";

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
