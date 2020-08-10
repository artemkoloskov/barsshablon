using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
	public class Строковый
	{
		public string Разделитель = ";";
		public bool МногострочныйРедактор = false;
		public string МаскаВвода = "";
		public string ВсплывающаяПодсказка = "";
		public bool ОбязательноДляЗаполнения = false;
		public bool ТолькоЧтение = false;
		public string Комментарий = "";
		public bool ЯвляетсяКлючевым = false;
		public string ЗначениеПоУмолчанию = "";
		public string ДействиеСПолем = "БезИтогов";

		public string ToXML()
		{
			using (var stringwriter = new System.IO.StringWriter())
			{
				var serializer = new XmlSerializer(this.GetType());
				serializer.Serialize(stringwriter, this);
				return stringwriter.ToString();
			}
		}
	}
}
