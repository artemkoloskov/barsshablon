using System.Xml.Serialization;

namespace БАРСШаблон.DataTypes
{
    public class Логический
    {
        public bool ОбязательноДляЗаполнения = false;
        public bool ТолькоЧтение = false;
        public string Комментарий = "";
        public bool ЯвляетсяКлючевым = true;
        public bool ЗначениеПоУмолчанию;
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
