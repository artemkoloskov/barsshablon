using System.Xml.Serialization;
using System.Xml.Schema;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class ОписаниеФормы
    {
        public ОписаниеФормы ()
        {
        }

        private Мета мета;
        private Структура структура;
        private string меню = "";
        private Справочник[] справочники;

        [XmlElement("Мета", Form = XmlSchemaForm.Unqualified)]
        public Мета Мета
        {
            get
            {
                return мета;
            }
            set
            {
                мета = value;
            }
        }

        [XmlElement("Структура", Form = XmlSchemaForm.Unqualified)]
        public Структура Структура
        {
            get
            {
                return структура;
            }
            set
            {
                структура = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Меню
        {
            get
            {
                return меню;
            }
            set
            {
                меню = value;
            }
        }

        [XmlArray(Form = XmlSchemaForm.Unqualified)]
        [XmlArrayItem("Справочник", typeof(Справочник), Form = XmlSchemaForm.Unqualified, IsNullable = false)]
        public Справочник[] Справочники
        {
            get
            {
                return справочники;
            }
            set
            {
                справочники = value;
            }
        }
    }
}