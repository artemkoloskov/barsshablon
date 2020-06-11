using System.Xml.Serialization;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class ЭлементСправочника
    {
        private string код;
        private string наименование;

        [XmlAttribute()]
        public string Код
        {
            get
            {
                return код;
            }
            set
            {
                код = value;
            }
        }

        [XmlAttribute()]
        public string Наименование
        {
            get
            {
                return наименование;
            }
            set
            {
                наименование = value;
            }
        }
    }
}