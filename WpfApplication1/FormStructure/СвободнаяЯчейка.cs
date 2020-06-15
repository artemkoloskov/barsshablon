using System.Xml.Serialization;
using System.Configuration;
using System;
using БАРСШаблон.DataTypes;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class СвободнаяЯчейка
    {
        public СвободнаяЯчейка()
        {
        }

        public СвободнаяЯчейка(string кодЯчейки, string типЯчейки)
        {
            идентификатор = кодЯчейки;
            код = кодЯчейки;
            тип = типЯчейки;
            тег = ConfigurationManager.AppSettings.Get("СвободнаяЯчейкаТегПрефикс") + CommonMethods.GetTagName(идентификатор);
            описание = CommonMethods.GetSerializedType(тип);
        }

        private string Serialize(object датаВремя)
        {

            throw new NotImplementedException();
        }

        private string идентификатор;
        private string код;
        private string наименованиеЭлемента;
        private string тип;
        private string описание;
        private string тег;

        [XmlAttribute()]
        public string Идентификатор
        {
            get
            {
                return идентификатор;
            }
            set
            {
                идентификатор = value;
            }
        }

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
        public string НаименованиеЭлемента
        {
            get
            {
                return наименованиеЭлемента;
            }
            set
            {
                наименованиеЭлемента = value;
            }
        }

        [XmlAttribute()]
        public string Тип
        {
            get
            {
                return тип;
            }
            set
            {
                тип = value;
            }
        }

        [XmlAttribute()]
        public string Описание
        {
            get
            {
                return описание;
            }
            set
            {
                описание = value;
            }
        }

        [XmlAttribute()]
        public string Тег
        {
            get
            {
                return тег;
            }
            set
            {
                тег = value;
            }
        }
    }
}
