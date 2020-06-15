using System.Xml.Serialization;
using System.Xml.Schema;
using System;
using System.Configuration;

namespace БАРСШаблон
{
    [System.Serializable()]
    [XmlType(AnonymousType = true)]
    public partial class Мета
    {
        public Мета()
        {
            группа = группа + DateTime.Today.Year;
            датаНачалаДействия = датаНачалаДействия.Replace("0001", DateTime.Today.Year.ToString());
            датаОкончанияДействия = датаОкончанияДействия.Replace("9999", DateTime.Today.Year.ToString());
            датаПоследнегоИзменения = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
            хост = Environment.MachineName;
        }

        private string версияМетаописания = ConfigurationManager.AppSettings.Get("МетаВерсияМетаописания");
        private string идентификатор = ConfigurationManager.AppSettings.Get("МетаИдентификатор");
        private string наименование = "";
        private string группа = ConfigurationManager.AppSettings.Get("МетаГруппа");
        private string датаНачалаДействия = ConfigurationManager.AppSettings.Get("МетаДатаНачалаДействия");
        private string датаОкончанияДействия = ConfigurationManager.AppSettings.Get("МетаДатаОкончанияДействия");
        private string авторство = ConfigurationManager.AppSettings.Get("МетаАвторство");
        private string датаПоследнегоИзменения = "";
        private string номерВерсии = ConfigurationManager.AppSettings.Get("МетаНомерВерсии");
        private string расположениеШапки = ConfigurationManager.AppSettings.Get("МетаРасположениеШапки");
        private string хост = "";
        private string ссылкаНаМетодическийСправочник = "";
        private string ссылкаНаВнешнююСправку = "";
        private string версияФорматаМетаструктуры = ConfigurationManager.AppSettings.Get("МетаВерсияФорматаМетаструктуры");
        private string тег = "";

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ВерсияМетаописания
        {
            get
            {
                return версияМетаописания;
            }
            set
            {
                версияМетаописания = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Группа
        {
            get
            {
                return группа;
            }
            set
            {
                группа = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаНачалаДействия
        {
            get
            {
                return датаНачалаДействия;
            }
            set
            {
                датаНачалаДействия = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаОкончанияДействия
        {
            get
            {
                return датаОкончанияДействия;
            }
            set
            {
                датаОкончанияДействия = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Авторство
        {
            get
            {
                return авторство;
            }
            set
            {
                авторство = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ДатаПоследнегоИзменения
        {
            get
            {
                return датаПоследнегоИзменения;
            }
            set
            {
                датаПоследнегоИзменения = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string НомерВерсии
        {
            get
            {
                return номерВерсии;
            }
            set
            {
                номерВерсии = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string РасположениеШапки
        {
            get
            {
                return расположениеШапки;
            }
            set
            {
                расположениеШапки = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string Хост
        {
            get
            {
                return хост;
            }
            set
            {
                хост = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string СсылкаНаМетодическийСправочник
        {
            get
            {
                return ссылкаНаМетодическийСправочник;
            }
            set
            {
                ссылкаНаМетодическийСправочник = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string СсылкаНаВнешнююСправку
        {
            get
            {
                return ссылкаНаВнешнююСправку;
            }
            set
            {
                ссылкаНаВнешнююСправку = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
        public string ВерсияФорматаМетаструктуры
        {
            get
            {
                return версияФорматаМетаструктуры;
            }
            set
            {
                версияФорматаМетаструктуры = value;
            }
        }

        [XmlElement(Form = XmlSchemaForm.Unqualified)]
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