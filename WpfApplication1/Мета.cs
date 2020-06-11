using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БАРСШаблон
{
    public class Мета
    {
        public string ВерсияМетаописания = "1.0";
        public string Идентификатор = "ДЗПК_";
        public string Наименование = "";
        public string Группа = "Региональный ";
        public string ДатаНачалаДействия = "01.01.0001 0:00:00";
        public string ДатаОкончанияДействия = "31.12.9999 0:00:00";
        public string Авторство = "ГАУЗ ПК МИАЦ";
        public string ДатаПоследнегоИзменения = "";
        public int НомерВерсии = 1;
        public string РасположениеШапки = "Сверху";
        public string Хост = "";
        public string СсылкаНаМетодическийСправочник = "";
        public string СсылкаНаВнешнююСправку = "";
        public string ВерсияФорматаМетаструктуры = "1,0";
        public string Тег = "";

        public Мета ()
        {  
            Группа = Группа + DateTime.Today.Year;
            ДатаНачалаДействия = ДатаНачалаДействия.Replace("0001", DateTime.Today.Year.ToString());
            ДатаОкончанияДействия = ДатаОкончанияДействия.Replace("9999", DateTime.Today.Year.ToString());
            ДатаПоследнегоИзменения = DateTime.Today.ToString("dd.MM.yyyy HH:mm:ss");
            Хост = Environment.MachineName;
        }
    }
}
