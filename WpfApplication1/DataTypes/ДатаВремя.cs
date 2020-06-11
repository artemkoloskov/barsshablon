using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БАРСШаблон.DataTypes
{
    public class ДатаВремя
    {
        public string ФорматОтображения = "";
        public string DateAttributes = "";
        public string DateRangeBegin = "";
        public string DateRangeEnd = "";
        public bool ОбязательноДляЗаполнения = false;
        public bool ТолькоЧтение = false;
        public string Комментарий = "";
        public bool ЯвляетсяКлючевым = false;
        public string ЗначениеПоУмолчанию = "";
        public string ДействиеСПолем = "БезИтогов";
    }
}
