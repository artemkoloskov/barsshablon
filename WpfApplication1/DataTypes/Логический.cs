using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
