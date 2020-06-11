using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
