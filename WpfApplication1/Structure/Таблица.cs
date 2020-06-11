using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БАРСШаблон.Structure
{
    public class Таблица
    {
        public string Идентификатор = "";
        public string Код = "";
        public string Наименование = "";
        public string Тег = "";
        public string СсылкаНаМетодическийСправочник = "";
        public bool РучноеДобавлениеСтрок = false;

        public ICollection<СвободнаяЯчейка> СвободныеЯчейки;
        public ICollection<Строка> Строки;
        public ICollection<Столбец> Столбцы;
    }
}
