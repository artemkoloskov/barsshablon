using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace БАРСШаблон.Structure
{
    public class СвободнаяЯчейка
    {
        public string Идентификатор = "";
        public string Код = "";
        public string НаименованиеЭлемента = "";
        public object Тип;
        public string Описание = "";
        public string Тег = "";

        public СвободнаяЯчейка(string кодЯчейки, object тип)
        {
            Идентификатор = кодЯчейки;
            Код = кодЯчейки;
            Тип = тип;
            Тег = "СвобЯч" + кодЯчейки;
        }
    }
}
