using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_Insight
{
    class Event_Handler
    {
        public delegate void RaiseEvent(bool save);
        public static event RaiseEvent DB_EXCEL_LOAD;

        public delegate void RaiseEvent_2();
        public static event RaiseEvent_2 WEEK_READ_AFTER_SAVE;

        public static void Db_Excel_Load(bool save)
        {
            DB_EXCEL_LOAD(save);
        }

        public static void Week_Read_after_save()
        {
            WEEK_READ_AFTER_SAVE();
        }

    }
}
