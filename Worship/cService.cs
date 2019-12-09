using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worship
{
    public class cService
    {
        public string service_rid = "";
        public string service_id = "";
        public string service_year_id = "";
        public string service_year = "";
        public string song_id = "";
        public string service_date = "";
        public string service_type = "";
        public string service_note = "";
        public string service_description = ""; // not using
        public string song_last_sang = "";
        public string service_sys_date = DateTime.Now.ToString("MM/dd/yyyy H:mm:ss");
        public string jrn_system_date = DateTime.Now.ToString("MM/dd/yyyy H:mm:ss");
        public string service_sys_type = "";
        public string service_sys_cnt = "";
        public object obj = null;
    }
}
