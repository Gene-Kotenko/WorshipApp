using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worship
{
    public class cSong
    {
        public string song_name = "";
        public string song_name_path = "";      // to save later in the right directory
        public string orig_song_path = "";
        public string source_chords_path = "";
        public string dest_chords_path = "";
        public string song_id = "";
        public string song_cat_id = "";
        public string song_type_id = "";
        public string song_worship_prd_id = "";
        public string song_text_path_id = "";   // new element
        public string song_text_path = "";
        public string phonogram_path = "";
        public string song_worship_num = "";
        public string song_general_num = "";
        public string song_key = "";
        public string song_phonogram = "";
        public string song_note = "";
        public string song_state = "1";         // 1-Active, 9-Inactive/Removed
        public string song_sys_date = DateTime.Now.ToString("MM/dd/yyyy H:mm:ss");
        public string sort_by_order = "";
        public bool move_song = false;
    }
}
