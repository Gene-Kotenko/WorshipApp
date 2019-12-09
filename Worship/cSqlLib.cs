using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worship
{
    public class cSqlLib
    {
        public cSqlLib()
        {
        }

        public string get_song_cat()
        {
            string sql = "select SONG_CAT_ID as id, SONG_CAT_DESCRIPTION as val" +
                " from song_cat" +
                " order by SONG_CAT_ID asc";
            return sql;
        }

        public string get_song_type()
        {
            string sql = "select SONG_TYPE_ID as id, SONG_TYPE_DESCRIPTION as val" +
                " from song_type" +
                " order by SONG_TYPE_ID asc";
            return sql;
        }

        public string get_song_worship_time()
        {
            string sql = "select SONG_SUBTYPE_ID as id, SONG_SUBTYPE_DESCRIPTION as val" +
                " from song_subtype" +
                " order by SONG_SUBTYPE_ID asc";
            return sql;
        }

        public string get_songs(string song_type)
        {
            string sql = "";

            sql = "select s.song_name, s.song_id, s.song_note," +
                " sd.song_general_number, sd.song_worship_number," +
                " sd.song_key, sd.phonogram, p2.path_location as song_text_path," +
                " p.path_location as phonogram_path, sc.song_cat_description," +
                " st.song_type_description, ss.song_subtype_description" +
                " from song s," +
                " song_detail sd," +
                " song_cat sc," +
                " song_type st," +
                " song_subtype ss," +
                " paths p," +
                " paths p2" +
                " where sd.song_id = s.song_id";

            switch (song_type)
            {
                case "ОБЩЕЕ ПЕНИЕ":
                    sql += " and sd.song_type_id = 1";
                    break;
                case "ГРУППОВОЕ ПЕНИЕ":
                    sql += " and sd.song_type_id = 2";
                    break;
                case "ДЕТСКИЕ ПЕСНИ":
                    sql += " and sd.song_type_id = 3";
                    break;
            }

            sql += " and sd.song_cat_id = sc.song_cat_id" +
            " and sd.song_type_id = st.song_type_id" +
            " and sd.song_subtype_id = ss.song_subtype_id" +
            " and sd.song_state = '1'" +
            " and p.path_type_id = 4" +
            " and p2.path_type_id = sd.song_text_path_id" +
            " order by s.song_name asc";

            return sql;
        }

        public string get_song_detail(string song_id)
        {
            string sql = "select p2.path_location as song_text_path," +
                " p.path_location as phonogram_path," +
                " s.song_note, sd.song_general_number," +
                " sd.song_worship_number, sd.song_key, sd.phonogram" +
                " from song s, song_detail sd," +
                " paths p," +
                " paths p2" +
                " where sd.song_id = s.song_id" +
                " and s.song_id = " + song_id +
                " and p2.path_type_id = sd.song_text_path_id" +
                " and p.path_type_id = 4";
            return sql;
        }

        public string get_specific_songs(string[] key)
        {
            string sql = "";

            sql = "SELECT DISTINCT" +
                    " Song.SONG_ID," +
                    " Song.SONG_NAME," +
                    " Song.SONG_NOTE," +
                    " x.last_dt" +
                " FROM" +
                    " (Song INNER JOIN Song_detail ON Song.SONG_ID = Song_detail.SONG_ID) LEFT JOIN" +
                    " (select svr.song_id, max(svr.song_last_sang) as last_dt" +
                    " FROM service svr GROUP by svr.song_id) AS x" +
                    " ON Song.SONG_ID = x.song_id" +
                " WHERE Song_detail.song_cat_id = " + key[2] +
                " AND Song_detail.song_type_id = " + key[4] +
                " AND Song_detail.song_subtype_id = " + key[6];

            if (key[8].Length == 0) // default
                sql += " ORDER BY x.last_dt asc";
            else
                sql += " ORDER BY " + key[8];
            
            return sql;
        }

        public string save_new_service(cService serv, string song_id, int order)
        {
            string sql = "";

            sql = "INSERT into service (" +
                " SERVICE_ID," +
                " SERVICE_YEAR_ID," +
                " SONG_ID," +
                " SERVICE_DATE," +
                " SERVICE_TYPE," +
                " SERVICE_NOTE," +
                " SONG_LAST_SANG," +
                " SONG_ORDER," +
                " SYSTEM_DATE," +
                " SYSTEM_TYPE," +
                " SYSTEM_CNT)" +
                " values (" +
                serv.service_id + "," +
                serv.service_year_id + "," +
                "'" + song_id + "'," +
                "#" + serv.service_date + "#," +
                serv.service_type + "," +
                "'" + serv.service_note + "'," +
                "#" + serv.song_last_sang + "#, " +
                order + "," +
                "#" + serv.service_sys_date + "#," +
                "'I'," +
                "0)";
            return sql;
        }

        public string update_service(cService serv, string song_id, string song_rid, int order)
        {
            string sql = "";

            sql = "UPDATE service " +
                " set SERVICE_YEAR_ID = '" + serv.service_year_id + "'," +
                " SONG_ID = '" + song_id + "'," +
                " SERVICE_DATE = #" + serv.service_date + "#," +
                " SERVICE_TYPE = " + serv.service_type + "," +
                " SERVICE_NOTE = '" + serv.service_note + "'," +
                " SONG_LAST_SANG = #" + serv.song_last_sang + "#," +
                " SONG_ORDER = " + order + "," +
                " SYSTEM_DATE = #" + serv.service_sys_date + "#," +
                " SYSTEM_TYPE = 'U'," +
                " SYSTEM_CNT = SYSTEM_CNT+1" +
                " where SERVICE_ID = " + serv.service_id +
                " and rid = " + song_rid;

            return sql;
        }

        public string get_latest_spevka()
        {
            string sql = "select s.rid, s.service_id, " +
                " s.song_id, s.service_date, s.service_note, so.song_note," +
                " sd.song_general_number" +
                " from service s," +
                " song_detail sd," +
                " song so" +
                " where sd.song_id = s.song_id" +
                " and so.song_id = s.song_id" +
                " and s.service_date = (select Max(service_date) from service)" +
                " order by song_order asc, s.rid asc";

            return sql;
        }

        public string get_latest_servise_id()
        {
            string sql = "select Max(service_id) as id from service";
            return sql;
        }

        public string get_song(string song_id)
        {
            string sql = "select s.song_name" +
                " from song s" +
                " where s.song_id = " + song_id;
            return sql;
        }

        public string get_song_by_name(string song_name)
        {
            string sql = "select s.song_id, s.song_name, s.song_note," +
                " sd.song_general_number, sd.song_worship_number," +
                " sd.song_key, sd.phonogram, p2.path_location as song_text_path," +
                " p.path_location as phonogram_path, sc.song_cat_description," +
                " st.song_type_description, ss.song_subtype_description" +
                " from song s, " +
                " song_detail sd," +
                " song_cat sc," +
                " song_type st," +
                " song_subtype ss," +
                " paths p," +
                " paths p2" +
                " where s.song_name like '%" + song_name + "%'" +
                " and sd.song_id = s.song_id" +
                " and sc.song_cat_id = sd.song_cat_id" +
                " and st.song_type_id = sd.song_type_id" +
                " and ss.song_subtype_id = sd.song_subtype_id" +
                " and p2.path_type_id = sd.song_text_path_id" +
                " and p.path_type_id = 4";
            return sql;
        }

        public string get_song_by_id(string worship_id, string general_id)
        {
            string sql = "";

            sql = "select s.song_name, st.song_type_description" +
                " from song s, " +
                " song_detail sd, " +
                " song_type st ";

            if (worship_id.Length > 0)
                sql += " where sd.song_worship_number = '" + worship_id + "'";
            else
                sql += " where sd.song_general_number = '" + general_id + "'";

            sql += " and sd.song_id = s.song_id" +
                " and st.song_type_id = sd.song_type_id";

            return sql;
        }

        public string get_services(string crit_1, string crit_2, bool all_serv, string serv_crit)
        {
            string sql = "";

            sql = "select s.service_id, s.song_id, s.service_note," +
                " st.song_type_description, sc.song_cat_description," +
                " s.service_date, so.song_name, so.song_note" +
                " from service s," +
                " song so," +
                " song_cat sc," +
                " song_type st," +
                " song_detail sd" +
                " where so.song_id = s.song_id" +
                " and sd.song_id = s.song_id" +
                " and st.song_type_id = sd.song_type_id" +
                " and sc.song_cat_id = s.service_type";

            if (serv_crit == "Lord_supp" || serv_crit == "Holidays")
                sql += " and s.service_type in (" + crit_1 + "," + crit_2 + ")";
            else
            {
                if (all_serv == false)
                    sql += " and service_date between #" + crit_1 + "# AND #" + crit_2 + "#";
            }
            sql += " order by service_id desc, song_order asc";

            return sql;
        }

        public string jrn_service(cService s)
        {
            string sql = "insert into service_jrn (" +
                " RID," +
                " SERVICE_ID," +
                " SERVICE_YEAR_ID," +
                " SONG_ID," +
                " SERVICE_DATE," +
                " SERVICE_TYPE," +
                " SERVICE_NOTE," +
                " SONG_LAST_SANG," +
                " SERVICE_SYSTEM_DATE," +
                " SYSTEM_DATE," +
                " SYSTEM_TYPE," +
                " SYSTEM_CNT)" +
                " values (" +
                s.service_rid + "," +
                s.service_id + "," +
                s.service_year_id + "," +
                s.song_id + "," +
                "#" + s.service_date + "#," +
                s.service_type + "," +
                "'" + s.service_note + "'," +
                "#" + s.song_last_sang + "#," +
                "#" + s.service_sys_date + "#," +
                "#" + s.jrn_system_date + "#," +
                "'" + s.service_sys_type + "'," +
                s.service_sys_cnt + ")";

            return sql;
        }

        public string delete_song_or_service(string serv_id, string song_id)
        {
            string sql = "delete * from service" +
                " where service_id = " + serv_id;

            if (song_id.Length != 0)
                sql += " and song_id = " + song_id;

            return sql;
        }

        public string get_service_song_data(string serv_id, string song_id)
        {
            string sql = "select * from service" +
                " where service_id = " + serv_id;

            if (song_id.Length != 0)
                sql += " and song_id = " + song_id;

            return sql;
        }

        public string get_max_song_id()
        {
            string sql = "select Max(song_id) as id from song";
            return sql;
        }

        public string get_path(string path_type_id)
        {
            string sql = "select path_location" +
                " from paths" +
                " where path_type_id = " + path_type_id;

            return sql;
        }

        public string get_path_id(string path_type_path)
        {
            string sql = "select path_type_id" +
                " from paths" +
                " where path_location = '" + path_type_path + "'";

            return sql;
        }

        public string add_new_song(cSong s)
        {
            string sql = "insert into song (" +
                " SONG_NAME," +
                " SONG_NOTE," +
                " SYSTEM_DATE," +
                " SYSTEM_TYPE," +
                " SYSTEM_CNT)" +
                " values (" +
                "'" + s.song_name + "'," +
                "'" + s.song_note + "'," +
                "#" + s.song_sys_date + "#," +
                "'I'," +
                "0)";

            return sql;
        }

        public string add_new_song_dtl(cSong s)
        {
            string sql = "insert into song_detail (" +
                " SONG_ID," +
                " SONG_CAT_ID," +
                " SONG_TYPE_ID," +
                " SONG_SUBTYPE_ID," +
                " SONG_TEXT_PATH_ID," +
                " SONG_GENERAL_NUMBER," +
                " SONG_WORSHIP_NUMBER," +
                " SONG_KEY," +
                " PHONOGRAM," +
                " SONG_STATE," +
                " SYSTEM_DATE," +
                " SYSTEM_TYPE," +
                " SYSTEM_CNT)" +
                " values (" +
                s.song_id + "," +
                s.song_cat_id + "," +
                s.song_type_id + "," +
                s.song_worship_prd_id + "," +
                s.song_text_path_id + "," +
                "'" + s.song_general_num + "'," +
                "'" + s.song_worship_num + "'," +
                "'" + s.song_key + "'," +
                "'" + s.song_phonogram + "'," +
                "'" + s.song_state + "'," +
                "#" + s.song_sys_date + "#," +
                "'I'," +
                "0)";

            return sql;
        }

        public string update_song(cSong s)
        {
            string sql = "UPDATE song " +
                " set SONG_NAME = '" + s.song_name + "'," +
                " SONG_NOTE = '" + s.song_note + "'," +
                " SYSTEM_DATE = #" + s.song_sys_date + "#," +
                " SYSTEM_TYPE = 'U'," +
                " SYSTEM_CNT = SYSTEM_CNT+1" +
                " where song_id = " + s.song_id;

            return sql;
        }

        public string update_song_dtl(cSong s)
        {
            string sql = "UPDATE song_detail " +
                " set SONG_CAT_ID = " + s.song_cat_id + "," +
                " SONG_TYPE_ID = " + s.song_type_id + "," +
                " SONG_SUBTYPE_ID = " + s.song_worship_prd_id + "," +
                " SONG_TEXT_PATH_ID = " + s.song_text_path_id + "," +
                " SONG_GENERAL_NUMBER = '" + s.song_general_num + "'," +
                " SONG_WORSHIP_NUMBER = " + s.song_id + "," +
                " SONG_KEY = '" + s.song_key + "'," +
                " PHONOGRAM = '" + s.song_phonogram + "'," +
                " SONG_STATE = '" + s.song_state + "'," +
                " SYSTEM_DATE = #" + s.song_sys_date + "#," +
                " SYSTEM_TYPE = 'U'," +
                " SYSTEM_CNT = SYSTEM_CNT+1" +
                " where song_id = " + s.song_id;

            return sql;
        }

        public string remove_song(cSong s)
        {
            string sql = "UPDATE song_detail " +
                " set SONG_STATE = '" + s.song_state + "'," +
                " SYSTEM_DATE = #" + s.song_sys_date + "#," +
                " SYSTEM_TYPE = 'D'," +
                " SYSTEM_CNT = SYSTEM_CNT+1" +
                " where song_id = " + s.song_id;

            return sql;
        }

        public string get_start_date()
        {
            string sql = "select Max(service_date) as latest_date from service"; // FORMAT(service_date, 'mm/dd/yyyy')
            return sql;
        }

    }
}
