using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;

namespace Worship
{
    public class cUtil
    {
        private cSqlLib msql_lib = null;
        private cDb mdb = null;
        
        private cSqlLib sql_lib
        {
            get
            {
                if (msql_lib == null)
                    msql_lib = new cSqlLib();
                return msql_lib;
            }
            set
            {
                msql_lib = value;
            }
        }

        private cDb db
        {
            get { return mdb; }
            set { mdb = value; }
        }

        public cUtil()
        {
            //string db_path = Application.StartupPath + "\\db\\Worship.accdb"; // Didn't work for "like search" for some reason
            string db_path = Application.StartupPath + "\\db\\Worship.mdb";
            db = new cDb(db_path);
            sql_lib = new cSqlLib();
        }

        public void init_Tree(ref TreeView tv, string view_type)
        {
            TreeNode n = null;
            TreeNode sub = null;

            tv.Nodes.Clear();
            tv.BeginUpdate();

            if (view_type == "Песни_Ноты")
            {
                n = new TreeNode("ПЕСНИ", 56, 56); // 3, 3
                n.Tag = "ПЕСНИ;";

                sub = new TreeNode("ОБЩЕЕ ПЕНИЕ", 2, 2);
                sub.Tag = "ОБЩЕЕ ПЕНИЕ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                sub = new TreeNode("ГРУППОВОЕ ПЕНИЕ", 2, 2);
                sub.Tag = "ГРУППОВОЕ ПЕНИЕ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                sub = new TreeNode("ДЕТСКИЕ ПЕСНИ", 2, 2);
                sub.Tag = "ДЕТСКИЕ ПЕСНИ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                tv.Nodes.Add(n);
                n.Expand();

                // =====================================
                n = new TreeNode("ФОНОГРАММЫ", 56, 56); // 2, 2
                n.Tag = "ФОНОГРАММЫ;";
                n.Nodes.Add("...");
                tv.Nodes.Add(n);

                // =====================================
                n = new TreeNode("ПРОГРАММЫ", 56, 56); // 2, 2
                n.Tag = "ПРОГРАММЫ;";                

                sub = new TreeNode("РОЖДЕСТВЕНСКИЕ", 2, 2);
                sub.Tag = "РОЖДЕСТВЕНСКИЕ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                sub = new TreeNode("ПАСХАЛЬНЫЕ", 2, 2);
                sub.Tag = "ПАСХАЛЬНЫЕ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                sub = new TreeNode("ЖАТВЕННЫЕ", 2, 2);
                sub.Tag = "ЖАТВЕННЫЕ;";
                sub.Nodes.Add("...");
                n.Nodes.Add(sub);

                tv.Nodes.Add(n);
            }
            else if (view_type == "Спевка")
            {
                n = new TreeNode("ПЕСНИ ПО КАТЕГОРИЯМ", 2, 2);
                n.Tag = "ПЕСНИ_ПО_КАТЕГОРИЯМ;";
                n.Nodes.Add("...");
                tv.Nodes.Add(n);
                n.Expand();
            }
            tv.EndUpdate();
        }

        public void fill_cbo(ref ComboBox cbo, string type)
        {
            string[] data = get_cbo_data(type).Split("^".ToCharArray());
            string[] itm = null;

            cbo.Items.Clear();
            cbo.Tag = "";

            foreach (string rec in data)
            {
                itm = rec.Split("~".ToCharArray());
                if (itm.Length == 1)
                    break;
                cbo.Items.Add(itm[0]);
                cbo.Tag += itm[1];
            }
            if (cbo.Items.Count > 0)
                cbo.SelectedIndex = 0;
        }

        public string get_cbo_data(string type)
        {
            ADODB.Recordset rs = null;
            string sql = "";
            string data = "";

            switch (type)
            {
                case "song_cat":
                    sql = sql_lib.get_song_cat();
                    break;
                case "song_type":
                    sql = sql_lib.get_song_type();
                    break;
                case "song_worship_time":
                    sql = sql_lib.get_song_worship_time();
                    break;
            }

            if (sql.Length == 0)
                return data;

            if (mdb.RsOpen(out rs, sql) == false)
                return "";
            while (rs.EOF == false)
            {
                data += rs.Fields["val"].Value + "~" + rs.Fields["id"].Value + ";" + rs.Fields["val"].Value + "|^";
                rs.MoveNext();
            }
            rs.Close();
            rs = null;
            return data;
        }

        public string get_cbo_itm_key(ComboBox cbo)
        {
            return get_cbo_itm_key(cbo, 0);
        }

        public string get_cbo_itm_key(ComboBox cbo, int tag_itm_index)
        {
            string tag = "";
            string ret_val = "";
            int selected_ind = cbo.SelectedIndex;

            if (cbo.Tag == null)
                return ret_val;
            if (cbo.SelectedItem == null)
                return ret_val;
            tag = cbo.Tag.ToString();
            ret_val = get_itm_key(tag, selected_ind, tag_itm_index);
            return ret_val;
        }

        public string get_cbo_itm_key(ToolStripComboBox cbo)
        {
            if (cbo.Tag == null)
                return "";
            if (cbo.SelectedItem == null)
                return "";

            int selected_ind = cbo.SelectedIndex;
            string val = get_itm_key(cbo.Tag.ToString(), selected_ind, 0);
            return val;
        }

        private string get_itm_key(string cbo_tag, int selected_index, int tag_itm_index)
        {
            string[] lst = cbo_tag.Split("|".ToCharArray());
            string ret_val = "";
            string[] itm = null;
            int cnt = -1;

            foreach (string row in lst)
            {
                cnt++;
                itm = row.Split(";".ToCharArray());
                if (cnt == selected_index)
                {
                    ret_val = itm[tag_itm_index];
                    break;
                }
            }
            return ret_val;
        }

        public int set_cbo_itm(string cbo_tag, string itm_id)
        {
            string[] lst = cbo_tag.Split("|".ToCharArray());
            int ret_val = -1;
            string[] itm;
            bool stat = false;

            foreach (string row in lst)
            {
                itm = row.Split(";".ToCharArray());
                ret_val++;
                if (itm[1] == itm_id)
                {
                    stat = true;
                    break;
                }
            }
            if (stat == false)
                ret_val = 0;
            return ret_val;
        }

        public void process_node(ref TreeNode n)
        {
            string[] key = n.Tag.ToString().Split(';');

            switch (key[0])
            {
                case "ОБЩЕЕ ПЕНИЕ":
                case "ГРУППОВОЕ ПЕНИЕ":
                case "ДЕТСКИЕ ПЕСНИ":
                case "ТЕКСТ_ПЕСНИ":
                case "ДЕТАЛИ_ПЕСНИ":
                case "ПЕСНИ_ПО_КАТЕГОРИЯМ":
                case "ФОНОГРАММЫ":
                case "РОЖДЕСТВЕНСКИЕ":
                case "ПАСХАЛЬНЫЕ":
                case "ЖАТВЕННЫЕ":

                    OpenGroup(ref n);
                    break;
            }
        }

        private void OpenGroup(ref TreeNode n)
        {
            string sql = "";
            string[] key = n.Tag.ToString().Split(';');
            string node_text = "";
            DateTime date_last_song;
            string d_cnt = "";
            string date = "";
            string path = "";
            string path_type_id = "";
            int i = 0;
            TreeNode nGrp = null;
            ADODB.Recordset rs = null;
            DirectoryInfo di = null;

            n.Nodes.Clear();

            try
            {
                switch (key[0])
                {
                    case "ОБЩЕЕ ПЕНИЕ":
                    case "ГРУППОВОЕ ПЕНИЕ":
                    case "ДЕТСКИЕ ПЕСНИ":
                        sql = sql_lib.get_songs(key[0]);
                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        while (rs.EOF == false)
                        {
                            node_text = rs.Fields["song_name"].Value.ToString();
                            nGrp = new TreeNode(node_text, 47, 47);
                            nGrp.Tag = "ТЕКСТ_ПЕСНИ;song_name=;" + rs.Fields["song_name"].Value + ";song_id=;" + rs.Fields["song_id"].Value +
                                ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                                ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                                ";phonogram=;" + rs.Fields["phonogram"].Value + ";song_path=;" + rs.Fields["song_text_path"].Value +
                                ";phonogram_path=;" + rs.Fields["phonogram_path"].Value + ";song_cat_desc=;" + rs.Fields["song_cat_description"].Value +
                                ";song_type_desc=;" + rs.Fields["song_type_description"].Value + ";song_sub_desc=;" + rs.Fields["song_subtype_description"].Value;

                            nGrp.Nodes.Add("...");
                            n.Nodes.Add(nGrp);
                            rs.MoveNext();
                        }
                        break;

                    case "ТЕКСТ_ПЕСНИ":

                        if (key.Length < 22)
                            return;

                        if (key[22] == "Общее пение")
                            path_type_id = "6";
                        else if (key[22] == "Групповое пение")
                            path_type_id = "7";
                        else if (key[22] == "Детские песни")
                            path_type_id = "8";

                        sql = sql_lib.get_path(path_type_id);
                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        if (rs.EOF == false)
                            path = rs.Fields["path_location"].Value.ToString();

                        sql = sql_lib.get_song_detail(key[4]);
                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        while (rs.EOF == false)
                        {   // Song text                           
                            node_text = key[2];
                            nGrp = new TreeNode(node_text, 46, 46);
                            nGrp.Tag = "ДЕТАЛИ_ПЕСНИ;song_path=;" + rs.Fields["song_text_path"].Value + ";phonogram_path=;" + rs.Fields["phonogram_path"].Value +
                                ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                                 ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                                 ";phonogram=;" + rs.Fields["phonogram"].Value;
                            n.Nodes.Add(nGrp);

                            // Song chords (piano)
                            node_text = key[2];
                            nGrp = new TreeNode(node_text, 51, 51);
                            nGrp.Tag = "ДЕТАЛИ_ПЕСНИ;song_path=;" + path + ";phonogram_path=;" + rs.Fields["phonogram_path"].Value +
                                ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                                 ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                                 ";phonogram=;" + rs.Fields["phonogram"].Value;
                            n.Nodes.Add(nGrp);

                            //// Extra for future
                            //node_text = key[2];
                            //nGrp = new TreeNode(node_text, 52, 52);
                            //nGrp.Tag = "ДЕТАЛИ_ПЕСНИ;song_path=;" + rs.Fields["song_text_path"].Value + ";phonogram_path=;" + rs.Fields["phonogram_path"].Value +
                            //    ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                            //     ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                            //     ";phonogram=;" + rs.Fields["phonogram"].Value;
                            //n.Nodes.Add(nGrp);

                            //node_text = key[2];
                            //nGrp = new TreeNode(node_text, 53, 53);
                            //nGrp.Tag = "ДЕТАЛИ_ПЕСНИ;song_path=;" + rs.Fields["song_text_path"].Value + ";phonogram_path=;" + rs.Fields["phonogram_path"].Value +
                            //    ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                            //     ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                            //     ";phonogram=;" + rs.Fields["phonogram"].Value;
                            //n.Nodes.Add(nGrp);

                            rs.MoveNext();
                        }
                        break;

                    case "ДЕТАЛИ_ПЕСНИ":
                        sql = sql_lib.get_song_detail(key[4]);
                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        while (rs.EOF == false)
                        {
                            node_text = n.Text + " (Note: " + key[6] + ")";
                            nGrp = new TreeNode(node_text, 44, 44);
                            nGrp.Tag = "ДЕТАЛИ_ПЕСНИ;song_path=;" + rs.Fields["song_text_path"].Value + ";phonogram_path=;" + rs.Fields["phonogram_path"].Value +
                                ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                                 ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                                 ";phonogram=;" + rs.Fields["phonogram"].Value;
                            n.Nodes.Add(nGrp);
                        }
                        break;

                    case "ПЕСНИ_ПО_КАТЕГОРИЯМ":
                        sql = sql_lib.get_specific_songs(n.Tag.ToString().Split(';'));

                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        while (rs.EOF == false)
                        {
                            i++;
                            date = "";
                            d_cnt = "0";

                            if (rs.Fields["last_dt"].Value.ToString().Length == 0)
                                node_text = rs.Fields["song_name"].Value.ToString();
                            else
                            {
                                date_last_song = Convert.ToDateTime(rs.Fields["last_dt"].Value.ToString());
                                d_cnt = DateTime.Now.Subtract(date_last_song).Days.ToString();
                                date = date_last_song.ToString("MM/dd/yyyy");
                                node_text = rs.Fields["song_name"].Value.ToString() + " (" + date + ")";
                            }

                            nGrp = new TreeNode(node_text, 47, 47);
                            nGrp.Tag = "ТЕКСТ_ПЕСНИ;song_name=;" + rs.Fields["song_name"].Value + ";song_id=;" + rs.Fields["song_id"].Value +
                                ";song_note=;" + rs.Fields["song_note"].Value;

                            if (date.Length != 0 && Convert.ToInt32(d_cnt) < 120)
                                nGrp.ForeColor = Color.Red;

                            nGrp.Nodes.Add("...");
                            n.Nodes.Add(nGrp);
                            rs.MoveNext();
                        }

                        if (rs.EOF == true && i == 0)
                        {
                            node_text = "Не найдено...";
                            nGrp = new TreeNode(node_text, 50, 50);
                            n.Nodes.Add(nGrp);
                        }
                        break;

                    case "РОЖДЕСТВЕНСКИЕ":
                    case "ПАСХАЛЬНЫЕ":
                    case "ЖАТВЕННЫЕ":
                        sql = sql_lib.get_path("5");

                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        if (rs.EOF == false)
                            path = rs.Fields["path_location"].Value.ToString();

                        di = new DirectoryInfo(path + "\\" + key[0]);

                        foreach (FileInfo f in di.GetFiles()) // *.docx, *.docm
                        {
                            node_text = f.Name.ToString().Replace(".docx", "");
                            nGrp = new TreeNode(node_text, 44, 44); // 46
                            nGrp.Tag = "ПРОГРАММА;prog_path=;" + path + "\\" + key[0];
                            n.Nodes.Add(nGrp);
                        }
                        break;

                    case "ФОНОГРАММЫ":
                        sql = sql_lib.get_path("4");

                        if (mdb.RsOpen(out rs, sql) == false)
                            return;

                        if (rs.EOF == false)
                            path = rs.Fields["path_location"].Value.ToString();

                        di = new DirectoryInfo(path);
                        foreach (FileInfo f in di.GetFiles("*.mp3"))
                        {
                            node_text = f.Name.ToString();
                            nGrp = new TreeNode(node_text, 54, 54); // 48, 48
                            nGrp.Tag = "ФОНОГРАММА;prog_path=;" + path;
                            n.Nodes.Add(nGrp);
                        }
                        break;
                }
                if (rs != null)
                {
                    if (rs.State == 1)
                        rs.Close();
                }
                rs = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void open_song_file(string file_name, string open_with_app, bool check_file)
        {
            if (check_file == true)
            {
                if (file_name.Length == 0)
                    return;
            }
            ProcessStartInfo pInfo = new ProcessStartInfo();
            Process pr = new Process();
            bool stat = false;

            if (open_with_app.Length > 0)
            {
                if (File.Exists(file_name) == true)
                {
                    pInfo.FileName = file_name;
                    pInfo.Arguments = open_with_app;
                    pInfo.WindowStyle = ProcessWindowStyle.Normal;
                    stat = true;
                }
            }

            if (stat == true)
            {
                try
                {
                    Process.Start(pInfo);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }

        public void preview_song(ref RichTextBox rtb, string prev_file)
        {
            if (prev_file.Length == 0)
                return;

            try
            {
                if (File.Exists(prev_file) == true)
                {
                    rtb.LoadFile(prev_file);
                    rtb.SelectAll();
                    rtb.SelectionAlignment = HorizontalAlignment.Center;                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public string get_next_service_id()
        {
            string latest_serv_id = "";
            string next_serv_id = "";
            string sql = sql_lib.get_latest_servise_id();
            ADODB.Recordset rs = null;

            try
            {
                if (mdb.RsOpen(out rs, sql) == false)
                    return next_serv_id;

                if (rs.EOF == false)
                {
                    latest_serv_id = rs.Fields["id"].Value.ToString();
                    next_serv_id = (Convert.ToInt32(latest_serv_id) + 1).ToString();
                }
                rs.Close();
                rs = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return next_serv_id;
        }

        public bool save_spevka(cService serv)
        {
            bool stat = false;
            string sql = "";
            string[] key = null;
            string song_id = "";
            int i = 0;
            TreeNode n = (TreeNode)serv.obj;

            try
            {
                foreach (TreeNode nd in n.Nodes)
                {
                    i++;
                    key = nd.Tag.ToString().Split(';');
                    song_id = key[4];
                    // saving spevka into service
                    sql = sql_lib.save_new_service(serv, song_id, i);
                    stat = mdb.Execute(sql);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                stat = false;
            }
            return stat;
        }

        public bool update_spevka(cService serv)
        {
            bool stat = false;
            string sql = "";
            string[] key = null;
            string song_id = "";
            string song_rid = "";
            int i = 0;
            TreeNode n = (TreeNode)serv.obj;

            try
            {
                foreach (TreeNode nd in n.Nodes)
                {
                    i++;
                    key = nd.Tag.ToString().Split(';');
                    song_id = key[4];
                    song_rid = key[6];

                    // updating spevka into service
                    if (IsNumeric(song_rid) == true)
                        sql = sql_lib.update_service(serv, song_id, song_rid, i);
                    else // saving spevka into service
                        sql = sql_lib.save_new_service(serv, song_id, i);

                    stat = mdb.Execute(sql);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                stat = false;
            }
            return stat;
        }

        public bool IsNumeric(string val)
        {
            return Regex.IsMatch(val, "^\\d+(\\.\\d+)?$");
        }

        public void get_latest_spevka(ref TreeView tv, TextBox txt, DateTimePicker dtp)
        {
            string[] spevka_info = get_spevka_info().Split('|');
            string service_id = spevka_info[0];
            string[] song_ids = spevka_info[1].Split(';');
            string[] song_rid = spevka_info[2].Split(';');
            string service_note = spevka_info[3];
            string[] song_genr_num = spevka_info[4].Split(';');
            string[] song_note = spevka_info[5].Split(';');
            string sql = "";
            string song_name = "";
            int i = 0;
            ADODB.Recordset rs = null;
            TreeNode n = null;
            TreeNode sub = null;

            dtp.Value = Convert.ToDateTime(spevka_info[6]);

            tv.Nodes.Clear();
            tv.BeginUpdate();

            try
            {
                foreach (string id in song_ids)
                {
                    if (id == "0" || id == "")
                        break;

                    sql = sql_lib.get_song(id);
                    if (mdb.RsOpen(out rs, sql) == false)
                        return;

                    if (rs.EOF == false)
                    {
                        i++;
                        song_name = rs.Fields["song_name"].Value.ToString();

                        if (i == 1)
                        {
                            n = new TreeNode("ВЫБРАННЫЕ ПЕСНИ", 3, 3);
                            n.Tag = "ВЫБРАННЫЕ_ПЕСНИ;";
                            sub = new TreeNode(song_name, 47, 47);
                            sub.Tag = "ТЕКСТ_ПЕСНИ;song_name=;" + song_name + ";song_id=;" + id + ";song_rid=;" + song_rid[i - 1] +
                                ";song_genr_num=;" + song_genr_num[i - 1] + ";song_note=;" + song_note[i - 1];
                            n.Nodes.Add(sub);
                            tv.Nodes.Add(n);
                            n.Expand();
                        }
                        else
                        {
                            n = new TreeNode(song_name, 47, 47);
                            n.Tag = "ТЕКСТ_ПЕСНИ;song_name=;" + song_name + ";song_id=;" + id + ";song_rid=;" + song_rid[i - 1] +
                                ";song_genr_num=;" + song_genr_num[i - 1] + ";song_note=;" + song_note[i - 1];
                            tv.Nodes[0].Nodes.Add(n);
                        }
                    }
                    rs.Close();
                }
                
                rs = null;
                tv.Tag = service_id;
                txt.Text = service_note;
                tv.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string get_spevka_info()
        {
            string info = "";
            string rid = "";
            string song_ids = "0;"; // in case there is no songs
            string service_id = "";
            string service_date = "";
            string service_note = "";
            string song_note = "";
            string song_genr_num = "";
            string sql = sql_lib.get_latest_spevka();
            ADODB.Recordset rs = null;

            try
            {
                if (mdb.RsOpen(out rs, sql) == false)
                    return song_ids;
                // clean if we have songs
                song_ids = ""; 
                
                while (rs.EOF == false)
                {
                    rid += rs.Fields["rid"].Value.ToString() + ";";
                    song_ids += rs.Fields["song_id"].Value.ToString() + ";";
                    song_genr_num += rs.Fields["song_general_number"].Value.ToString() + ";";
                    service_id = rs.Fields["service_id"].Value.ToString();
                    service_note = rs.Fields["service_note"].Value.ToString();
                    song_note += rs.Fields["song_note"].Value.ToString() + ";";
                    service_date = Convert.ToDateTime(rs.Fields["service_date"].Value.ToString()).ToShortDateString();
                    rs.MoveNext();
                }
                rs.Close();
                rs = null;
                info = service_id + "|" + song_ids + "|" + rid + "|" + service_note + "|" + song_genr_num + "|" + song_note + "|" + service_date;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return info;
        }

        public void get_song_by_name(ref TreeView tv, string song_name)
        {
            string sql = "";
            string node_text = "";
            int i = 0;
            TreeNode n = null;
            TreeNode nGrp;
            ADODB.Recordset rs = null;

            tv.Nodes.Clear();
            tv.BeginUpdate();

            try
            {
                sql = sql_lib.get_song_by_name(song_name);

                if (mdb.RsOpen(out rs, sql) == false)
                    return;

                n = new TreeNode("ПЕСНИ ПО ЗАДАННОМУ ПОИСКУ", 2, 2);
                n.Tag = "ПЕСНИ_ПО_ЗАДАННОМУ_ПОИСКУ;";
                tv.Nodes.Add(n);                

                while (rs.EOF == false)
                {
                    i++;
                    node_text = rs.Fields["song_name"].Value.ToString();
                    nGrp = new TreeNode(node_text, 47, 47);
                    nGrp.Tag = "ТЕКСТ_ПЕСНИ;song_name=;" + rs.Fields["song_name"].Value + ";song_id=;" + rs.Fields["song_id"].Value +
                        ";song_note=;" + rs.Fields["song_note"].Value + ";general_number=;" + rs.Fields["song_general_number"].Value +
                        ";worship_number=;" + rs.Fields["song_worship_number"].Value + ";song_key=;" + rs.Fields["song_key"].Value +
                        ";phonogram=;" + rs.Fields["phonogram"].Value + ";song_path=;" + rs.Fields["song_text_path"].Value +
                        ";phonogram_path=;" + rs.Fields["phonogram_path"].Value + ";song_cat_desc=;" + rs.Fields["song_cat_description"].Value +
                        ";song_type_desc=;" + rs.Fields["song_type_description"].Value + ";song_sub_desc=;" + rs.Fields["song_subtype_description"].Value;

                    nGrp.Nodes.Add("...");
                    n.Nodes.Add(nGrp);
                    rs.MoveNext();
                }

                if (rs.EOF == true && i == 0)
                {
                    node_text = "Не найдено...";
                    nGrp = new TreeNode(node_text, 50, 50);
                    n.Nodes.Add(nGrp);
                }

                if (rs != null)
                {
                    if (rs.State == 1)
                        rs.Close();
                }
                rs = null;
                n.Expand();
                tv.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void get_song_by_id(ref TreeView tv, string worship_id, string general_id)
        {
            string sql = "";
            string song_name = "";
            string song_type_desc = "";
            bool stat = false;
            ADODB.Recordset rs = null;
            
            tv.BeginUpdate();

            try
            {
                sql = sql_lib.get_song_by_id(worship_id, general_id);

                if (mdb.RsOpen(out rs, sql) == false)
                    return;

                if (rs.EOF == false)
                {
                    song_name = rs.Fields["song_name"].Value.ToString();
                    song_type_desc = rs.Fields["song_type_description"].Value.ToString().ToUpper();
                }

                foreach (TreeNode n in tv.Nodes[0].Nodes)
                {
                    if (stat == true)
                        break;

                    if (n.Text == song_type_desc)
                    {
                        if (n.IsExpanded == false)
                            n.Expand();

                        foreach (TreeNode song in n.Nodes)
                        {
                            if (song.Text == song_name)
                            {
                                song.Expand();
                                stat = true;
                                break;
                            }
                        }
                    }
                }

                if (rs != null)
                {
                    if (rs.State == 1)
                        rs.Close();
                }
                rs = null;
                tv.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void get_services(ref ListView lv, string service_crit)
        {
            string[] criteria = { "", "" };

            if (service_crit != "All")
                criteria = set_criteria(service_crit).Split(';');

            string sql = sql_lib.get_services(criteria[0], criteria[1], service_crit == "All", service_crit);
            string song_cat_desc = "";
            int cur_serv_id = 0;
            int next_serv_id = 0;
            int i = 0;
            DateTime d = new DateTime();
            ListViewItem lvi = null;
            ADODB.Recordset rs = null;

            lv.Items.Clear();

            if (mdb.RsOpen(out rs, sql) == false)
                return;

            while (rs.EOF == false)
            {
                next_serv_id = Convert.ToInt32(rs.Fields["service_id"].Value);

                if (cur_serv_id != next_serv_id & i != 0)
                {
                    lvi = new ListViewItem();

                    lvi.BackColor = Color.FromArgb(241, 245, 248);
                    lv.Items.Add(lvi);
                }

                cur_serv_id = Convert.ToInt32(rs.Fields["service_id"].Value);
                d = Convert.ToDateTime(rs.Fields["service_date"].Value.ToString());

                lvi = new ListViewItem("   " + d.ToString("MM/dd/yyyy"));
                lvi.SubItems.Add(rs.Fields["song_name"].Value.ToString());
                lvi.SubItems.Add(rs.Fields["song_type_description"].Value.ToString());
                song_cat_desc = rs.Fields["song_cat_description"].Value.ToString();
                lvi.SubItems.Add(song_cat_desc);
                lvi.SubItems.Add(rs.Fields["song_note"].Value.ToString());
                lvi.SubItems.Add(rs.Fields["service_note"].Value.ToString());
                lvi.Tag = "service_id=;" + rs.Fields["service_id"].Value + ";song_id=;" + rs.Fields["song_id"].Value +
                    ";song_name=;" + rs.Fields["song_name"].Value;

                rs.MoveNext();
                i++;
                
                if (song_cat_desc == "Рождество" || song_cat_desc == "Пасха" || song_cat_desc == "Жатва")
                    lvi.ForeColor = Color.SeaGreen;
                else if (song_cat_desc == "Вечеря Господня")
                    lvi.ForeColor = Color.Sienna;

                lv.Items.Add(lvi);
            }            
        }

        private string set_criteria(string crit)
        {
            string criteria = "";
            string start_date = get_start_date(); // DateTime.Now.Date.ToShortDateString();
            DateTime d = new DateTime();

            switch (crit)
            {
                case "last_mo":
                    d = Convert.ToDateTime(start_date).AddDays(-30.4);
                    criteria = start_date + ";" + d.ToShortDateString();
                    break;
                case "last_3_mo":
                    d = Convert.ToDateTime(start_date).AddDays(-91.2);
                    criteria = start_date + ";" + d.ToShortDateString();
                    break;
                case "last_6_mo":
                    d = Convert.ToDateTime(start_date).AddDays(-182.5);
                    criteria = start_date + ";" + d.ToShortDateString();
                    break;
                case "Lord_supp":
                    criteria = "2;0";
                    break;
                case "Holidays":
                    criteria = "3,4;5";
                    break;
                case "last_12_mo":
                    d = Convert.ToDateTime(start_date).AddDays(-365);
                    criteria = start_date + ";" + d.ToShortDateString();
                    break;
            }            
            return criteria;
        }

        private string get_start_date()
        {
            string today_date = DateTime.Now.Date.ToShortDateString(); // default
            string latest_date = "";
            string sql = sql_lib.get_start_date();
            ADODB.Recordset rs = null;

            if (mdb.RsOpen(out rs, sql) == false)
                return latest_date;

            if (rs.EOF == false)
                latest_date = Convert.ToDateTime(rs.Fields["latest_date"].Value.ToString()).ToShortDateString();

            if (Convert.ToDateTime(today_date) > Convert.ToDateTime(latest_date))
                latest_date = today_date;

            return latest_date;
        }

        public void delete_song_or_service(string serv_id, string song_id)
        {
            bool stat = jrn_service(serv_id, song_id);

            if (stat == false)
            {
                MessageBox.Show("Стереть не удалось!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sql = sql_lib.delete_song_or_service(serv_id, song_id);
            mdb.Execute(sql); 
        }

        private bool jrn_service(string serv_id, string song_id)
        {
            bool stat = false;
            string sql = sql_lib.get_service_song_data(serv_id, song_id);
            DateTime d = new DateTime();
            ADODB.Recordset rs = null;

            if (mdb.RsOpen(out rs, sql) == false)
                return stat;

            while (rs.EOF == false)
            {
                cService s = new cService();

                s.service_rid = rs.Fields["rid"].Value.ToString();
                s.service_id = rs.Fields["service_id"].Value.ToString();
                s.service_year_id = rs.Fields["service_year_id"].Value.ToString();
                s.song_id = rs.Fields["song_id"].Value.ToString();
                d = Convert.ToDateTime(rs.Fields["service_date"].Value.ToString());
                s.service_date = d.ToShortDateString();
                s.service_type = rs.Fields["service_type"].Value.ToString();
                s.service_note = rs.Fields["service_note"].Value.ToString();
                d = Convert.ToDateTime(rs.Fields["song_last_sang"].Value.ToString());
                s.song_last_sang = d.ToShortDateString();
                s.service_sys_date = rs.Fields["system_date"].Value.ToString();
                s.service_sys_type = rs.Fields["system_type"].Value.ToString();
                s.service_sys_cnt = rs.Fields["system_cnt"].Value.ToString();

                sql = sql_lib.jrn_service(s);
                stat = mdb.Execute(sql);

                rs.MoveNext();
            }

            rs = null;
            return stat;
        }

        public string get_max_song_id()
        {
            int max_somg_id = 0;
            string sql = sql_lib.get_max_song_id();
            ADODB.Recordset rs = null;

            if (mdb.RsOpen(out rs, sql) == false)
                return max_somg_id.ToString();

            if (rs.EOF == false)
                max_somg_id = Convert.ToInt32(rs.Fields["id"].Value) + 1;

            return max_somg_id.ToString();
        }

        public string get_path(string song_type_id)
        {
            string path = "";
            string sql = sql_lib.get_path(song_type_id);
            ADODB.Recordset rs = null;

            if (mdb.RsOpen(out rs, sql) == false)
                return path;

            if (rs.EOF == false)
                path = rs.Fields["path_location"].Value.ToString();

            return path;
        }

        public string get_path_id(string song_type_path)
        {
            string path_id = "";
            string sql = sql_lib.get_path_id(song_type_path);
            ADODB.Recordset rs = null;

            if (mdb.RsOpen(out rs, sql) == false)
                return path_id;

            if (rs.EOF == false)
                path_id = rs.Fields["path_type_id"].Value.ToString();

            return path_id;
        }

        public bool add_new_song(cSong s)
        {
            bool stat = false;
            string sql = sql_lib.add_new_song(s);

            if (mdb.Execute(sql) == true)
                sql = sql_lib.add_new_song_dtl(s);
            else
                return stat;

            if (mdb.Execute(sql) == true)
            {
                stat = move_song_home(s.song_name_path, s.song_text_path, s.song_text_path + "\\" + s.song_name + ".docx", false); // docm

                if (stat == true)
                    stat = create_preview_file(s.song_text_path + "\\" + s.song_name + ".docx", s.song_text_path + "\\Preview\\" + s.song_name + ".rtf");

                if (stat == true)
                    stat = create_chords_file(s.song_text_path + "\\" + s.song_name + ".docx");
            }            
            return stat;
        }

        private bool move_song_home(string source_file_name, string dest_folder_path, string dest_file, bool delete_source)
        {
            bool stat = false;

            if (Directory.Exists(dest_folder_path) == false)
                Directory.CreateDirectory(dest_folder_path);

            if (Directory.Exists(dest_folder_path) == true)
            {
                stat = File.Exists(source_file_name);
                if (stat == true)
                {
                    File.Copy(source_file_name, dest_file, true);
                    stat = File.Exists(dest_file);
                }
            }

            if (stat == true && delete_source == true)
                File.Delete(source_file_name);

            return stat;
        }

        public bool edit_song(cSong s, bool song_location_changed)
        {
            bool stat = false;
            string sql = sql_lib.update_song(s);

            if (mdb.Execute(sql) == true)
                sql = sql_lib.update_song_dtl(s);
            else
                return stat;

            stat = mdb.Execute(sql);

            if (stat == true)
            {
                if (song_location_changed == true)
                {
                    // move song
                    stat = move_song_home(s.orig_song_path + "\\" + s.song_name + ".docx", s.song_text_path, s.song_text_path +
                        "\\" + s.song_name + ".docx", true); // .docm

                    // move preview
                    if (stat == true)
                        stat = move_song_home(s.orig_song_path + "\\Preview\\" + s.song_name + ".rtf", s.song_text_path +
                        "\\Preview", s.song_text_path + "\\Preview\\" + s.song_name + ".rtf", true);

                    // move chords
                    if (stat == true)
                    {
                        s.source_chords_path = get_source_chords_path(s.orig_song_path);
                        s.dest_chords_path = get_source_chords_path(s.song_text_path);

                        if (s.source_chords_path.Length > 0)
                            stat = move_song_home(s.source_chords_path + "\\" + s.song_name + ".docx", s.dest_chords_path, s.dest_chords_path +
                            "\\" + s.song_name + ".docx", true);
                    }
                }
            }
            return stat;
        }

        private bool create_preview_file(string sourse_path, string dest_path)
        {
            bool stat = true;
            Word.Application newApp = new Word.Application();
            object source = sourse_path;
            object target = dest_path;
            object o = Type.Missing;
            object save = false;
            object format = Word.WdSaveFormat.wdFormatRTF;

            try
            {
                newApp.Documents.Open(ref source, ref o, ref o, ref o, ref o,
                    ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);

                newApp.ActiveDocument.SaveAs(ref target, ref format, ref o, ref o, ref o,
                    ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);

                newApp.Quit(ref o, ref o, ref o);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newApp);
                newApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                stat = false;
            }
            return stat;
        }

        private bool create_chords_file(string source_path)
        {
            bool stat = true;
            string[] key = source_path.Split('\\');
            string path = get_source_chords_path(source_path);

            move_song_home(source_path, path, path + "\\" + key[4], false);

            return stat;
        }

        private string get_source_chords_path(string work_path)
        {
            string path = "";
            string[] key = work_path.Split('\\');
            string path_id = "";

            if (key[3] == "Worship_songs")
                path_id = "6";
            else if (key[3] == "Group_songs")
                path_id = "7";
            else if (key[3] == "Kids_songs")
                path_id = "8";

            path = get_path(path_id);

            return path;
        }

        public bool remove_song(cSong s)
        {
            bool stat = false;
            string sql = sql_lib.remove_song(s);

            stat = mdb.Execute(sql);

            return stat;
        }

    }
}
