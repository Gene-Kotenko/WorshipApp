using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Worship
{
    public partial class fNewSong : Form
    {
        cUtil utl = null;
        private string key_mod = "";
        private char ru_let = ' ';
        private XmlDocument xml_doc = new XmlDocument();
        private bool add_new_song = false;
        private bool song_changed = false;
        private bool loading_form = false;
        private string[] song_key = null;
        private string orig_song_type = "";
        private string orig_song_path = "";
        private bool RU = true;

        public fNewSong(cUtil u, bool add_new, string[] key, bool ru)
        {
            InitializeComponent();
            utl = u;
            RU = ru;
            add_new_song = add_new;
            song_key = key;
            init_form();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void init_form()
        {
            loading_form = true;
            utl.fill_cbo(ref cboSongCat, "song_cat");
            utl.fill_cbo(ref cboSongType, "song_type");
            utl.fill_cbo(ref cboWorshipTime, "song_worship_time");
            set_ru_btn();

            if (add_new_song == true)
            {
                txtWorshipSongNum.Text = utl.get_max_song_id();
                txtWorshipSongNum.Tag = "new_song";
                btnDelSong.Visible = false;
                tsSep2.Visible = false;
            }
            else
            {
                if (song_key != null)
                    populate_form_with_song_info();
            }

            if (init_xml(Application.StartupPath + "\\data\\user.dat") == false)
                MessageBox.Show("XML не загрузился.", "XML Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            loading_form = false;
        }

        private void set_ru_btn()
        {
            if (RU == true)
                btnRuEn.Text = "RU";
            else
                btnRuEn.Text = "EN";
        }

        private void populate_form_with_song_info()
        {
            txtSongName.Text = song_key[2];
            txtWorshipSongNum.Text = song_key[10];
            txtGenrSongNum.Text = song_key[8];
            txtSongKey.Text = song_key[12];
            txtSongPhono.Text = song_key[14];
            txtSongNote.Text = song_key[6];
            orig_song_path = song_key[16];

            cboSongCat.SelectedIndex = utl.set_cbo_itm(cboSongCat.Tag.ToString(), song_key[20]);
            cboSongType.SelectedIndex = utl.set_cbo_itm(cboSongType.Tag.ToString(), song_key[22]);
            orig_song_type = (cboSongType.SelectedIndex + 1).ToString(); // to compare later
            cboWorshipTime.SelectedIndex = utl.set_cbo_itm(cboWorshipTime.Tag.ToString(), song_key[24]);            

            txtWorshipSongNum.Tag = "edit_song";
            btnDelSong.Visible = true;
            tsSep2.Visible = true;
            this.Icon = new Icon(Application.StartupPath + "\\data\\pic\\edit_song.ico");            
            this.Text = "Изменить Песню";
            lblHader.Text = " ИЗМЕНИТЬ ПЕСНЮ";
        }

        private void btnGetSong_Click(object sender, EventArgs e)
        {
            string path = "";
            string[] temp = null;
            string song_name = "";

            dlgFile = new OpenFileDialog();
            dlgFile.ShowDialog();
            path = dlgFile.FileName.ToString().Trim();
            temp = path.Split('\\');
            song_name = temp[temp.Length - 1].Replace(".docm", "");
            song_name = temp[temp.Length - 1].Replace(".docx", "");

            txtSongName.Text = song_name;
            lblGetSongPath.Text = path;
        }

        private void btnSaveSong_Click(object sender, EventArgs e)
        {
            if (validate_song() == true)
            {
                cSong s = null;

                if (add_new_song == true)
                {
                    if (MessageBox.Show("Вы хотите сохранить новую песню?", "Сохранить", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        s = refresh_song();

                        if (utl.add_new_song(s) == true)
                            MessageBox.Show("Песня успешно сохранена.", "Сохранить Песню", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            MessageBox.Show("Песня не сохранилась.", "Сохранить Песню", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                        return;
                }
                else
                {
                    if (song_changed == true)
                    {
                        if (MessageBox.Show("Вы хотите изменить выбранную песню?", "Изменить", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            s = refresh_song();

                            if (utl.edit_song(s, s.move_song) == true)
                                MessageBox.Show("Песня успешно изменена.", "Изменить Песню", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("Изменения не сохранились.", "Изменить Песню", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                            return;
                    }
                    song_changed = false; // reset
                }
            }
            else
                MessageBox.Show("Пожалуйста заполните все \nнеобходимые детали песни.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private bool validate_song()
        {
            bool stat = false;

            if (txtSongName.Text.Length == 0)
            {
                lblSongName.ForeColor = Color.Red;
                txtSongName.Tag = lblSongName;
                stat = false;
                return stat;
            }
            else if (txtSongKey.Text.Length == 0)
            {
                lblSongKey.ForeColor = Color.Red;
                txtSongKey.Tag = lblSongKey;
                stat = false;
                return stat;
            }
            else if (txtSongPhono.Text.Length == 0)
            {
                lblSongPhono.ForeColor = Color.Red;
                txtSongPhono.Tag = lblSongPhono;
                stat = false;
                return stat;
            }
            else if (txtSongNote.Text.Length == 0)
            {
                lblSongNote.ForeColor = Color.Red;
                txtSongNote.Tag = lblSongNote;
                stat = false;
                return stat;
            }
            else
            {
                stat = true;
            }

            return stat;
        }

        private cSong refresh_song()
        {
            cSong s = new cSong();

            s.song_id = txtWorshipSongNum.Text.Trim();
            s.song_name = txtSongName.Text.Trim();
            s.song_name_path = lblGetSongPath.Text; // to save later in the right directory
            s.song_general_num = txtGenrSongNum.Text.Trim();
            if (s.song_general_num == "")
                s.song_general_num = "-";
            
            s.song_worship_num = txtWorshipSongNum.Text.Trim();
            s.song_key = txtSongKey.Text.Trim();
            s.song_phonogram = txtSongPhono.Text.Trim();
            s.song_cat_id = utl.get_cbo_itm_key(cboSongCat);
            s.song_type_id = utl.get_cbo_itm_key(cboSongType);
            s.song_worship_prd_id = utl.get_cbo_itm_key(cboWorshipTime);
            s.song_note = txtSongNote.Text.Trim();
            s.song_text_path = utl.get_path(s.song_type_id);
            s.song_text_path_id = utl.get_path_id(s.song_text_path);

            //s.phonogram_path = utl.get_path("4"); // not necessary

            if (orig_song_type != s.song_type_id && orig_song_type != "")
            {
                s.orig_song_path = orig_song_path;
                s.move_song = true;
            }

            return s;
        }

        private void txt_changed_Click(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;

            if (txt.Tag != null)
            {
                Label l = (Label)txt.Tag;

                if (l.ForeColor == Color.Red)
                    l.ForeColor = Color.Black;
            }

            if (loading_form == false)
                song_changed = true;
        }

        private void btnDelSong_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы хотите удалить выбранную песню?", "Удалить", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                cSong s = new cSong();
                s.song_state = "9";
                s.song_id = song_key[4];

                if (utl. remove_song(s) == true )
                    MessageBox.Show("Песня успешно удалена.", "Удалить", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("Песня не удалена.", "Удалить", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                this.Close();
            }
            else
                return;
        }

        private bool init_xml(string file_name)
        {
            StreamReader read = null;
            string xml = "";
            string[] ln = null;
            char dq = '"';
            bool stat = false;

            if (file_name.Length == 0)
            {
                xml_doc = null;
                return stat;
            }

            try
            {
                read = new StreamReader(file_name);
                if (read != null)
                {
                    while (read.EndOfStream == false)
                    {
                        ln = read.ReadLine().Split(";".ToCharArray());
                        xml += "\n <ru_en ru_let=" + dq + ln[0] + dq + " en_let=" + dq + ln[1] + dq + " key=" + dq + ln[2] + dq + " />";
                    }
                }

                if (xml.Length > 0)
                {
                    xml = "<?xml version=\"1.0\"?>" +
                        "<root>" + xml + "</root>";
                    xml_doc = new XmlDocument();
                    xml_doc.LoadXml(xml);
                    stat = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return stat;
        }

        private void txtSongName_KeyDown(object sender, KeyEventArgs e)
        {
            string key = e.Modifiers.ToString().ToLower();
            string key2 = e.KeyCode.ToString().ToLower();

            key_mod = "";
            if (key.StartsWith("shift"))
                key_mod = "Shift";
            if (key.StartsWith("control"))
                key_mod = "Ctrl";
            else if (key.ToString().StartsWith("alt"))
                key_mod = "Alt";
        }

        private void txtSongName_KeyPress(object sender, KeyPressEventArgs e)
        {
            string en_let = e.KeyChar.ToString();

            switch (en_let)
            {
                case ";":
                    en_let = "cc";
                    break;
                case "\'":
                    en_let = "sq";
                    break;
                case "\"":
                    en_let = "dq";
                    break;
            }

            if (RU == true)
            {
                if (set_let(en_let) == true)
                    e.KeyChar = ru_let;
            }
        }

        private bool set_let(string en_let)
        {
            XmlNode e = null;
            string path = "";
            bool stat = false;
            ru_let = ' ';

            if (xml_doc == null)
                return stat;

            try
            {
                path = "//ru_en[@en_let='" + en_let + "' and @key='" + key_mod + "']";
                e = xml_doc.SelectSingleNode(path);
                if (e != null)
                {
                    stat = true;
                    ru_let = e.Attributes["ru_let"].Value.ToCharArray()[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return stat;
        }

        private void cboSongCat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading_form == false)
                song_changed = true;
        }

        private void cboSongType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading_form == false)
                song_changed = true;
        }

        private void cboWorshipTime_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading_form == false)
                song_changed = true;
        }

        private void btnRuEn_Click(object sender, EventArgs e)
        {
            if (btnRuEn.Text == "RU")
            {
                btnRuEn.Text = "EN";
                RU = false;
            }
            else
            {
                btnRuEn.Text = "RU";
                RU = true;
            }
        }

    }
}
