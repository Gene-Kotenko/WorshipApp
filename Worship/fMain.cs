using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Worship
{
    public partial class fMain : Form
    {
        public enum FormView { Служения, Спевка, Песни_Ноты };
        private FormView f_view = FormView.Песни_Ноты;
        private ToolStripMenuItem cur_item = null;
        private XmlDocument xml_doc = new XmlDocument();
        private char ru_let = ' ';
        private string key_mod = "";
        private string service_crit = "last_mo"; // Default
        private string spevka_orig_note = "";
        private bool spevka_changed = false;
        private bool song_removed_from_spevka = false;
        private bool RU = true;
        
        cUtil utl = null;

        public fMain()
        {
            InitializeComponent();
            utl = new cUtil();
            init_form();
        }

        private void init_form()
        {
            this.DoubleBuffered = true;
            set_pnl(FormView.Служения);

            utl.fill_cbo(ref cboSongCat, "song_cat");
            utl.fill_cbo(ref cboSongType, "song_type");
            utl.fill_cbo(ref cboWorshipTime, "song_worship_time");

            utl.init_Tree(ref tvSong, FormView.Песни_Ноты.ToString());
            utl.init_Tree(ref tvSpevka, FormView.Спевка.ToString());

            utl.get_services(ref lvService, service_crit);

            if (init_xml(Application.StartupPath + "\\data\\user.dat") == false)
                MessageBox.Show("XML не загрузился.", "XML Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            pnlToolBar.Width = 320;
            lblToday.Text = "Today: " + DateTime.Today.ToString("MMMM") + " " + DateTime.Today.ToString("dd") + ", " + DateTime.Today.ToString("yyyy");
        }

        private void set_pnl(FormView view_type)
        {
            if (view_type == FormView.Служения)
            {                
                pnlService.Dock = DockStyle.Fill;
                pnlSong.Visible = false;
                pnlMainLeft.Visible = false;
                pnlSpevka.Visible = false;
                pnlMainTopLeft.Height = 0;
                //btnPrint.Visible = false;
                //spSep5.Visible = false;
                lblInfo.Text = " Служения";
                f_view = FormView.Служения;
                pnlService.Visible = true;
            }
            else if (view_type == FormView.Спевка)
            {
                pnlMainLeft.Visible = true;
                pnlSpevka.Dock = DockStyle.Fill;
                pnlSong.Dock = DockStyle.None;
                lblInfo.Text = " Спевка";
                pnlSong.Visible = false;
                pnlMainTopLeft.Size = new Size(1026, 90);
                btnSearchSong.Visible = false;
                mnuSearch.Visible = false;
                mnuAscDesc.Visible = true;
                tvSpevka.Dock = DockStyle.Fill;
                tvSpevka.BringToFront();
                //btnPrint.Visible = false;
                //spSep5.Visible = false;
                f_view = FormView.Спевка;
                pnlService.Visible = false;
                pnlSpevka.Visible = true;
            }
            else if (view_type == FormView.Песни_Ноты)
            {
                btnSearchSong.Visible = true;
                pnlSongChords.Visible = true;
                pnlMainLeft.Visible = true;
                pnlService.Visible = false;
                pnlSong.Visible = true;
                pnlSong.Dock = DockStyle.Fill;
                pnlSpevka.Visible = false;
                pnlSpevka.Dock = DockStyle.None;
                lblInfo.Text = " Песни/Ноты";
                pnlMainTopLeft.Size = new Size(1026, 0);
                mnuSearch.Visible = true;
                mnuAscDesc.Visible = false;
                txtSearch.Text = "";
                tvSong.Dock = DockStyle.Fill;
                tvSong.BringToFront();
                //btnPrint.Visible = true;
                //spSep5.Visible = true;
                f_view = FormView.Песни_Ноты; 
            }
        }

        private void fMain_Activated(object sender, EventArgs e)
        {
            txtSearch.Focus();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnСлужения_Click(object sender, EventArgs e)
        {
            set_pnl(FormView.Служения);

            if (lvService.Items.Count == 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                utl.get_services(ref lvService, service_crit);
                Cursor.Current = Cursors.Default;
            }
            lblInfo.Text = " Служения";
        }

        private void btnПесниНоты_Click(object sender, EventArgs e)
        {
            if (pnlSong.Visible == true && tvSong.Nodes.Count > 0)
                return;
            set_pnl(FormView.Песни_Ноты);
        }

        private void tvSong_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.ImageIndex == 3)
            {
                e.Node.ImageIndex = 2;
                e.Node.SelectedImageIndex = 2;
            }
        }

        private void tvSong_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode n = e.Node;
            string[] key = n.Tag.ToString().Split(';');
            tvSong.BeginUpdate();

            if (n.ImageIndex == 2)
            {
                n.ImageIndex = 3;
                n.SelectedImageIndex = 3;
            }

            utl.process_node(ref n);

            if (f_view == FormView.Песни_Ноты && key[0] == "ТЕКСТ_ПЕСНИ")
            {
                n.Nodes[0].Text += " (Текст)";
                n.Nodes[1].Text += " (Пиано)";
            }

            tvSong.EndUpdate();
        }

        private string get_song_params()
        {
            string song_params = "";
            cSong s = new cSong();

            s.song_cat_id = utl.get_cbo_itm_key(cboSongCat);
            s.song_type_id = utl.get_cbo_itm_key(cboSongType);
            s.song_worship_prd_id = utl.get_cbo_itm_key(cboWorshipTime);

            if (cur_item != null)
            {
                if (cur_item.Name == "mAscByName")
                    s.sort_by_order = mAscByName.Tag.ToString();
                else if (cur_item.Name == "mDescByName")
                    s.sort_by_order = mDescByName.Tag.ToString();
                else if (cur_item.Name == "mAscByDate")
                    s.sort_by_order = mAscByDate.Tag.ToString();
                else if (cur_item.Name == "mDescByDate")
                    s.sort_by_order = mDescByDate.Tag.ToString();
            }

            song_params = "ПЕСНИ_ПО_КАТЕГОРИЯМ;" + "song_cat_id=;" + s.song_cat_id + ";song_type_id=;" +
                s.song_type_id + ";song_worship_prd_id=;" + s.song_worship_prd_id + ";song_sort_by_order=;" + s.sort_by_order;

            return song_params;
        }

        private void tvSong_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (tvSong.Nodes.Count == 0)
                return;
            if (e.Node == null)
                return;
            int ico = e.Node.SelectedImageIndex;
            string[] key = null;
            string file_name = "";

            try
            {
                if (f_view == FormView.Песни_Ноты)
                {
                    if (ico == 46 || ico == 51)
                        return;

                    if (ico == 47)
                    {
                        richTbPrv.Clear();

                        key = e.Node.Tag.ToString().Split(';');
                        lblSongWorshipNum.Text = key[10];
                        lblSongGenrlNum.Text = key[8];
                        lblSongDetNote.Text = key[6];
                        lblSongKey.Text = key[12];
                        lblSongPhonogram.Text = key[14];
                        lblSongDetlCat.Text = key[20];
                        lblSongDetlType.Text = key[22];
                        lblSongDetlSubType.Text = key[24];

                        file_name = key[16] + "\\Preview\\" + e.Node.Text + ".rtf";
                        utl.preview_song(ref richTbPrv, file_name);
                    }
                    else
                        clean_song_det_info();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void clean_song_det_info()
        {
            lblSongWorshipNum.Text = "-";
            lblSongGenrlNum.Text = "-";
            lblSongDetNote.Text = "-";
            lblSongKey.Text = "-";
            lblSongPhonogram.Text = "-";
            lblSongDetlCat.Text = "-";
            lblSongDetlType.Text = "-";
            lblSongDetlSubType.Text = "-";
            richTbPrv.Clear();
        }

        private void tvSong_DoubleClick(object sender, EventArgs e)
        {
            if (tvSong.Nodes.Count == 0)
                return;
            if (tvSong.SelectedNode == null)
                return;
            TreeNode n = tvSong.SelectedNode;
            string[] key = null;
            string file_name = "";

            if (n.SelectedImageIndex == 54) // 48
            {
                key = n.Tag.ToString().Split(';');
                file_name = key[2] + "\\" + n.Text;
                utl.open_song_file(file_name, "Windows Media Player", true);
            }
            else
            {
                if (n.SelectedImageIndex == 46 || n.SelectedImageIndex == 44 || n.SelectedImageIndex == 51)
                {
                    key = n.Tag.ToString().Split(';');
                    string[] tmp = n.Text.Split('(');
                    string song_name = tmp[0].Trim();
                    file_name = key[2] + "\\" + song_name + ".docx";
                    utl.open_song_file(file_name, "word", true);
                }
            }
        }

        private void pnlToolBar_Resize(object sender, EventArgs e)
        {
            txtSearch.Width = pnlToolBar.Width - mnuSearch.Width - btnSearchSong.Width - 9;
        }

        private void cboSongCat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tvSong.Nodes.Count > 0)
                tvSpevka.Nodes[0].Collapse();
        }

        private void cboSongType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tvSong.Nodes.Count > 0)
                tvSpevka.Nodes[0].Collapse();
        }

        private void cboWorshipTime_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tvSong.Nodes.Count > 0)
                tvSpevka.Nodes[0].Collapse();
        }

        private void tvBuildService_MouseDown(object sender, MouseEventArgs e)
        {
            if (tvBuildSpevka.SelectedNode == null)
                return;
            int ico = 0;
            string txt = "";

            if (e.Button == MouseButtons.Right)
            {
                ico = tvBuildSpevka.SelectedNode.SelectedImageIndex;

                if (f_view == FormView.Спевка)
                {
                    if (ico == 3)
                    {
                        txt = tvBuildSpevka.SelectedNode.Text;
                        btnRemove.Image = (Image)imgTvMember2.Images[tvBuildSpevka.SelectedNode.ImageIndex];
                        btnSongUp.Visible = false;
                        btnSongDown.Visible = false;
                    }
                    else if (ico == 47)
                    {
                        txt = tvBuildSpevka.SelectedNode.Text;
                        btnRemove.Image = (Image)imgTvMember2.Images[55];
                        btnSongUp.Visible = true;
                        btnSongDown.Visible = true;
                    }

                    if (txt.Length > 0)
                    {
                        txt = "Удалить: " + txt;
                        btnRemove.Text = txt;
                        btnRemove.Tag = tvBuildSpevka.SelectedNode;
                        cmSelection.Show(tvBuildSpevka, new Point(e.X, e.Y));
                    }
                }
            }
        }

        private void tvBuildService_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.ImageIndex == 3)
            {
                e.Node.ImageIndex = 2;
                e.Node.SelectedImageIndex = 2;
            }
        }

        private void tvBuildService_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode n = e.Node;
            Cursor.Current = Cursors.WaitCursor;
            tvSong.BeginUpdate();

            if (n.ImageIndex == 2)
            {
                n.ImageIndex = 3;
                n.SelectedImageIndex = 3;
            }

            utl.process_node(ref n);
            tvSong.EndUpdate();
            Cursor.Current = Cursors.Default;
        }

        private void mnuAdd_Click(object sender, EventArgs e)
        {
            TreeNode n = null;
            TreeNode sub = null;
            TreeNode nd = (TreeNode)mnuAdd.Tag;
            spevka_changed = true;

            if (nd != null)
            {
                if (tvBuildSpevka.Nodes.Count == 0)
                {
                    n = new TreeNode("ВЫБРАННЫЕ ПЕСНИ", 3, 3);
                    n.Tag = "ВЫБРАННЫЕ_ПЕСНИ;";
                    sub = new TreeNode(nd.Text, 47, 47);
                    sub.Tag = nd.Tag;
                    n.Nodes.Add(sub);
                    tvBuildSpevka.Nodes.Add(n);
                    n.Expand();
                }
                else
                {   // prevent duplicates
                    foreach (TreeNode node in tvBuildSpevka.Nodes[0].Nodes)
                    {
                        if (node.Text == nd.Text)
                            return;
                    }

                    n = new TreeNode(nd.Text, 47, 47);
                    n.Tag = nd.Tag;
                    tvBuildSpevka.Nodes[0].Nodes.Add(n);
                }
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (tvBuildSpevka.Nodes.Count == 0 || tvBuildSpevka.SelectedNode == null)
                return;

            TreeNode n = tvBuildSpevka.SelectedNode;
            tvBuildSpevka.Nodes.Remove(n);
            song_removed_from_spevka = true;
           
            if (tvBuildSpevka.Nodes.Count > 0)
            {
                if (tvBuildSpevka.Nodes[0].Nodes.Count == 0)
                    tvBuildSpevka.Nodes.Clear();
            }
        }

        private void btnRefreshSongs_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (txtSearch.Text.Trim().Length > 0)
                utl.get_song_by_name(ref tvSpevka, txtSearch.Text.Trim());
            else
            {
                if (tvSpevka.Nodes.Count > 0)
                    tvSpevka.Nodes[0].Expand();
            }
            Cursor.Current = Cursors.Default;
        }

        private void btnSearchSong_Click(object sender, EventArgs e)
        {
            string criteria = txtSearch.Text.Trim();

            if (criteria.Length == 0)
                return;
            if (tvSong.Nodes[0].IsExpanded == false)
                tvSong.Nodes[0].Expand();

            if (mnuSearch.Tag.ToString() == "song_name")
                searsh_song_by_name(criteria);
            else
            {
                if (utl.IsNumeric(criteria) == true)
                    search_song_by_id(criteria, mnuSearch.Tag.ToString());
            }
        }

        private void searsh_song_by_name(string song_name)
        {
            utl.get_song_by_name(ref tvSong, song_name);
        }

        private void search_song_by_id(string id, string type)
        {
            if (id.Length == 0 || type == "")
                return;

            string worship_id = "";
            string general_id = "";

            if (type == "worship_id")
                worship_id = id;
            else if (type == "genr_id")
                general_id = id;

            utl.get_song_by_id(ref tvSong, worship_id, general_id);
        }

        private void btnNewService_Click(object sender, EventArgs e)
        {
            tvBuildSpevka.Nodes.Clear();
            tvBuildSpevka.Tag = null;    // reset for the new spevka
            dtPick.Value = DateTime.Now; // reset for the new spevka
            txtServiceNote.Text = "";
            label10.Text = "Песни для следующего служения";
            song_removed_from_spevka = false;
        }

        private void btnSaveService_Click(object sender, EventArgs e)
        {
            if (validate_spevka() == false)
                return;

            bool stat = true;
            Cursor.Current = Cursors.WaitCursor;
            cService serv = new cService();

            serv.service_type = utl.get_cbo_itm_key(cboSongCat);
            serv.service_year = DateTime.Now.Year.ToString();
            //serv.service_date = DateTime.Now.ToShortDateString();
            serv.service_date = dtPick.Value.ToShortDateString(); // Changed logic so I can set the date if needed
            serv.service_year_id = serv.service_year.Substring(2, 2);
            serv.song_last_sang = serv.service_date;
            serv.service_note = txtServiceNote.Text.Trim();
            serv.obj = tvBuildSpevka.Nodes[0];

            if (tvBuildSpevka.Tag == null) // Save new spevka
            {
                serv.service_id = utl.get_next_service_id();
                stat = utl.save_spevka(serv);
                utl.get_latest_spevka(ref tvBuildSpevka, txtServiceNote, dtPick);

                if (stat == false)
                    MessageBox.Show("Служение не сохранено.", "Сохранить Служение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show("Служение успешно сохранено.", "Сохранить Служение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else // update spevka
            {
                serv.service_id = tvBuildSpevka.Tag.ToString();
                stat = utl.update_spevka(serv);                

                if (stat == false)
                    MessageBox.Show("Служение не пересохранино.", "Пересохранить Служение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                {
                    MessageBox.Show("Служение успешно пересохранино.", "Пересохранить Служение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    spevka_changed = false;
                    spevka_orig_note = txtServiceNote.Text; // reset note
                }
            }
            Cursor.Current = Cursors.Default;
        }

        private bool validate_spevka()
        {
            bool stat = false;

            if (tvBuildSpevka.Nodes.Count == 0)
                return stat;

            string msg1 = "Служение не пересохранино. \nПесню необходимо удалять в отделе 'Служения'.";
            string msg2 = "Песен не достаточно для полного служения. \nХотите ли вы всеравно сохранить?";

            if (spevka_changed == true || txtServiceNote.Text != spevka_orig_note)
                stat = true;

            if (tvBuildSpevka.Nodes[0].Nodes.Count < 4)
                stat = (MessageBox.Show(msg2, "Save Service", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes);

            if (spevka_changed == false && song_removed_from_spevka == true )
                MessageBox.Show(msg1, "Пересохранить Служение", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            return stat;
        }

        private void mnu_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem itm = (ToolStripMenuItem)sender;
            cur_item = itm;
            mnuAscDesc.Tag = itm.Tag;
            mnuAscDesc.Image = itm.Image;

            if (tvSpevka.Nodes.Count > 0)
                tvSpevka.Nodes[0].Collapse();
        }

        private void mnu2_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem itm = (ToolStripMenuItem)sender;
            cur_item = itm;
            mnuSearch.Tag = itm.Tag;
            mnuSearch.Image = itm.Image;

            if (tvSong.Nodes[0].Text == "ПЕСНИ ПО ЗАДАННОМУ ПОИСКУ")
                utl.init_Tree(ref tvSong, FormView.Песни_Ноты.ToString());
        }

        private bool init_xml(string file_name)
        {
            StreamReader read;
            string xml = "";
            string[] ln;
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

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            string key = e.Modifiers.ToString().ToLower();
            string key2 = e.KeyCode.ToString().ToLower();
            System.EventArgs ea = new EventArgs();

            key_mod = "";
            if (key.StartsWith("shift"))
                key_mod = "Shift";
            if (key.StartsWith("control"))
                key_mod = "Ctrl";
            else if (key.ToString().StartsWith("alt"))
                key_mod = "Alt";

            if (key2 == "return" && f_view == FormView.Спевка || key2 == "enter" && f_view == FormView.Спевка
                || key2 == "return" && f_view == FormView.Песни_Ноты || key2 == "enter" && f_view == FormView.Песни_Ноты)
            {
                if (f_view == FormView.Спевка)
                    btnRefreshSongs_Click(sender, ea);

                else if (f_view == FormView.Песни_Ноты)
                    btnSearchSong_Click(sender, ea);
            }
        }

        private void txtSearch_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text.Trim().Length == 0 && tvSpevka.Nodes[0].Text == "ПЕСНИ ПО ЗАДАННОМУ ПОИСКУ"
                || txtSearch.Text.Trim().Length == 0 && tvSong.Nodes[0].Text == "ПЕСНИ ПО ЗАДАННОМУ ПОИСКУ")
            {

                if (f_view == FormView.Спевка)
                {
                    tvSpevka.Nodes[0].Collapse();
                    tvSpevka.Nodes[0].Expand();
                }
                else                
                    utl.init_Tree(ref tvSong, FormView.Песни_Ноты.ToString());
            }
        }

        public bool PrintFile(string f_name) // for future
        {
            Process proc = new Process();
            bool stat = false;

            try
            {
                proc.StartInfo.FileName = f_name;
                proc.StartInfo.Verb = "print";
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                stat = true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return stat;
        }

        private void btnSpevka_Click(object sender, EventArgs e)
        {
            if (pnlMainTopLeft.Size.Height == 0)
            {
                txtSearch.Text = "";
                if (tvBuildSpevka.Nodes.Count == 0)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    utl.get_latest_spevka(ref tvBuildSpevka, txtServiceNote, dtPick);
                    Cursor.Current = Cursors.Default;
                }
               
                set_pnl(FormView.Спевка);
                spevka_changed = false;
                spevka_orig_note = txtServiceNote.Text;
                label10.Text = "Песни с последнего служения";
                song_removed_from_spevka = false;
            }
            else
            {
                if (tvBuildSpevka.Nodes.Count == 0)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    utl.get_latest_spevka(ref tvBuildSpevka, txtServiceNote, dtPick);
                    spevka_changed = false;
                    spevka_orig_note = txtServiceNote.Text;
                    song_removed_from_spevka = false;
                    Cursor.Current = Cursors.Default;
                }
            }
        }

        private void tvSpevka_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.ImageIndex == 3)
            {
                e.Node.ImageIndex = 2;
                e.Node.SelectedImageIndex = 2;
            }
        }

        private void tvSpevka_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode n = e.Node;
            string[] key = n.Tag.ToString().Split(';');
            tvSong.BeginUpdate();

            if (n.ImageIndex == 2)
            {
                n.ImageIndex = 3;
                n.SelectedImageIndex = 3;
            }

            if (n.ImageIndex == 3 && txtSearch.Text.Length == 0)
            {
                // setting back
                n.Text = "ПЕСНИ ПО КАТЕГОРИЯМ";
                n.Tag = get_song_params();
            }

            utl.process_node(ref n);
            tvSong.EndUpdate();
        }

        private void tvSpevka_DoubleClick(object sender, EventArgs e)
        {
            if (tvSpevka.Nodes.Count == 0)
                return;
            if (tvSpevka.SelectedNode == null)
                return;
            TreeNode n = tvSpevka.SelectedNode;

            if (n.SelectedImageIndex != 46)
                return;

            string[] key = n.Tag.ToString().Split(';');
            string[] tmp = n.Text.Split('(');
            string song_name = tmp[0].Trim();

            string file_name = key[2] + "\\" + song_name + ".docx";
            utl.open_song_file(file_name, "word", true);
        }

        private void tvSpevka_MouseDown(object sender, MouseEventArgs e)
        {
            if (tvSpevka.SelectedNode == null)
                return;
            int ico = 0;
            string txt = "";

            if (e.Button == MouseButtons.Right)
            {
                ico = tvSpevka.SelectedNode.SelectedImageIndex;

                if (f_view == FormView.Спевка)
                {
                    if (ico == 47)
                        txt = tvSpevka.SelectedNode.Text;

                    if (txt.Length > 0)
                    {
                        txt = "Добавить: " + txt;
                        mnuAdd.Text = txt;
                        mnuAdd.Tag = tvSpevka.SelectedNode;
                        //mnuAdd.Image = (Image)imgTvMember2.Images[tvSpevka.SelectedNode.ImageIndex];
                        cmAdd.Show(tvSpevka, new Point(e.X, e.Y));
                    }
                }
            }
        }

        private void btnSelectPrd_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem btn = (ToolStripMenuItem)sender;
            btnPrdSelection.Text = btn.Text;
            service_crit = btn.Tag.ToString();
        }

        private void btnRefreshServices_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            utl.get_services(ref lvService, service_crit);
            Cursor.Current = Cursors.Default;
        }

        private void lvService_MouseDown(object sender, MouseEventArgs e)
        {
            if (lvService.Items.Count == 0 || lvService.SelectedItems.Count == 0)
                return;
            if (lvService.SelectedItems[0] == null || lvService.SelectedItems[0].Text == "")
                return;
            ListViewItem lvi = null;

            if (e.Button == MouseButtons.Right)
            {
                lvi = lvService.SelectedItems[0];
                btnRemoveSong.Text = "Удалить: " + lvi.SubItems[1].Text;
                btnRemoveSong.Tag = lvi.Tag;
                btnRemoveServ.Tag = lvi.Tag;
                cmService.Show(lvService, new Point(e.X, e.Y));
            }
        }

        private void btnDelete_Song_or_Service_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem btn = (ToolStripMenuItem)sender;
            string[] key = btnRemoveSong.Tag.ToString().Split(';');
            string msg = "";
            string caption = "";

            if (btn.Name == "btnRemoveSong")
            {
                msg = "Вы хотите удалить: " + key[5] + "?";
                caption = "Удалить Песню";
            }
            else
            {
                msg = "Вы хотите удалить всё Служение?";
                caption = "Удалить Служение";
            }

            if (MessageBox.Show(msg, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (btn.Name == "btnRemoveSong")
                    utl.delete_song_or_service(key[1], key[3]);
                else
                    utl.delete_song_or_service(key[1], "");

                utl.get_services(ref lvService, service_crit);
                tvBuildSpevka.Nodes.Clear();
            }
        }

        private void btnAddNewSong_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string[] key = null; // nothing if new song

            fNewSong f = new fNewSong(this.utl, true, key, RU);
            f.ShowDialog();
            f = null;
            utl.init_Tree(ref tvSong, FormView.Песни_Ноты.ToString());
            richTbPrv.Clear();
            Cursor.Current = Cursors.Default;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // TO DO
        }

        private void btnSongUp_Click(object sender, EventArgs e)
        {
            move_doc(true);
        }

        private void btnSongDown_Click(object sender, EventArgs e)
        {
            move_doc(false);
        }

        private void move_doc(bool up)
        {
            if (tvBuildSpevka.Nodes.Count == 0)
                return;
            TreeNode doc = null;
            int index = 0;
            spevka_changed = true;

            foreach (TreeNode n in tvBuildSpevka.Nodes[0].Nodes)
            {
                if (n.IsSelected == true)
                {
                    doc = new TreeNode(n.Text, 47, 47);
                    index = n.Index;
                    if (up == true)
                    {
                        index--;
                        if (index < 0)
                            index = 0;
                    }
                    else
                    {
                        index++;
                        if (index >= tvBuildSpevka.Nodes[0].Nodes.Count)
                            index = tvBuildSpevka.Nodes[0].Nodes.Count - 1;
                    }

                    doc.Tag = n.Tag;
                    tvBuildSpevka.Nodes[0].Nodes.Remove(n);
                    tvBuildSpevka.Nodes[0].Nodes.Insert(index, doc);
                    return;
                }
            }
        }

        private void btnComposeEmail_Click(object sender, EventArgs e)
        {
            if (tvBuildSpevka.Nodes.Count == 0)
                return;
            txtBuildEmail.Text = "";
            compose_email_from_spevka();
        }

        private void compose_email_from_spevka()
        {
            TreeNode n = tvBuildSpevka.Nodes[0];
            string line = "";
            string[] key = null;
            int i = 0;

            foreach (TreeNode nd in n.Nodes)
            {
                i++;
                key = nd.Tag.ToString().Split(';');

                if (key.Length <= 7)
                {
                    MessageBox.Show("Сохраните спевку перед тем как составить Email.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    if (i == 1)
                        line += "\r\n" + " " + i.ToString() + ". " + key[8] + " " + key[2] + "\r\n"; // + " (" + key[10] + ")"
                    else
                        line += " " + i.ToString() + ". " + key[8] + " " + key[2] + "\r\n"; // + " (" + key[10] + ")"
                }
            }

            line += "\r\n" + "\r\n" + " " + "Благослови вас Господь.";
            txtBuildEmail.Text = line;
        }

        private void tvSong_MouseDown(object sender, MouseEventArgs e)
        {
            if (tvSong.SelectedNode == null)
                return;
            int ico = 0;
            string txt = "";

            if (e.Button == MouseButtons.Right)
            {
                ico = tvSong.SelectedNode.SelectedImageIndex;

                if (f_view == FormView.Песни_Ноты)
                {
                    if (ico == 47)
                        txt = tvSong.SelectedNode.Text;

                    if (txt.Length > 0)
                    {
                        txt = "Изменить: " + txt;
                        btnChangeSong.Text = txt;
                        btnChangeSong.Tag = tvSong.SelectedNode.Tag;
                        //mnuAdd.Image = (Image)imgTvMember2.Images[tvSong.SelectedNode.ImageIndex];
                        cmChangeSong.Show(tvSong, new Point(e.X, e.Y));
                    }
                }
            }
        }

        private void btnChangeSong_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string[] key = btnChangeSong.Tag.ToString().Split(';');

            fNewSong f = new fNewSong(this.utl, false, key, RU);
            f.ShowDialog();
            f = null;
            utl.init_Tree(ref tvSong, FormView.Песни_Ноты.ToString());
            richTbPrv.Clear();
            Cursor.Current = Cursors.Default;
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
