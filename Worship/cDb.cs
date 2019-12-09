using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Worship
{
    public class cDb
    {
        private string cnn_str = "";
        private ADODB.Connection cnn = null;

        public cDb(string dsn)
        {
            try
            {
                //dsn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dsn; // original one (but its not working properly when querying)
                dsn = "driver={Microsoft Access Driver (*.mdb)}" + ";dbq=" + dsn;
                cnn = new ADODB.Connection();
                cnn.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
                cnn.Open(dsn, "", "", 0);
                cnn_str = dsn;
                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public bool Connected()
        {
            bool stat = false;
            stat = (cnn_str.Length > 0);
            return stat;
        }

        public bool RsOpen(out ADODB.Recordset rs, string sql)
        {
            bool stat = false;
            rs = new ADODB.Recordset();

            try
            {
                if (Connected() == true)
                {
                    rs.Open(sql, cnn_str, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, -1);             
                    stat = (rs.State == 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return stat;
        }

        public bool Execute(string sql)
        {
            bool stat = false;
            object obj;
            try
            {
                if (Connected() == true)
                {
                    if (cnn == null)
                        cnn = new ADODB.Connection();
                    if (cnn.State == 0)
                        cnn.Open(cnn_str, "", "", 0);
                    if (cnn.State == 1)
                    {
                        cnn.Execute(sql, out obj, 0);
                        cnn.Close();
                        stat = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return stat;
        }
    }
}
