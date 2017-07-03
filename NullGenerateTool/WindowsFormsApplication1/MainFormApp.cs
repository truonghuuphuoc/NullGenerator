using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace NULL_is_my_son
{
    public partial class MainAppForm : Form
    {
        private Database databaseInfor;
        private ParkingListParser packingParser;
        private PackingListGenerator packingGenerator;

        public MainAppForm()
        {
            InitializeComponent();
        }

        private void btnInputFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Exel Files (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = false;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;

                edtFileInput.Text = sFileName;
            }
        }

        private void btnBowseDatabase_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Exel Files (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = false;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                bool error = false;
                string sFileName = choofdlog.FileName;

                edtDatabasePath.Text = sFileName;

                this.databaseInfor = new Database();

                error = this.databaseInfor.CreateListDatabase(sFileName);

                if (!error)
                {
                    txtLog.Text += "Import data base err: " + this.databaseInfor.GetErrorCode() + "\r\n";
                }
            }
        }

        private void btnBrowseOutFile_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK)
                {
                    String time = DateTime.Now.ToString("yyyy_MM_dd_HH_mm");
                    String file = fbd.SelectedPath + "\\PackingList_" + time + ".xlsx";
                    edtOutFilePath.Text = file;
                }
            }
        }

        private void cbbCommandType_SelectedIndexChanged(object sender, EventArgs e)
        {
            intCommandType = cbbCommandType.SelectedIndex;
        }

        private void btnGenerateFile_Click(object sender, EventArgs e)
        {
            bool parse;

            if (this.databaseInfor == null)
            {
                txtLog.Text += "Database error:  Database file is not exits\r\n";

                return;
            }

            if (edtFileInput.Text == "")
            {
                txtLog.Text += "Output file error:  Path file is empty\r\n";

                return;
            }


            packingParser = new ParkingListParser(this.databaseInfor);

            String file = edtFileInput.Text;

            parse = packingParser.PackingListParser(file);

            if (!parse)
            {
                txtLog.Text += "Parse packing data error: " + this.packingParser.GetErrorCode() + "\r\n";

                return;
            }
            else
            {
                txtLog.Text += "Parse packing data information: " + this.packingParser.GetErrorCode() + "\r\n";
            }
            packingParser.MergerPackingList();

            if (intCommandType == 0)
            {
                packingGenerator = new PackingListGenerator(packingParser, databaseInfor, edtOutFilePath.Text);

                parse = packingGenerator.GeneratePackingListFile();

                if (!parse)
                {
                    txtLog.Text += "Generate file error: " + this.packingParser.GetErrorCode() + "\r\n";

                    return;
                }
                else
                {
                    txtLog.Text += "Generate file log: " + this.packingParser.GetErrorCode() + "\r\n";


                    txtLog.Text += "\r\n\r\n---------------------------------\r\n";
                    txtLog.Text += "---------Sucessfully-------------\r\n";
                }
            }

        }

        private void btnClearLog_Click(object sender, EventArgs e)
        {
            txtLog.Text = "";
        }
    }
}
