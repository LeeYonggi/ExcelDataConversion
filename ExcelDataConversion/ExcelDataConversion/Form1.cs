using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDataConversion
{
    public partial class Form1 : Form
    {
        DataTable nowDataTable = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Load_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            
            fileDialog.Title = "엑셀파일 오픈 창";
            fileDialog.FileName = "C:/Users";
            fileDialog.Filter = "엑셀 파일 (*.*) | *.xlsx*";
            
            DialogResult dr = fileDialog.ShowDialog();
            
            if(dr == DialogResult.OK)
            {
                string fileName = fileDialog.FileName;

                nowDataTable = ExcelDataParser.ExcelFileRead(fileName);
            }
        }

        private void Save_Json_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();

            saveDialog.Title = "Json 저장 창";
            saveDialog.Filter = "Json 파일 (*.*) | *.json*";

            DialogResult dr = saveDialog.ShowDialog();

            if (dr == DialogResult.OK)
            {
                string fileName = saveDialog.FileName;

                ExcelDataParser.DataTableToJson(fileName + ".json", nowDataTable);
            }
        }
    }
}
