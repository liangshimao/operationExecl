using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Core;
using MSword = Microsoft.Office.Interop.Excel;

namespace HelloWorld
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(textBox1.Text.ToString());
            string path = textBox1.Text.ToString();

            DirectoryInfo TheFolder = new DirectoryInfo(path);
            if (!TheFolder.Exists) {
                MessageBox.Show("文件夹不存在！", "Warning");
                return;
            }

            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (Path.GetExtension(NextFile.Name) != ".xls" && Path.GetExtension(NextFile.Name) != ".xlsx")
                {
                    continue;
                }

                label2.Text = "正在处理" + NextFile.Name;
                // 获取文件完整路径
                string fileName = NextFile.FullName;
                //string extension = Path.GetExtension(heatmappath);
                //MessageBox.Show(heatmappath, "Warning");
                MSword.Application EXC1 = new MSword.Application();
                //EXC1.Visible = true;
                MSword.Workbooks wbs = EXC1.Workbooks;
                MSword.Workbook wb = wbs.Add(fileName);//打开并显示EXCEL文件
                MSword.Worksheet Exsheet;
                Exsheet = wb.Sheets[1];
                //Exsheet.Cells[1,1] = "liang";
                //Console.WriteLine(Exsheet.Name);
                int sRowcount = Exsheet.UsedRange.Columns.Count;
                Console.WriteLine(sRowcount);
                Exsheet.Activate();

                Exsheet.Cells[1, ++sRowcount] = "是否党员";
                Exsheet.Cells[1, ++sRowcount] = "交通工具";
                Exsheet.Cells[1, ++sRowcount] = "未住户情况";
                //Exsheet.Cells[1, 1] = "liang111";

                //object MisValue = Type.Missing;
                //string[] insertName = new string[] { "是否党员", "交通工具", "未住户情况" };
                /*
                for (int i = 0;i < 3;i++) 
                {
                    MSword.Range xlsColumns = (MSword.Range)Exsheet.Columns[2, MisValue];
                    xlsColumns.EntireColumn.Insert(MSword.XlDirection.xlToRight, MSword.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                */
                //xlsColumns.Insert();


                EXC1.DisplayAlerts = false;
                wb.SaveAs(fileName, Type.Missing, "", "", Type.Missing, Type.Missing, MSword.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);



                label2.Text = "已完成" + NextFile.Name;
                wb.Close();//关闭文档
                wbs.Close();//关闭工作簿

                EXC1.Quit();//关闭EXCEL应用程序

                //释放EXCEL应用程序的进程
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXC1);

            }

            label2.Text = "已完成全部";


        }




        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dilog = new FolderBrowserDialog();
            dilog.Description = "请选择文件夹";
            if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
            {
                textBox1.Text = dilog.SelectedPath;
            }
        }

        
    }
}
