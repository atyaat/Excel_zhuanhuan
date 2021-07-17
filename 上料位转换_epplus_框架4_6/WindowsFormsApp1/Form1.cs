using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [DllImport("SCHEDULE_DLL.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern string getGroupmoInfo(string Str_Tnsname, StringBuilder Str_TestData);

        private void Form1_Load(object sender, EventArgs e)
        {
            butt_group.Enabled = false;
            richTextBox_group.Visible = false;
        }
       

        DataTable dd;
        DataTable gg; // 放group数据
        DataTable aa;       // 数据放置的表名
        DataRow dr;
        DataSet ds = new DataSet();

        string group;  //工单号码导出的组合十码
        string[] groups;
        string[] groups_mianb;
        string[] file_name ;  //导入的文件名
        string[] file_names;
        string save_name;   //保存的文件名
        string save_lij; // 保存文件的路径
        string ten_ma;   // 获取导入的组合十码+面别
        string[] ten_mas;
        string[] Main_Slot_ID = new string[400];     // 
        string[] Sub_Slot_ID = new string[400];
        string[] ten_ma_ex;  //记入Excel表中的十码
        int n, g = 0;      // g.Tray group 行数。
        int flag_s = 0;    // 标志区分单个，多个导入
        int flie_num = 0; // 记录需要导入的文件数
        int file_num_in = 0; //记入已经到导入的数量
        int flag_AB = 0;  // 标志输出文件名AB面
       
        int flag_saixuan = 0; //标志导入数据的组合十码和面别的筛选情况。0 :不是需要的文件  1: 文件重复  2:是需要的文件，新建表 3:是需要的文件，但有相同的组合十码，面别不同。

        public static DataTable ReadExcelToTable(string path)//excel存放的路径
        {
            try
            {
                //连接字符串
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串                    //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set);
                    return set.Tables[0];
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void excel_datatabe( string M) // l 为表名  将数据放入表中
        {
                aa.Rows.RemoveAt(0);
                aa.Rows.RemoveAt(0);

                int a = aa.Columns.Count;  // 获取列数
                int b = aa.Rows.Count;     //行數
                int C = 0;

                for (int x = 1; x < b - 1; x++)
                {
                    string pos = aa.Rows[x]["F1"].ToString();
                    string liaoh = aa.Rows[x]["F2"].ToString();
                    string guig = aa.Rows[x]["F3"].ToString();
                    string shuliang = aa.Rows[x]["F5"].ToString();
                    string weiz = aa.Rows[x]["F6"].ToString();
                    string daiyong = aa.Rows[x]["F7"].ToString();
                if (liaoh != ""  )
                {
                    pos = pos.Replace(" ", "");
                    string[] p = pos.Split('-');
                    var num = pos;
                    int num1 = Regex.Matches(num, "-").Count;

                    switch (num1)
                    {
                        case 0:
                            if (p[0].Equals("00TPCB"))
                            {
                                guig = "PCB";
                                weiz = "PWB";
                                shuliang = "1";
                            }
                            else if (p[0].Equals("00TSOL"))
                            {
                                weiz = "錫膏";
                                shuliang = "1";
                            }
                            break;
                        case 1:
                            if (Convert.ToInt32(p[0]) < 10)
                                p[0] = "0" + p[0];
                            if (Convert.ToInt32(p[1]) < 10)
                                p[1] = "00" + p[1];
                            else if (Convert.ToInt32(p[1]) < 100)
                                p[1] = "0" + p[1];
                            pos = p[0] + "T" + p[1];
                            break;
                        case 2:
                            if (Convert.ToInt32(p[0]) < 10)
                                p[0] = "0" + p[0];
                        if (Convert.ToInt32(p[2]) < 10)
                                p[1] = "0" + p[2];
                            else if (Convert.ToInt32(p[2]) < 100)
                                p[1] = p[2];
                            pos = p[0] + "M3" + p[1];
                            break;

                        case 3:
                            if (Convert.ToInt32(p[0]) < 10)
                                p[0] = "0" + p[0];

                            if (p[1].Equals("A"))
                                p[2] = ((Convert.ToInt32(p[2]) - 1) * 2 + 1).ToString();
                            else if (p[1].Equals("B"))
                                p[2] = ((Convert.ToInt32(p[2]) - 1) * 2 + 25).ToString(); ;

                            if (Convert.ToInt32(p[3]) == 1)
                                p[1] = "L";
                            else if (Convert.ToInt32(p[3]) == 2)
                            {
                                p[1] = "R";
                                p[2] = (Convert.ToInt32(p[2]) + 1).ToString();
                            }
                            if (Convert.ToInt32(p[2]) < 10)
                                p[2] = "0" + p[2];
                            pos = p[0] + p[1] + "9" + p[2];
                            break;
                    }
                   if (flag_saixuan == 2||flag_saixuan ==3)
                   {
                    dr = dd.NewRow();
                    dr["Machinetype"] = "NXT"; dr["Slotid"] = pos; dr["ItemCode"] = liaoh; dr["QTY"] = shuliang; dr["ItemSize"] = guig; dr["Location"] = weiz; dr["SIDE"] = ten_ma.Substring(ten_ma.Length - 1, 1); dr["daiyong"] = daiyong;
                    //把创建的行插入到数据表“Table1”中
                    dd.Rows.Add(dr);
                        
                   }
                    int t = (x * 100 / b) + 5;
                    if (t >= 100)
                        t = 100;
                    progressBar1.Value = t;
                    }
                }
                int hang = dd.Rows.Count;
                int h = 0 ;
            
                int[] del = new int[hang];
                
                for (int x = 0; x < hang - 1; x++)      // 遍历每一行
                {
                    string s = dd.Rows[x]["QTY"].ToString();
                    string L = dd.Rows[x]["ItemCode"].ToString();
                    string D = dd.Rows[x]["daiyong"].ToString();
                    string lz = dd.Rows[x]["Slotid"].ToString();
                    string wz = dd.Rows[x]["Location"].ToString();

                    
                    if (s!="" &&Convert.ToInt32(s) == 0)    // 找到数量为0 的行
                    {
                        int flog = 0;
                        for (int y = 0; y < hang - 1; y++)   // 遍历料站找他的主用料
                        {
                            string s2 = dd.Rows[y]["QTY"].ToString();
                            string L2 = dd.Rows[y]["ItemCode"].ToString();
                            string D2 = dd.Rows[y]["daiyong"].ToString();
                            string lz2 = dd.Rows[y]["Slotid"].ToString();
                            string wz2 = dd.Rows[y]["Location"].ToString();

                            if (y != x)
                            {
                                if (lz.Equals(lz2)  && Convert.ToInt32(s2) > 0)   // 找了到料站相同。（代用料找主用料）
                                {
                                    dd.Rows[x]["QTY"] = s2;
                                    flog = 1;                                 // 标记这一行不删除
                                    y = hang;         //跳出循环
                                }
                            }
                        }
                        if (flog == 0)                   // 将要删除的行标记--将行数计入到数组里。全部处理完后一起删除。不然行号会变。
                        {
                            del[h] = x;
                            h++;
                        }
                    }
                    // --在将找到的子料站的数量改成1--
                    if (!lz.Substring(2, 1).Equals("T") && Convert.ToInt32(s) > 0) //主料站找子料站。
                    {
                        for (int j = x; j < hang - 1; j++)
                        {
                            int flog_2 = 0;
                            int flog_3 = 0;

                            string s2 = dd.Rows[j]["QTY"].ToString();
                            string L2 = dd.Rows[j]["ItemCode"].ToString();
                            string D2 = dd.Rows[j]["daiyong"].ToString();
                            string lz2 = dd.Rows[j]["Slotid"].ToString();
                            string wz2 = dd.Rows[j]["Location"].ToString();

                            if (L.Equals(L2) && Convert.ToInt32(s2) == 0 && wz.Equals(wz2)) //找到其子料站
                            {
                            if (checkBox1.Checked)
                            {
                                dd.Rows[j]["QTY"] = s;
                            }
                            for (int i = 0; i < g; i++)       // 遍历数组。如果没有主料站就加入  找到的主料站应该在前的
                                {

                                    if (lz.Equals(Sub_Slot_ID[i]))   // 
                                    {
                                        flog_2 = 1;
                                    }
                                    if (lz2.Equals(Sub_Slot_ID[i]))
                                    {
                                        flog_3 = 1;
                                    }
                                }
                                //Sub_Slot_ID[g] = lz;g++;
                                if (flog_2 == 0)     //没有重复的主料站
                                {
                                    Sub_Slot_ID[g] = lz;
                                    Main_Slot_ID[g] = lz;
                                    g++;

                                }
                                if (flog_3 == 0)   // 没有重复的子料站。
                                {
                                    Sub_Slot_ID[g] = lz2;
                                    Main_Slot_ID[g] = Main_Slot_ID[g - 1];
                                    g++;

                                }
                            }
                        }
                    }
                }
                for (int i = 0; i < h; i++)      // 将表里的数据要删除的行统一删除。Tray Tray Tray Tray Tray Tray Tray TrayTrayTrayTray Tray Tray Tray Tray
            {
                int flag = 0;
                if (checkBox1.Checked)
                {
                    for (int j = 0; j < g; j++)
                    {
                        string c = dd.Rows[del[i] - i]["Slotid"].ToString();
                        if (dd.Rows[del[i] - i]["Slotid"].ToString().Equals(Sub_Slot_ID[j]) && Convert.ToInt32(dd.Rows[del[i] - i]["QTY"].ToString())>0)
                        {
                            flag = 1;
                        }
                    }
                }
                if(flag==0)
                    dd.Rows.RemoveAt(del[i] - i);
                }
            int u= dd.Rows.Count;
            richTextBox1.Text = richTextBox1.Text + "\n" + "The table was successfully imported: " + ten_ma;

            if (flag_saixuan == 2)
            {
                ds.Tables.Add(dd);
            } 
            else if (flag_saixuan == 3)
            {
                for (int i = 0; i < u; i++)
                {
                    object[] obj = new object[8];
                    dd.Rows[i].ItemArray.CopyTo(obj, 0);
                    ds.Tables[M].Rows.Add(obj);

                }
                for (int i = 0; i < ds.Tables[M].Rows.Count; i++)
                {
                    string bb = ds.Tables[M].Rows[i]["Slotid"].ToString();
                    if (bb.Equals("00TSOL"))
                        C++;
                }
                if (C == 1)
                {
                    dr = dd.NewRow();
                    dr["Machinetype"] = "NXT"; dr["Slotid"] = "00TSOL"; dr["ItemCode"] = "4090211600"; dr["QTY"] = "1"; dr["ItemSize"] = ""; dr["Location"] = "錫膏";dr["SIDE"] = ten_ma.Substring(ten_ma.Length-1,1);
                    dd.Rows.Add(dr);
                }
            }
                if (file_num_in == flie_num || flag_s ==0)
            {
                // 将单个表生成group和tray group？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？？
                
                ten_ma_ex = new string[ds.Tables.Count];
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    object[] obj = new object[8];
                    object[] ob = new object[8];

                    for (int j = 0; j < ds.Tables[i].Rows.Count; j++)
                    {
                        string dddd = save_name.Substring(save_name.Length - 1, 1);
                        ds.Tables[i].Rows[j].ItemArray.CopyTo(obj, 0);
                        if (dddd.Equals("A") || dddd.Equals("B"))
                            obj[6] = dddd;
                        else
                            obj[6] = "G";
                            gg.Rows.Add(obj);
                    }
                    ten_ma_ex[i] = ds.Tables[i].TableName.ToString();
                    ds.Tables[i].TableName = "M" + (i+1);
                }
                gg.DefaultView.Sort = "Slotid";     //选择排序的列
                //buil_test("tray group"); // 
                gg = gg.DefaultView.ToTable();
                  if (flag_s == 1)
                  {
                    expor_excel();
                    MessageBox.Show("文件全部导入完成，正在导出");
                  }
                }
        }
        private void lead_Excel_Click(object sender, EventArgs e)
        {
           
            if (file_num_in == flie_num && flag_s == 1 && richTextBox2.Text != "")
            {
                MessageBox.Show("文件全部导出完成，再次导入请先清除数据");
               
            }
            else if (file_num_in > 0 && flag_s == 0)
            {
                MessageBox.Show("再次导入前请先清除数据");
            }
            else
            {
                int r = 0;
                richTextBox1.Text = richTextBox1.Text + "\n" + "To Leading.....";

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    file_name = openFileDialog1.FileNames;
                     r = file_name.Length;
                }
               
                ///////////////////////////////////////////////////////////////判断是否有相同的组合十码不同面别的。如果有就不新建表，加入到相同组合十码的表里

                    if (r!=0 && file_name[0] != "" && r == 1)
                    {

                    try
                    {
                        string[] arr1 = file_name[0].Split('\\'); // 以'\\'字符对字符串进行分割，返回字符串数组 。。将导入Excel文件的路径作为导出的路径。
                            int c = arr1.Length; ;
                            aa = ReadExcelToTable(file_name[0]);
                            ten_ma = aa.Rows[1]["F2"].ToString() + "-" + aa.Rows[1]["F7"].ToString().Substring(aa.Rows[1]["F7"].ToString().Length - 1, 1);
                            for (int i = 1; i < c - 1; i++)
                            {
                                arr1[0] += "\\" + arr1[i];
                            }
                            save_lij = arr1[0] + "\\";   //?????
                                                         //----------------------------------------------------------------------------------------------
                            int a = 0; // 记录字符数
                            if (flag_s == 0)
                            {
                                flag_saixuan = 2;
                               if(textBox1.Text =="")
                                save_name = ten_ma;
                               else
                                save_name = textBox1.Text.ToString();
                                ten_mas = new string[1];
                                buil_test_g("GROUP");
                            }
                            else
                            {
                                for (int i = 0; i < flie_num; i++)    // 
                                {
                                    if (ten_ma.Equals(groups[i]))
                                    {
                                        for (int x = 0; x < i; x++)
                                        {
                                            a = a + richTextBox2.Lines[x].Length + 1;
                                        }
                                        richTextBox2.Select(a, richTextBox2.Lines[i].Length);
                                        if (richTextBox2.SelectionBackColor.Name.Equals("DeepSkyBlue"))
                                        {

                                            flag_saixuan = 1;
                                        }
                                        else
                                        {
                                            richTextBox2.SelectionBackColor = Color.DeepSkyBlue;
                                            flag_saixuan = 2;
                                            //导入的文件正确
                                        }
                                    }
                                    if (flag_saixuan == 2 && file_num_in > 0)
                                    {
                                        for (int j = 0; j < file_num_in; j++)
                                        {
                                            if (ten_mas[j].Substring(0, ten_mas[j].Length - 2).Equals(ten_ma.Substring(0, ten_ma.Length - 2)))
                                                flag_saixuan = 3;

                                        }
                                    }

                                }

                            }
                            string M = ten_ma.Substring(0, ten_ma.Length - 2);
                            switch (flag_saixuan)
                            {

                                case 0:

                                    MessageBox.Show("导入失败—请选择正确的文件");
                                    break;
                                case 1:

                                    MessageBox.Show("检测到导入重复");
                                    flag_saixuan = 0;
                                    break;
                                case 2:

                                    richTextBox1.Text = richTextBox1.Text + "\n" + ten_ma + "导入成功";


                                    ten_mas[file_num_in] = ten_ma;
                                    file_num_in++;
                                    //--------

                                    //---------换表名： 先将组合十码的作为表名，在数据都弄好后再讲表从新命名。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。

                                    buil_test(M);
                                    excel_datatabe(M);
                                    if (flag_s == 0)
                                        expor_excel();
                                    flag_saixuan = 0;
                                    break;
                                case 3:

                                    richTextBox1.Text = richTextBox1.Text + "\n" + ten_ma + "导入成功,找到相同组合十码....";

                                    ten_mas[file_num_in] = ten_ma;
                                    file_num_in++;
                                    buil_test(M);
                                    excel_datatabe(M);
                                    flag_saixuan = 0;
                                    break;
                            }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                   // -----------------处理导入文件的路径-------------------------------- -
                    }
                    else if (r > 1)
                    {
                        for (int ri = 0; ri < r; ri++)
                        {
                        if (file_num_in > 0 && flag_s == 0)
                        {
                            MessageBox.Show("再次导入前请先清除数据");
                        }
                        else
                        {
                            //try
                            //{
                                string[] arr1 = file_name[ri].Split('\\'); // 以'\\'字符对字符串进行分割，返回字符串数组 。。将导入Excel文件的路径作为导出的路径。
                                int c = arr1.Length; ;
                                aa = ReadExcelToTable(file_name[ri]);
                                ten_ma = aa.Rows[1]["F2"].ToString() + "-" + aa.Rows[1]["F7"].ToString().Substring(aa.Rows[1]["F7"].ToString().Length - 1, 1);
                                for (int i = 1; i < c - 1; i++)
                                {
                                    arr1[0] += "\\" + arr1[i];
                                }
                                save_lij = arr1[0] + "\\";   //?????
                                                             //----------------------------------------------------------------------------------------------
                                int a = 0; // 记录字符数
                                if (flag_s == 0)
                                {
                                    flag_saixuan = 2;
                                    save_name = ten_ma;
                                    ten_mas = new string[1];
                                    buil_test_g("GROUP");
                                }
                                else
                                {
                                    for (int i = 0; i < flie_num; i++)    // 
                                    {
                                        if (ten_ma.Equals(groups[i]))
                                        {
                                            for (int x = 0; x < i; x++)
                                            {
                                                a = a + richTextBox2.Lines[x].Length + 1;
                                            }
                                            richTextBox2.Select(a, richTextBox2.Lines[i].Length);
                                            if (richTextBox2.SelectionBackColor.Name.Equals("DeepSkyBlue"))
                                            {

                                                flag_saixuan = 1;
                                            }
                                            else
                                            {
                                                richTextBox2.SelectionBackColor = Color.DeepSkyBlue;
                                                flag_saixuan = 2;
                                                //导入的文件正确
                                            }
                                        }
                                        if (flag_saixuan == 2 && file_num_in > 0)
                                        {
                                            for (int j = 0; j < file_num_in; j++)
                                            {
                                                if (ten_mas[j].Substring(0, ten_mas[j].Length - 2).Equals(ten_ma.Substring(0, ten_ma.Length - 2)))
                                                    flag_saixuan = 3;

                                            }
                                        }

                                    }

                                }
                                string M = ten_ma.Substring(0, ten_ma.Length - 2);
                                switch (flag_saixuan)
                                {

                                    case 0:

                                        MessageBox.Show("导入失败—请选择正确的文件");
                                        break;
                                    case 1:

                                        MessageBox.Show("检测到导入重复");
                                        flag_saixuan = 0;
                                        break;
                                    case 2:

                                        richTextBox1.Text = richTextBox1.Text + "\n" + ten_ma + "导入成功";


                                        ten_mas[file_num_in] = ten_ma;
                                        file_num_in++;
                                        //--------

                                        //---------换表名： 先将组合十码的作为表名，在数据都弄好后再讲表从新命名。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。。

                                        buil_test(M);
                                        excel_datatabe(M);
                                        if (flag_s == 0)
                                            expor_excel();
                                        flag_saixuan = 0;
                                        break;
                                    case 3:

                                        richTextBox1.Text = richTextBox1.Text + "\n" + ten_ma + "导入成功,找到相同组合十码....";

                                        ten_mas[file_num_in] = ten_ma;
                                        file_num_in++;
                                        buil_test(M);
                                        excel_datatabe(M);
                                        flag_saixuan = 0;
                                        break;
                                }
                            //}
                            //catch (Exception ex)
                            //{
                            //    MessageBox.Show(ex.Message);
                            //}
                            //-----------------处理导入文件的路径---------------------------------
                            
                        }
                        }
                       
                    }
                    else
                    {
                        MessageBox.Show("请选择需要导入的文件");
                    }
            }
        }
        private void buil_test(string test_name)   //新建表
        {
            dd = new DataTable(test_name); //创建数据表 
            DataColumn dcMachinetype = new DataColumn("Machinetype", typeof(string)); //创建ID列
            DataColumn dcSlotid = new DataColumn("Slotid", typeof(string));//创建姓名列
            DataColumn dcItemCode = new DataColumn("ItemCode", typeof(string));//创建状态列
            DataColumn dcQTY = new DataColumn("QTY", typeof(string));//创建状态列
            DataColumn dcItemSize = new DataColumn("ItemSize", typeof(string));//创建状态列
            DataColumn dcLocation = new DataColumn("Location", typeof(string));//创建状态列
            DataColumn dcmianb = new DataColumn("SIDE", typeof(string));//创建状态列
            DataColumn daiyongliao = new DataColumn("daiyong", typeof(string));//创建状态列

            dd.Columns.Add(dcMachinetype);
            dd.Columns.Add(dcSlotid);
            dd.Columns.Add(dcItemCode);
            dd.Columns.Add(dcQTY);
            dd.Columns.Add(dcItemSize);
            dd.Columns.Add(dcLocation);
            dd.Columns.Add(dcmianb);
            dd.Columns.Add(daiyongliao);
        }
        private void buil_test_g(string test_name)   //新建表
        {
            gg = new DataTable(test_name); //创建数据表 
            DataColumn dcMachinetype = new DataColumn("Machinetype", typeof(string)); //创建ID列
            DataColumn dcSlotid = new DataColumn("Slotid", typeof(string));//创建姓名列
            DataColumn dcItemCode = new DataColumn("ItemCode", typeof(string));//创建状态列
            DataColumn dcQTY = new DataColumn("QTY", typeof(string));//创建状态列
            DataColumn dcItemSize = new DataColumn("ItemSize", typeof(string));//创建状态列
            DataColumn dcLocation = new DataColumn("Location", typeof(string));//创建状态列
            DataColumn dcmianb = new DataColumn("SIDE", typeof(string));//创建状态列
            DataColumn daiyongliao = new DataColumn("daiyong", typeof(string));//创建状态列

            gg.Columns.Add(dcMachinetype);
            gg.Columns.Add(dcSlotid);
            gg.Columns.Add(dcItemCode);
            gg.Columns.Add(dcQTY);
            gg.Columns.Add(dcItemSize);
            gg.Columns.Add(dcLocation);
            gg.Columns.Add(dcmianb);
            gg.Columns.Add(daiyongliao);
        }
        private void buil_excel(DataSet dgv)
        {

            
            int x = ds.Tables.Count;    //获取有多少个表
           
            save_name = save_lij + save_name;
            FileInfo newFile = new FileInfo(@save_name + ".xls");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@save_name + ".xls");
            }
            using (ExcelPackage package = new ExcelPackage(newFile))         //新建Excel
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(gg.TableName);   //新建表格test


                string c = AppDomain.CurrentDomain.BaseDirectory;    // 获取文件路径
                ExcelPicture picture = worksheet.Drawings.AddPicture("logo", Image.FromFile(@c + "g.png"));//插入图片
                picture.SetPosition(3, 0);//设置图片的位置
                picture.SetSize(150, 50);//设置图片的大小

                //ExcelPicture picture3 = worksheet.Drawings.AddPicture("log", Image.FromFile(@c + "k.png"));//插入图片
                //picture3.SetPosition(0, 382);//设置图片的位置
                //picture3.SetSize(170, 50);//设置图片的大小

                worksheet.Row(1).Height = 40;//设置行高

                worksheet.Cells[1, 1, 1, 7].Merge = true;//合并单元格
                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                worksheet.Cells[1, 1].Value = "料站分配表";

                //worksheet.Cells[1, 1].Style.Font.Bold = true;//字体为粗体
                worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Black);//字体颜色
                worksheet.Cells[1, 1].Style.Font.Name = "微软雅黑";//字体
                worksheet.Cells[1, 1].Style.Font.Size = 20;//字体大小


                worksheet.Cells[3, 1].Value = "工单";
                worksheet.Cells[4, 1].Value = "面别";
                //worksheet.Cells[5, 1].Value = "组合十码";
                //worksheet.Cells[5, 2].Value = ten_mas[i];
                worksheet.Cells[6, 1].Value = "Machinetype";
                worksheet.Cells[6, 2].Value = "Slotid";
                worksheet.Cells[6, 3].Value = "ItemCode";
                worksheet.Cells[6, 4].Value = "QTY";
                worksheet.Cells[6, 5].Value = "ItemSize";
                worksheet.Cells[6, 6].Value = "Location";
                worksheet.Cells[6, 7].Value = "SIDE";

                worksheet.Cells.Style.WrapText = true;//自动换行
                worksheet.Column(1).Width = 15;//设置列宽
                worksheet.Column(3).Width = 15;
                worksheet.Column(5).Width = 45;
                worksheet.Column(6).Width = 60;//设置列宽
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中

                worksheet.Cells[6, 1, 6, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[6, 1, 6, 7].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 204));//设置单元格背景色

                using (ExcelRange r = worksheet.Cells[1, 1, gg.Rows.Count + 6, 7])   // 设置单元格边框 
                {
                    r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                }


                for (int k = 0; k < gg.Rows.Count; k++)   //行数
                {

                    for (int j = 0; j < 7; j++)   //列
                    {

                        try
                        {
                            worksheet.Cells[k + 7, j + 1].Value = gg.Rows[k][j].ToString();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString() + (k + 7).ToString());
                        }

                    }

                }

                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Tray Group");   //新建表格test 
                worksheet2.Cells[1, 1].Value = "Main Slot ID";
                worksheet2.Cells[1, 2].Value = "Sub Slot ID";
                for (int i = 0; i < Main_Slot_ID.Length; i++)
                {
                    try
                    {
                        worksheet2.Cells[i + 2, 1].Value = Main_Slot_ID[i];

                    }
                    catch
                    {
                    }
                }
                for (int i = 0; i < Sub_Slot_ID.Length; i++)
                {
                    try
                    {
                        worksheet2.Cells[i + 2, 2].Value = Sub_Slot_ID[i];

                    }
                    catch
                    {
                    }
                }
                if (flag_s == 1)
                {

                    for (int i = 0; i < x; i++)
                    {
                        
                        worksheet = package.Workbook.Worksheets.Add(dgv.Tables[i].TableName);   //新建表格test


                        c = AppDomain.CurrentDomain.BaseDirectory;     // 获取文件路径
                        picture = worksheet.Drawings.AddPicture("logo", Image.FromFile(@c + "g.png"));//插入图片
                        picture.SetPosition(3, 0);//设置图片的位置
                        picture.SetSize(150, 50);//设置图片的大小

                        //picture3 = worksheet.Drawings.AddPicture("log", Image.FromFile(@c + "k.png"));//插入图片
                        //picture3.SetPosition(0, 382);//设置图片的位置
                        //picture3.SetSize(170, 50);//设置图片的大小

                        worksheet.Row(1).Height = 40;//设置行高

                        worksheet.Cells[1, 1, 1, 7].Merge = true;//合并单元格
                        worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直居中
                        worksheet.Cells[1, 1].Value = "料站分配表";

                        //worksheet.Cells[1, 1].Style.Font.Bold = true;//字体为粗体
                        worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.Black);//字体颜色
                        worksheet.Cells[1, 1].Style.Font.Name = "微软雅黑";//字体
                        worksheet.Cells[1, 1].Style.Font.Size = 20;//字体大小


                        worksheet.Cells[3, 1].Value = "工单";
                        worksheet.Cells[4, 1].Value = "面别";
                        worksheet.Cells[3, 4].Value = "Model";
                        worksheet.Cells[3, 5].Value =ten_ma_ex[i];

                        //worksheet.Cells[5, 1].Value = "组合十码";
                        //worksheet.Cells[5, 2].Value = ten_mas[i];
                        worksheet.Cells[6, 1].Value = "Machinetype";
                        worksheet.Cells[6, 2].Value = "Slotid";
                        worksheet.Cells[6, 3].Value = "ItemCode";
                        worksheet.Cells[6, 4].Value = "QTY";
                        worksheet.Cells[6, 5].Value = "ItemSize";
                        worksheet.Cells[6, 6].Value = "Location";
                        worksheet.Cells[6, 7].Value = "SIDE";

                        worksheet.Cells.Style.WrapText = true;//自动换行
                        worksheet.Column(1).Width = 15;//设置列宽
                        worksheet.Column(3).Width = 15;
                        worksheet.Column(5).Width = 45;
                        worksheet.Column(6).Width = 60;//设置列宽
                        worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中
                        worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平居中

                        worksheet.Cells[6, 1, 6, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[6, 1, 6, 7].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 204));//设置单元格背景色

                        using (ExcelRange r = worksheet.Cells[1, 1, dgv.Tables[i].Rows.Count + 6, 7])   // 设置单元格边框 
                        {
                            r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        }


                        for (int k = 0; k < dgv.Tables[i].Rows.Count; k++)   //行数
                        {

                            for (int j = 0; j < 7; j++)   //列
                            {

                                try
                                {
                                    worksheet.Cells[k + 7, j + 1].Value = dgv.Tables[i].Rows[k][j].ToString();

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString() + (k + 7).ToString());
                                }

                            }

                        }

                    }


                }
                package.Save();                                    // 保存文件

            }
            richTextBox1.Text = richTextBox1.Text + "\n" + "Successfully exported:" + save_name;


        }

        private void export_Excel_Click(object sender, EventArgs e)
        {
            
            
        }
        private void expor_excel()
        {
            richTextBox1.Text = richTextBox1.Text + "\n" + "exporting....";
            buil_excel(ds);
        }
        private void butt_group_Click(object sender, EventArgs e)
        {
           
            if (flag_s == 1)
            {

                try
                {
                    if (richTextBox2.Text == "" || radioButton2.Checked)
                    {
                        if (richTextBox2.Text == "" && !radioButton2.Checked)
                        {
                            if (textBox1.Text != "")
                            {
                                StringBuilder stringBuilder = new StringBuilder(1024);
                                stringBuilder.Append(textBox1.Text);
                                group = getGroupmoInfo("SFISM4", stringBuilder);


                                textBox2.Text = group; //   0}A,3941311405,3941311405;B,3941311405,3941311405;}
                            }
                            else
                                MessageBox.Show("请手动输入工单号码");
                        }
                        else if (radioButton2.Checked && richTextBox2.Text == "")
                        {
                            if (richTextBox_group.Text != "" && textBox1.Text != "")
                            {
                                group = richTextBox_group.Text.ToString();
                                textBox2.Text = group;
                            }
                            else
                                MessageBox.Show("请手动输入工单号码和Group");
                        }
                        
                        group = group.TrimStart('0').Trim('}');
                        groups = group.Split(';');
                       
                        radioButton1.Enabled = false;
                        radioButton2.Enabled = false;
                        radioButton3.Enabled = false;

                        butt_group.Enabled = false;
                        butt_group.BackColor = Color.DarkGray;
                        butt_group.ForeColor = Color.LightGray;

                        lead_Excel.Enabled = true;
                        lead_Excel.BackColor = Color.DeepSkyBlue;
                        lead_Excel.ForeColor = Color.Black;

                        richTextBox_group.Visible = false;
                        for (int i = 0; i < groups.Length - 1; i++)
                        {
                            groups_mianb = groups[i].Split(',');
                            groups[i] = groups_mianb[2] + "-" + groups_mianb[0];
                          
                            richTextBox2.Text = richTextBox2.Text + groups[i] + "\n";
                            if (groups_mianb[0].Equals("A"))
                            {
                                flag_AB++;
                            }
                            else if (groups_mianb[0].Equals("B"))
                            {
                                flag_AB += 2;
                            }
                        }

                        file_names = new string[groups.Length];
                        ten_mas = new string[groups.Length];
                        flie_num = groups.Length - 1;
                        if (flag_AB == flie_num)
                            save_name = textBox1.Text + "-A";
                        else if (flag_AB == flie_num * 2)
                            save_name = textBox1.Text + "-B";
                        else
                            save_name = textBox1.Text;
                        buil_test_g("GROUP");
                    }
                    else
                        MessageBox.Show("请先清除数据");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

               
            }
            else
                MessageBox.Show("导入多个工单模式下使用");
            

            
        }

        private void clean_Click(object sender, EventArgs e)
        {

           
            richTextBox2.Text = "";
            textBox2.Text = "";
            try
            {
                //清除所有表中的数据和数组--------------------------------------------------------------------------

                ds = new DataSet();
                gg = new DataTable();
                dd =new DataTable();
                n = 0;
                g = 0;

                flie_num = 0;
                file_num_in = 0;
                flag_AB = 0;
               
                flag_saixuan = 0;

                group = null;  //工单号码导出的组合十码
                groups = null;
                groups_mianb = null;
                file_name = null;  //导入的文件名
                file_names = null;
                save_name = null;   //保存的文件名
                save_lij = null; // 保存文件的路径
                ten_ma = null;   // 获取导入的组合十码+面别

               
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                radioButton3.Checked = true;
                richTextBox_group.Visible = false;
                richTextBox_group.Text = "A,（机种）,3941311405;B,（机种）,3941311405";
                
                richTextBox1.Text = richTextBox1.Text + "\n" + " 数据已被清除可继续导入........................." ;
            }
            catch
            { }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                flag_s = 1;
                butt_group.BackColor = Color.DeepSkyBlue;
                butt_group.ForeColor = Color.Black;
                butt_group.Enabled = true;
                lead_Excel.Enabled = false;
                lead_Excel.BackColor = Color.DarkGray;
                lead_Excel.ForeColor = Color.LightGray;
                richTextBox_group.Visible = false;
                
            }
            else
            {
                flag_s = 0;
                butt_group.BackColor = Color.DarkGray;
                butt_group.ForeColor = Color.LightGray;
               
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                flag_s = 1;
                butt_group.BackColor = Color.DeepSkyBlue;
                butt_group.ForeColor = Color.Black;
                butt_group.Enabled = true;
                lead_Excel.Enabled = false;

                lead_Excel.BackColor = Color.DarkGray;
                lead_Excel.ForeColor = Color.LightGray;

                richTextBox_group.Visible = true;
                richTextBox_group.BringToFront();

                
            }
            else
            {
                flag_s = 0;
                butt_group.BackColor = Color.DarkGray;
                butt_group.ForeColor = Color.LightGray;
                
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            butt_group.Enabled = false;
            lead_Excel.Enabled = true;
            lead_Excel.BackColor = Color.DeepSkyBlue;
            lead_Excel.ForeColor = Color.Black;
            richTextBox_group.Visible = false;
        }

    
    }
}
