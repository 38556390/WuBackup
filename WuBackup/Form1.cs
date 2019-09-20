using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Win32;

namespace WuBackup
{
    public partial class Form1 : Form
    {
        private static string Torun = "NO";//备份是否执行
        //private string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/backupdata.mdb";
       // private static string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=backupdata.mdb";
        private string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @Application.StartupPath +  "/backupdata.mdb";
        //private string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data/backupdata.mdb";
  
        public Form1()
        {
            InitializeComponent();
        }
        //初始化函数
        private void dbind()
        {
            timer1.Start();
            #region//在没有设置过时显示初始值
            dateTimePicker1.Value = Convert.ToDateTime(DateTime.Now.ToShortTimeString());
            comboBox1.SelectedItem = "星期一";//初始选项
            panel2.Visible = false;
            panel3.Visible = false;
            radioButton1.Checked = true;
            radioButton4.Checked = true;
            radioButton7.Checked = true;

            numericUpDown5.Enabled = false;
            comboBox1.Enabled = false;
            numericUpDown6.Enabled = false;

            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;

            dateTimePicker3.Enabled = false;
            #endregion
            #region//读取设置记录
            string Hertz = "";//发生频率
            OleDbConnection con = new OleDbConnection(constr);
            string select = "select Hertz,Daycount,Weekcount,Weekdate,Monthtype,Whenday,Amonthcount,Bweekcount,Bweekdate,Bmonthcount," +
                "Backuptimes,Startdate,Enddate,Datetype,Sourcepath,Backuppath,Startingup,Torun" +
                " from setbackup";
            con.Open();
            OleDbCommand com = new OleDbCommand(select, con);
            OleDbDataReader dir = com.ExecuteReader();
            if (dir.Read())
            {
                Torun = Convert.ToString(dir["Torun"]);//是否执行备份
                if (Torun == "NO")
                {
                    label16.Text = "你的设置为：不进行自动备份。";
                    label17.Text = "";
                }
                else
                {
                    Hertz = Convert.ToString(dir["Hertz"]);
                    #region//每天类型时读取
                    if (Hertz == "D")
                    {
                        radioButton1.Checked = true;
                        numericUpDown1.Value = Convert.ToInt32(dir["Daycount"]);//每多少天发生一次
                    }
                    #endregion
                    #region//W每周类型时读取
                    if (Hertz == "W")
                    {
                        radioButton2.Checked = true;
                        numericUpDown2.Value = Convert.ToInt32(dir["Weekcount"]);//每多少周发生一次
                        string Weekdate = Convert.ToString(dir["Weekdate"]);//具体星期几
                        if (Weekdate.IndexOf("星期一") > -1)//查找是否包含星期一这个字符串
                        {
                            checkBox2.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期二") > -1)
                        {
                            checkBox3.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期三") > -1)
                        {
                            checkBox4.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期四") > -1)
                        {
                            checkBox5.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期五") > -1)
                        {
                            checkBox6.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期六") > -1)
                        {
                            checkBox7.Checked = true;
                        }
                        if (Weekdate.IndexOf("星期日") > -1)
                        {
                            checkBox8.Checked = true;
                        }
                    }
                    #endregion
                    #region//M每月类型时读取
                    if (Hertz == "M")
                    {
                        radioButton3.Checked = true;
                        string Monthtype = Convert.ToString(dir["Monthtype"]);//月份发生类型
                        if (Monthtype == "A")
                        {
                            radioButton4.Checked = true;
                            numericUpDown3.Value = Convert.ToInt32(dir["Whenday"]);//第几天
                            numericUpDown4.Value = Convert.ToInt32(dir["Amonthcount"]);//几个月
                        }
                        if (Monthtype == "B")
                        {
                            radioButton5.Checked = true;
                            numericUpDown5.Value = Convert.ToInt32(dir["Bweekcount"]);//第几个星期
                            comboBox1.Text = Convert.ToString(dir["Bweekdate"]);//具体星期几
                            numericUpDown6.Value = Convert.ToInt32(dir["Bmonthcount"]);//几个月
                        }
                    }
                    #endregion
                    #region//读取其它信息
                    dateTimePicker1.Value = Convert.ToDateTime(dir["Backuptimes"]);//备份时间
                    dateTimePicker2.Value = Convert.ToDateTime(dir["Startdate"]);//开始日期
                    string Datetype = Convert.ToString(dir["Datetype"]);//结束时间类型
                    if (Datetype == "Y")
                    {
                        radioButton6.Checked = true;
                        dateTimePicker3.Value = Convert.ToDateTime(dir["Enddate"]);//结束日期
                    }
                    else
                    {
                        radioButton7.Checked = true;
                        dateTimePicker3.Enabled = false;
                    }
                    textBox2.Text = Convert.ToString(dir["Sourcepath"]);//目标路径
                    textBox1.Text = Convert.ToString(dir["Backuppath"]);//备份路径
                    string Startingup = Convert.ToString(dir["Startingup"]);//开机是否启动
                    if (Startingup == "YES")
                    {
                        checkBox1.Checked = true;
                    }
                    Torun = Convert.ToString(dir["Torun"]);//备份是否执行
                    #endregion
                }
            }
            dir.Close();
            con.Close();
            #endregion
        }
        private void dbind2()
        {

            #region//读取设置记录
            string Hertz = "";//发生频率
            string record01 = "你的设置为：";
            string record02 = "";
            OleDbConnection con = new OleDbConnection(constr);
            string select = "select Hertz,Daycount,Weekcount,Weekdate,Monthtype,Whenday,Amonthcount,Bweekcount,Bweekdate,Bmonthcount," +
                "Backuptimes,Startdate,Enddate,Datetype,Sourcepath,Backuppath,Startingup,Torun" +
                " from setbackup";
     
            con.Open();
            OleDbCommand com = new OleDbCommand(select, con);
            OleDbDataReader dir = com.ExecuteReader();
            if (dir.Read())
            {
                Torun = Convert.ToString(dir["Torun"]);//是否执行备份
                if (Torun=="NO")
                {
                    label16.Text = "你的设置为：不进行自动备份。";
                    label17.Text = "";
                }
                else
                {
                    Hertz = Convert.ToString(dir["Hertz"]);
                    #region//每天类型时读取
                    if (Hertz == "D")
                    {
                        record01 += " 每" + Convert.ToString(dir["Daycount"]) + "天";
                    }
                    #endregion
                    #region//W每周类型时读取
                    if (Hertz == "W")
                    {
                        string Weekdate = Convert.ToString(dir["Weekdate"]);//具体星期几
                        record01 += " 每" + Convert.ToString(dir["Weekcount"]) + "周 " + Weekdate;
                    }
                    #endregion
                    #region//M每月类型时读取
                    if (Hertz == "M")
                    {
                        radioButton3.Checked = true;
                        string Monthtype = Convert.ToString(dir["Monthtype"]);//月份发生类型
                        if (Monthtype == "A")
                        {
                            record01 += " 每" + Convert.ToString(dir["Amonthcount"]) + "月 " + Convert.ToString(dir["Whenday"]) + "天";
                        }
                        if (Monthtype == "B")
                        {
                            record01 += " 每" + Convert.ToString(dir["Bmonthcount"]) + "月 " + "第" + Convert.ToString(dir["Bweekcount"]) + "个星期的" + Convert.ToString(dir["Bweekdate"]);
                        }
                    }
                    #endregion
                    #region//读取其它信息
                    record01 += " " + Convert.ToDateTime(dir["Backuptimes"]).ToString("HH:mm:ss");
                    string Datetype = Convert.ToString(dir["Datetype"]);//结束时间类型
                    if (Datetype == "Y")
                    {
                        record01 += " 结束时间:" + Convert.ToString(dir["Enddate"]);

                    }
                    else
                    {
                        record01 += " 无结束时间";
                    }
                    label16.Text = record01;
                    string Startingup = Convert.ToString(dir["Startingup"]);//开机是否启动

                    if (Startingup == "YES")
                    {

                        record02 += "开机自动运行";
                    }
                    else
                    {
                        record02 += "开机不自动运行";
                    }
                    record02 += " 备份路径" + Convert.ToString(dir["Backuppath"]);
                    label17.Text = record02;
                    #endregion
                }
               
            }
            dir.Close();
            con.Close();
            #endregion
        }
        private void Form1_Load(object sender, EventArgs e)
        {
          
            dbind();
            dbind2();
       

        }
        private void CopyDirectory(string sourcePath, string destPath)
        {
            DirectoryInfo dir = new DirectoryInfo(sourcePath); //实例化
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();//获取文件夹中所中目录
            foreach (FileSystemInfo i in fileinfo)
            {
                if (i is DirectoryInfo)//判断是文件夹
                {
                    Directory.CreateDirectory(destPath + "\\" + i.Name);//创建文件夹
                    CopyDirectory(sourcePath + "\\" + i.Name, destPath + "\\" + i.Name);//递归调用
                }
                else
                {
                    if (File.Exists(destPath + "\\" + i.Name) == false)
                    {
                        File.Copy(i.FullName, destPath + "\\" + i.Name);//复制文件
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();//实例化文件夹浏览对话框
            if (folder.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folder.SelectedPath;//获取文件夹路径
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();//实例化文件夹浏览对话框
            if (folder.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folder.SelectedPath;//获取文件夹路径
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Torun == "YES")
            {
                DateTime now_time = DateTime.Now;//当前时间
                string end_backup_time = "";//最后一次备份时间
                #region//读取设置记录
                string Hertz = "";//发生频率
                OleDbConnection con = new OleDbConnection(constr);
                string select = "select Hertz,Daycount,Weekcount,Weekdate,Monthtype,Whenday,Amonthcount,Bweekcount,Bweekdate,Bmonthcount," +
                    "Backuptimes,Startdate,Enddate,Datetype,Sourcepath,Backuppath,Startingup,Torun" +
                    " from setbackup";
                string select_backup = "select max(Backuptimes) as Backuptimes from backupdate";
                con.Open();
                OleDbCommand com_backup = new OleDbCommand(select_backup, con);
                OleDbDataReader dir_backup = com_backup.ExecuteReader();
                if (dir_backup.Read())
                {
                    end_backup_time = Convert.ToString(dir_backup["Backuptimes"]);
                }
                dir_backup.Close();
                OleDbCommand com = new OleDbCommand(select, con);
                OleDbDataReader dir = com.ExecuteReader();
                if (dir.Read())
                {
                    Hertz = Convert.ToString(dir["Hertz"]);
                    if (end_backup_time == "")//若还没有备份过则用设置的开始日期
                    {
                        end_backup_time = Convert.ToString(dir["Startdate"]);//开始日期
                    }
                    TimeSpan ts = Convert.ToDateTime(now_time.ToShortDateString()) - Convert.ToDateTime(Convert.ToDateTime(end_backup_time).AddDays(-1).ToShortDateString());//计算两个时间差
                    DateTime Backuptimes = Convert.ToDateTime(dir["Backuptimes"]);//具体备份时间
                    string short_time = Convert.ToString(Backuptimes.ToString("HH:mm:ss"));//具体备份时间（只含时分秒）
                    string now_short_time = Convert.ToString(now_time.ToString("HH:mm:ss"));//具体备份时间（只含时分秒）
                    string Datetype = Convert.ToString(dir["Datetype"]);//结束时间类型
                    DateTime Enddate = Convert.ToDateTime(dir["Enddate"]);//结束日期
                    TimeSpan endts = Enddate - now_time;//计算是否没有过结束时间
                    string Sourcepath = Convert.ToString(dir["Sourcepath"]);//目标路径
                    string Backuppath = Convert.ToString(dir["Backuppath"]);//备份路径
                    string Weekdate = Convert.ToString(dir["Weekdate"]);//具体星期几
                    string week_now = Convert.ToString(DateTime.Now.DayOfWeek);
                    string Monthtype = Convert.ToString(dir["Monthtype"]);//月份发生类型
                    string Whenday = Convert.ToString(dir["Whenday"]);//第几天
                    int ys = now_time.Year - Convert.ToDateTime(end_backup_time).Year;
                    int ms = now_time.Month - Convert.ToDateTime(end_backup_time).Month;
                    int span = ys * 12 + ms;//当前时间与最生备份时间相差月份
                    int Amonthcount = Convert.ToInt32(dir["Amonthcount"]);//几个月(A类)
                    int Bmonthcount = Convert.ToInt32(dir["Bmonthcount"]);//几个月(B类)
                    string Bweekdate = Convert.ToString(dir["Bweekdate"]);//具体星期几(B类)
                    int Bweekcount = Convert.ToInt32(dir["Bweekcount"]);//第几个星期(B类)
                    #region//每天类型时备份
                    if (Hertz == "D")
                    {
                        int Daycount = 0;
                        Daycount = Convert.ToInt32(dir["Daycount"]);//每多少天发生一次
                        if (ts.Days >= Daycount)//ts.Days获取的为天
                        {
                            backup(short_time, now_short_time, Datetype, Backuppath, Sourcepath, endts);
                        }
                    }
                    #endregion
                    #region//W每周类型时备份
                    if (Hertz == "W")
                    {
                        int Weekcount = 0;
                        Weekcount = Convert.ToInt32(dir["Weekcount"]);//每多少周发生一次
                        if (ts.Days >= (Weekcount * 7))
                        {
                            if (Weekdate.IndexOf(week_now) > -1)
                            {
                                backup(short_time, now_short_time, Datetype, Backuppath, Sourcepath, endts);
                            }

                        }
                    }
                    #endregion
                    #region//M每月类型时备份
                    if (Hertz == "M")
                    {
                        if (Monthtype == "A")
                        {
                            if (Convert.ToString(now_time.Day) == Whenday)
                            {
                                if (span == Amonthcount)
                                {
                                    backup(short_time, now_short_time, Datetype, Backuppath, Sourcepath, endts);
                                }
                            }
                        }
                        if (Monthtype == "B")
                        {
                            if (span >= Bmonthcount)
                            {
                                if (Bweekdate == str(now_time.DayOfWeek.ToString()))
                                {
                                    if (Bweekcount == WeekOfMonth(now_time, 1))
                                    {
                                        backup(short_time, now_short_time, Datetype, Backuppath, Sourcepath, endts);
                                    }

                                }

                            }

                        }
                    }
                    #endregion
                
                }
                dir.Close();
                con.Close();
                #endregion

            }

        }
        private void backup(string short_time, string now_short_time, string Datetype, string Backuppath, string Sourcepath, TimeSpan endts)
        {
            DateTime Backuptimes = DateTime.Now;

            string backname = Convert.ToString(DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString());
            if (short_time == now_short_time)
            {
                if (Datetype == "N")
                {
                    if (Backuppath != "" && Sourcepath != "")
                    {
                        if (Directory.Exists(Backuppath + "\\" + DateTime.Now.Month.ToString()) == false)
                        {
                            Directory.CreateDirectory(Backuppath + "\\" + backname);//创建备份文件夹
                        }
                        // CopyDirectory(textBox2.Text, textBox1.Text + "\\" + DateTime.Now.Month.ToString() + "\\" + DateTime.Now.Date.ToShortDateString());//复制文件
                        CopyDirectory(Sourcepath, Backuppath + "\\" + backname);//复制文件

                    }
                }
                else
                {
                    if (endts.Milliseconds > 0)
                    {
                        if (Backuppath != "" && Sourcepath != "")
                        {
                            if (Directory.Exists(Backuppath + "\\" + DateTime.Now.Month.ToString()) == false)
                            {
                                Directory.CreateDirectory(Backuppath + "\\" + backname);//创建备份件夹
                            }
                            // CopyDirectory(textBox2.Text, textBox1.Text + "\\" + DateTime.Now.Month.ToString() + "\\" + DateTime.Now.Date.ToShortDateString());//复制文件
                            CopyDirectory(Sourcepath, Backuppath + "\\" + backname);//复制文件

                        }
                    }
                }
                OleDbConnection con = new OleDbConnection(constr);
                con.Open();
                string insert = "insert into backupdate(Backuptimes) values ('" + Backuptimes + "')";
                OleDbCommand com = new OleDbCommand(insert, con);
                com.ExecuteNonQuery();
                con.Close();
            }
        }

 
        private void panel_visible(Panel panel)
        {
            panel.Visible = true;
            this.groupBox2.Controls.Add(panel);
            panel.Location = new System.Drawing.Point(6, 18);
            panel.Name = "xx";
            panel.Size = new System.Drawing.Size(293, 87);
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Text = "每天";
            panel2.Visible = false;
            panel3.Visible = false;
            //  panel1.Visible = true;
            panel_visible(panel1);
            //this.groupBox2.Controls.Add(this.panel1);
            //this.panel1.Location = new System.Drawing.Point(22, 21);
            //this.panel1.Name = "panel1";
            //this.panel1.Size = new System.Drawing.Size(200, 65);
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Text = "每周";
            panel1.Visible = false;
            panel3.Visible = false;
            panel_visible(panel2);
            //this.groupBox2.Controls.Add(this.panel2);
            //this.panel2.Location = new System.Drawing.Point(22, 21);
            //this.panel2.Name = "panel2";
            //this.panel2.Size = new System.Drawing.Size(200, 65);
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            groupBox2.Text = "每月";
            panel1.Visible = false;
            panel2.Visible = false;
            panel_visible(panel3);
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                numericUpDown5.Enabled = false;
                comboBox1.Enabled = false;
                numericUpDown6.Enabled = false;

                numericUpDown3.Enabled = true;
                numericUpDown4.Enabled = true;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true)
            {
                numericUpDown5.Enabled = true;
                comboBox1.Enabled = true;
                numericUpDown6.Enabled = true;

                numericUpDown3.Enabled = false;
                numericUpDown4.Enabled = false;
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked == true)
            {
                dateTimePicker3.Enabled = true;
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked == true)
            {
                dateTimePicker3.Enabled = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(constr);
            con.Open();
            string delete = "delete from setbackup";
            string delete_backupdate = "delete from backupdate";
            OleDbCommand com_de = new OleDbCommand(delete, con);
            OleDbCommand com_de_backupdate = new OleDbCommand(delete_backupdate, con);
            com_de_backupdate.ExecuteNonQuery();
            com_de.ExecuteNonQuery();
            string Hertz = "";//发生频率
            string Weekdate = "";//具体星期几
            string Monthtype = "";//月份发生类型
            string Datetype = "";//结束时间类型
            string Startingup = "";//是否开机启动
            #region//各选项赋值
            if (checkBox1.Checked == true)
            {
                Startingup = "YES";
            }
            else
            {
                Startingup = "NO";
            }
            if (radioButton6.Checked == true)
            {
                Datetype = "Y";
            }
            if (radioButton7.Checked == true)
            {
                Datetype = "N";
            }
            if (radioButton4.Checked == true)
            {
                Monthtype = "A";
            }
            if (radioButton5.Checked == true)
            {
                Monthtype = "B";
            }
            if (radioButton1.Checked == true)
            {
                Hertz = "D";
            }
            if (radioButton2.Checked == true)
            {
                Hertz = "W";
            }
            if (radioButton3.Checked == true)
            {
                Hertz = "M";
            }
            if (checkBox2.Checked == true)
            {
                Weekdate += "  Monday星期一";
            }
            if (checkBox3.Checked == true)
            {
                Weekdate += "  Tuesday星期二";
            }
            if (checkBox4.Checked == true)
            {
                Weekdate += "  Wednesday星期三";
            }
            if (checkBox5.Checked == true)
            {
                Weekdate += "  Thursday星期四";
            }
            if (checkBox6.Checked == true)
            {
                Weekdate += "  Friday星期五";
            }
            if (checkBox7.Checked == true)
            {
                Weekdate += "  Saturday星期六";
            }
            if (checkBox8.Checked == true)
            {
                Weekdate += "  Sunday星期天";
            }
            #endregion
            string insert = "insert into setbackup(Hertz,Daycount,Weekcount,Weekdate,Monthtype,Whenday,Amonthcount,Bweekcount,Bweekdate," +
                "Bmonthcount,Backuptimes,Startdate,Enddate,Datetype,Sourcepath,Backuppath,Startingup,Torun) values " +
                "('" + Hertz + "','" + numericUpDown1.Value + "','" + numericUpDown2.Value + "','" + Weekdate + "','" + Monthtype + "','" + numericUpDown3.Value + "'," +
                "'" + numericUpDown4.Value + "','" + numericUpDown5.Value + "','" + comboBox1.Text + "','" + numericUpDown6.Value + "','" + dateTimePicker1.Value + "'," +
                "'" + dateTimePicker2.Value + "','" + dateTimePicker3.Value + "','" + Datetype + "','" + textBox2.Text + "','" + textBox1.Text + "'," +
                "'" + Startingup + "','YES')";
            OleDbCommand com_insert = new OleDbCommand(insert, con);
            com_insert.ExecuteNonQuery();
            con.Close();
            if (checkBox1.Checked == true)
            {
                computer_run(1);
            }
            else
            {
                computer_run(0);
            }
            dbind2();
            Torun = "YES";//设置执行备份
            MessageBox.Show("设置成功！");
        }
        private string str(string week)
        {
            string value = "";
            if (week == "Monday")
            {
                value = "星期一";
            }
            if (week == "Tuesday")
            {
                value = "星期二";
            }
            if (week == "Wednesday")
            {
                value = "星期三";
            }
            if (week == "Thursday")
            {
                value = "星期四";
            }
            if (week == "Friday")
            {
                value = "星期五";
            }
            if (week == "Saturday")
            {
                value = "星期六";
            }
            if (week == "Sunday")
            {
                value = "星期天";
            }
            return value;
        }
        //参数说明：day:要判断的日期,WeekStart：1周一为一周的开始，2周日为一周的开始
        private int WeekOfMonth(DateTime day, int WeekStart)
        {
            //WeekStart
            //1表示 周一至周日 为一周
            //2表示 周日至周六 为一周
            DateTime FirstofMonth;
            FirstofMonth = Convert.ToDateTime(day.Date.Year + "-" + day.Date.Month + "-" + 1);

            int i = (int)FirstofMonth.Date.DayOfWeek;
            if (i == 0)
            {
                i = 7;
            }

            if (WeekStart == 1)
            {
                return (day.Date.Day + i - 2) / 7 + 1;
            }
            if (WeekStart == 2)
            {
                return (day.Date.Day + i - 1) / 7;

            }
            return 0;
            //错误返回值0
        }
        //开机是否启动1启动，0不启动
        private void computer_run(int aa)
        {
            //获取程序执行路径
            string starupPath = Application.ExecutablePath;
            RegistryKey local = Registry.LocalMachine;
            RegistryKey run = local.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run");//注册表启动地址
           
                //SetValue:存储值的名称
                if (aa == 1)
                {
                     try
                     {
                         run.SetValue("WuBackup", starupPath);
                     }
                     catch (Exception ex)
                     {
                         MessageBox.Show(ex.Message.ToString(), "开机启动/取消设置出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                     }
                }
                else
                {
                    try
                    {
                        run.DeleteValue("WuBackup");// 取消开机启动
                    }
                    catch
                    {
                        //什么都不做
                    }
                  

                }
                local.Close();

          
        }
        private void button5_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(constr);
            con.Open();
            string update = "update setbackup set Torun='NO'";
            OleDbCommand com = new OleDbCommand(update, con);
            com.ExecuteNonQuery();
            con.Close();
            Torun = "NO";
            label16.Text = "你的设置为：不进行自动备份。";
            label17.Text = "";
            MessageBox.Show("成功取消！");
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if(e.Button==MouseButtons.Left)//左键双击
            {
                //this.Show()必须重复写两次才能正常显示
                this.Show();//窗口显示
                this.WindowState = FormWindowState.Normal;
                this.Activate();
                this.Show();//窗口显示
               
            }
        }

        private void 测试一ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //this.Show()必须重复写两次才能正常显示
            this.Show();//窗口显示
            this.WindowState = FormWindowState.Normal;
            this.Activate();
            this.Show();//窗口显示
           
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;//桌面图标隐藏
            this.Close();
            this.Dispose();//消除资源
            Application.Exit();//程序退出
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", "http://www.zhnetfans.com");

        }


    }
}
