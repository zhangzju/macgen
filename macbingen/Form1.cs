using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
namespace macbingen
{
    public partial class Form1 : Form
    {
        static bool isUsername;
        static decimal usernameLength;
        static bool isPassword;
        static decimal passwordLength;
        static bool isPppoeUsername;
        static decimal pppoeUsernameLength;
        static bool isPppoePassword;
        static decimal pppoePasswordLength;
        static decimal count;
        static string filePath = Application.StartupPath + "\\MACBIN.xlsx";
        static List<string> titleList = new List<string>();
        static List<string> pinList = new List<string>();
        static List<string> usernameList = new List<string>();
        static List<string> passwordList = new List<string>();
        static List<string> pppoeUsernameList = new List<string>();
        static List<string> pppoePaswordList = new List<string>();

        public Form1()
        {
            InitializeComponent();
            this.numericUpDown2.Enabled = false;
            this.numericUpDown2.Value = 8;
            this.numericUpDown3.Enabled = false;
            this.numericUpDown3.Value = 8;
            this.numericUpDown4.Enabled = false;
            this.numericUpDown4.Value = 8;
            this.numericUpDown5.Enabled = false;
            this.numericUpDown5.Value = 8;
            this.numericUpDown6.Value = 8;
            this.textBox1.Text = filePath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            isUsername = this.checkBox2.Checked;
            usernameLength = this.numericUpDown2.Value;
            isPassword = this.checkBox3.Checked;
            passwordLength = this.numericUpDown3.Value;
            isPppoeUsername = this.checkBox4.Checked;
            pppoeUsernameLength = this.numericUpDown4.Value;
            isPppoePassword = this.checkBox5.Checked;
            pppoePasswordLength = this.numericUpDown5.Value;
            count = this.numericUpDown6.Value;

            titleList.Add("MAC Address");
             titleList.Add("Username");
            titleList.Add("Password");
            titleList.Add("WirelessKey");
            titleList.Add("PPPoE Username");
            titleList.Add("PPPoE Password");
            titleList.Add("IP Address");
            titleList.Add("Mask");
            titleList.Add("Gateway");
            titleList.Add("DNS");
            titleList.Add("2.4GHZ SSID");
            titleList.Add("2.4GHZ Guest SSID");
            titleList.Add("5GHZ SSID");
            titleList.Add("5GHZ Guest SSID");
            titleList.Add("Generate Flag");

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("CONFIG INFO");
            IRow row;
            ICell cell;

            row = sheet.CreateRow(0);
            int tempcount = 0;
            foreach(var title in titleList)
            {
                cell = row.CreateCell(tempcount);
                cell.SetCellValue(title);
                sheet.SetColumnWidth(tempcount, 20 * 256);
                tempcount++;
            }

            
            for (decimal index = 0; index < count; index++)
            {
                row = sheet.CreateRow(Convert.ToInt32(index) + 1);
              
                if(isUsername)
                {
                    cell = row.CreateCell(1);
                    cell.SetCellValue(getRandomizer(Convert.ToInt32(usernameLength), true, true, true, false));
                }
                
                if(isPassword)
                {
                    cell = row.CreateCell(2);
                    cell.SetCellValue(getRandomizer(Convert.ToInt32(passwordLength), true, true, true, false));
                }

                if(isPppoeUsername)
                {
                    cell = row.CreateCell(4);
                    cell.SetCellValue(getRandomizer(Convert.ToInt32(pppoeUsernameLength), true, true, true, false));
                }
                 
                if(isPppoePassword)
                {
                    cell = row.CreateCell(5);
                    cell.SetCellValue(getRandomizer(Convert.ToInt32(pppoePasswordLength), true, true, true, false));
                }

                if(true)
                {
                    cell = row.CreateCell(14);
                    cell.SetCellValue(0);
                }
            }

            if(File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            FileStream sw = File.Create(filePath);
            workbook.Write(sw);
            
            sw.Close();
            MessageBox.Show("Generate success in" + filePath);

        }

        public string getRandomPin(int length)
        {
            string pin;
            if(length != 8)
            {
                return getRandomizer(length, true, false, false, false);
            }

            pin = getRandomizer(7, true, false, false, false);

            int pinTemp = Convert.ToInt32(pin[0]) * 3 + Convert.ToInt32(pin[1]) * 1 + Convert.ToInt32(pin[2]) * 3 + Convert.ToInt32(pin[3]) * 1 + Convert.ToInt32(pin[4]) * 3 + Convert.ToInt32(pin[5]) * 1 + Convert.ToInt32(pin[6]) * 3;

            pin = pin + Convert.ToString(pinTemp%10);
            //MessageBox.Show(pin + ";" + Convert.ToString(pinTemp % 10));
            return pin;
        }

        public string getRandomizer(int length, bool useNum, bool useLow, bool useUpp, bool useSpe)
        {
            byte[] b = new byte[4];
            new System.Security.Cryptography.RNGCryptoServiceProvider().GetBytes(b);
            Random r = new Random(BitConverter.ToInt32(b, 0));
            string s = null, str = null;
            if (useNum == true) { str += "0123456789"; }
            if (useLow == true) { str += "abcdefghijklmnopqrstuvwxyz"; }
            if (useUpp == true) { str += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; }
            if (useSpe == true) { str += "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"; }
            for (int i = 0; i < length; i++)
            {
                s += str.Substring(r.Next(0, str.Length - 1), 1);
            }
            return s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog configFilePath = new OpenFileDialog();
            configFilePath.Title = "Choose the Configuration file";
            configFilePath.InitialDirectory = Application.StartupPath;
            configFilePath.Filter = "xlsx files(*.xls*)|*.xlsx|All files (*.*)|*.*";
            configFilePath.FilterIndex = 1;

            if(configFilePath.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = configFilePath.FileName;
                filePath = configFilePath.FileName;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.Text = "TP-Link ISP MACBIN generator";
            about.Show();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox2.Checked)
            {
                this.numericUpDown2.Enabled = true;
            }
            else
            {
                this.numericUpDown2.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox3.Checked)
            {
                this.numericUpDown3.Enabled = true;
            }
            else
            {
                this.numericUpDown3.Enabled = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox4.Checked)
            {
                this.numericUpDown4.Enabled = true;
            }
            else
            {
                this.numericUpDown4.Enabled = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox5.Checked)
            {
                this.numericUpDown5.Enabled = true;
            }
            else
            {
                this.numericUpDown5.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string outputpath = Application.StartupPath + "/output";
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "MAC.bin output path";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                outputpath = dialog.SelectedPath + @"\Output";
            }
            this.textBox2.Text = outputpath;
            //MessageBox.Show(filePath + ";" + count.ToString());
            IWorkbook workbook = null;
            try
            {
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    count = 0;
                    //把xls文件读入workbook变量里，之后就可以关闭了  
                    workbook = new XSSFWorkbook(fs);
                    fs.Close();

                    ISheet sheet = workbook.GetSheet("CONFIG INFO");

                    int index = 0;
                    string macAddress;
                    string macbinPath;
                    string username;
                    string password;
                    string wirelesskey;
                    string pppoeusername;
                    string pppoepassword;
                    string ipaddress;
                    string mask;
                    string gateway;
                    string dns;
                    string ssid2;
                    string ssid2guest;
                    string ssid5;
                    string ssid5guest;


                    while (sheet.GetRow(index) != null && sheet.GetRow(index).GetCell(0) != null && !string.IsNullOrEmpty(sheet.GetRow(index).GetCell(0).ToString()))
                    {
                        count++;
                        index++;
                    }

                    if (!Directory.Exists(outputpath))
                    {
                        Directory.CreateDirectory(outputpath);
                    }

                    for (index = 1; index < Convert.ToInt32(count); index++)
                    {
                        IRow row = sheet.GetRow(index);

                        ICell cell = row.GetCell(0);
                        macAddress = cell.ToString();
                        macbinPath = outputpath + "/" + macAddress + ".bin";

                        if (string.IsNullOrWhiteSpace(macAddress))
                        {
                            MessageBox.Show("Illegal mac address!");
                        }

                        FileStream macbinStream = File.OpenWrite(macbinPath);
                        StreamWriter writer = new StreamWriter(macbinStream);

                        cell = row.GetCell(1);
                        if (cell != null)
                        {
                            username = cell.ToString();
                            writer.WriteLine("username:" + username);
                        }

                        cell = row.GetCell(2);
                        if (cell != null)
                        {
                            password = cell.ToString();
                            writer.WriteLine("password:" + password);
                        }

                        cell = row.GetCell(3);
                        if (cell != null)
                        {
                            wirelesskey = cell.ToString();
                            writer.WriteLine("wirelesskey:" + wirelesskey);
                        }

                        cell = row.GetCell(4);
                        if (cell != null)
                        {
                            pppoeusername = cell.ToString();
                            writer.WriteLine("PPPOE4_username:" + pppoeusername);
                        }

                        cell = row.GetCell(5);
                        if (cell != null)
                        {
                            pppoepassword = cell.ToString();
                            writer.WriteLine("PPPOE4_password:" + pppoepassword);
                        }

                        cell = row.GetCell(6);
                        if (cell != null)
                        {
                            ipaddress = cell.ToString();
                            writer.WriteLine("static_IP4:" + ipaddress);
                        }

                        cell = row.GetCell(7);
                        if (cell != null)
                        {
                            mask = cell.ToString();
                            writer.WriteLine("static_Mask4:" + mask);
                        }

                        cell = row.GetCell(8);
                        if (cell != null)
                        {
                            gateway = cell.ToString();
                            writer.WriteLine("static_GW4:" + gateway);
                        }

                        cell = row.GetCell(9);
                        if (cell != null)
                        {
                            dns = cell.ToString();
                            writer.WriteLine("static_DNS4:" + dns);
                        }

                        cell = row.GetCell(10);
                        if (cell != null)
                        {
                            ssid2 = cell.ToString();
                            writer.WriteLine("SSID_2G_0:" + ssid2);
                        }

                        cell = row.GetCell(11);
                        if (cell != null)
                        {
                            ssid2guest = cell.ToString();
                            writer.WriteLine("SSID_2G_1:" + ssid2guest);
                        }

                        cell = row.GetCell(12);
                        if (cell != null)
                        {
                            ssid5 = cell.ToString();
                            writer.WriteLine("SSID_5G_0:" + ssid5);
                        }

                        cell = row.GetCell(13);
                        if (cell != null)
                        {
                            ssid5guest = cell.ToString();
                            writer.WriteLine("SSID_5G_1:" + ssid5guest);
                        }

                        cell = row.GetCell(14);
                        if (cell != null)
                        {
                            string flag = cell.ToString();
                            if (Convert.ToInt32(flag) == 0)
                            {
                                cell.SetCellValue(1);
                                //MessageBox.Show(cell.ToString());
                            }
                            else if (flag == "1")
                            {
                                writer.Close();
                                File.Delete(macbinPath);
                                continue;
                            }
                        }

                        writer.Close();
                    }

                    MessageBox.Show("Generate success! All file is in" + outputpath);
                }
            }
            catch (Exception error)
            {

                Console.WriteLine(error.Message); //输出错误提示
            }

         

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Introduction introduction = new Introduction();
            introduction.Show();
        }

        private void specificationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
