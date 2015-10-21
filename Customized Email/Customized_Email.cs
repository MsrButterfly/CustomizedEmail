using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Mime;
using System.Net.Mail;

namespace Customized_Email
{
    public partial class Customized_Email : Form
    {
        int Amount = 0;
        string DataFile = String.Format("{0}/Profiles.ini", AppDomain.CurrentDomain.BaseDirectory);
        string SettingFile = String.Format("{0}/Settings.ini", AppDomain.CurrentDomain.BaseDirectory);

        #region [初始化]
        public Customized_Email()
        {
            InitializeComponent();
            INIOperator iniFile = new INIOperator(DataFile);
            String Year = iniFile.ReadString("Time", "Year", "");
            String Month = iniFile.ReadString("Time", "Month", "");
            String Day = iniFile.ReadString("Time", "Day", "");
            if (Year != Convert.ToString(DateTime.Today.Year) || Month != Convert.ToString(DateTime.Today.Month) || Day != Convert.ToString(DateTime.Today.Day))
            {
                for (int i = 0; i <= 19; i++)
                {
                    iniFile.WriteString("Status", Convert.ToString(i), "Failed");
                }
            }
            iniFile.WriteString("Time", "Year", Convert.ToString(DateTime.Today.Year));
            iniFile.WriteString("Time", "Month", Convert.ToString(DateTime.Today.Month));
            iniFile.WriteString("Time", "Day", Convert.ToString(DateTime.Today.Day));
        }
        #endregion
        
        #region [定义控件组]
        String[] Names = new String[20];
        String[] Accounts = new String[20];
        String[] Status = new String[20];
        Label[] No = new Label[10];
        ComboBox[] Receiver = new ComboBox[10];
        TextBox[] Email = new TextBox[10];
        TextBox[] File = new TextBox[10];
        TextBox[] Time = new TextBox[10];
        ProgressBar[] Progress = new ProgressBar[10];
        PictureBox[] Flag = new PictureBox[10];
        Button[] Send = new Button[10];
        #endregion

        #region [开启窗口时生成控件]
        private void Customized_Email_Load(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(SettingFile);

            Customized_Email_Activated(sender, e);

            #region [生成Labels]
            {
                Label Numbers_Label = new Label();
                Label Names_Label = new Label();
                Label Accounts_Label = new Label();
                Label Files_Label = new Label();
                Label Time_Label = new Label();
                Label Status_Label = new Label();
                Numbers_Label.Text = "序号";
                Numbers_Label.SetBounds(22, 40, 29, 12);
                Names_Label.Text = "姓名";
                Names_Label.SetBounds(60, 40, 29, 12);
                Accounts_Label.Text = "电子邮件地址";
                Accounts_Label.SetBounds(130, 40, 77, 12);
                Files_Label.Text = "文件";
                Files_Label.SetBounds(245, 40, 77, 12);
                Time_Label.Text = "最后一次成功发送时间";
                Time_Label.SetBounds(420, 40, 125, 12);
                Status_Label.Text = "发送状态";
                Status_Label.SetBounds(546, 40, 53, 12);
                this.Controls.Add(Numbers_Label);
                this.Controls.Add(Names_Label);
                this.Controls.Add(Accounts_Label);
                this.Controls.Add(Files_Label);
                this.Controls.Add(Time_Label);
                this.Controls.Add(Status_Label);
                for (int i = 0; i <= 9; i++)
                {
                    No[i] = new Label();
                    No[i].Text = (i + 1).ToString().PadLeft(2, '0');
                    No[i].SetBounds(28, 64 + i * 30, 17, 12);
                    this.Controls.Add(No[i]);
                }
            }
            #endregion

            #region[生成姓名ComboBox]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Receiver[i] = new ComboBox();
                    Receiver[i].SetBounds(60, 59 + i * 30, 65, 20);
                    Receiver[i].DropDownStyle = ComboBoxStyle.DropDownList;
                    Receiver[i].Tag = Convert.ToString(i);
                    Receiver[i].SelectedIndexChanged += new System.EventHandler(this.MatchData);
                    this.Controls.Add(Receiver[i]);
                }
            }
            #endregion

            #region [生成电子邮件地址TextBox]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Email[i] = new TextBox();
                    Email[i].SetBounds(130, 59 + i * 30, 110, 21);
                    Email[i].ReadOnly = true;
                    this.Controls.Add(Email[i]);
                }
            }
            #endregion

            #region [生成文件TextBox]
            {
                for (int i = 0; i <= 9; i++)
                {
                    File[i] = new TextBox();
                    File[i].SetBounds(245, 59 + i * 30, 170, 21);
                    File[i].ReadOnly = true;
                    this.Controls.Add(File[i]);
                }
            }
            #endregion

            #region [生成最后一次成功发送时间TextBox]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Time[i] = new TextBox();
                    Time[i].SetBounds(420, 59 + i * 30, 120, 21);
                    Time[i].ReadOnly = true;
                    this.Controls.Add(Time[i]);
                }
            }
            #endregion

            #region [生成发送ProgressBar]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Progress[i] = new ProgressBar();
                    Progress[i].SetBounds(546, 59 + i * 30, 120, 20);
                    this.Controls.Add(Progress[i]);
                }
            }
            #endregion

            #region [生成发送PictureBox]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Flag[i] = new PictureBox();
                    Flag[i].SetBounds(666, 59 + i * 30, 20, 20);
                    Flag[i].BackColor = Color.Red;
                    this.Controls.Add(Flag[i]);
                }
            }
            #endregion

            #region [生成发送Button]
            {
                for (int i = 0; i <= 9; i++)
                {
                    Send[i] = new Button();
                    Send[i].SetBounds(697, 58 + i * 30, 38, 22);
                    Send[i].Text = "发送";
                    Send[i].Tag = Convert.ToString(i);
                    Send[i].Click += new System.EventHandler(this.SendEmail);
                    this.Controls.Add(Send[i]);
                }
            }
            #endregion

            #region [载入上次保存界面]
            {
                ID_I.Text = iniFile.ReadString("ID", "ID", "");
                Password_I.Text = iniFile.ReadString("Password", "Password", "");
                Server_I.Text = iniFile.ReadString("Server", "Server", "");
                for (int i = 0; i <= 9; i++)
                {
                    Receiver[i].Items.Clear();
                    Receiver[i].Items.Add("");
                    for (int j = 0; j <= Amount - 1; j++)
                    {
                        Receiver[i].Items.Add(Names[j]);
                    }
                    int x = Convert.ToInt32(iniFile.ReadString("Number", Convert.ToString(i), ""));
                    Receiver[i].SelectedIndex = x;
                    MatchData(Receiver[i], e);
                }
            }
            #endregion
        }
        #endregion

        #region [窗口获取焦点时刷新收件人列表]
        private void Customized_Email_Activated(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(DataFile);
            Amount = Convert.ToInt32(iniFile.ReadString("Amount", "Amount", ""));
            for (int i = 0; i <= 19; i++)
            {
                Names[i] = iniFile.ReadString("Names", Convert.ToString(i), "");
                Accounts[i] = iniFile.ReadString("Accounts", Convert.ToString(i), "");
                Status[i] = iniFile.ReadString("Status", Convert.ToString(i), "");
            }
        }
        #endregion

        #region [收件人设定]
        private void 键入收件人及邮箱地址IToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("此操作将导致主界面和发送状态被重置，继续吗？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK)
            {
                INIOperator iniFile = new INIOperator(DataFile);
                恢复默认值RToolStripMenuItem_Click(sender, e);
                for (int i = 0; i <= 19; i++)
                {
                    iniFile.WriteString("Status", Convert.ToString(i), "Failed");
                }
                Email_Accounts Email = new Email_Accounts();
                Email.ShowDialog();
            }
            else
            {
                return;
            }
        }
        #endregion

        #region [匹配数据]
        private void MatchData(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((ComboBox)sender).Tag);
            int j = ((ComboBox)sender).SelectedIndex - 1;
            if (j == -1)
            {
                Email[i].Text = "";
                File[i].Text = "";
                Time[i].Text = "";
                Progress[i].Value = 0;
                Flag[i].BackColor = Color.Red;
                return;
            }
            Email[i].Text = Accounts[j];
            File[i].Text = String.Concat("日常工作记录表(", Names[j], ").xls");
            if (Status[j] == "Succeeded")
            {
                Progress[i].Value = 100;
                Flag[i].BackColor = Color.Green;
            }
            else
            {
                Progress[i].Value = 0;
                Flag[i].BackColor = Color.Red;
            }
        }
        #endregion        

        private void SendEmail(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((Button)sender).Tag);
            String ID = ID_I.Text;
            String Password = Password_I.Text;
            NetworkCredential myCredentials = new NetworkCredential(ID, Password);
            Progress[i].Value = 10;
            MailAddress from = new MailAddress("x@x.com");
            try
            {
                from = new MailAddress(ID);
            }
            catch(Exception ex)
            {
                MessageBox.Show("请输入正确的发件人名称。\n\n详细：\n" + Convert.ToString(ex));
                Progress[i].Value = 0;
                return;
            }
            MailAddress to = new MailAddress("x@x.com");
            Progress[i].Value = 20;
            try
            {
                to = new MailAddress(Email[i].Text);
            }
            catch(Exception ex)
            {
                MessageBox.Show("请选择收件人并确保收件人邮箱地址输入正确。\n\n详细：\n" + Convert.ToString(ex));
                Progress[i].Value = 0;
                return;
            }
            MailMessage Message = new MailMessage(from, to);            
            Message.Subject = File[i].Text;
            Message.SubjectEncoding = System.Text.Encoding.UTF8;
            Message.Body = "";
            Message.BodyEncoding = System.Text.Encoding.UTF8;
            Progress[i].Value = 30;
            Attachment attachFile = null;
            try
            {
                attachFile = new Attachment(File[i].Text);
                Message.Attachments.Add(attachFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("请确定您已经把本程序放在与邮件附件相同的文件夹下，且附件名称相符。\n\n详细：\n" + Convert.ToString(ex));
                Progress[i].Value = 0;
                return;
            }
            String Server = Server_I.Text;
            SmtpClient Client = new SmtpClient(Server);
            Progress[i].Value = 40;
            Client.Credentials = myCredentials;
            Progress[i].Value = 50;
            try
            {
                Client.Send(Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("无法发送邮件，请检查用户名、密码与服务器输入是否有误，并确定网络连接正常。\n\n详细：\n" + Convert.ToString(ex));
                Progress[i].Value = 0;
                return;
            }
            Progress[i].Value = 100;
            Flag[i].BackColor = Color.Green;
            Time[i].Text = DateTime.Now.ToString();
            int j = Receiver[Convert.ToInt32(((Button)sender).Tag)].SelectedIndex-1;
            Status[j] = "Succeeded";
            INIOperator iniFile = new INIOperator(DataFile);
            iniFile.WriteString("Status", Convert.ToString(j), "Succeeded");
        }

        private void 恢复默认值RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= 9; i++)
            {
                Receiver[i].Items.Clear();
                Receiver[i].Items.Add("");
                for (int j = 0; j <= Amount - 1; j++)
                {
                    Receiver[i].Items.Add(Names[j]);
                }
                Email[i].Text = "";
                File[i].Text = "";
                Time[i].Text = "";
                Progress[i].Value = 0;
                Flag[i].BackColor = Color.Red;
            }
        }

        private void 保存SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(SettingFile);
            iniFile.WriteString("ID", "ID", ID_I.Text);
            iniFile.WriteString("Password", "Password", Password_I.Text);
            iniFile.WriteString("Server", "Server", Server_I.Text);
            int tmp = 0;
            for(int i=0;i<=9;i++)
            {
                if (Receiver[i].SelectedIndex == -1)
                {
                    tmp = 0;
                }
                else
                {
                    tmp = Receiver[i].SelectedIndex;
                }
                iniFile.WriteString("Number", Convert.ToString(i), Convert.ToString(tmp));
            }
        }

        private void 读取LToolStripMenuItem_Click(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(SettingFile);
            ID_I.Text = iniFile.ReadString("ID", "ID", "");
            Password_I.Text = iniFile.ReadString("Password", "Password", "");
            Server_I.Text = iniFile.ReadString("Server", "Server", "");
            for (int i = 0; i <= 9; i++)
            {
                Receiver[i].Items.Clear();
                Receiver[i].Items.Add("");
                for (int j = 0; j <= Amount - 1; j++)
                {
                    Receiver[i].Items.Add(Names[j]);
                }
                int x = Convert.ToInt32(iniFile.ReadString("Number", Convert.ToString(i), ""));
                Receiver[i].SelectedIndex = x;                
                MatchData(Receiver[i], e);
            }
        }

        private void 退出XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
