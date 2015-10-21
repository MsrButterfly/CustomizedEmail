using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Customized_Email
{
    public partial class Email_Accounts : Form
    {
        string DataFile = String.Format("{0}/Profiles.ini", AppDomain.CurrentDomain.BaseDirectory);
        TextBox[] NameTextBox = new TextBox[20];
        TextBox[] AccountTextBox = new TextBox[20];
        Button[] Add = new Button[20];
        Button[] Delete = new Button[20];
        
        public Email_Accounts()
        {
            InitializeComponent();
        }

        private void Email_Address_Load(object sender, EventArgs e)
        {
            this.Controls.Clear();
            Label Names = new Label();
            Names.Text = "姓名";
            Names.SetBounds(5, 8, 29, 12);
            Label Accounts = new Label();
            Accounts.Text = "电子邮件地址";
            Accounts.SetBounds(90, 8, 77, 12);
            this.Controls.Add(Names);
            this.Controls.Add(Accounts);

            INIOperator iniFile = new INIOperator(DataFile);
            int Amount = Convert.ToInt32(iniFile.ReadString("Amount", "Amount", ""));
            //MessageBox.Show(Convert.ToString(Amount));

            this.Height = 21 * Amount + 55;

            for (int i = 0; i <= 19; i++)
            {
                NameTextBox[i] = new TextBox();
                NameTextBox[i].SetBounds(5, 25 + 21 * i, 80, 21);
                NameTextBox[i].Text = iniFile.ReadString("Names", Convert.ToString(i), "");
                NameTextBox[i].TextChanged += new System.EventHandler(this.SaveProfiles);
                AccountTextBox[i] = new TextBox();
                AccountTextBox[i].SetBounds(90, 25 + 21 * i, 160, 21);
                AccountTextBox[i].Text = iniFile.ReadString("Accounts", Convert.ToString(i), "");
                AccountTextBox[i].TextChanged += new System.EventHandler(this.SaveProfiles);
                Add[i] = new Button();
                Add[i].Text = "+";
                Add[i].Click += new System.EventHandler(this.AddObjects);
                Add[i].SetBounds(255, 25 + 21 * i, 21, 21);
                if (i <= Amount - 2)
                {
                    Add[i].Enabled = false;
                }
                else
                {
                    Add[i].Enabled = true;
                }
                Delete[i] = new Button();
                Delete[i].Text = "-";
                Delete[i].Click += new System.EventHandler(this.DeleteObjects);
                Delete[i].SetBounds(275, 25 + 21 * i, 21, 21);
                if (i <= Amount - 2)
                {
                    Delete[i].Enabled = false;
                }
                else
                {
                    Delete[i].Enabled = true;
                }
            }

            for (int i = 0; i <= Amount - 1; i++)
            {
                this.Controls.Add(NameTextBox[i]);
                this.Controls.Add(AccountTextBox[i]);
                this.Controls.Add(Add[i]);
                this.Controls.Add(Delete[i]);
            }
        }
        
        public void AddObjects(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(DataFile);
            int Amount = Convert.ToInt32(iniFile.ReadString("Amount", "Amount", ""));
            if (Amount <= 19)
            {
                Amount++;
            }
            else
            {
                MessageBox.Show("抱歉，目前本程序只能添加20个收件人，请等待更新版本。");
            }
            iniFile.WriteString("Amount", "Amount", Convert.ToString(Amount));
            this.Email_Address_Load(sender, e);
        }

        public void DeleteObjects(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(DataFile);
            int Amount = Convert.ToInt32(iniFile.ReadString("Amount", "Amount", ""));
            if (Amount >= 2)
            {
                Amount--;
            }
            else
            {
                MessageBox.Show("至少也要存在1个收件人。");
            }
            iniFile.WriteString("Amount", "Amount", Convert.ToString(Amount));
            this.Email_Address_Load(sender, e);
        }

        public void SaveProfiles(object sender, EventArgs e)
        {
            INIOperator iniFile = new INIOperator(DataFile);
            int Amount = Convert.ToInt32(iniFile.ReadString("Amount", "Amount", ""));
            for (int i = 0; i <= 19; i++)
            {
                iniFile.WriteString("Names", Convert.ToString(i), NameTextBox[i].Text);
                iniFile.WriteString("Accounts", Convert.ToString(i), AccountTextBox[i].Text);
            }
        }

    }
}
