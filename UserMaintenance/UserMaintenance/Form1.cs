using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using UserMaintenance.Entities;

namespace UserMaintenance
{
    public partial class Form1 : Form
    {
        BindingList<User> users = new BindingList<User>();
        public Form1()
        {
            InitializeComponent();
            label1.Text = Resource1.FullName;
            button1.Text = Resource1.Add;
            button2.Text = Resource1.WriteOut;
            listBox1.DataSource = users;
            listBox1.ValueMember = "ID";
            listBox1.DisplayMember = "FullName";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var u = new User() { FullName = textBox1.Text };
            users.Add(u);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName);
                foreach (var item in users)
                {
                    streamWriter.WriteLine(item.ID + "," + item.FullName);
                }
                streamWriter.Close();
            }

        }
    }
}
