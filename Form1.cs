using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace AccessEng
{
    public partial class Form1 : Form
    {
        private readonly OleDbConnection connection = new OleDbConnection();
        private readonly string connectString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                Path.Combine(Application.StartupPath, "WorkDB.mdb");


        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = connectString;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckConnection();
        }
        void CheckConnection()
        {
            try
            {
                connection.Open();
                checkConnection.Text = "Соединение с базой данных установлено";
                connection.Close();
            }
            catch (Exception ex)
            {
                checkConnection.Text = "Ошибка подключения к базе данных\nОбратитесь к разработчику";
                Clipboard.SetText(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "admin" && textBox2.Text == "12345")
            {
                MessageBox.Show("Вы успешно вошли на аккаунт", "Успешный вход",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                Hide();
                Form2 f2 = new Form2();
                f2.Show();
            }
            else
            {
                MessageBox.Show("Введены неправильные данные", "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                textBox1.Clear();
                textBox2.Clear();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1_Click(sender, e);
        }
    }
}
