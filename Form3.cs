using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace AccessEng
{
    public partial class Form3 : Form
    {
        private readonly string connectString =
           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
               Path.Combine(Application.StartupPath, "WorkDB.mdb");

        private OleDbConnection myConnection;
        public Form3()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
        }
        void ViewDataGrid(OleDbCommand command)
        {
            DataSet dataSet = new DataSet();
            DataTable dataTable = dataSet.Tables.Add("Table");
            dataSet.Load(command.ExecuteReader(), LoadOption.OverwriteChanges, dataTable);
            request_DB.DataSource = dataSet.Tables["Table"];
        }
        void First()
        {
            request_DB.Columns[0].HeaderText = "Количество часов";
            request_DB.Columns[1].HeaderText = "Заказчик";
            request_DB.Columns[2].HeaderText = "Кому оказывается услуга";
            request_DB.Columns[3].HeaderText = "Дата заключения";
            request_DB.Columns[4].HeaderText = "Срок действия";
            request_DB.Columns[5].HeaderText = "Краткое описание договора";

            request_DB.Columns[0].Width = 140;
            request_DB.Columns[1].Width = 200;
            request_DB.Columns[2].Width = 200;
            request_DB.Columns[3].Width = 120;
            request_DB.Columns[4].Width = 120;
            request_DB.Columns[5].Width = 160;
        }
        void Second()
        {
            request_DB.Columns[0].HeaderText = "ФИО";
            request_DB.Columns[1].HeaderText = "Дата рождения";
            request_DB.Columns[2].HeaderText = "Место жительства";
            request_DB.Columns[3].HeaderText = "Место регистрации";
            request_DB.Columns[4].HeaderText = "Количество заключенных договоров";
            request_DB.Columns[5].HeaderText = "Контакты";
            request_DB.Columns[6].HeaderText = "Административное правонарушение";

            request_DB.Columns[0].Width = 180;
            request_DB.Columns[1].Width = 80;
            request_DB.Columns[2].Width = 155;
            request_DB.Columns[3].Width = 155;
            request_DB.Columns[4].Width = 100;
            request_DB.Columns[5].Width = 120;
            request_DB.Columns[6].Width = 150;
        }
        void Third()
        {
            request_DB.Columns[0].HeaderText = "ФИО";
            request_DB.Columns[1].HeaderText = "Место жительства";
            request_DB.Columns[2].HeaderText = "Место регистрации";
            request_DB.Columns[3].HeaderText = "Группа инвалидности";
            request_DB.Columns[4].HeaderText = "Краткое описание ИПР";

            request_DB.Columns[0].Width = 220;
            request_DB.Columns[1].Width = 150;
            request_DB.Columns[2].Width = 150;
            request_DB.Columns[3].Width = 150;
            request_DB.Columns[4].Width = 220;
        }
        void Fourth1()
        {
            request_DB.Columns[0].HeaderText = "Наименование организации";
            request_DB.Columns[1].HeaderText = "Юридический адрес";
            request_DB.Columns[2].HeaderText = "ФИО персонального ассистента";
            request_DB.Columns[3].HeaderText = "Контакты";
            request_DB.Columns[4].HeaderText = "Перечень услуг";

            request_DB.Columns[0].Width = 180;
            request_DB.Columns[1].Width = 150;
            request_DB.Columns[2].Width = 200;
            request_DB.Columns[3].Width = 160;
            request_DB.Columns[4].Width = 220;
        }
        void Fourth2()
        {
            request_DB.Columns[0].HeaderText = "ФИО";
            request_DB.Columns[1].HeaderText = "Место жительства";
            request_DB.Columns[2].HeaderText = "Контакты";
            request_DB.Columns[3].HeaderText = "Перечень услуг";

            request_DB.Columns[0].Width = 210;
            request_DB.Columns[1].Width = 160;
            request_DB.Columns[2].Width = 180;
            request_DB.Columns[3].Width = 220;
        }
        void Fifth()
        {
            request_DB.Columns[0].HeaderText = "ФИО";
            request_DB.Columns[1].HeaderText = "Место жительства";
            request_DB.Columns[2].HeaderText = "Место регистрации";
            request_DB.Columns[3].HeaderText = "Контактный номер";
            request_DB.Columns[4].HeaderText = "e-mail";

            request_DB.Columns[0].Width = 210;
            request_DB.Columns[1].Width = 170;
            request_DB.Columns[2].Width = 170;
            request_DB.Columns[3].Width = 150;
            request_DB.Columns[4].Width = 160;
        }
        void Sixth()
        {
            request_DB.Columns[0].HeaderText = "ФИО";
            request_DB.Columns[1].HeaderText = "Контакты";
            request_DB.Columns[2].HeaderText = "Номер договора";
            request_DB.Columns[3].HeaderText = "Количество часов в неделю";
            request_DB.Columns[4].HeaderText = "Количество заключенных договоров";

            request_DB.Columns[0].Width = 220;
            request_DB.Columns[1].Width = 200;
            request_DB.Columns[2].Width = 160;
            request_DB.Columns[3].Width = 180;
            request_DB.Columns[4].Width = 140;
        }
        public void UnCheckAllForm3Buttons()
        {
            request1_button.Enabled = false;
            request1_button.Visible = false;

            request2_button.Enabled = false;
            request2_button.Visible = false;

            request3_button.Enabled = false;
            request3_button.Visible = false;

            request4_button.Enabled = false;
            request4_button.Visible = false;

            request5_button.Enabled = false;
            request5_button.Visible = false;

            request6_button.Enabled = false;
            request6_button.Visible = false;

            request_comboBoxFor5.Enabled = false;
            request_comboBoxFor5.Visible = false;

            request_textBox_for6.Enabled = false;
            request_textBox_for6.Visible = false;
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (request_comboBox.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Personal_assistents.hours, Contracts.customer, Contracts.to_whom_spec_service, Contracts.date_of_contract, Contracts.date_of_end, Contracts.description_contract 
                                 FROM Contracts, Personal_assistents, Physical_disableds
                                 WHERE (Physical_disableds.num_passport = @Passport) and(Physical_disableds.number_of_contract = Contracts.number_of_contract) and (Contracts.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);  
                command.Parameters.AddWithValue("@Passport", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                First();
            }
            else if (request_comboBox.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Personal_assistents.hours, Contracts.customer, Contracts.to_whom_spec_service, Contracts.date_of_contract, Contracts.date_of_end, Contracts.description_contract 
                                 FROM Contracts, Personal_assistents, Physical_disableds
                                 WHERE (Physical_disableds.number_of_contract = @NumberContract) and(Physical_disableds.number_of_contract = Contracts.number_of_contract) and (Contracts.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@NumberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                First();
            }
            else
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void request2_button_Click(object sender, EventArgs e)
        {
            if(request_comboBox.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Count(Personal_assistents.number_of_contract), Personal_assistents.contacts, Criminal_record.admin_offense
                                 FROM Physical_pers_assists, Personal_assistents, Criminal_record
                                 WHERE (Physical_pers_assists.num_passport = @NumberPassport) and (Physical_pers_assists.fio = Personal_assistents.fio) and (Criminal_record.fio = Physical_pers_assists.fio)
                                 GROUP BY Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Personal_assistents.contacts, Criminal_record.admin_offense";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@NumberPassport", request_textBox.Text);
                ViewDataGrid(command);

                myConnection.Close();
                Second();
            }
            else if(request_comboBox.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Count(Personal_assistents.number_of_contract), Personal_assistents.contacts, Criminal_record.admin_offense
                                 FROM Physical_pers_assists, Personal_assistents, Criminal_record
                                 WHERE (Physical_pers_assists.fio = @FIO) and (Physical_pers_assists.fio = Personal_assistents.fio) and (Criminal_record.fio = Physical_pers_assists.fio)
                                 GROUP BY Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Personal_assistents.contacts, Criminal_record.admin_offense";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@FIO", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Second();
            }
            else if (request_comboBox.SelectedIndex == 2)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Count(Personal_assistents.number_of_contract), Personal_assistents.contacts, Criminal_record.admin_offense
                                 FROM Physical_pers_assists, Personal_assistents, Criminal_record
                                 WHERE (Physical_pers_assists.number_of_contract = @numberContract) and (Physical_pers_assists.fio = Personal_assistents.fio) and (Criminal_record.fio = Physical_pers_assists.fio)
                                 GROUP BY Physical_pers_assists.fio, Physical_pers_assists.date_birth, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Personal_assistents.contacts, Criminal_record.admin_offense";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Second();
            }
            else
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void request3_button_Click(object sender, EventArgs e)
        {
            if(request_comboBox.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Physical_disableds.fio, Physical_disableds.place_of_living, Physical_disableds.place_of_registration, Info_from_medical_cart.disability_group, Info_from_medical_cart.brief_description_of_IPR
                                 FROM Physical_disableds, Info_from_medical_cart
                                 WHERE (Physical_disableds.num_passport = @numberPassport) and (Physical_disableds.fio = Info_from_medical_cart.fio)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberPassport", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Third();
            }
            else if (request_comboBox.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Physical_disableds.fio, Physical_disableds.place_of_living, Physical_disableds.place_of_registration, Info_from_medical_cart.disability_group, Info_from_medical_cart.brief_description_of_IPR
                                 FROM Physical_disableds, Info_from_medical_cart
                                 WHERE (Physical_disableds.fio = @FIO) and (Physical_disableds.fio = Info_from_medical_cart.fio)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@FIO", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Third();
            }
            else if (request_comboBox.SelectedIndex == 2)
            {
                myConnection.Open();
                string query = @"SELECT Physical_disableds.fio, Physical_disableds.place_of_living, Physical_disableds.place_of_registration, Info_from_medical_cart.disability_group, Info_from_medical_cart.brief_description_of_IPR
                                 FROM Physical_disableds, Info_from_medical_cart
                                 WHERE (Physical_disableds.number_of_contract = @numberContract) and (Physical_disableds.fio = Info_from_medical_cart.fio)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Third();
            }
            else
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void request4_button_Click(object sender, EventArgs e)
        {
            if(request_comboBox.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Legal_executors.name_organisation, Legal_executors.legal_address, Personal_assistents.fio, Personal_assistents.contacts, Personal_assistents.kind_of_services
                                 FROM Legal_executors, Personal_assistents
                                 WHERE (Legal_executors.number_of_contract = @numberContract) and (Legal_executors.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();

                Fourth1();
            }
            else if (request_comboBox.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.place_of_living, Personal_assistents.contacts, Personal_assistents.kind_of_services
                                 FROM Physical_pers_assists, Personal_assistents
                                 WHERE (Physical_pers_assists.number_of_contract = @numberContract) and (Physical_pers_assists.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();

                Fourth2();
            }
            else if (request_comboBox.SelectedIndex == 2)
            {
                myConnection.Open();
                string query = @"SELECT Legal_executors.name_organisation, Legal_executors.legal_address, Personal_assistents.fio, Personal_assistents.contacts, Personal_assistents.kind_of_services
                                 FROM Legal_executors, Personal_assistents
                                 WHERE (Legal_executors.name_organisation = @nameOrganisation) and (Legal_executors.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@nameOrganisation", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();

                Fourth1();
            }
            else if (request_comboBox.SelectedIndex == 3)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.place_of_living, Personal_assistents.contacts, Personal_assistents.kind_of_services
                                 FROM Physical_pers_assists, Personal_assistents
                                 WHERE (Physical_pers_assists.fio = @FIO) and (Physical_pers_assists.number_of_contract = Personal_assistents.number_of_contract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@FIO", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();

                Fourth2();
            }
            else    
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void request5_button_Click(object sender, EventArgs e)
        {
            if (request_comboBox.SelectedIndex == 0 && request_comboBoxFor5.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Physical_disableds.fio, Physical_disableds.place_of_living, Physical_disableds.place_of_registration, Physical_disableds.contact_numbers, Physical_disableds.email
                                 FROM Physical_disableds
                                 WHERE (Physical_disableds.number_of_contract = @numberContract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Fifth();
            }
            else if (request_comboBox.SelectedIndex == 0 && request_comboBoxFor5.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Physical_pers_assists.contact_numbers, Physical_pers_assists.email
                                 FROM Physical_pers_assists
                                 WHERE (Physical_pers_assists.number_of_contract = @numberContract)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@numberContract", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Fifth();
            }
            else if (request_comboBox.SelectedIndex == 1 && request_comboBoxFor5.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Physical_disableds.fio, Physical_disableds.place_of_living, Physical_disableds.place_of_registration, Physical_disableds.contact_numbers, Physical_disableds.email
                                 FROM Physical_disableds
                                 WHERE (Physical_disableds.fio = @FIO)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@FIO", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Fifth();
            }
            else if (request_comboBox.SelectedIndex == 1 && request_comboBoxFor5.SelectedIndex == 1)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Physical_pers_assists.place_of_living, Physical_pers_assists.place_of_registration, Physical_pers_assists.contact_numbers, Physical_pers_assists.email
                                 FROM Physical_pers_assists
                                 WHERE (Physical_pers_assists.fio = @FIO)";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@FIO", request_textBox.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Fifth();
            }
            else
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void request6_button_Click(object sender, EventArgs e)
        {
            if(request_comboBox.SelectedIndex == 0)
            {
                myConnection.Open();
                string query = @"SELECT Physical_pers_assists.fio, Personal_assistents.contacts, Personal_assistents.number_of_contract, Personal_assistents.hours, Count(Personal_assistents.number_of_contract)
                                 FROM Personal_assistents, Physical_pers_assists, Contracts
                                 WHERE (Contracts.date_of_contract BETWEEN @date1 AND @dat2) and (Physical_pers_assists.number_of_contract = Personal_assistents.number_of_contract) and (Contracts.number_of_contract = Physical_pers_assists.number_of_contract)
                                 GROUP BY Physical_pers_assists.fio, Personal_assistents.number_of_contract, Personal_assistents.contacts, Personal_assistents.hours";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.Parameters.AddWithValue("@date1", request_textBox.Text);
                command.Parameters.AddWithValue("@date2", request_textBox_for6.Text);
                ViewDataGrid(command);
                myConnection.Close();
                Sixth();
            }
            else
                MessageBox.Show("Не выбраны все параметры для поиска!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
