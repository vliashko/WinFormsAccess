using System;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace AccessEng
{
    public partial class Form2 : Form
    {
        private OleDbConnection connection = new OleDbConnection();

        private readonly static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                Path.Combine(Application.StartupPath, "WorkDB.mdb");

        private OleDbConnection myConnection;

        public string CurrentTable { get; private set; } = "";
        public Form2()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            this.capacityTableAdapter.Fill(this.dBDataSet.Capacity);
            LoadAllDBFromAccess();
            CloseAllDB();
        }

        public void CheckCriminalRecord()
        {

        }
        void LoadAllDBFromAccess()
        {
            this.physical_representativesTableAdapter.Fill(this.dBDataSet.Physical_representatives);
            this.physical_pers_assistsTableAdapter.Fill(this.dBDataSet.Physical_pers_assists);
            this.physical_disabledsTableAdapter.Fill(this.dBDataSet.Physical_disableds);
            this.personal_assistentsTableAdapter.Fill(this.dBDataSet.Personal_assistents);
            this.legal_executorsTableAdapter.Fill(this.dBDataSet.Legal_executors);
            this.legal_customersTableAdapter.Fill(this.dBDataSet.Legal_customers);
            this.info_from_medical_cartTableAdapter.Fill(this.dBDataSet.Info_from_medical_cart);
            this.individual_programTableAdapter.Fill(this.dBDataSet.Individual_program);
            this.customer_services_personTableAdapter.Fill(this.dBDataSet.Customer_services_person);
            this.customer_services_organisationsTableAdapter.Fill(this.dBDataSet.Customer_services_organisations);
            this.criminal_recordTableAdapter.Fill(this.dBDataSet.Criminal_record);
            this.contractsTableAdapter.Fill(this.dBDataSet.Contracts);
            this.capacityTableAdapter.Fill(this.dBDataSet.Capacity);
        }
        void UpdateAllDB(string table)
        {
            if (table == "DB_Customer_services_person")
            {
                SaveChangesControl.Focus();
                this.customer_services_personTableAdapter.Update(this.dBDataSet.Customer_services_person);
                DB_Customer_services_person.Focus();
            }
            else if (table == "DB_Customer_services_organisations")
            {
                SaveChangesControl.Focus();
                this.customer_services_organisationsTableAdapter.Update(this.dBDataSet.Customer_services_organisations); 
                DB_Customer_services_organisations.Focus();
            }
            else if (table == "DB_Capacity")
            {
                SaveChangesControl.Focus();
                this.capacityTableAdapter.Update(this.dBDataSet.Capacity);
                DB_Capacity.Focus();
            }
            else if(table == "DB_Info_from_medical_cart")
            {
                SaveChangesControl.Focus();
                this.info_from_medical_cartTableAdapter.Update(this.dBDataSet.Info_from_medical_cart);
                DB_Info_from_medical_cart.Focus();
            }
            else if(table == "DB_Individual_program")
            {
                SaveChangesControl.Focus();
                this.individual_programTableAdapter.Update(this.dBDataSet.Individual_program);
                DB_Individual_program.Focus();
            }
            else if(table == "DB_Criminal_record")
            {
                SaveChangesControl.Focus();
                this.criminal_recordTableAdapter.Update(this.dBDataSet.Criminal_record);
                DB_Criminal_record.Focus();
            }
            else if(table == "DB_Contracts")
            {
                SaveChangesControl.Focus();
                this.contractsTableAdapter.Update(this.dBDataSet.Contracts);
                DB_Contracts.Focus();
            }
            else if(table == "DB_Personal_assistents")
            {
                SaveChangesControl.Focus();
                this.personal_assistentsTableAdapter.Update(this.dBDataSet.Personal_assistents);
                DB_Personal_assistents.Focus();
            }
            else if(table == "DB_Physical_representatives")
            {
                SaveChangesControl.Focus();
                this.physical_representativesTableAdapter.Update(this.dBDataSet.Physical_representatives);
                DB_Physical_representatives.Focus();
            }
            else if(table == "DB_Physical_pers_assists")
            {
                SaveChangesControl.Focus();
                this.physical_pers_assistsTableAdapter.Update(this.dBDataSet.Physical_pers_assists);
                DB_Physical_pers_assists.Focus();
            }
            else if(table == "DB_Physical_disableds")
            {
                SaveChangesControl.Focus();
                this.physical_disabledsTableAdapter.Update(this.dBDataSet.Physical_disableds);
                DB_Physical_disableds.Focus();
            }
            else if(table == "DB_Legal_executors")
            {
                SaveChangesControl.Focus();
                this.legal_executorsTableAdapter.Update(this.dBDataSet.Legal_executors);
                DB_Legal_executors.Focus();
            }
            else if(table == "DB_Legal_customers")
            {
                SaveChangesControl.Focus();
                this.legal_customersTableAdapter.Update(this.dBDataSet.Legal_customers);
                DB_Legal_customers.Focus();
            }

            MessageBox.Show("Внесенные в таблицу данные были успешно сохранены!", "Сохранение", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        void CloseAllDB()
        {
            // Дееспособность
            DB_Capacity.Enabled = false;
            DB_Capacity.Visible = false;
            // Договор
            DB_Contracts.Enabled = false;
            DB_Contracts.Visible = false;
            // Судимость
            DB_Criminal_record.Enabled = false;
            DB_Criminal_record.Visible = false;
            // Заказчик услуг - физическое лицо
            DB_Customer_services_person.Enabled = false;
            DB_Customer_services_person.Visible = false;
            // Заказчик услуг - организация
            DB_Customer_services_organisations.Enabled = false;
            DB_Customer_services_organisations.Visible = false;
            // Индивидуальная программа реабилитации инвалида
            DB_Individual_program.Enabled = false;
            DB_Individual_program.Visible = false;
            // Сведения из медицинской карты
            DB_Info_from_medical_cart.Enabled = false;
            DB_Info_from_medical_cart.Visible = false;
            // Юридическое лицо - заказчик
            DB_Legal_customers.Enabled = false;
            DB_Legal_customers.Visible = false;
            // Юридическое лицо - исполнитель
            DB_Legal_executors.Enabled = false;
            DB_Legal_executors.Visible = false;
            // Персональный ассистент
            DB_Personal_assistents.Enabled = false;
            DB_Personal_assistents.Visible = false;
            // Физическое лицо - инвалид
            DB_Physical_disableds.Enabled = false;
            DB_Physical_disableds.Visible = false;
            // Физическое лицо - персональный ассистент
            DB_Physical_pers_assists.Enabled = false;
            DB_Physical_pers_assists.Visible = false;
            // Физическое лицо - организация
            DB_Physical_representatives.Enabled = false;
            DB_Physical_representatives.Visible = false;
        }
        void infoCurrentTableInLabel(string table)
        {

            if (table == "DB_Customer_services_person")
                infoAboutCurrentTable.Text = "Заказчик услуг - Физическое лицо";
            else if (table == "DB_Customer_services_organisations")
                infoAboutCurrentTable.Text = "Заказчик услуг - Организация";
            else if (table == "DB_Capacity")
                infoAboutCurrentTable.Text = "Дееспособность";
            else if (table == "DB_Info_from_medical_cart")
                infoAboutCurrentTable.Text = "Сведения из медицинской карты";
            else if (table == "DB_Individual_program")
                infoAboutCurrentTable.Text = "Индивидуальная программа реабилитации инвалида";
            else if (table == "DB_Criminal_record")
                infoAboutCurrentTable.Text = "Судимость";
            else if (table == "DB_Contracts")
                infoAboutCurrentTable.Text = "Договора";
            else if (table == "DB_Personal_assistents")
                infoAboutCurrentTable.Text = "Персональные ассистенты";
            else if (table == "DB_Physical_representatives")
                infoAboutCurrentTable.Text = "Список заказчиков и исполнителей - Законные представители";
            else if (table == "DB_Physical_pers_assists")
                infoAboutCurrentTable.Text = "Список заказчиков и исполнителей - Персональные ассистенты";
            else if (table == "DB_Physical_disableds")
                infoAboutCurrentTable.Text = "Список заказчиков и исполнителей - Инвалиды";
            else if (table == "DB_Legal_executors")
                infoAboutCurrentTable.Text = "Список заказчиков и исполнителей - Исполнитель";
            else if (table == "DB_Legal_customers")
                infoAboutCurrentTable.Text = "Список заказчиков и исполнителей - Заказчик";

            currentTableLabel.Visible = true;
            infoAboutCurrentTable.Visible = true;
        }
        void DeleteCurrentRow(string table)
        {
            try
            {
                DialogResult dr = MessageBox.Show("Вы точно хотите удалить запись?", "Удаление записи",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr == DialogResult.Yes)
                {
                    if (table == "DB_Customer_services_person")
                    {
                        int CurrentRow = DB_Customer_services_person.SelectedCells[0].RowIndex;
                        DB_Customer_services_person.Rows.RemoveAt(CurrentRow);
                    }
                    else if (table == "DB_Customer_services_organisations")
                    {
                        int CurrentRow = DB_Customer_services_organisations.SelectedCells[0].RowIndex;
                        DB_Customer_services_organisations.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Capacity")
                    {
                        int CurrentRow = DB_Capacity.SelectedCells[0].RowIndex;
                        DB_Capacity.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Info_from_medical_cart")
                    {
                        int CurrentRow = DB_Info_from_medical_cart.SelectedCells[0].RowIndex;
                        DB_Info_from_medical_cart.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Individual_program")
                    {
                        int CurrentRow = DB_Individual_program.SelectedCells[0].RowIndex;
                        DB_Individual_program.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Criminal_record")
                    {
                        int CurrentRow = DB_Criminal_record.SelectedCells[0].RowIndex;
                        DB_Criminal_record.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Contracts")
                    {
                        int CurrentRow = DB_Contracts.SelectedCells[0].RowIndex;
                        DB_Contracts.Rows.RemoveAt(CurrentRow);
                    }
                    if (table == "DB_Physical_representatives")
                    {
                        int CurrentRow = DB_Physical_representatives.SelectedCells[0].RowIndex;
                        DB_Physical_representatives.Rows.RemoveAt(CurrentRow);
                    }
                    if (table == "DB_Physical_pers_assists")
                    {
                        int CurrentRow = DB_Physical_pers_assists.SelectedCells[0].RowIndex;
                        DB_Physical_pers_assists.Rows.RemoveAt(CurrentRow);
                    }
                    if (table == "DB_Physical_disableds")
                    {
                        int CurrentRow = DB_Physical_disableds.SelectedCells[0].RowIndex;
                        DB_Physical_disableds.Rows.RemoveAt(CurrentRow);
                    }
                    if (table == "DB_Legal_executors")
                    {
                        int CurrentRow = DB_Legal_executors.SelectedCells[0].RowIndex;
                        DB_Legal_executors.Rows.RemoveAt(CurrentRow);
                    }
                    if (table == "DB_Legal_customers")
                    {
                        int CurrentRow = DB_Legal_customers.SelectedCells[0].RowIndex;
                        DB_Legal_customers.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Personal_assistents")
                    {
                        int CurrentRow = DB_Personal_assistents.SelectedCells[0].RowIndex;
                        DB_Personal_assistents.Rows.RemoveAt(CurrentRow);
                    }
                    else if (table == "DB_Physical_representatives")
                    {
                        int CurrentRow = DB_Physical_representatives.SelectedCells[0].RowIndex;
                        DB_Physical_representatives.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Physical_pers_assists")
                    {
                        int CurrentRow = DB_Physical_pers_assists.SelectedCells[0].RowIndex;
                        DB_Physical_pers_assists.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Physical_disableds")
                    {
                        int CurrentRow = DB_Physical_disableds.SelectedCells[0].RowIndex;
                        DB_Physical_disableds.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Legal_executors")
                    {
                        int CurrentRow = DB_Legal_executors.SelectedCells[0].RowIndex;
                        DB_Legal_executors.Rows.RemoveAt(CurrentRow);
                    }
                    else if(table == "DB_Legal_customers")
                    {
                        int CurrentRow = DB_Legal_customers.SelectedCells[0].RowIndex;
                        DB_Legal_customers.Rows.RemoveAt(CurrentRow);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Возможно Вы пытаетесь удалить несуществующую строку...\nЕсли это не так, обратитесь к разработчику.", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void CleanerForm3ComboBox()
        {
            Form3 f3 = new Form3();
            f3.request_comboBox.Items.Clear();
            f3.request_comboBoxFor5.Items.Clear();
        }
        void DataErrorFixer(DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception != null && e.Context == DataGridViewDataErrorContexts.Commit)
            {
                DialogResult dr = MessageBox.Show("Скорее всего Вы не заполнили ключевое поле и попытались сохранить данные!\nСкопировать код ошибки в буфер обмена?", "Ошибка", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dr == DialogResult.Yes)
                    Clipboard.SetText(e.Exception.Message);
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Данная программа разработана для учета персональных ассистентов и лиц с инвалидностью", 
                "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void разработчикToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Владимир Ляшко, 2020 год", "Разработчик", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
        private void физическоеЛицоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Customer_services_person.Enabled = true;
            DB_Customer_services_person.Visible = true;
            CurrentTable = "DB_Customer_services_person";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void организацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Customer_services_organisations.Enabled = true;
            DB_Customer_services_organisations.Visible = true;
            CurrentTable = "DB_Customer_services_organisations";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void дееспособностьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Capacity.Enabled = true;
            DB_Capacity.Visible = true;
            CurrentTable = "DB_Capacity";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void сведенияИзМедицинскойКартыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Info_from_medical_cart.Enabled = true;
            DB_Info_from_medical_cart.Visible = true;
            CurrentTable = "DB_Info_from_medical_cart";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void индивидуальнойПрограммаРеабилитацииИнвалидаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Individual_program.Enabled = true;
            DB_Individual_program.Visible = true;
            CurrentTable = "DB_Individual_program";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void судимостьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Criminal_record.Enabled = true;
            DB_Criminal_record.Visible = true;
            CurrentTable = "DB_Criminal_record";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void договорToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Contracts.Enabled = true;
            DB_Contracts.Visible = true;
            CurrentTable = "DB_Contracts";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void персональныйАссистентToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Personal_assistents.Enabled = true;
            DB_Personal_assistents.Visible = true;
            CurrentTable = "DB_Personal_assistents";
            infoCurrentTableInLabel(CurrentTable);
        }
        private void сохранениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateAllDB(CurrentTable);
        }
        private void удалениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteCurrentRow(CurrentTable);
            UpdateAllDB(CurrentTable);
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void запрос1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request_comboBox.Size = new Size(335, 24);

            f3.request1_button.Enabled = true;
            f3.request1_button.Visible = true;

            f3.request_comboBox.Items.Add("Номер паспорта лица с инвалидностью");
            f3.request_comboBox.Items.Add("Номер договора");
        }
        private void запрос2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request_comboBox.Size = new Size(335, 24);

            f3.request2_button.Enabled = true;
            f3.request2_button.Visible = true;
            f3.request_comboBox.Items.Add("Номер паспорта персонального ассистента");
            f3.request_comboBox.Items.Add("ФИО персонального ассистента");
            f3.request_comboBox.Items.Add("Номер договора");
        }

        private void запрос3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request_comboBox.Size = new Size(335, 24);

            f3.request3_button.Enabled = true;
            f3.request3_button.Visible = true;
            f3.request_comboBox.Items.Add("Номер паспорта лица с инвалидностью");
            f3.request_comboBox.Items.Add("ФИО лица с инвалидностью");
            f3.request_comboBox.Items.Add("Номер договора");
        }

        private void запрос4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request_comboBox.Size = new Size(335, 24);

            f3.request4_button.Enabled = true;
            f3.request4_button.Visible = true;
            f3.request_comboBox.Items.Add("Номер договора (юридический)");
            f3.request_comboBox.Items.Add("Номер договора (физический)");
            f3.request_comboBox.Items.Add("Наименование организации");
            f3.request_comboBox.Items.Add("ФИО");
        }

        private void запрос5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request5_button.Enabled = true;
            f3.request5_button.Visible = true;

            f3.request_comboBoxFor5.Enabled = true;
            f3.request_comboBoxFor5.Visible = true;

            f3.request_comboBox.Items.Add("Номер договора");
            f3.request_comboBox.Items.Add("ФИО");

            f3.request_comboBox.Size = new Size(335, 24);

            f3.request_comboBoxFor5.Items.Add("Инвалиды");
            f3.request_comboBoxFor5.Items.Add("Персональные ассистенты");
        }

        private void запрос6ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            CleanerForm3ComboBox();
            f3.Show();
            f3.UnCheckAllForm3Buttons();

            f3.request6_button.Enabled = true;
            f3.request6_button.Visible = true;

            f3.request_textBox_for6.Enabled = true;
            f3.request_textBox_for6.Visible = true;

            f3.request_comboBox.Size = new Size(420, 24);

            f3.request_comboBox.Items.Add("Ввод промежутка времени в формате XX.XX.XXXX");
        }

        private void инвалидыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();
            DB_Physical_disableds.Enabled = true;
            DB_Physical_disableds.Visible = true;
            CurrentTable = "DB_Physical_disableds";
            infoCurrentTableInLabel(CurrentTable);
        }

        private void персональныеАссистентыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllDB();
            DB_Physical_pers_assists.Enabled = true;
            DB_Physical_pers_assists.Visible = true;
            CurrentTable = "DB_Physical_pers_assists";
            infoCurrentTableInLabel(CurrentTable);
        }

        private void законныеПредставителиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            CloseAllDB();
            DB_Physical_representatives.Enabled = true;
            DB_Physical_representatives.Visible = true;
            CurrentTable = "DB_Physical_representatives";
            infoCurrentTableInLabel(CurrentTable);
        }

        private void заказчикToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Legal_customers.Enabled = true;
            DB_Legal_customers.Visible = true;
            CurrentTable = "DB_Legal_customers";
            infoCurrentTableInLabel(CurrentTable);
        }

        private void исполнительToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            CloseAllDB();

            DB_Legal_executors.Enabled = true;
            DB_Legal_executors.Visible = true;
            CurrentTable = "DB_Legal_executors";
            infoCurrentTableInLabel(CurrentTable);
        }

        private void DB_Physical_representatives_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Physical_pers_assists_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Physical_disableds_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Personal_assistents_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Legal_executors_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Legal_customers_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Info_from_medical_cart_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Individual_program_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Customer_services_person_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Customer_services_organisations_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Criminal_record_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Contracts_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }

        private void DB_Capacity_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            DataErrorFixer(e);
        }
    }
}
