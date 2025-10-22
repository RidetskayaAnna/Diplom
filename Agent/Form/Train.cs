using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Agent.Form
{
    public partial class Train : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        public Train()
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
        }
        int id = 0;
        public void train()
        {
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select idtrain, name as [Название транспорта],type as [Тип транспорта],nomer as [Номер транспорта], year as [Год постройки], certificate as [Сертификат] from train", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = true;
            label6.Text = "Добавить транспорт";
            button11.Text = "Добавить";
            button11.Width = 198;
            button11.Left = 247;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (31).png");
            dataGridView1.Enabled = true;
            clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = true;
            label6.Text = "Редактировать транспорт";
            button11.Width = 272;
            button11.Left = 210;
            button11.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (29).png");
            button11.Text = "Редактировать";
            textBox2.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            textBox3.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox2.Visible = true;
            comboBox2.SelectedItem = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            dataGridView1.Enabled = false;
        }
        public void clear()
        {
            textBox4.Text = "";
            textBox3.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.Text = "";
            comboBox2.Visible = false;
            textBox2.Text = "";
        }
        private void Train_Load(object sender, EventArgs e)
        {
            dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            train();
            panel2.Visible = false;
            panel3.Visible = false;
        }
        int k = 0; int j = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                k = 0;
                j = 0;
                if (label6.Text == "Добавить транспорт")
                {
                    if (textBox2.Text != "" || comboBox2.SelectedIndex != -1 || textBox3.Text != "" || textBox3.Text != "")
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (comboBox1.Text.ToLower() + comboBox2.Text.ToLower() + textBox2.Text.ToLower() + textBox3.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower() + dataGridView1[3, i].Value.ToString().ToLower() + dataGridView1[5, i].Value.ToString().ToLower())
                            {
                                k++;
                            }
                        }
                        if (k == 0)
                        {
                            k = 0;
                            if (Convert.ToInt32(textBox4.Text) <= 2023 && Convert.ToInt32(textBox4.Text) >= 1800)
                            {
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (textBox2.Text.ToLower() == dataGridView1[3, i].Value.ToString().ToLower())
                                    {
                                        k++;
                                    }
                                }
                                if (k == 0)
                                {
                                    k = 0;
                                    sqlConnection.Open();
                                    SqlCommand command = new SqlCommand($@"INSERT INTO [train](name,type,nomer,year,certificate) VALUES (@n,@t,@no,@y,@c);", sqlConnection);
                                    command.Parameters.AddWithValue("@n", (comboBox1.Text));
                                    command.Parameters.AddWithValue("@t", (comboBox2.Text));
                                    command.Parameters.AddWithValue("@no", (textBox2.Text));
                                    command.Parameters.AddWithValue("@y", (textBox4.Text));
                                    command.Parameters.AddWithValue("@c", (textBox3.Text));
                                    command.ExecuteNonQuery();
                                    sqlConnection.Close();
                                    train();
                                    clear();
                                    panel2.Visible = false;
                                }
                                else
                                {
                                    MessageBox.Show("Номер транспорта неуникален!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Введите год постройки корректно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Такой транспорт уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (textBox2.Text != "" || comboBox2.SelectedIndex != -1 || textBox3.Text != "" || textBox3.Text != "")
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            if (comboBox1.Text.ToLower() + comboBox2.Text.ToLower() + textBox2.Text.ToLower() + textBox3.Text.ToLower() == dataGridView1[1, i].Value.ToString().ToLower() + dataGridView1[2, i].Value.ToString().ToLower() + dataGridView1[3, i].Value.ToString().ToLower() + dataGridView1[5, i].Value.ToString().ToLower())
                            {
                                k++; j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                            }
                        }
                        if (k == 0 || j == id)
                        {
                            if (Convert.ToInt32(textBox4.Text) <= 2023 && Convert.ToInt32(textBox4.Text) >= 1800)
                            {
                                k = 0; j = 0;
                                for (int i = 0; i < dataGridView1.RowCount; i++)
                                {
                                    if (textBox2.Text.ToLower() == dataGridView1[3, i].Value.ToString().ToLower())
                                    {
                                        k++; j = Convert.ToInt32(dataGridView1[0, i].Value.ToString());
                                    }
                                }
                                if (k == 0 || j == id)
                                {
                                    k = 0; j = 0;
                                    sqlConnection.Open();
                                    SqlCommand command = new SqlCommand($@"UPDATE train SET name=@n,type=@t, nomer=@no,
                            year=@y, certificate=@c WHERE idtrain=@id", sqlConnection);
                                    command.Parameters.AddWithValue("@n", (comboBox1.Text));
                                    command.Parameters.AddWithValue("@t", (comboBox2.Text));
                                    command.Parameters.AddWithValue("@no", (textBox2.Text));
                                    command.Parameters.AddWithValue("@y", (textBox4.Text));
                                    command.Parameters.AddWithValue("@c", (textBox3.Text));
                                    command.Parameters.AddWithValue("@id", (id));
                                    command.ExecuteNonQuery();
                                    sqlConnection.Close();
                                    train();

                                    panel2.Visible = false;
                                    dataGridView1.Enabled = true;
                                }
                                else
                                {
                                    MessageBox.Show("Номер транспорта неуникален!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Введите год постройки корректно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Такой транспорт уже есть!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch { }
        }
        public void one()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                count = 0;
                if (dataGridView1[1, i].Value.ToString() == comboBox1.Text)
                {
                    for (int j = 0; j < comboBox2.Items.Count; j++)
                    {
                        if (dataGridView1[2, i].Value.ToString() != comboBox2.Items[j].ToString())
                        {
                            count++;

                        }
                    }
                    if (count == 6)
                    {
                        comboBox2.Items.Add(dataGridView1[2, i].Value.ToString());
                    }
                }
            }
        }
        int count = 0;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Электровозы");
                comboBox2.Items.Add("Тепловозы");
                comboBox2.Items.Add("Паровозы");
                comboBox2.Items.Add("Газотрубовозы");
                comboBox2.Items.Add("Мотовозы");
                comboBox2.Items.Add("Автомотрисы");
                one();
                comboBox2.SelectedIndex = -1;

            }
            else if (comboBox1.SelectedIndex == 1)
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Электропоезда");
                comboBox2.Items.Add("Дизельные поезда");
                comboBox2.Items.Add("Турбопоезда");
                comboBox2.Items.Add("Аккумуляторные поезда");
                comboBox2.Items.Add("Специальные самоходные подвижные составы");
                comboBox2.Items.Add("Дрезины");
                one();
                comboBox2.SelectedIndex = -1;
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Вагоны крытые");
                comboBox2.Items.Add("Полувагоны");
                comboBox2.Items.Add("Платформы");
                comboBox2.Items.Add("Окатышевозы");
                comboBox2.Items.Add("Вагоны для перевозки автомабилей");
                comboBox2.Items.Add("Думпкары");
                comboBox2.Items.Add("Цистерны");
                comboBox2.Items.Add("Вагоны-зерновозы");
                comboBox2.Items.Add("Вагоны-минераловозы");
                comboBox2.Items.Add("Фитинговые платформы");
                comboBox2.Items.Add("Садовозы");
                comboBox2.Items.Add("Вагоны-цементовозы");
                comboBox2.Items.Add("Контейнеровозы");
                comboBox2.Items.Add("Транспортеры");
                comboBox2.Items.Add("Рефрижераторные вагоны");
                comboBox2.Items.Add("Вагоны-термосы");
                one();
                comboBox2.SelectedIndex = -1;
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                comboBox2.Visible = true;
                comboBox2.Text = "";
                comboBox2.Items.Clear();
                comboBox2.Items.Add("Пассажирские");
                comboBox2.Items.Add("Багажные");
                comboBox2.Items.Add("Почтовые");
                comboBox2.Items.Add("Багажно-почтовые");
                comboBox2.Items.Add("Служебные");
                comboBox2.Items.Add("Вагоны-рестораны (кафе)");
                one();
                comboBox2.SelectedIndex = -1;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            dataGridView1.Enabled = true;
            clear();
            panel2.Visible = false;
            if (id != 0)
            {
                try
                {
                    //Фрагмент кода кдаления данных о транспорте из БД
                    if (MessageBox.Show($@"Вы уверены что хотите удалить транспорт 
                    {dataGridView1.CurrentRow.Cells[1].Value.ToString() + ' ' 
                    + dataGridView1.CurrentRow.Cells[2].Value.ToString()}?",
                        "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        sqlConnection.Open();
                        string query = $@"DELETE FROM [train] WHERE [idtrain] = 
                             {dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                        SqlCommand command = new SqlCommand(query, sqlConnection);
                        command.ExecuteNonQuery();
                        sqlConnection.Close();
                        train();
                    }
                }
                catch { MessageBox.Show("Удаление невозможно, данные все еще нужны!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else { MessageBox.Show("Выберите строку для удаления!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (textBox2.Text.Length == 0)
            {
                char c = e.KeyChar;
                e.Handled = !(c == '3' || c == '9' || c == 8);

            }
            else
            {
                if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                    e.Handled = true;
            }
        }

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            e.Handled = !((c >= 'а' && c <= 'я') || (c >= 'А' && c <= 'Я') || (c == 'ё' || c == '-' || c == 8 || c == 32));
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text.Length == 1)
            {
                comboBox2.Text = comboBox2.Text.ToUpper();
                comboBox2.SelectionStart = 2;
            }
          
        }
        public void two()
        {
            try
            {

                if (comboBox2.Text.ToLower().Contains("крыты"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "2");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.Remove(0, 1).Insert(0, "2");
                    }
                }
                else if (comboBox2.Text.ToLower().Contains("платформ"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "4");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.Remove(0, 1).Insert(0, "4");
                    }
                }
                else if (comboBox2.Text.ToLower().Contains("полувагон"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "6");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.Remove(0, 1).Insert(0, "6");
                    }
                }
                else if (comboBox2.Text.ToLower().Contains("цистерн"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "7");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.Remove(0, 1).Insert(0, "7");
                    }
                }
                else if (comboBox2.Text.ToLower().Contains("изотерми"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "8");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.Remove(0, 1).Insert(0, "8");
                    }
                }
                else if (comboBox2.Text.ToLower().Contains("собствен"))
                {
                    if (textBox2.Text.Length == 0)
                    {
                        textBox2.Text = textBox2.Text.Insert(0, "5");
                    }
                    else
                    {
                        textBox2.Text = textBox2.Text.ToLower().Remove(0, 1).Insert(0, "5");
                    }
                }

            }
            catch { }
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            two();

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }
        int visible = 0;
        int y = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
            clear();
            panel2.Visible = false;
            visible = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    visible++;
                }
            }
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Columns.NumberFormat = "General";
            ExcelWorkSheet.StandardWidth = 30;
            ExcelWorkSheet.Columns.ColumnWidth = 20;
            ExcelApp.Rows[1].Columns[3] = "Транспорт";
            ExcelApp.Rows[visible + 3].Columns[3] = "Ридецкая Анна Михайловна";
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i] = dataGridView1.Columns[i].HeaderText;

            }
            
            for (int j = 1; j < dataGridView1.ColumnCount; j++)
            {y = 0;
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                      
                            ExcelApp.Cells[y + 3, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        y++;
                    }
                }
            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:E{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:E{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
         
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            two();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            clear();
            dataGridView1.Enabled = true;
            if (panel3.Visible == true)
            {
                panel3.Visible = false;
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();
                comboBox5.Text = "";
                train();
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (48).png");
            }
            else
            {
                checkBox1.Checked = false;
                checkBox3.Checked = false;
                comboBox4.Items.Clear();
                comboBox5.Items.Clear();

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (comboBox4.Items.Contains(dataGridView1[1, i].Value.ToString()))
                    {
                    }
                    else
                    {

                        comboBox4.Items.Add(dataGridView1[1, i].Value.ToString());
                    }
                }
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (comboBox5.Items.Contains(dataGridView1[2, i].Value.ToString()))
                    {
                    }
                    else
                    {

                        comboBox5.Items.Add(dataGridView1[2, i].Value.ToString());
                    }
                }
                comboBox4.SelectedIndex = -1;
                comboBox5.SelectedIndex = -1;
                panel3.Visible = true;
                button7.Image = new Bitmap(@"/Diplom/proga/Agent/Agent/Resources/pngwing.com (47).png");
            }
            panel2.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox5.Text == dataGridView1[2, i].Value.ToString() && comboBox4.Text == dataGridView1[1, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else

            ////////////////
            ///// 2 по 1 ////

            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox4.Text == dataGridView1[1, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else

                if (checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if (comboBox5.Text == dataGridView1[2, i].Value.ToString())
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else if (checkBox1.Checked == false && checkBox3.Checked == false)
            {
                train();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()) && textBox1.Text != "")
                        {
                            dataGridView1.Rows[i].Selected = true;
                            dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                            dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                            break;
                        }
                        else
                        {
                            dataGridView1.Rows[i].Selected = false;
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;


                        }
                    }
                }
            }
        }
    }
}
