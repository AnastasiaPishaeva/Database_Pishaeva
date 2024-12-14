using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Лабораторная_2_С_
{
    public partial class ButtonSaveXcel : Form
    {
        private DataBaseManager manager;
        public ButtonSaveXcel()
        {
            InitializeComponent();
            manager = new DataBaseManager();
            try
            {
                manager.Load();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}");
            }
        }

        private void LoadDataToGrid()
        {
            dataGridView1.Rows.Clear(); // Очищаем таблицу

            if (!File.Exists("database.bin"))
            {
                return;
            }

            try
            {
                using (var reader = new BinaryReader(File.Open("database.bin", FileMode.Open)))
                {
                    while (reader.BaseStream.Position < reader.BaseStream.Length)
                    {
                        try
                        {
                            // Попытка прочитать поля
                            int id = reader.ReadInt32();
                            string name = reader.ReadString();
                            string author = reader.ReadString();
                            double price = reader.ReadDouble();
                            string dateString = reader.ReadString();

                            if (!DateTime.TryParse(dateString, out DateTime date))
                            {
                                MessageBox.Show("Некорректный формат даты.");
                                break;
                            }

                            if (id != -1) // Пропускаем удаленные записи
                            {
                                dataGridView1.Rows.Add(id, name, author, price, date.ToString("yyyy-MM-dd"));
                            }
                        }
                        catch (EndOfStreamException)
                        {
                            MessageBox.Show("Достигнут конец файла при чтении данных.");
                            break;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ошибка при чтении записи: {ex.Message}");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла: {ex.Message}");
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear(); // Очищаем предыдущие колонки
            dataGridView1.Columns.Add("ID", "ID");
            dataGridView1.Columns.Add("Name", "Название");
            dataGridView1.Columns.Add("Author", "Автор");
            dataGridView1.Columns.Add("Price", "Цена");
            dataGridView1.Columns.Add("Date", "Дата");

            comboBox.Items.AddRange(new string[] { "ID", "Название", "Автор", "Цена", "Дата издания" });
            comboBox.SelectedIndex = 0;

            LoadDataToGrid();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string selectedField = comboBox.SelectedItem?.ToString(); // смотрим по какому параметру ищем
            List<BookRecord?> results = new List<BookRecord?>();
            string searchValueString;
            double searchValueDouble;
            int searchValueInt;
            DateTime searchValueDateTime;
            switch (selectedField)
            {
                case "Название":
                    searchValueString = textBoxName.Text.Trim();
                    if (!string.IsNullOrEmpty(searchValueString))
                    {
                        results = manager.SearchRecordsByName(searchValueString);
                    }
                    else
                    {
                        MessageBox.Show("Выбранное поле пусто.");
                    }
                    break;
                case "Автор":
                    searchValueString = textBoxAuthor.Text.Trim();
                    if (!string.IsNullOrEmpty(searchValueString))
                    {
                        results = manager.SearchRecordsByAuthor(searchValueString);
                    }
                    else
                    {
                        MessageBox.Show("Выбранное поле пусто.");
                    }
                    break;
                case "ID":
                    if (!int.TryParse(textBoxID.Text, out searchValueInt))
                    {
                        MessageBox.Show("Введите корректный числовой ID.");
                        return;
                    }
                    results = manager.SearchRecordById(searchValueInt);
                    break;
                case "Цена":
                    if (!double.TryParse(textBoxPrice.Text, out searchValueDouble))
                    {
                        MessageBox.Show("Введите корректную цену.");
                        return;
                    }
                    results = manager.SearchRecordsByPrice(searchValueDouble);
                    break;
                case "Дата издания":
                    if (!DateTime.TryParse(dateTimePicker.Text, out searchValueDateTime))
                    {
                        MessageBox.Show("Ошибка при чтении даты.");
                        return;
                    }
                    results = manager.SearchRecordsByDate(searchValueDateTime.Date.ToString("yyyy-MM-dd"));
                    break;

                default:
                    MessageBox.Show("Выбранное поле не поддерживается для поиска.");
                    return;
            }
            if (results.Count > 0)
            {
                string res = $"Найдена запись:  \n";
                int i = 1;
                foreach (BookRecord record in results)
                {
                    res += $"{i}) Название: {record.Name}, автор: {record.Author}," +
                        $" цена: {record.Price}, дата издания: {record.Date.ToString("yyyy-MM-dd")}\n";
                    i++;
                }
                MessageBox.Show(res);
            }
            else
            {
                MessageBox.Show("Запись не найдена!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // Считываем данные из текстовых полей
                if (!int.TryParse(textBoxID.Text, out int id))
                {
                    MessageBox.Show("Введите корректный ID.");
                    return;
                }

                string name = textBoxName.Text.Trim();
                string author = textBoxAuthor.Text.Trim();

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(author))
                {
                    MessageBox.Show("Название и автор книги не могут быть пустыми.");
                    return;
                }

                if (!double.TryParse(textBoxPrice.Text, out double price) || price <= 0)
                {
                    MessageBox.Show("Введите корректную цену.");
                    return;
                }

                DateTime date = dateTimePicker.Value;

                // Проверяем, существует ли запись с данным ID
                var existingRecord = manager.SearchRecordById(id);

                if (existingRecord.Count > 0)
                {
                    MessageBox.Show("Запись с таким ID уже существует.");
                    return;
                }

                // Создаем новую запись
                var newRecord = new BookRecord(id, name, author, price, date);

                // Добавляем запись в файл
                manager.WriteRecord(newRecord, "database.bin");
                LoadDataToGrid(); // Обновляем таблицу
            }
            catch (FormatException)
            {
                MessageBox.Show("Некорректный формат данных. Проверьте введённые значения.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void EnterID_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxID_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string selectedField = comboBox.SelectedItem?.ToString();
            string searchValueString;
            double searchValueDouble;
            int searchValueInt;
            DateTime searchValueDate;
            switch (selectedField)
            {
                case "Название":
                    searchValueString = textBoxName.Text.Trim();
                    if (!string.IsNullOrEmpty(searchValueString))
                    {
                        manager.DeleteRecordbyName(searchValueString);
                    }
                    else
                    {
                        MessageBox.Show("Выбранное поле пусто.");
                    }
                    break;

                case "Автор":
                    searchValueString = textBoxAuthor.Text.Trim();
                    if (!string.IsNullOrEmpty(searchValueString))
                    {
                        manager.DeleteRecordbyAuthor(searchValueString);
                    }
                    else
                    {
                        MessageBox.Show("Выбранное поле пусто.");
                    }
                    break;

                case "ID":
                    if (!int.TryParse(textBoxID.Text, out searchValueInt))
                    {
                        MessageBox.Show("Введите корректный числовой ID.");
                        return;
                    }
                    manager.DeleteRecordbyId(searchValueInt);
                    break;

                case "Цена":
                    if (!double.TryParse(textBoxPrice.Text, out searchValueDouble))
                    {
                        MessageBox.Show("Введите корректную цену.");
                        return;
                    }
                    manager.DeleteRecordbyPrice(searchValueDouble);
                    break;
                case "Дата издания":
                    if (!DateTime.TryParse(dateTimePicker.Text, out searchValueDate))
                    {
                        MessageBox.Show("Некорректная дата.");
                        return;
                    }
                    manager.DeleteRecordbyDate(searchValueDate.Date);
                    break;

                default:
                    MessageBox.Show("Выбранное поле не поддерживается для удаления.");
                    return;
            }
            LoadDataToGrid(); // Обновляем таблицу
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear(); // Очищаем строки
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("ID", "ID");
            dataGridView1.Columns.Add("Name", "Название");
            dataGridView1.Columns.Add("Author", "Автор");
            dataGridView1.Columns.Add("Price", "Цена");
            dataGridView1.Columns.Add("Date", "Дата");
            manager.Clear();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void EnterPrice_Click(object sender, EventArgs e)
        {

        }

        private void buttonEdit_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("В поля для записи ввдите ID (обязательно) и необходимые поля записи, которую хотите отредактировать.");
            try
            {
                if (!int.TryParse(textBoxID.Text, out int id))
                {
                    MessageBox.Show("Введите корректный ID.");
                    return;
                }
                string? newName = string.IsNullOrWhiteSpace(textBoxName.Text) ? null : textBoxName.Text; // Поле для нового имени
                string? newAuthor = string.IsNullOrWhiteSpace(textBoxAuthor.Text) ? null : textBoxAuthor.Text; // Поле для нового автора
                double? newPrice = string.IsNullOrWhiteSpace(textBoxPrice.Text) ? null : double.Parse(textBoxPrice.Text); // Поле для новой цены
                DateTime? newDate = string.IsNullOrWhiteSpace(dateTimePicker.Text) ? null : DateTime.Parse(dateTimePicker.Text); // Поле для новой даты

                bool result = manager.EditRecordById(id, newName, newAuthor, newPrice, newDate);

                if (result)
                {
                    LoadDataToGrid(); // Обновляем таблицу
                    MessageBox.Show("Запись успешно отредактирована.");
                }
                else
                {
                    MessageBox.Show("Не удалось отредактировать запись.");
                }
            }
            catch (FormatException ex)
            {
                MessageBox.Show($"Ошибка ввода данных: {ex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}");
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            manager.SaveToXlsx();
        }
    }
}
