using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Spire.Xls;


namespace Лабораторная_2_С_
{
    internal class DataBaseManager
    {
        private string filePath = "database.bin";
        private Dictionary<int, long> indexTable = new Dictionary<int, long>();
        private Dictionary<string, List<long>> nameTable = new Dictionary<string, List<long>>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, List<long>> authorTable = new Dictionary<string, List<long>>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<double, List<long>> priceTable = new Dictionary<double, List<long>>();
        private Dictionary<DateTime, List<long>> dateTable = new Dictionary<DateTime, List<long>>();

        private void SaveIndex()
        {
            try
            {
                using (var writer = new BinaryWriter(File.Open("index.dat", FileMode.Create)))
                {
                    foreach (var kvp in indexTable)
                    {
                        writer.Write(kvp.Key);
                        writer.Write(kvp.Value);
                    }
                }

            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при сохранении индекса: {ex.Message}");
            }
        }

        private void SaveName()
        {
            try
            {
                using (var writer = new BinaryWriter(File.Open("index_name.dat", FileMode.Create)))
                {
                    foreach (var kvp in nameTable)
                    {
                        writer.Write(kvp.Key);
                        writer.Write(kvp.Value.Count);
                        foreach (var position in kvp.Value)
                        {
                            writer.Write(position);
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при сохранении индекса Name: {ex.Message}");
            }
        }

        private void SaveAuthor()
        {
            try
            {
                using (var writer = new BinaryWriter(File.Open("index_author.dat", FileMode.Create)))
                {
                    foreach (var kvp in authorTable)
                    {
                        writer.Write(kvp.Key);
                        writer.Write(kvp.Value.Count);
                        foreach (var position in kvp.Value)
                        {
                            writer.Write(position);
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при сохранении индекса Author: {ex.Message}");
            }
        }

        private void SavePrice()
        {
            try
            {
                using (var writer = new BinaryWriter(File.Open("index_price.dat", FileMode.Create)))
                {
                    foreach (var kvp in priceTable)
                    {
                        writer.Write(kvp.Key);
                        writer.Write(kvp.Value.Count);
                        foreach (var position in kvp.Value)
                        {
                            writer.Write(position);
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при сохранении индекса Name: {ex.Message}");
            }
        }

        private void SaveDate()
        {
            try
            {
                using (var writer = new BinaryWriter(File.Open("index_date.dat", FileMode.Create)))
                {
                    foreach (var kvp in dateTable)
                    {
                        if (kvp.Key == default(DateTime))
                        {
                            Debug.WriteLine("Попытка записать некорректную дату.");
                            continue; // Пропускаем некорректные записи
                        }

                        writer.Write(kvp.Key.Date.ToString("yyyy-MM-dd")); // Явный формат для DateTime
                        writer.Write(kvp.Value.Count);
                        foreach (var position in kvp.Value)
                        {
                            writer.Write(position);
                        }
                    }
                }

            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при сохранении индекса: {ex.Message}");
            }
        }

        private void SaveAll()
        {
            SaveAuthor();
            SavePrice();
            SaveName();
            SaveIndex();
            SaveDate();
        }

        public void WriteRecord(BookRecord record, string filePath) // запись в бинарный файл
        {
            if (!File.Exists("index.dat"))
            {
                using (var writer = new BinaryWriter(File.Open(filePath, FileMode.OpenOrCreate)))
                {
                    // Позиция, на которую будем записывать
                    long position = writer.BaseStream.Position;

                    writer.Write(record.ID);
                    writer.Write(record.Name);
                    writer.Write(record.Author);
                    writer.Write(record.Price);
                    writer.Write(record.Date.Date.ToString());

                    // Обновляем индексы
                    indexTable[record.ID] = position;

                    if (!nameTable.ContainsKey(record.Name))
                    {
                        nameTable[record.Name] = new List<long>();
                    }
                    nameTable[record.Name].Add(position);

                    if (!authorTable.ContainsKey(record.Author))
                    {
                        authorTable[record.Author] = new List<long>();
                    }
                    authorTable[record.Author].Add(position);

                    if (!priceTable.ContainsKey(record.Price))
                    {
                        priceTable[record.Price] = new List<long>();
                    }
                    priceTable[record.Price].Add(position);
                    if (!dateTable.ContainsKey(record.Date.Date))
                    {
                        dateTable[record.Date.Date] = new List<long>();
                    }
                    dateTable[record.Date.Date].Add(position);
                    // Сохранение индексов
                    SaveAll();
                }
            }
            else
            {
                using (var writer = new BinaryWriter(File.Open(filePath, FileMode.Append)))
                {
                    // Позиция, на которую будем записывать
                    long position = writer.BaseStream.Position;

                    writer.Write(record.ID);
                    writer.Write(record.Name);
                    writer.Write(record.Author);
                    writer.Write(record.Price);
                    writer.Write(record.Date.Date.ToString());

                    // Обновляем индексы
                    indexTable[record.ID] = position;

                    if (!nameTable.ContainsKey(record.Name))
                    {
                        nameTable[record.Name] = new List<long>();
                    }
                    nameTable[record.Name].Add(position);

                    if (!authorTable.ContainsKey(record.Author))
                    {
                        authorTable[record.Author] = new List<long>();
                    }
                    authorTable[record.Author].Add(position);

                    if (!priceTable.ContainsKey(record.Price))
                    {
                        priceTable[record.Price] = new List<long>();
                    }
                    priceTable[record.Price].Add(position);
                    if (!dateTable.ContainsKey(record.Date.Date))
                    {
                        dateTable[record.Date.Date] = new List<long>();
                    }
                    dateTable[record.Date.Date].Add(position);
                    // Сохранение индексов
                    SaveAll();
                }
            }
        }
        public void LoadName()
        {
            if (!File.Exists("index_name.dat"))
                return;

            nameTable.Clear();
            using (var reader = new BinaryReader(File.Open("index_name.dat", FileMode.Open)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    string name = reader.ReadString();
                    int count = reader.ReadInt32();
                    var positions = new List<long>();
                    for (int i = 0; i < count; i++)
                    {
                        positions.Add(reader.ReadInt64());
                    }
                    nameTable[name] = positions;

                }
            }
        }

        public void LoadIndex()
        {
            if (!File.Exists("index.dat"))
                return;

            indexTable.Clear();
            using (var reader = new BinaryReader(File.Open("index.dat", FileMode.Open)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    int id = reader.ReadInt32();
                    long position = reader.ReadInt64();
                    indexTable[id] = position;
                }
            }
        }

        public void LoadAuthor()
        {
            if (!File.Exists("index_author.dat"))
                return;

            nameTable.Clear();
            using (var reader = new BinaryReader(File.Open("index_author.dat", FileMode.Open)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    string author = reader.ReadString();
                    int count = reader.ReadInt32();
                    var positions = new List<long>();
                    for (int i = 0; i < count; i++)
                    {
                        positions.Add(reader.ReadInt64());
                    }
                    authorTable[author] = positions;

                }
            }
        }

        public void LoadPrice()
        {
            if (!File.Exists("index_price.dat"))
                return;

            nameTable.Clear();
            using (var reader = new BinaryReader(File.Open("index_price.dat", FileMode.Open)))
            {
                while (reader.BaseStream.Position < reader.BaseStream.Length)
                {
                    double price = reader.ReadInt64();
                    int count = reader.ReadInt32();
                    var positions = new List<long>();
                    for (int i = 0; i < count; i++)
                    {
                        positions.Add(reader.ReadInt64());
                    }
                    priceTable[price] = positions;

                }
            }
        }
        public void LoadDate()
        {
            if (!File.Exists("index_date.dat"))
                return;

            dateTable.Clear();
            try
            {
                using (var reader = new BinaryReader(File.Open("index_date.dat", FileMode.Open)))
                {
                    while (reader.BaseStream.Position < reader.BaseStream.Length)
                    {
                        string datastr = reader.ReadString();
                        if (!DateTime.TryParse(datastr, out DateTime date))
                        {
                            MessageBox.Show($"Некорректный формат даты: {datastr}");
                            continue; // Пропускаем некорректные записи
                        }

                        int count = reader.ReadInt32();
                        var positions = new List<long>();
                        for (int i = 0; i < count; i++)
                        {
                            positions.Add(reader.ReadInt64());
                        }

                        dateTable[date.Date] = positions;
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при загрузке индекса: {ex.Message}");
            }
        }

        public void Load()
        {
            LoadAuthor();
            LoadPrice();
            LoadIndex();
            LoadName();
            LoadDate();
        }

        public void Clear()
        {
            indexTable.Clear();
            nameTable.Clear();
            authorTable.Clear();
            priceTable.Clear();
            dateTable.Clear();
            if (File.Exists("index.dat"))
            {
                try
                {
                    File.Delete("index.dat");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла index.dat: {ex.Message}");
                }
            }
            if (File.Exists("index_price.dat"))
            {
                try
                {
                    File.Delete("index_price.dat");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла index_price.dat: {ex.Message}");
                }
            }
            if (File.Exists("index_name.dat"))
            {
                try
                {
                    File.Delete("index_name.dat");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла index_name.dat: {ex.Message}");
                }
            }
            if (File.Exists("index_author.dat"))
            {
                try
                {
                    File.Delete("index_author.dat");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла index_author.dat: {ex.Message}");
                }
            }
            if (File.Exists("index_date.dat"))
            {
                try
                {
                    File.Delete("index_date.dat");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла index_date.dat: {ex.Message}");
                }
            }
            if (File.Exists("database.bin"))
            {
                try
                {
                    File.Delete("database.bin");
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла database.bin: {ex.Message}");
                }
            }
            if (!File.Exists("index.dat") && !File.Exists("database.bin")
                && !File.Exists("index_name.dat") && !File.Exists("index_author.dat")
                && !File.Exists("index_price.dat") && !File.Exists("index_data.dat"))
            {
                MessageBox.Show("Файлы базы данных удалены.");
            }
        }
        public List<BookRecord?> SearchRecordById(int idToFind)
        {
            List<BookRecord?> results = new List<BookRecord?>();
            // Проверяем наличие файла
            if (!File.Exists(filePath))
            {
                return results; // Возвращаем пцстой, если файла не существует
            }

            // Проверяем наличие записи в индексе
            if (indexTable.TryGetValue(idToFind, out long position))
            {
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open, FileAccess.Read)))
                {
                    // Переходим к позиции записи
                    reader.BaseStream.Seek(position, SeekOrigin.Begin);

                    // Считываем запись
                    int id = reader.ReadInt32();
                    string name = reader.ReadString();
                    string author = reader.ReadString();
                    double price = reader.ReadDouble();
                    DateTime date = DateTime.Parse(reader.ReadString()).Date;

                    // Возвращаем найденную запись
                    results.Add(new BookRecord(id, name, author, price, date));
                }
            }
            return results;
        }



        public List<BookRecord?> SearchRecordsByAuthor(string authorToFind)
        {
            List<BookRecord?> results = new List<BookRecord?>();

            // Проверяем наличие файла и индекса
            if (!File.Exists(filePath)) return results; // если нет, то пустой список 
            if (authorTable.TryGetValue(authorToFind, out List<long> positions))
            {
                // Открываем файл и читаем записи по смещениям
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open)))
                {
                    foreach (long position in positions)
                    {
                        reader.BaseStream.Seek(position, SeekOrigin.Begin);

                        // Читаем запись
                        int id = reader.ReadInt32();
                        string name = reader.ReadString();
                        string author = reader.ReadString();
                        double price = reader.ReadDouble();
                        DateTime date = DateTime.Parse(reader.ReadString()).Date;

                        // Добавляем запись в результаты
                        results.Add(new BookRecord(id, name, author, price, date));
                    }
                }
            }
            return results;
        }

        public List<BookRecord?> SearchRecordsByName(string nameToFind)
        {
            List<BookRecord?> results = new List<BookRecord?>();

            // Проверяем наличие файла и индекса
            if (!File.Exists(filePath)) return results; // если нет, то пустой список 
            if (nameTable.TryGetValue(nameToFind, out List<long> positions))
            {
                // Открываем файл и читаем записи по смещениям
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open)))
                {
                    foreach (long position in positions)
                    {
                        reader.BaseStream.Seek(position, SeekOrigin.Begin);

                        // Читаем запись
                        int id = reader.ReadInt32();
                        string name = reader.ReadString();
                        string author = reader.ReadString();
                        double price = reader.ReadDouble();
                        DateTime date = DateTime.Parse(reader.ReadString()).Date;

                        // Добавляем запись в результаты
                        results.Add(new BookRecord(id, name, author, price, date));
                    }
                }
            }
            return results;
        }

        public List<BookRecord?> SearchRecordsByPrice(double priceToFind)
        {
            List<BookRecord?> results = new List<BookRecord?>();
            // Проверяем наличие файла и индекса
            if (!File.Exists(filePath)) return results; // если нет, то пустой список 
            if (priceTable.TryGetValue(priceToFind, out List<long> positions))
            {
                // Открываем файл и читаем записи по смещениям
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open)))
                {
                    foreach (long position in positions)
                    {
                        reader.BaseStream.Seek(position, SeekOrigin.Begin);

                        // Читаем запись
                        int id = reader.ReadInt32();
                        string name = reader.ReadString();
                        string author = reader.ReadString();
                        double price = reader.ReadDouble();
                        DateTime date = DateTime.Parse(reader.ReadString()).Date;

                        // Добавляем запись в результаты
                        results.Add(new BookRecord(id, name, author, price, date));
                    }
                }
            }
            return results;
        }
        public List<BookRecord?> SearchRecordsByDate(string dateToFind)
        {
            List<BookRecord?> results = new List<BookRecord?>();

            // Проверяем наличие файла и индекса
            if (!File.Exists(filePath)) return results; // если нет, то пустой список
            DateTime searchDate = DateTime.Parse(dateToFind).Date;
            if (dateTable.TryGetValue(DateTime.Parse(dateToFind).Date, out List<long> positions))
            {
                // Открываем файл и читаем записи по смещениям
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open)))
                {
                    foreach (long position in positions)
                    {
                        reader.BaseStream.Seek(position, SeekOrigin.Begin);

                        // Читаем запись
                        int id = reader.ReadInt32();
                        string name = reader.ReadString();
                        string author = reader.ReadString();
                        double price = reader.ReadDouble();
                        DateTime date = DateTime.Parse(reader.ReadString()).Date;

                        if (date == searchDate) // Проверяем только дату, игнорируя время
                        {
                            results.Add(new BookRecord(id, name, author, price, date));
                        }
                    }
                }
            }
            return results;
        }
        public BookRecord? SearchRecordByPosition(FileStream file, long position)
        {
            try
            {
                BookRecord record = new BookRecord();
                using (BinaryReader reader = new BinaryReader(file, Encoding.Default, true)) // Передаем существующий поток
                {
                    reader.BaseStream.Seek(position, SeekOrigin.Begin);
                    int id = reader.ReadInt32();
                    string name = reader.ReadString();
                    string author = reader.ReadString();
                    double price = reader.ReadDouble();
                    DateTime date = DateTime.Parse(reader.ReadString()).Date;

                    if (id == -1)
                        return null;

                    record.Author = author;
                    record.Date = date;
                    record.Price = price;
                    record.Name = name;
                    record.ID = id;
                }
                return record;
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при поиске записи по позиции: {ex.Message}");
                return null;
            }
        }


        public void DeleteRecordbyId(int id) // удаление записи 
        {
            if (!indexTable.ContainsKey(id))
            {
                MessageBox.Show("Запись не найдена в индексе.");
                return;
            }
            else
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    long position = indexTable[id];
                    file.Seek(position, SeekOrigin.Begin);
                    // Можно либо "обнулить" данные на месте, либо переопределить запись как удаленную.
                    // Здесь очищаем поля как удаление
                    using (var writer = new BinaryWriter(file, Encoding.Default, true))
                    {
                        writer.Write(-1); // Помечаем ID как удаленный, например, ID = -1.
                    }

                    indexTable.Remove(id);
                    BookRecord? record = SearchRecordByPosition(file, position);
                    if (record.HasValue)
                    {
                        // Удаляем позицию из индекса Name
                        if (nameTable.ContainsKey(record.Value.Name))
                        {
                            nameTable[record.Value.Name].Remove(position);
                            if (nameTable[record.Value.Name].Count == 0)
                            {
                                nameTable.Remove(record.Value.Name);
                            }
                        }

                        // Удаляем позицию из индекса Author
                        if (authorTable.ContainsKey(record.Value.Author))
                        {
                            authorTable[record.Value.Author].Remove(position);
                            if (authorTable[record.Value.Author].Count == 0)
                            {
                                authorTable.Remove(record.Value.Author);
                            }
                        }
                        // Удаляем позицию из индекса price
                        if (priceTable.ContainsKey(record.Value.Price))
                        {
                            priceTable[record.Value.Price].Remove(position);
                            if (priceTable[record.Value.Price].Count == 0)
                            {
                                priceTable.Remove(record.Value.Price);
                            }
                        }
                        if (dateTable.ContainsKey(record.Value.Date.Date))
                        {
                            dateTable[record.Value.Date.Date].Remove(position);
                            if (dateTable[record.Value.Date.Date].Count == 0)
                            {
                                dateTable.Remove(record.Value.Date.Date);
                            }
                        }
                    }
                    SaveAll();
                }
            }
        }

        public void DeleteRecordbyName(string name) // удаление записи 
        {
            if (!nameTable.ContainsKey(name))
            {
                MessageBox.Show("Запись не найдена.");
                return;
            }
            else
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    List<long> positions = nameTable[name];
                    using (var writer = new BinaryWriter(file, Encoding.Default, true))
                    {
                        foreach (long position in positions)
                        {
                            BookRecord? record = SearchRecordByPosition(file, position);
                            if (record.HasValue)
                            {
                                // Удаляем позицию из индекса Name
                                if (indexTable.ContainsKey(record.Value.ID))
                                {
                                    indexTable.Remove(record.Value.ID);
                                }

                                // Удаляем позицию из индекса Author
                                if (authorTable.ContainsKey(record.Value.Author))
                                {
                                    authorTable[record.Value.Author].Remove(position);
                                    if (authorTable[record.Value.Author].Count == 0)
                                    {
                                        authorTable.Remove(record.Value.Author);
                                    }
                                }
                                // Удаляем позицию из индекса price
                                if (priceTable.ContainsKey(record.Value.Price))
                                {
                                    priceTable[record.Value.Price].Remove(position);
                                    if (priceTable[record.Value.Price].Count == 0)
                                    {
                                        priceTable.Remove(record.Value.Price);
                                    }
                                }
                                if (dateTable.ContainsKey(record.Value.Date.Date))
                                {
                                    dateTable[record.Value.Date.Date].Remove(position);
                                    if (dateTable[record.Value.Date.Date].Count == 0)
                                    {
                                        dateTable.Remove(record.Value.Date.Date);
                                    }
                                }
                            }
                            file.Seek(position, SeekOrigin.Begin);
                            writer.Write(-1); // Помечаем ID как удаленный
                        }
                    }
                    nameTable.Remove(name);
                    SaveAll();
                }
            }
        }

        public void DeleteRecordbyAuthor(string author) // удаление записи 
        {
            if (!authorTable.ContainsKey(author))
            {
                MessageBox.Show("Запись не найдена.");
                return;
            }
            else
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    List<long> positions = authorTable[author];
                    using (var writer = new BinaryWriter(file, Encoding.Default, true))
                    {
                        foreach (long position in positions)
                        {
                            BookRecord? record = SearchRecordByPosition(file, position);
                            if (record.HasValue)
                            {
                                // Удаляем позицию из индекса Name
                                if (indexTable.ContainsKey(record.Value.ID))
                                {
                                    indexTable.Remove(record.Value.ID);
                                }

                                // Удаляем позицию из индекса Author
                                if (nameTable.ContainsKey(record.Value.Name))
                                {
                                    nameTable[record.Value.Name].Remove(position);
                                    if (nameTable[record.Value.Name].Count == 0)
                                    {
                                        nameTable.Remove(record.Value.Name);
                                    }
                                }
                                // Удаляем позицию из индекса price
                                if (priceTable.ContainsKey(record.Value.Price))
                                {
                                    priceTable[record.Value.Price].Remove(position);
                                    if (priceTable[record.Value.Price].Count == 0)
                                    {
                                        priceTable.Remove(record.Value.Price);
                                    }
                                }
                                if (dateTable.ContainsKey(record.Value.Date.Date))
                                {
                                    dateTable[record.Value.Date.Date].Remove(position);
                                    if (dateTable[record.Value.Date.Date].Count == 0)
                                    {
                                        dateTable.Remove(record.Value.Date.Date);
                                    }
                                }
                            }
                            file.Seek(position, SeekOrigin.Begin);
                            writer.Write(-1); // Помечаем ID как удаленный
                        }
                        authorTable.Remove(author);
                        SaveAll();
                    }
                }
            }
        }

        public void DeleteRecordbyPrice(double price) // удаление записи 
        {
            if (!priceTable.ContainsKey(price))
            {
                MessageBox.Show("Запись не найдена.");
                return;
            }
            else
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    List<long> positions = priceTable[price];
                    using (var writer = new BinaryWriter(file, Encoding.Default, true))
                    {
                        foreach (long position in positions)
                        {

                            BookRecord? record = SearchRecordByPosition(file, position);
                            if (record.HasValue)
                            {
                                // Удаляем позицию из индекса Name
                                if (indexTable.ContainsKey(record.Value.ID))
                                {
                                    indexTable.Remove(record.Value.ID);
                                }

                                // Удаляем позицию из индекса Author
                                if (nameTable.ContainsKey(record.Value.Name))
                                {
                                    nameTable[record.Value.Name].Remove(position);
                                    if (nameTable[record.Value.Name].Count == 0)
                                    {
                                        nameTable.Remove(record.Value.Name);
                                    }
                                }
                                // Удаляем позицию из индекса price
                                if (authorTable.ContainsKey(record.Value.Author))
                                {
                                    authorTable[record.Value.Author].Remove(position);
                                    if (authorTable[record.Value.Author].Count == 0)
                                    {
                                        authorTable.Remove(record.Value.Author);
                                    }
                                }
                                if (dateTable.ContainsKey(record.Value.Date.Date))
                                {
                                    dateTable[record.Value.Date.Date].Remove(position);
                                    if (dateTable[record.Value.Date.Date].Count == 0)
                                    {
                                        dateTable.Remove(record.Value.Date.Date);
                                    }
                                }
                            }
                            file.Seek(position, SeekOrigin.Begin);
                            writer.Write(-1); // Помечаем ID как удаленный
                        }
                        priceTable.Remove(price);
                        SaveAll();
                    }
                }
            }
        }

        public void DeleteRecordbyDate(DateTime date) // удаление записи 
        {
            if (!dateTable.ContainsKey(date.Date))
            {
                MessageBox.Show("Запись не найдена.");
                return;
            }
            else
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    List<long> positions = dateTable[date.Date];
                    using (var writer = new BinaryWriter(file, Encoding.Default, true))
                    {
                        foreach (long position in positions)
                        {

                            BookRecord? record = SearchRecordByPosition(file, position);
                            if (record.HasValue)
                            {
                                // Удаляем позицию из индекса Name
                                if (indexTable.ContainsKey(record.Value.ID))
                                {
                                    indexTable.Remove(record.Value.ID);
                                }

                                // Удаляем позицию из индекса Author
                                if (nameTable.ContainsKey(record.Value.Name))
                                {
                                    nameTable[record.Value.Name].Remove(position);
                                    if (nameTable[record.Value.Name].Count == 0)
                                    {
                                        nameTable.Remove(record.Value.Name);
                                    }
                                }
                                // Удаляем позицию из индекса price
                                if (authorTable.ContainsKey(record.Value.Author))
                                {
                                    authorTable[record.Value.Author].Remove(position);
                                    if (authorTable[record.Value.Author].Count == 0)
                                    {
                                        authorTable.Remove(record.Value.Author);
                                    }
                                }
                                if (priceTable.ContainsKey(record.Value.Price))
                                {
                                    priceTable[record.Value.Price].Remove(position);
                                    if (priceTable[record.Value.Price].Count == 0)
                                    {
                                        priceTable.Remove(record.Value.Price);
                                    }
                                }
                            }
                            file.Seek(position, SeekOrigin.Begin);
                            writer.Write(-1); // Помечаем ID как удаленный
                        }
                        dateTable.Remove(date.Date);
                        SaveAll();
                    }
                }
            }
        }

        public bool EditRecordById(int id, string? newName = null, string? newAuthor = null, double? newPrice = null, DateTime? newDate = null)
        {
            if (!indexTable.ContainsKey(id))
            {
                MessageBox.Show("Запись не найдена.");
                return false;
            }

            // Получаем позицию записи в файле
            long position = indexTable[id];

            try
            {
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var writer = new BinaryWriter(file, Encoding.Default, true))
                using (var reader = new BinaryReader(file, Encoding.Default, true))
                {
                    // Считываем текущую запись
                    file.Seek(position, SeekOrigin.Begin);
                    int recordId = reader.ReadInt32();
                    string currentName = reader.ReadString();
                    string currentAuthor = reader.ReadString();
                    double currentPrice = reader.ReadDouble();
                    DateTime currentDate = DateTime.Parse(reader.ReadString());
                    if (recordId != id)
                    {
                        MessageBox.Show("Ошибка: ID записи не совпадает.");
                        return false;
                    }

                    // Обновляем индексные таблицы
                    if (newName != null && newName != currentName)
                    {
                        if (nameTable.ContainsKey(currentName))
                        {
                            nameTable[currentName].Remove(position);
                            if (nameTable[currentName].Count == 0)
                            {
                                nameTable.Remove(currentName);
                            }
                        }
                        if (!nameTable.ContainsKey(newName))
                        {
                            nameTable[newName] = new List<long>();
                        }
                        nameTable[newName].Add(position);
                    }

                    if (newAuthor != null && newAuthor != currentAuthor)
                    {
                        if (authorTable.ContainsKey(currentAuthor))
                        {
                            authorTable[currentAuthor].Remove(position);
                            if (authorTable[currentAuthor].Count == 0)
                            {
                                authorTable.Remove(currentAuthor);
                            }
                        }
                        if (!authorTable.ContainsKey(newAuthor))
                        {
                            authorTable[newAuthor] = new List<long>();
                        }
                        authorTable[newAuthor].Add(position);
                    }

                    if (newPrice != null && newPrice != currentPrice)
                    {
                        if (priceTable.ContainsKey(currentPrice))
                        {
                            priceTable[currentPrice].Remove(position);
                            if (priceTable[currentPrice].Count == 0)
                            {
                                priceTable.Remove(currentPrice);
                            }
                        }
                        if (!priceTable.ContainsKey(newPrice.Value))
                        {
                            priceTable[newPrice.Value] = new List<long>();
                        }
                        priceTable[newPrice.Value].Add(position);
                    }

                    if (newDate != null && newDate.Value.Date != currentDate.Date)
                    {
                        if (dateTable.ContainsKey(currentDate.Date))
                        {
                            dateTable[currentDate.Date].Remove(position);
                            if (dateTable[currentDate.Date].Count == 0)
                            {
                                dateTable.Remove(currentDate.Date);
                            }
                        }
                        if (!dateTable.ContainsKey(newDate.Value.Date))
                        {
                            dateTable[newDate.Value.Date] = new List<long>();
                        }
                        dateTable[newDate.Value.Date].Add(position);
                    }

                    // Перезапись записи в файл
                    file.Seek(position, SeekOrigin.Begin);
                    writer.Write(recordId); // ID остаётся прежним
                    writer.Write(newName ?? currentName);
                    writer.Write(newAuthor ?? currentAuthor);
                    writer.Write(newPrice ?? currentPrice);
                    writer.Write((newDate ?? currentDate).Date.ToString());

                    SaveAll(); // Сохраняем обновлённые индексы
                    return true;
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Ошибка при редактировании записи: {ex.Message}");
                return false;
            }
        }
        public void SaveToXlsx()
        {
            // Проверка, существует ли файл базы данных
            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл базы данных не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Создаем новую Excel-книгу
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0]; // Первый лист

            // Заголовки столбцов
            sheet.Range["A1"].Text = "ID";
            sheet.Range["B1"].Text = "Name";
            sheet.Range["C1"].Text = "Author";
            sheet.Range["D1"].Text = "Price";
            sheet.Range["E1"].Text = "Date";

            int row = 2; // Начинаем со второй строки для данных

            try
            {
                // Открываем файл для чтения данных
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var reader = new BinaryReader(fileStream, Encoding.Default, true))
                {
                    foreach (var kvp in indexTable)
                    {
                        int id = kvp.Key;
                        long position = kvp.Value;

                        // Переходим к позиции записи в файле
                        fileStream.Seek(position, SeekOrigin.Begin);

                        // Читаем данные
                        int recordId = reader.ReadInt32();
                        string name = reader.ReadString();
                        string author = reader.ReadString();
                        double price = reader.ReadDouble();
                        DateTime date = DateTime.Parse(reader.ReadString());

                        // Записываем данные в Excel
                        sheet.Range[$"A{row}"].NumberValue = recordId;
                        sheet.Range[$"B{row}"].Text = name;
                        sheet.Range[$"C{row}"].Text = author;
                        sheet.Range[$"D{row}"].NumberValue = price;
                        sheet.Range[$"E{row}"].Text = date.ToString("yyyy-MM-dd");

                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Сохраняем Excel-файл с использованием диалога
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                saveFileDialog.Title = "Сохранить базу данных как Excel файл";
                saveFileDialog.FileName = "Database.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        workbook.SaveToFile(saveFileDialog.FileName, ExcelVersion.Version2016);
                        MessageBox.Show("База данных успешно сохранена в файл Excel!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

        }
    }
    [Serializable]
    public struct BookRecord
    {
        public int ID;
        public string Name;
        public string Author;
        public double Price;
        public DateTime Date;

        public BookRecord(int id, string name, string author, double price, DateTime date)
        {
            ID = id;
            Name = name;
            Author = author;
            Price = price;
            Date = date.Date;
        }
    }
    
}

