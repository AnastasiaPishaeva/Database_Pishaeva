namespace Лабораторная_2_С_
{
    partial class ButtonSaveXcel
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            dataGridView1 = new System.Windows.Forms.DataGridView();
            buttonSearchRecord = new System.Windows.Forms.Button();
            buttonDeleteClick = new System.Windows.Forms.Button();
            buttonAddRecord = new System.Windows.Forms.Button();
            textBoxID = new System.Windows.Forms.TextBox();
            EnterID = new System.Windows.Forms.Label();
            textBoxName = new System.Windows.Forms.TextBox();
            textBoxAuthor = new System.Windows.Forms.TextBox();
            textBoxPrice = new System.Windows.Forms.TextBox();
            dateTimePicker = new System.Windows.Forms.DateTimePicker();
            EnterName = new System.Windows.Forms.Label();
            EnterAuthor = new System.Windows.Forms.Label();
            EnterPrice = new System.Windows.Forms.Label();
            EnterData = new System.Windows.Forms.Label();
            buttonClear = new System.Windows.Forms.Button();
            comboBox = new System.Windows.Forms.ComboBox();
            label1 = new System.Windows.Forms.Label();
            buttonEdit = new System.Windows.Forms.Button();
            button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new System.Drawing.Point(43, 13);
            dataGridView1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 24;
            dataGridView1.Size = new System.Drawing.Size(934, 461);
            dataGridView1.TabIndex = 0;
            // 
            // buttonSearchRecord
            // 
            buttonSearchRecord.Location = new System.Drawing.Point(365, 623);
            buttonSearchRecord.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            buttonSearchRecord.Name = "buttonSearchRecord";
            buttonSearchRecord.Size = new System.Drawing.Size(140, 82);
            buttonSearchRecord.TabIndex = 1;
            buttonSearchRecord.Text = "Поиск";
            buttonSearchRecord.UseVisualStyleBackColor = true;
            buttonSearchRecord.Click += button1_Click;
            // 
            // buttonDeleteClick
            // 
            buttonDeleteClick.Location = new System.Drawing.Point(202, 623);
            buttonDeleteClick.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            buttonDeleteClick.Name = "buttonDeleteClick";
            buttonDeleteClick.Size = new System.Drawing.Size(140, 82);
            buttonDeleteClick.TabIndex = 2;
            buttonDeleteClick.Text = "Удалить запись";
            buttonDeleteClick.UseVisualStyleBackColor = true;
            buttonDeleteClick.Click += button2_Click;
            // 
            // buttonAddRecord
            // 
            buttonAddRecord.Location = new System.Drawing.Point(43, 623);
            buttonAddRecord.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            buttonAddRecord.Name = "buttonAddRecord";
            buttonAddRecord.Size = new System.Drawing.Size(140, 82);
            buttonAddRecord.TabIndex = 3;
            buttonAddRecord.Text = "Добавить запись";
            buttonAddRecord.UseVisualStyleBackColor = true;
            buttonAddRecord.Click += button3_Click;
            // 
            // textBoxID
            // 
            textBoxID.Location = new System.Drawing.Point(43, 497);
            textBoxID.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textBoxID.Name = "textBoxID";
            textBoxID.Size = new System.Drawing.Size(132, 27);
            textBoxID.TabIndex = 6;
            textBoxID.TextChanged += textBoxID_TextChanged;
            // 
            // EnterID
            // 
            EnterID.AutoSize = true;
            EnterID.Location = new System.Drawing.Point(40, 528);
            EnterID.Name = "EnterID";
            EnterID.Size = new System.Drawing.Size(84, 20);
            EnterID.TabIndex = 7;
            EnterID.Text = "Введите ID";
            EnterID.Click += EnterID_Click;
            // 
            // textBoxName
            // 
            textBoxName.Location = new System.Drawing.Point(222, 497);
            textBoxName.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textBoxName.Name = "textBoxName";
            textBoxName.Size = new System.Drawing.Size(132, 27);
            textBoxName.TabIndex = 8;
            textBoxName.TextChanged += textBox3_TextChanged;
            // 
            // textBoxAuthor
            // 
            textBoxAuthor.Location = new System.Drawing.Point(407, 497);
            textBoxAuthor.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textBoxAuthor.Name = "textBoxAuthor";
            textBoxAuthor.Size = new System.Drawing.Size(132, 27);
            textBoxAuthor.TabIndex = 9;
            // 
            // textBoxPrice
            // 
            textBoxPrice.Location = new System.Drawing.Point(587, 497);
            textBoxPrice.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            textBoxPrice.Name = "textBoxPrice";
            textBoxPrice.Size = new System.Drawing.Size(132, 27);
            textBoxPrice.TabIndex = 10;
            // 
            // dateTimePicker
            // 
            dateTimePicker.Location = new System.Drawing.Point(774, 497);
            dateTimePicker.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            dateTimePicker.Name = "dateTimePicker";
            dateTimePicker.Size = new System.Drawing.Size(200, 27);
            dateTimePicker.TabIndex = 11;
            // 
            // EnterName
            // 
            EnterName.AutoSize = true;
            EnterName.Location = new System.Drawing.Point(222, 528);
            EnterName.Name = "EnterName";
            EnterName.Size = new System.Drawing.Size(183, 20);
            EnterName.TabIndex = 12;
            EnterName.Text = "Введите  название книги";
            EnterName.Click += label3_Click;
            // 
            // EnterAuthor
            // 
            EnterAuthor.AutoSize = true;
            EnterAuthor.Location = new System.Drawing.Point(407, 528);
            EnterAuthor.Name = "EnterAuthor";
            EnterAuthor.Size = new System.Drawing.Size(149, 20);
            EnterAuthor.TabIndex = 13;
            EnterAuthor.Text = "Введите имя автора";
            // 
            // EnterPrice
            // 
            EnterPrice.AutoSize = true;
            EnterPrice.Location = new System.Drawing.Point(587, 528);
            EnterPrice.Name = "EnterPrice";
            EnterPrice.Size = new System.Drawing.Size(102, 20);
            EnterPrice.TabIndex = 14;
            EnterPrice.Text = "Введите цену";
            EnterPrice.Click += EnterPrice_Click;
            // 
            // EnterData
            // 
            EnterData.AutoSize = true;
            EnterData.Location = new System.Drawing.Point(774, 528);
            EnterData.Name = "EnterData";
            EnterData.Size = new System.Drawing.Size(179, 20);
            EnterData.TabIndex = 15;
            EnterData.Text = "Выберете дату создания";
            // 
            // buttonClear
            // 
            buttonClear.Location = new System.Drawing.Point(526, 623);
            buttonClear.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            buttonClear.Name = "buttonClear";
            buttonClear.Size = new System.Drawing.Size(140, 82);
            buttonClear.TabIndex = 16;
            buttonClear.Text = "Очистить";
            buttonClear.UseVisualStyleBackColor = true;
            buttonClear.Click += buttonClear_Click;
            // 
            // comboBox
            // 
            comboBox.FormattingEnabled = true;
            comboBox.Location = new System.Drawing.Point(43, 559);
            comboBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            comboBox.Name = "comboBox";
            comboBox.Size = new System.Drawing.Size(469, 28);
            comboBox.TabIndex = 17;
            comboBox.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(43, 591);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(469, 20);
            label1.TabIndex = 18;
            label1.Text = "Выберите поле, по которому осуществляется поиск или удаление";
            // 
            // buttonEdit
            // 
            buttonEdit.Location = new System.Drawing.Point(687, 623);
            buttonEdit.Name = "buttonEdit";
            buttonEdit.Size = new System.Drawing.Size(140, 82);
            buttonEdit.TabIndex = 19;
            buttonEdit.Text = "Редактировать";
            buttonEdit.UseVisualStyleBackColor = true;
            buttonEdit.Click += buttonEdit_Click;
            // 
            // button1
            // 
            button1.Location = new System.Drawing.Point(849, 623);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(140, 82);
            button1.TabIndex = 20;
            button1.Text = " Сохранить в xcel";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click_1;
            // 
            // ButtonSaveXcel
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1032, 736);
            Controls.Add(button1);
            Controls.Add(buttonEdit);
            Controls.Add(label1);
            Controls.Add(comboBox);
            Controls.Add(buttonClear);
            Controls.Add(EnterData);
            Controls.Add(EnterPrice);
            Controls.Add(EnterAuthor);
            Controls.Add(EnterName);
            Controls.Add(dateTimePicker);
            Controls.Add(textBoxPrice);
            Controls.Add(textBoxAuthor);
            Controls.Add(textBoxName);
            Controls.Add(EnterID);
            Controls.Add(textBoxID);
            Controls.Add(buttonAddRecord);
            Controls.Add(buttonDeleteClick);
            Controls.Add(buttonSearchRecord);
            Controls.Add(dataGridView1);
            Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            Name = "ButtonSaveXcel";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button buttonSearchRecord;
        private System.Windows.Forms.Button buttonDeleteClick;
        private System.Windows.Forms.Button buttonAddRecord;
        private System.Windows.Forms.TextBox textBoxID;
        private System.Windows.Forms.Label EnterID;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.TextBox textBoxAuthor;
        private System.Windows.Forms.TextBox textBoxPrice;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.Label EnterName;
        private System.Windows.Forms.Label EnterAuthor;
        private System.Windows.Forms.Label EnterPrice;
        private System.Windows.Forms.Label EnterData;
        private System.Windows.Forms.Button buttonClear;
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonEdit;
        private System.Windows.Forms.Button button1;
    }
}

