namespace Otpuska
{
    partial class MenuScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MenuScreen));
            this.metroTabControl1 = new MetroFramework.Controls.MetroTabControl();
            this.metroTabPage1 = new MetroFramework.Controls.MetroTabPage();
            this.deleteSotrButton = new MetroFramework.Controls.MetroButton();
            this.deleteOtdelButton = new MetroFramework.Controls.MetroButton();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.button6 = new MetroFramework.Controls.MetroButton();
            this.button4 = new MetroFramework.Controls.MetroButton();
            this.button2 = new MetroFramework.Controls.MetroButton();
            this.button1 = new MetroFramework.Controls.MetroButton();
            this.FIOListView = new MetroFramework.Controls.MetroListView();
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.otdelsListView = new MetroFramework.Controls.MetroListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.metroTabPage2 = new MetroFramework.Controls.MetroTabPage();
            this.button5 = new MetroFramework.Controls.MetroButton();
            this.button3 = new MetroFramework.Controls.MetroButton();
            this.button7 = new MetroFramework.Controls.MetroButton();
            this.metroTabPage3 = new MetroFramework.Controls.MetroTabPage();
            this.createTable = new MetroFramework.Controls.MetroButton();
            this.inputYearTextBox = new MetroFramework.Controls.MetroTextBox();
            this.metroTabControl1.SuspendLayout();
            this.metroTabPage1.SuspendLayout();
            this.metroTabPage2.SuspendLayout();
            this.metroTabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroTabControl1
            // 
            this.metroTabControl1.Controls.Add(this.metroTabPage1);
            this.metroTabControl1.Controls.Add(this.metroTabPage2);
            this.metroTabControl1.Controls.Add(this.metroTabPage3);
            this.metroTabControl1.Location = new System.Drawing.Point(23, 63);
            this.metroTabControl1.Name = "metroTabControl1";
            this.metroTabControl1.SelectedIndex = 0;
            this.metroTabControl1.Size = new System.Drawing.Size(1052, 484);
            this.metroTabControl1.Style = MetroFramework.MetroColorStyle.Brown;
            this.metroTabControl1.TabIndex = 12;
            this.metroTabControl1.UseSelectable = true;
            // 
            // metroTabPage1
            // 
            this.metroTabPage1.Controls.Add(this.deleteSotrButton);
            this.metroTabPage1.Controls.Add(this.deleteOtdelButton);
            this.metroTabPage1.Controls.Add(this.metroLabel1);
            this.metroTabPage1.Controls.Add(this.button6);
            this.metroTabPage1.Controls.Add(this.button4);
            this.metroTabPage1.Controls.Add(this.button2);
            this.metroTabPage1.Controls.Add(this.button1);
            this.metroTabPage1.Controls.Add(this.FIOListView);
            this.metroTabPage1.Controls.Add(this.otdelsListView);
            this.metroTabPage1.HorizontalScrollbarBarColor = false;
            this.metroTabPage1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage1.HorizontalScrollbarSize = 10;
            this.metroTabPage1.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage1.Name = "metroTabPage1";
            this.metroTabPage1.Size = new System.Drawing.Size(1044, 442);
            this.metroTabPage1.Style = MetroFramework.MetroColorStyle.Brown;
            this.metroTabPage1.TabIndex = 0;
            this.metroTabPage1.Text = "Данные КЦ";
            this.metroTabPage1.VerticalScrollbar = true;
            this.metroTabPage1.VerticalScrollbarBarColor = true;
            this.metroTabPage1.VerticalScrollbarHighlightOnWheel = true;
            this.metroTabPage1.VerticalScrollbarSize = 10;
            this.metroTabPage1.Click += new System.EventHandler(this.metroTabPage1_Click);
            // 
            // deleteSotrButton
            // 
            this.deleteSotrButton.Location = new System.Drawing.Point(809, 265);
            this.deleteSotrButton.Name = "deleteSotrButton";
            this.deleteSotrButton.Size = new System.Drawing.Size(192, 27);
            this.deleteSotrButton.TabIndex = 18;
            this.deleteSotrButton.Text = "Удалить сотрудника";
            this.deleteSotrButton.UseSelectable = true;
            this.deleteSotrButton.Click += new System.EventHandler(this.deleteSotrButton_Click);
            // 
            // deleteOtdelButton
            // 
            this.deleteOtdelButton.Enabled = false;
            this.deleteOtdelButton.Location = new System.Drawing.Point(267, 265);
            this.deleteOtdelButton.Name = "deleteOtdelButton";
            this.deleteOtdelButton.Size = new System.Drawing.Size(192, 27);
            this.deleteOtdelButton.TabIndex = 17;
            this.deleteOtdelButton.Text = "Удалить отдел";
            this.deleteOtdelButton.UseSelectable = true;
            this.deleteOtdelButton.Click += new System.EventHandler(this.deleteOtdelButton_Click);
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.Location = new System.Drawing.Point(0, 317);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(418, 38);
            this.metroLabel1.TabIndex = 16;
            this.metroLabel1.Text = "    Если впервые пользуетесь программой, то список сотрудников \r\nвы можете добави" +
    "ть из Excel-файла";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(3, 371);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(192, 27);
            this.button6.TabIndex = 15;
            this.button6.Text = "Импорт сотрудников";
            this.button6.UseSelectable = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(545, 317);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(192, 27);
            this.button4.TabIndex = 14;
            this.button4.Text = "Редактировать сотрудника";
            this.button4.UseSelectable = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(545, 265);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(192, 27);
            this.button2.TabIndex = 13;
            this.button2.Text = "Добавить сотрудника";
            this.button2.UseSelectable = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(3, 265);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(192, 27);
            this.button1.TabIndex = 12;
            this.button1.Text = "Добавить отдел";
            this.button1.UseSelectable = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FIOListView
            // 
            this.FIOListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader2});
            this.FIOListView.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.FIOListView.FullRowSelect = true;
            this.FIOListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.FIOListView.HideSelection = false;
            this.FIOListView.Location = new System.Drawing.Point(545, 25);
            this.FIOListView.MultiSelect = false;
            this.FIOListView.Name = "FIOListView";
            this.FIOListView.OwnerDraw = true;
            this.FIOListView.ShowGroups = false;
            this.FIOListView.Size = new System.Drawing.Size(456, 211);
            this.FIOListView.Style = MetroFramework.MetroColorStyle.Brown;
            this.FIOListView.TabIndex = 11;
            this.FIOListView.UseCompatibleStateImageBehavior = false;
            this.FIOListView.UseSelectable = true;
            this.FIOListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "ФИО сотрудника";
            this.columnHeader2.Width = 434;
            // 
            // otdelsListView
            // 
            this.otdelsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.otdelsListView.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.otdelsListView.FullRowSelect = true;
            this.otdelsListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.otdelsListView.HideSelection = false;
            this.otdelsListView.Location = new System.Drawing.Point(3, 25);
            this.otdelsListView.MaximumSize = new System.Drawing.Size(456, 600);
            this.otdelsListView.MinimumSize = new System.Drawing.Size(456, 211);
            this.otdelsListView.MultiSelect = false;
            this.otdelsListView.Name = "otdelsListView";
            this.otdelsListView.OwnerDraw = true;
            this.otdelsListView.ShowGroups = false;
            this.otdelsListView.Size = new System.Drawing.Size(456, 211);
            this.otdelsListView.Style = MetroFramework.MetroColorStyle.Brown;
            this.otdelsListView.TabIndex = 10;
            this.otdelsListView.UseCompatibleStateImageBehavior = false;
            this.otdelsListView.UseSelectable = true;
            this.otdelsListView.View = System.Windows.Forms.View.Details;
            this.otdelsListView.SelectedIndexChanged += new System.EventHandler(this.otdelsListView_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Название отдела/бюро";
            this.columnHeader1.Width = 434;
            // 
            // metroTabPage2
            // 
            this.metroTabPage2.Controls.Add(this.button5);
            this.metroTabPage2.Controls.Add(this.button3);
            this.metroTabPage2.Controls.Add(this.button7);
            this.metroTabPage2.HorizontalScrollbarBarColor = true;
            this.metroTabPage2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.HorizontalScrollbarSize = 10;
            this.metroTabPage2.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage2.Name = "metroTabPage2";
            this.metroTabPage2.Size = new System.Drawing.Size(1044, 442);
            this.metroTabPage2.TabIndex = 1;
            this.metroTabPage2.Text = "Отпуска";
            this.metroTabPage2.VerticalScrollbarBarColor = true;
            this.metroTabPage2.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage2.VerticalScrollbarSize = 10;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(221, 26);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(192, 27);
            this.button5.TabIndex = 14;
            this.button5.Text = "Импорт отпусков";
            this.button5.UseSelectable = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(3, 26);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(192, 27);
            this.button3.TabIndex = 13;
            this.button3.Text = "Выбор отпуска";
            this.button3.UseSelectable = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(439, 26);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(192, 27);
            this.button7.TabIndex = 11;
            this.button7.Text = "Экспорт графика отпусков в Excel";
            this.button7.UseSelectable = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // metroTabPage3
            // 
            this.metroTabPage3.Controls.Add(this.createTable);
            this.metroTabPage3.Controls.Add(this.inputYearTextBox);
            this.metroTabPage3.HorizontalScrollbarBarColor = true;
            this.metroTabPage3.HorizontalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.HorizontalScrollbarSize = 10;
            this.metroTabPage3.Location = new System.Drawing.Point(4, 38);
            this.metroTabPage3.Name = "metroTabPage3";
            this.metroTabPage3.Size = new System.Drawing.Size(1044, 442);
            this.metroTabPage3.TabIndex = 2;
            this.metroTabPage3.Text = "Табель учёта времени";
            this.metroTabPage3.VerticalScrollbarBarColor = true;
            this.metroTabPage3.VerticalScrollbarHighlightOnWheel = false;
            this.metroTabPage3.VerticalScrollbarSize = 10;
            // 
            // createTable
            // 
            this.createTable.Location = new System.Drawing.Point(223, 28);
            this.createTable.Name = "createTable";
            this.createTable.Size = new System.Drawing.Size(192, 27);
            this.createTable.TabIndex = 13;
            this.createTable.Text = "Создать табель";
            this.createTable.UseSelectable = true;
            this.createTable.Click += new System.EventHandler(this.createTable_Click);
            // 
            // inputYearTextBox
            // 
            // 
            // 
            // 
            this.inputYearTextBox.CustomButton.Image = null;
            this.inputYearTextBox.CustomButton.Location = new System.Drawing.Point(166, 1);
            this.inputYearTextBox.CustomButton.Name = "";
            this.inputYearTextBox.CustomButton.Size = new System.Drawing.Size(25, 25);
            this.inputYearTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.inputYearTextBox.CustomButton.TabIndex = 1;
            this.inputYearTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.inputYearTextBox.CustomButton.UseSelectable = true;
            this.inputYearTextBox.CustomButton.Visible = false;
            this.inputYearTextBox.Lines = new string[0];
            this.inputYearTextBox.Location = new System.Drawing.Point(3, 28);
            this.inputYearTextBox.MaxLength = 32767;
            this.inputYearTextBox.Name = "inputYearTextBox";
            this.inputYearTextBox.PasswordChar = '\0';
            this.inputYearTextBox.PromptText = "Введите год";
            this.inputYearTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.inputYearTextBox.SelectedText = "";
            this.inputYearTextBox.SelectionLength = 0;
            this.inputYearTextBox.SelectionStart = 0;
            this.inputYearTextBox.ShortcutsEnabled = true;
            this.inputYearTextBox.Size = new System.Drawing.Size(192, 27);
            this.inputYearTextBox.TabIndex = 12;
            this.inputYearTextBox.UseSelectable = true;
            this.inputYearTextBox.WaterMark = "Введите год";
            this.inputYearTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.inputYearTextBox.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // MenuScreen
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1125, 582);
            this.Controls.Add(this.metroTabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(2000, 2180);
            this.MinimumSize = new System.Drawing.Size(435, 140);
            this.Name = "MenuScreen";
            this.Resizable = false;
            this.ShadowType = MetroFramework.Forms.MetroFormShadowType.AeroShadow;
            this.Style = MetroFramework.MetroColorStyle.Brown;
            this.Text = "Конструкторский центр ДАиР";
            this.metroTabControl1.ResumeLayout(false);
            this.metroTabPage1.ResumeLayout(false);
            this.metroTabPage1.PerformLayout();
            this.metroTabPage2.ResumeLayout(false);
            this.metroTabPage3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private MetroFramework.Controls.MetroTabControl metroTabControl1;
        private MetroFramework.Controls.MetroTabPage metroTabPage1;
        private MetroFramework.Controls.MetroTabPage metroTabPage2;
        private MetroFramework.Controls.MetroTabPage metroTabPage3;
        private MetroFramework.Controls.MetroTextBox inputYearTextBox;
        private MetroFramework.Controls.MetroListView otdelsListView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private MetroFramework.Controls.MetroListView FIOListView;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private MetroFramework.Controls.MetroButton button1;
        private MetroFramework.Controls.MetroButton button2;
        private MetroFramework.Controls.MetroButton button4;
        private MetroFramework.Controls.MetroButton button6;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private MetroFramework.Controls.MetroButton deleteOtdelButton;
        private MetroFramework.Controls.MetroButton deleteSotrButton;
        private MetroFramework.Controls.MetroButton button5;
        private MetroFramework.Controls.MetroButton button3;
        private MetroFramework.Controls.MetroButton button7;
        private MetroFramework.Controls.MetroButton createTable;
    }
}

