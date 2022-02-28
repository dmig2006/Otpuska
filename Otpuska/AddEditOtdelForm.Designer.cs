namespace Otpuska
{
    partial class AddEditOtdelForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddEditOtdelForm));
            this.nameTextBox = new MetroFramework.Controls.MetroTextBox();
            this.SaveButton = new MetroFramework.Controls.MetroButton();
            this.shortNameTextBox = new MetroFramework.Controls.MetroTextBox();
            this.SuspendLayout();
            // 
            // nameTextBox
            // 
            // 
            // 
            // 
            this.nameTextBox.CustomButton.Image = null;
            this.nameTextBox.CustomButton.Location = new System.Drawing.Point(445, 1);
            this.nameTextBox.CustomButton.Name = "";
            this.nameTextBox.CustomButton.Size = new System.Drawing.Size(25, 25);
            this.nameTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.nameTextBox.CustomButton.TabIndex = 1;
            this.nameTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.nameTextBox.CustomButton.UseSelectable = true;
            this.nameTextBox.CustomButton.Visible = false;
            this.nameTextBox.Lines = new string[0];
            this.nameTextBox.Location = new System.Drawing.Point(23, 78);
            this.nameTextBox.MaxLength = 32767;
            this.nameTextBox.Name = "nameTextBox";
            this.nameTextBox.PasswordChar = '\0';
            this.nameTextBox.PromptText = "Введите название отдела";
            this.nameTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.nameTextBox.SelectedText = "";
            this.nameTextBox.SelectionLength = 0;
            this.nameTextBox.SelectionStart = 0;
            this.nameTextBox.ShortcutsEnabled = true;
            this.nameTextBox.Size = new System.Drawing.Size(471, 27);
            this.nameTextBox.TabIndex = 0;
            this.nameTextBox.UseSelectable = true;
            this.nameTextBox.WaterMark = "Введите название отдела";
            this.nameTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.nameTextBox.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(302, 144);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(192, 27);
            this.SaveButton.TabIndex = 2;
            this.SaveButton.Text = "Сохранить";
            this.SaveButton.UseSelectable = true;
            this.SaveButton.Click += new System.EventHandler(this.deleteOtdelButton_Click);
            // 
            // shortNameTextBox
            // 
            // 
            // 
            // 
            this.shortNameTextBox.CustomButton.Image = null;
            this.shortNameTextBox.CustomButton.Location = new System.Drawing.Point(445, 1);
            this.shortNameTextBox.CustomButton.Name = "";
            this.shortNameTextBox.CustomButton.Size = new System.Drawing.Size(25, 25);
            this.shortNameTextBox.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.shortNameTextBox.CustomButton.TabIndex = 1;
            this.shortNameTextBox.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.shortNameTextBox.CustomButton.UseSelectable = true;
            this.shortNameTextBox.CustomButton.Visible = false;
            this.shortNameTextBox.Lines = new string[0];
            this.shortNameTextBox.Location = new System.Drawing.Point(23, 111);
            this.shortNameTextBox.MaxLength = 32767;
            this.shortNameTextBox.Name = "shortNameTextBox";
            this.shortNameTextBox.PasswordChar = '\0';
            this.shortNameTextBox.PromptText = "Введите сокращенное название отдела";
            this.shortNameTextBox.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.shortNameTextBox.SelectedText = "";
            this.shortNameTextBox.SelectionLength = 0;
            this.shortNameTextBox.SelectionStart = 0;
            this.shortNameTextBox.ShortcutsEnabled = true;
            this.shortNameTextBox.Size = new System.Drawing.Size(471, 27);
            this.shortNameTextBox.TabIndex = 1;
            this.shortNameTextBox.UseSelectable = true;
            this.shortNameTextBox.WaterMark = "Введите сокращенное название отдела";
            this.shortNameTextBox.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.shortNameTextBox.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // AddEditOtdelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 207);
            this.Controls.Add(this.shortNameTextBox);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.nameTextBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AddEditOtdelForm";
            this.Resizable = false;
            this.ShadowType = MetroFramework.Forms.MetroFormShadowType.AeroShadow;
            this.Style = MetroFramework.MetroColorStyle.Brown;
            this.Text = "AddEditOtdelForm";
            this.Load += new System.EventHandler(this.AddEditOtdelForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroTextBox nameTextBox;
        private MetroFramework.Controls.MetroButton SaveButton;
        private MetroFramework.Controls.MetroTextBox shortNameTextBox;
    }
}