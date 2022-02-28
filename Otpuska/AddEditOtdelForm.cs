using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otpuska
{
    public partial class AddEditOtdelForm : MetroFramework.Forms.MetroForm
    {
        List<Otdel> otdels = new List<Otdel>();
        public AddEditOtdelForm()
        {
            InitializeComponent();
            otdels = SQLClient.ReadAllOtdels();
        }
        //случайно название такое вышло, не обессудьте
        private void deleteOtdelButton_Click(object sender, EventArgs e)
        {
            int i = 0;
            if(nameTextBox.Text == "" || shortNameTextBox.Text == "")
            {
                MetroFramework.MetroMessageBox.Show(this, "Заполните все поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            else if(Text == "Добавить отдел")
            {
                foreach(Otdel otdel in otdels)
                {
                    if(otdel.OtdelName == nameTextBox.Text)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "Отдел с таким названием уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        i = 1;
                        break;
                    }
                    else if(otdel.OtdelShortName == shortNameTextBox.Text)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "Отдел с таким сокращенным названием уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Question);

                        i = 1;
                        break;
                    }
                }
                if(i == 0)
                {
                    SQLClient.AddOtdel(nameTextBox.Text, shortNameTextBox.Text);
                    MetroFramework.MetroMessageBox.Show(this, "Отдел добавлен", "Добавление", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    this.Close();


                }
            }

        }

        private void AddEditOtdelForm_Load(object sender, EventArgs e)
        {

        }
    }
}
