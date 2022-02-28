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
    public partial class AddPearsonScreen : MetroFramework.Forms.MetroForm
    {
        public AddPearsonScreen()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private bool checkTxt()
        {
            bool res = true;
            foreach(Control c in Controls)
            {
                if(c is TextBox)
                {
                    if(c.Text == "")
                    {
                        res = false;
                    }
                }
            }
            return res;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Pearson pearson = new Pearson();

            if(checkTxt() == false)
            {
                MessageBox.Show("Заполнены не все поля формы");
                return;
            }

            pearson.FIO = textBox1.Text;
            pearson.TableNum = textBox2.Text;
            pearson.Otdel = comboBox1.Text;

            string[] array = textBox3.Text.Split(',');
            pearson.AdditionalPearsonId = new List<string>();
            foreach (string str in array)
            {
                pearson.AdditionalPearsonId.Add(str);
            }
            pearson.Proffession = textBox4.Text;
            pearson.Age = Int32.Parse(textBox5.Text);
            pearson.DopDni = Int32.Parse(textBox7.Text);
            pearson.PrevYearDays = Int32.Parse(textBox6.Text);
            #region льготы
            if (checkBox1.Checked)
            {
                pearson.Dikret = 1;
            }
            else
            {
                pearson.Dikret = 0;
            }

            if (checkBox2.Checked)
            {
                pearson.Zhena_otpusk = 1;
            }
            else
            {
                pearson.Zhena_otpusk = 0;
            }

            if (checkBox3.Checked)
            {
                pearson.Zhena_much_voenn = 1;
            }
            else
            {
                pearson.Zhena_much_voenn = 0;
            }

            if (checkBox4.Checked && checkBox5.Checked)
            {
                pearson.Veteran = 2;
            }
            else
            {
                if (checkBox4.Checked ^ checkBox5.Checked)
                {
                    pearson.Veteran = 1;
                }
                else
                {
                    pearson.Veteran = 0;
                }
            }

            if (checkBox6.Checked)
            {
                pearson.Likvidator = 1;
            }
            else
            {
                pearson.Likvidator = 0;
            }

            if (checkBox7.Checked)
            {
                pearson.Zhena_2detei_menee12let = 1;
            }
            else
            {
                pearson.Zhena_2detei_menee12let = 0;
            }

            if (checkBox8.Checked)
            {
                pearson.Mnogodet = 1;
            }
            else
            {
                pearson.Mnogodet = 0;
            }
            #endregion
            pearson.SaveToDB();

            MessageBox.Show(pearson.ShowPearsonMsg());

            MessageBox.Show("Сотрудник добавлен в БД");
        }


    }
}
