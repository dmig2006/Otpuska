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
    public partial class ModifyPearsonScreen : MetroFramework.Forms.MetroForm
    {
        Pearson pearson = new Pearson();

        public ModifyPearsonScreen()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
            {
                pearson = SQLClient.ReadFromDB(Convert.ToInt32(textBox8.Text));
            }
            else
            {
                pearson = SQLClient.ReadFromDB(textBox2.Text);
            }
            textBox1.Text = pearson.FIO;
            textBox2.Text = pearson.TableNum;
            comboBox1.Text = pearson.Otdel;
            foreach(string str in pearson.AdditionalPearsonId)
            {
                textBox3.Text += str + ",";
            }
            textBox4.Text = pearson.Proffession;
            textBox5.Text = pearson.Age.ToString();
            textBox6.Text = pearson.PrevYearDays.ToString();
            textBox7.Text = pearson.DopDni.ToString();
            #region льготы
            if (pearson.Dikret == 1)
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }

            if (pearson.Zhena_otpusk == 1)
            {
                checkBox2.Checked = true;
            }
            else
            {
                checkBox2.Checked = false;
            }

            if (pearson.Zhena_much_voenn == 1)
            {
                checkBox3.Checked = true;
            }
            else
            {
                checkBox3.Checked = false;
            }

            if (pearson.Veteran == 2)
            {
                checkBox4.Checked = true;
                checkBox5.Checked = true;
            }
            else if (pearson.Veteran == 1)
            {
                checkBox4.Checked = true;
                checkBox5.Checked = false;
            }
            else if (pearson.Veteran == 0)
            {
                checkBox4.Checked = false;
                checkBox5.Checked = false;
            }

            if (pearson.Likvidator == 1)
            {
                checkBox6.Checked = true;
            }
            else
            {
                checkBox6.Checked = false;
            }

            if (pearson.Zhena_2detei_menee12let == 1)
            {
                checkBox7.Checked = true;
            }
            else
            {
                checkBox7.Checked = false;
            }

            if (pearson.Mnogodet == 1)
            {
                checkBox8.Checked = true;
            }
            else
            {
                checkBox8.Checked = false;
            }
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
        {
            pearson.FIO = textBox1.Text;
            pearson.TableNum = textBox2.Text;
            pearson.Otdel = comboBox1.Text;

            pearson.AdditionalPearsonId.Clear();

            string[] array = textBox3.Text.Split(',');
            foreach (string str in array)
            {
                pearson.AdditionalPearsonId.Add(str);
            }
            pearson.Proffession = textBox4.Text;
            pearson.Age = Convert.ToInt32(textBox5.Text);
            pearson.PrevYearDays = Convert.ToInt32(textBox6.Text);
            pearson.DopDni = Convert.ToInt32(textBox7.Text);

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
            else if (checkBox4.Checked || checkBox5.Checked)
            {
                pearson.Veteran = 1;
            }
            else 
            {
                pearson.Veteran = 0;
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
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
