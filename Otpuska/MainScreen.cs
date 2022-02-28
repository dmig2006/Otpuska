using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Otpuska
{
    public partial class MainScreen : MetroFramework.Forms.MetroForm
    {
        float pearson_day_cost;//Вес одного дня человека
        Pearson pearson;
        List<Pearson> ludi;

        enum VacationChooseStatus
        {
            PrevYDays,//Неотгулянные дни за прошлый год
            days_14,//2 недельный обязательный отпуск
            other_days,//Оставшиеся дни (14) + доп дни (могут быть выбраны в любое время)
            incorrect
        }

        struct Days
        {
            public int PrevYDays;
            public int Days14;//Обязательные 2 недели подряд
            public int otherDays;//оставшиеся 2 недели + доп дни
        }

        VacationChooseStatus status = VacationChooseStatus.incorrect;
        Days days = new Days();

        public MainScreen()
        {
            InitializeComponent();

            comboBox1.SelectedIndex = 0;
        }

        public static List<DateTime> getHolidaysList(string year)
        {
            XmlDocument xmlDoc = new XmlDocument();
            String filePath = year + ".xml";
            string sRequestData = filePath;
            xmlDoc.Load(filePath);
            List<DateTime> HolidaysList = new List<DateTime>();
            XmlNodeList nodeList = xmlDoc.GetElementsByTagName("day");
            foreach (XmlNode xmlnode in nodeList)
            {
                if (xmlnode.Attributes["t"].InnerText == "1")
                {
                    string[] helper;
                    helper = xmlnode.Attributes["d"].InnerText.Split('.');
                    //string a = helper[1] + "." + helper[0] + "." + year;
                    HolidaysList.Add(Convert.ToDateTime(helper[1] + "." + helper[0] + "." + year));
                }
            }
            return HolidaysList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime dt1 = new DateTime(Int32.Parse((comboBox1.SelectedItem as String)), 1, 1);
            DateTime dt2 = new DateTime(Int32.Parse((comboBox1.SelectedItem as String)), 12, 31);
            monthCalendar1.MinDate = new DateTime(2000, 1, 1);
            monthCalendar1.MaxDate = new DateTime(3000, 12, 20);
            monthCalendar1.MaxDate = dt2;
            monthCalendar1.MinDate = dt1;

            monthCalendar1.Holidays = getHolidaysList(comboBox1.SelectedItem as String);
            monthCalendar1.Repaint();

            ludi = SQLClient.ReadAllFromDB();
            foreach (Pearson p in ludi)
            {
                comboBox2.Items.Add(p.FIO+"|"+p.TableNum);
            }

            SQLClient.ReadKoeffs(Convert.ToInt32(comboBox1.SelectedItem)); 
            for(int i = 0;i<12;i++)
            {
                if(Constant.month_koeffs[i] > Constant.max_month_koeffs[i])
                {
                    int tmp = 0;
                    while(tmp<DateTime.DaysInMonth(Convert.ToInt32(comboBox1.SelectedItem),i+1))
                    {
                        DateTime t = new DateTime(Convert.ToInt32(comboBox1.SelectedItem), i + 1, tmp+1);
                        monthCalendar1.ClosedDays.Add(t);
                        tmp++;
                    }
                }
            }

            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String[] a = (comboBox2.SelectedItem as String).Split('|');
            pearson = SQLClient.ReadFromDB(a[1]);
            foreach (string str in pearson.AdditionalPearsonId)
            {
                if (str != "0")
                {
                    List<DateTime> tmp = SQLClient.ReadFromDB(str).Vacation;
                    monthCalendar1.AnotherIdDays.AddRange(tmp);
                }
            }

            //if (AdditionalPearson.Vacation != null)
            //{
            //    monthCalendar1.AnotherIdDays = AdditionalPearson.Vacation;
            //}
            //else
            //{
            //    monthCalendar1.AnotherIdDays = new List<DateTime>();
            //}

            textBox1.Text = pearson.ShowPearsonMsg();

            days.PrevYDays = pearson.PrevYearDays;
            days.Days14 = 14;
            days.otherDays = 14 + pearson.DopDni;

            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;

            monthCalendar1.Invalidate();
            monthCalendar1.Focus();

            Point pt = new Point(50, 50);
            Cursor.Position = pt;

            pearson_day_cost = (100.0f / comboBox2.Items.Count) / (28 + pearson.DopDni);
            if (pearson.Vacation != null)
            {
                foreach (DateTime dt in pearson.Vacation)
                {
                    Constant.month_koeffs[dt.Month - 1] -= pearson_day_cost;
                }
            }
            pearson.Vacation = new List<DateTime>();
        }

        private void monthCalendar1_MouseUp(object sender, MouseEventArgs e)
        {
            if(status == VacationChooseStatus.incorrect)
            {
                MessageBox.Show("Выберите тип выбираемых дней");
                return;
            }

            MonthCalendar.HitTestInfo hInfo = monthCalendar1.HitTest(e.Location);
            if(!(hInfo.HitArea == MonthCalendar.HitArea.Date))
            {
                return;
            }
            if (!monthCalendar1.Vacation.Contains(hInfo.Time))
            {
                if (monthCalendar1.AnotherIdDays.Contains(hInfo.Time)|| monthCalendar1.ClosedDays.Contains(hInfo.Time))
                {
                    MessageBox.Show("Вы не можете выбрать этот день");
                }
                else
                {

                    if (status == VacationChooseStatus.PrevYDays)
                    {
                        if (monthCalendar1.Holidays.Contains(hInfo.Time) || (hInfo.Time.DayOfWeek == DayOfWeek.Saturday)|| (hInfo.Time.DayOfWeek == DayOfWeek.Sunday))
                        {
                            MessageBox.Show("Выбран праздничный или выходной день");
                            return;
                        }
                        else
                        {
                            days.PrevYDays--;
                        }
                        label1.Text = "Осталось выбрать дней:" + days.PrevYDays.ToString();

                        monthCalendar1.Vacation.Add(hInfo.Time);
                        pearson.Vacation.Add(hInfo.Time);
                        //Constant.month_koeffs[hInfo.Time.Month-1] += pearson_day_cost;//TODO: Уточнить, учитывать ли обязательные дни отпуска

                        if (days.PrevYDays <= 0)
                        {
                            MessageBox.Show("Обязательные дни отпуска(неотгулянные за прошлый год) выбраны");
                            status = VacationChooseStatus.incorrect;
                            button3.Enabled = false;
                        }
                        return;
                    }

                    if(status == VacationChooseStatus.days_14)
                    {
                        //if (monthCalendar1.Holidays.Contains(hInfo.Time))
                        //{
                        //    MessageBox.Show("Выбран праздничный день");
                        //    return;
                        //}
                        //else
                        //{
                        //    days.Days14--;
                        //}
                        //label1.Text = "Осталось выбрать дней:" + days.Days14.ToString();
                        int counter = 0;
                        while (days.Days14 != 0)
                        {
                            DateTime dt = hInfo.Time.AddDays(counter);
                            if(!monthCalendar1.Holidays.Contains(dt))
                            {
                                days.Days14--;
                                Constant.month_koeffs[dt.Month - 1] += pearson_day_cost;
                            }
                            monthCalendar1.Vacation.Add(dt);
                            pearson.Vacation.Add(dt);
                            counter++;
                        }
                        MessageBox.Show("Выбор обязательных 14 дней подряд окончен");
                        label1.Text = "Осталось выбрать дней: 0";
                        status = VacationChooseStatus.incorrect;
                        button4.Enabled = false;
                        return;
                    }

                    if (status == VacationChooseStatus.other_days)
                    {
                        if (!monthCalendar1.Holidays.Contains(hInfo.Time))
                        {
                            days.otherDays--;
                        }
                        else
                        {
                            monthCalendar1.Holidays.Remove(hInfo.Time);
                        }
                        label1.Text = "Осталось выбрать дней:" + days.otherDays.ToString();

                        monthCalendar1.Vacation.Add(hInfo.Time);
                        pearson.Vacation.Add(hInfo.Time);
                        Constant.month_koeffs[hInfo.Time.Month - 1] += pearson_day_cost;

                        if (days.otherDays <= 0)
                        {
                            MessageBox.Show("Выбор оставшихся дней окончен");
                            status = VacationChooseStatus.incorrect;
                            button5.Enabled = false;
                        }
                        return;
                    }
                }
            }
            else if (monthCalendar1.Vacation.Contains(hInfo.Time))
            {
                if(status == VacationChooseStatus.days_14)
                {
                    return;
                }
                monthCalendar1.Vacation.Remove(hInfo.Time);
                pearson.Vacation.Remove(hInfo.Time);
                switch(status)
                {
                    case VacationChooseStatus.other_days:
                        days.otherDays++;
                        Constant.month_koeffs[hInfo.Time.Month - 1] -= pearson_day_cost;
                        label1.Text = "Осталось выбрать дней:" + days.otherDays.ToString();
                        break;
                    case VacationChooseStatus.PrevYDays:
                        days.PrevYDays++;
                        label1.Text = "Осталось выбрать дней:" + days.PrevYDays.ToString();
                        break;
                }
                MessageBox.Show("Выбраный день исключен из отпуска");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (days.PrevYDays != 0)
            {
                status = VacationChooseStatus.PrevYDays;
                label1.Text = "Осталось выбрать дней:" + days.PrevYDays.ToString();
            }
            else
            {
                MessageBox.Show("Выбор обязательных дней отпуска окончен");
                button3.Enabled = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            status = VacationChooseStatus.days_14;
            label1.Text = "Осталось выбрать дней:" + days.Days14.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            status = VacationChooseStatus.other_days;
            label1.Text = "Осталось выбрать дней:"+days.otherDays.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            String str = "";
            foreach(DateTime dt in pearson.Vacation)
            {
                str += dt.ToString("dd.MM.yyyy") + Environment.NewLine;
            }
            str += "Число невыбранных дней:" + (days.Days14 + days.otherDays).ToString();
            if(MessageBox.Show(str,"Запись в БД",MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                SQLClient.WriteKoeffs(Convert.ToInt32(comboBox1.Text));
                pearson.PrevYearDays = days.Days14 + days.otherDays;
                pearson.SaveToDB();
                this.Close();
            }

        }
    }
}
