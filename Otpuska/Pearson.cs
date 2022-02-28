using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Otpuska
{
    class Pearson
    {
        private int id;//id сотрудника в базе данных
        private string fio;//ФИО сотрудника
        private int age;//Возраст
        private string proffession;//Профессия(должность)
        private string tableNum;//Табельный номер сотрудника
        private string otdel;//Отдел сотрудника
        private List<string> additionalPearsonId;//Табельный номер сотрудника одновременно с которым данный сотрудник не может взять отпуск
        private int dikret;//Дикрет
        private int zhena_otpusk;//Жена отпуск
        private int zhena_much_voenn;//Жена/муж военнослужащие
        private int veteran;//Ветеран военных действий или труда (0-ничего,1-одна льгота,2-обе)
        private int likvidator;//Ликвидатор
        private int zhena_2detei_menee12let;//Жена имеет двух детей меньше 12 лет
        private int mnogodet;//Многодетные
        private int dopdni;//Дополнительные дни отпуска
        private int koeff;//Коэффицент приоритета сотрудника
        private int prevYearDays;//Число неотгулянных дней за предыдущий год
        private List<DateTime> vacation = new List<DateTime>();//Числа(конкретные даты) отпуска в текущем году
        private DateTime firstWorkDay;//дата трудоустройства

        private SqlConnection sqlConnection; //Подключение базы SQL
        private List<string[]> strList = new List<string[]>(); //Данные сотрудников для работы с SQL 
        private String connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Project\Отпуска\С ТАБЕЛЕМ\Otpusk_git\Otpuska\Otpuska\OtpuskBase.mdf;
                                                                           Integrated Security=True;Connect Timeout=30";


        public int Id { get => id; set => id = value; }
        public string FIO { get => fio; set => fio = value; }
        public string TableNum { get => tableNum; set => tableNum = value; }
        public int Koeff { get => koeff; set => koeff = value; }
        public int PrevYearDays { get => prevYearDays; set => prevYearDays = value; }
        public List<DateTime> Vacation { get => vacation; set => vacation = value; }
        public string Otdel { get => otdel; set => otdel = value; }
        public List<string> AdditionalPearsonId { get => additionalPearsonId; set => additionalPearsonId = value; }
        public int Age { get => age; set => age = value; }
        public string Proffession { get => proffession; set => proffession = value; }
        public int Dikret { get => dikret; set => dikret = value; }
        public int Zhena_otpusk { get => zhena_otpusk; set => zhena_otpusk = value; }
        public int Zhena_much_voenn { get => zhena_much_voenn; set => zhena_much_voenn = value; }
        public int Veteran { get => veteran; set => veteran = value; }
        public int Likvidator { get => likvidator; set => likvidator = value; }
        public int Zhena_2detei_menee12let { get => zhena_2detei_menee12let; set => zhena_2detei_menee12let = value; }
        public int Mnogodet { get => mnogodet; set => mnogodet = value; }
        public int DopDni { get => dopdni; set => dopdni = value; }
        public DateTime FirstWorkDay { get => firstWorkDay; set => firstWorkDay = value; }

        public void SaveToDB()
        {
            //SQLClient.connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\d.t.gilmutdinov\Documents\Visual Studio 2017\Projects\Otpusk_git\Otpuska\Otpuska\OtpuskBase.mdf;
            //                                                               Integrated Security=True;Connect Timeout=30";
            SQLClient.Save(this);
        }

        public Pearson LoadFromDB(String TableNum)
        {
            return SQLClient.ReadFromDB(TableNum);
        }

        public String ShowPearsonMsg()
        {
            String msg = "";
            msg += "ФИО:" + FIO + Environment.NewLine;
            msg += "Табельный:" + TableNum + Environment.NewLine;
            string str = "";
            foreach(string tmp in AdditionalPearsonId)
            {
                str += tmp + ",";
            }
            str = str.Remove(str.Length - 1);
            msg += "Дополнительный табельный:" + str + Environment.NewLine;
            //msg += "Возраст:" + Age + Environment.NewLine;
            msg += "Должность:" + Proffession + Environment.NewLine;
            //msg += "Отдел:" + Otdel + Environment.NewLine;
            //msg += "Дикрет:" + Dikret + Environment.NewLine;
            //msg += "Жена отпуск:" + Zhena_otpusk + Environment.NewLine;
            //msg += "Жена/муж военные:" + Zhena_much_voenn + Environment.NewLine;
            //msg += "Ветеран:" + Veteran + Environment.NewLine;
            //msg += "Ликвидатор:" + Likvidator + Environment.NewLine;
            //msg += "Имеет 2 детей младше 12 лет:" + Zhena_2detei_menee12let + Environment.NewLine;
            //msg += "Многодетные:" + Mnogodet + Environment.NewLine;
            msg += "Льготы:" + Environment.NewLine;

            if (Dikret != 0)
            {
                msg += "Дикрет" + Environment.NewLine;
            }

            if (Zhena_otpusk != 0)
            {
                msg += "Отпуск жены" + Environment.NewLine;
            }

            if (Zhena_much_voenn != 0)
            {
                msg += "Жена/муж военные" + Environment.NewLine;
            }

            if (Veteran != 0)
            {
                msg += "Ветеран" + Environment.NewLine;
            }

            if (Likvidator != 0)
            {
                msg += "Ликвидатор" + Environment.NewLine;
            }

            if (Zhena_2detei_menee12let != 0)
            {
                msg += "2 детей младше 12 лет" + Environment.NewLine;
            }

            if (Mnogodet != 0)
            {
                msg += "Многодетный" + Environment.NewLine;
            }

            msg += "Дополнительные дни отпуска:" + DopDni + Environment.NewLine;
            msg += "Неотгулянные дни отпуска за прошлый год:" + PrevYearDays + Environment.NewLine;

            return msg;
        }

        public int koefficient()
        {
            List<string[]> strOtpusk = new List<string[]>();
            int koef = 0;
            List<int> intKoef = new List<int>(); //Коэффициенты   
            List<int[]> intOtpusk = new List<int[]>(); //Льготы
            List<DateTime> listOtpusk = new List<DateTime>(); //Otpuska
            String[] words;
            int number = 0;
            Constant cons = new Constant();

            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            SqlCommand command = new SqlCommand("SELECT Коэффициент, Отпуск, Возраст, Декрет, Жена_Отпуск, Жена_Муж_военнослужащие, Ветеран_боевых_дейcтвий, " +
                            "Ветеран_труда, Ликвидатор, ребенка_2_меньше_12_лет, Многодетные FROM [Sotr] WHERE [Табельный_номер] = @Табельный_номер", sqlConnection);

            command.Parameters.AddWithValue("Табельный_номер", tableNum);

            SqlDataReader sqlReader = command.ExecuteReader();
            while (sqlReader.Read())
            {
                strOtpusk.Add(new string[11]);
                intKoef.Add(new int());
                intOtpusk.Add(new int[11]);

                intOtpusk[intOtpusk.Count - 1][2] = Convert.ToInt32(sqlReader["Возраст"]);
                intOtpusk[intOtpusk.Count - 1][3] = Convert.ToInt32(sqlReader["Декрет"]);
                intOtpusk[intOtpusk.Count - 1][4] = Convert.ToInt32(sqlReader["Жена_Отпуск"]);
                intOtpusk[intOtpusk.Count - 1][5] = Convert.ToInt32(sqlReader["Жена_Муж_военнослужащие"]);
                intOtpusk[intOtpusk.Count - 1][6] = Convert.ToInt32(sqlReader["Ветеран_боевых_дейcтвий"]);
                intOtpusk[intOtpusk.Count - 1][7] = Convert.ToInt32(sqlReader["Ветеран_труда"]);
                intOtpusk[intOtpusk.Count - 1][8] = Convert.ToInt32(sqlReader["Ликвидатор"]);
                intOtpusk[intOtpusk.Count - 1][9] = Convert.ToInt32(sqlReader["ребенка_2_меньше_12_лет"]);
                intOtpusk[intOtpusk.Count - 1][10] = Convert.ToInt32(sqlReader["Многодетные"]);

                //strOtpusk[strOtpusk.Count - 1][2] = Convert.ToString(sqlReader["Возраст"]);
                //if (strOtpusk[number][2] == null) intOtpusk[number][2] = 1;

                //strOtpusk[strOtpusk.Count - 1][3] = Convert.ToString(sqlReader["Декрет"]);
                //if (strOtpusk[number][3] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][4] = Convert.ToString(sqlReader["Жена_Отпуск"]);
                //if (strOtpusk[number][4] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][5] = Convert.ToString(sqlReader["Жена_Муж_военнослужащие"]);
                //strOtpusk[strOtpusk.Count - 1][6] = Convert.ToString(sqlReader["Ветеран_боевых_дейcтвий"]);
                //if (strOtpusk[number][6] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][7] = Convert.ToString(sqlReader["Ветеран_труда"]);
                //if (strOtpusk[number][7] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][8] = Convert.ToString(sqlReader["Ликвидатор"]);
                //if (strOtpusk[number][8] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][9] = Convert.ToString(sqlReader["ребенка_2_меньше_12_лет"]);
                //if (strOtpusk[number][9] == null) intOtpusk[number][2] = 1;
                //strOtpusk[strOtpusk.Count - 1][10] = Convert.ToString(sqlReader["Многодетные"]);
                //if (strOtpusk[number][10] == null) intOtpusk[number][2] = 1;


                /*Коэффициент по отпуску*/
                koef = raschetOtpuska(vacation);

                /*Возраст меньше 18*/
                if (intOtpusk[number][2] < 18)
                    intKoef[number] = intKoef[number] + Constant.age;

                /*Женщина по беременности*/
                if (intOtpusk[number][3] == 1)
                    intKoef[number] = intKoef[number] + Constant.dekr;

                /*Муж, когда жена в отпуске и беремена*/
                if (intOtpusk[number][4] == 1)
                    intKoef[number] = intKoef[number] + Constant.f_ot;

                /*Супруг(а) когда жена(муж) на военной службе в отпуске*/
                if (intOtpusk[number][5] == 1)
                    intKoef[number] = intKoef[number] + Constant.воен;

                /*Ветеран труда или вооруженных сил*/
                if (intOtpusk[number][6] == 1)
                    intKoef[number] = intKoef[number] + Constant.ветеран_1;
                else if (intOtpusk[number][6] == 2)
                    intKoef[number] = intKoef[number] + Constant.ветеран_2;

                /*АЭС*/
                if (intOtpusk[number][7] == 1)
                    intKoef[number] = intKoef[number] + Constant.аэс;

                /*Женщина, у которой 2 и больше ребенка меньше 12 лет*/
                if (intOtpusk[number][8] == 1)
                    intKoef[number] = intKoef[number] + Constant.жен_2_ребен;

                /*Многодетные семьи*/
                if (intOtpusk[number][9] == 1)
                    intKoef[number] = intKoef[number] + Constant.многодет;

                koef = koef + intKoef[number];
                //strList[number][16] = Convert.ToString(intKoef[number]);
                number++;
            }
            number = 0;
            sqlReader.Close();
            sqlConnection.Close();
            return koef;
        }

        //Функция рассчета коэфициента по месецам
        private int raschetOtpuska(List<DateTime> listDate)
        {
            Constant cons = new Constant();
            int k = 0;

            for (int i = 0; i < listDate.Count; i++)
            {
                if (listDate[i].Month == 1)
                {
                    k = k + cons.munth[0];
                }
                else if (listDate[i].Month == 2)
                {
                    k = k + cons.munth[1];
                }
                else if (listDate[i].Month == 3)
                {
                    k = k + cons.munth[2];
                }
                else if (listDate[i].Month == 4)
                {
                    k = k + cons.munth[3];
                }
                else if (listDate[i].Month == 5)
                {
                    k = k + cons.munth[4];
                }
                else if (listDate[i].Month == 6)
                {
                    k = k + cons.munth[5];
                }
                else if (listDate[i].Month == 7)
                {
                    k = k + cons.munth[6];
                }
                else if (listDate[i].Month == 8)
                {
                    k = k + cons.munth[7];
                }
                else if (listDate[i].Month == 9)
                {
                    k = k + cons.munth[8];
                }
                else if (listDate[i].Month == 10)
                {
                    k = k + cons.munth[9];
                }
                else if (listDate[i].Month == 11)
                {
                    k = k + cons.munth[10];
                }
                else if (listDate[i].Month == 12)
                {
                    k = k + cons.munth[11];
                }
            }

            return k;
        }


    }
}
