using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Otpuska
{
    static class SQLClient
    {
        private static SqlConnection sqlConnection; //Подключение базы SQL
        private static List<string[]> strList = new List<string[]>(); //Данные сотрудников для работы с SQL 
        private static String connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Project\Отпуска\С ТАБЕЛЕМ\Otpusk_git\Otpuska\Otpuska\OtpuskBase.mdf;
                                                                           Integrated Security=True;Connect Timeout=30";

        static public int Save(Pearson pearson)
        {
            int test = 0;
            SqlDataReader sqlReader = null;
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            SqlCommand command = new SqlCommand("SELECT * FROM [Sotr] WHERE [Табельный_номер] = @Табельный_номер", sqlConnection);
            command.Parameters.AddWithValue("Табельный_номер", pearson.TableNum);
            sqlReader = command.ExecuteReader();
            while (sqlReader.Read())
            {
                strList.Add(new String[1]);
                strList[strList.Count - 1][0] = Convert.ToString(sqlReader["Табельный_номер"]);
                test = Convert.ToInt32(strList[strList.Count - 1][0] = Convert.ToString(sqlReader["Табельный_номер"]));
            }
            SqlCommand cmd;

            if (test == 0)
            {
                cmd = new SqlCommand("INSERT INTO [Sotr] (ФИО, Возраст, Должность, Табельный_номер, Отдел, Доп_табельный_номер, Декрет, Жена_Отпуск," +
                "Жена_Муж_военнослужащие, Ветеран_боевых_дейcтвий, Ветеран_труда,  Ликвидатор, ребенка_2_меньше_12_лет, Многодетные, Доп_Дни, Отпуск, Не_отгул_за_пГод, Коэффициент)" +
                "VALUES(@ФИО, @Возраст, @Должность, @Табельный_номер, @Отдел, @Доп_табельный_номер, @Декрет, @Жена_Отпуск, @Жена_Муж_военнослужащие," +
                "@Ветеран_боевых_дейcтвий, @Ветеран_труда, @Ликвидатор, @ребенка_2_меньше_12_лет, @Многодетные, @Доп_Дни, @Отпуск, @Не_отгул_за_пГод, @Коэффициент)");
            }
            else
            {
                cmd = new SqlCommand("UPDATE [Sotr] SET [ФИО] = @ФИО, [Возраст] = @Возраст, [Должность] = @Должность, [Отдел] = @Отдел," +
                " [Доп_табельный_номер] = @Доп_табельный_номер, [Декрет] = @Декрет, [Жена_Отпуск] = @Жена_Отпуск, [Жена_Муж_военнослужащие] = @Жена_Муж_военнослужащие," +
                "[Ветеран_боевых_дейcтвий] = @Ветеран_боевых_дейcтвий, [Ветеран_труда] = @Ветеран_труда ,[Ликвидатор] = @Ликвидатор," +
                "[ребенка_2_меньше_12_лет] = @ребенка_2_меньше_12_лет, [Многодетные] = @Многодетные, [Доп_Дни] = @Доп_Дни, [Отпуск] = @Отпуск," +
                "[Не_отгул_за_пГод] = @Не_отгул_за_пГод, [Коэффициент] = @Коэффициент WHERE [Табельный_номер] = @Табельный_номер");
            }
            EditToDB(cmd, pearson);
            return 0;
        }

        static private void EditToDB(SqlCommand command, Pearson pearson)
        {
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            
            //SqlCommand command = new SqlCommand("UPDATE [Sotr] SET [ФИО] = @ФИО, [Возраст] = @Возраст, [Должность] = @Должность, [Отдел] = @Отдел," +
            //    " [Доп_табельный_номер] = @Доп_табельный_номер, [Декрет] = @Декрет, [Жена_Отпуск] = @Жена_Отпуск, [Жена_Муж_военнослужащие] = @Жена_Муж_военнослужащие," +
            //    "[Ветеран_боевых_дейтвий] = @Ветеран_боевых_дейтвий, [Ветеран_труда] = @Ветеран_труда ,[Ликвидатор] = @Ликвидатор," +
            //    "[ребенка_2_меньше_12_лет] = @ребенка_2_меньше_12_лет, [Многодетные] = @Многодетные, [Доп_Дни] = @Доп_Дни, [Отпуск] = @Отпуск," +
            //    "[Не_отгул_за_пГод] = @Не_отгул_за_пГод, [Коэффициент] = @Коэффициент WHERE [Табельный_номер] = @Табельный_номер", sqlConnection);
            command.Connection = sqlConnection;

            command.Parameters.AddWithValue("ФИО", pearson.FIO);
            command.Parameters.AddWithValue("Возраст", pearson.Age);
            command.Parameters.AddWithValue("Должность", pearson.Proffession);
            command.Parameters.AddWithValue("Табельный_номер", pearson.TableNum);
            command.Parameters.AddWithValue("Отдел", pearson.Otdel);
            string str = "";
            foreach(string tmp in pearson.AdditionalPearsonId)
            {
                str += tmp + ",";
            }
            str = str.Remove(str.Length - 1);
            command.Parameters.AddWithValue("Доп_табельный_номер", str);
            command.Parameters.AddWithValue("Декрет", pearson.Dikret);
            command.Parameters.AddWithValue("Жена_Отпуск", pearson.Zhena_otpusk);
            command.Parameters.AddWithValue("Жена_Муж_военнослужащие", pearson.Zhena_much_voenn);
            command.Parameters.AddWithValue("Ветеран_боевых_дейcтвий", pearson.Veteran);
            command.Parameters.AddWithValue("Ветеран_труда", pearson.Veteran);
            command.Parameters.AddWithValue("Ликвидатор", pearson.Likvidator);
            command.Parameters.AddWithValue("ребенка_2_меньше_12_лет", pearson.Zhena_2detei_menee12let);
            command.Parameters.AddWithValue("Многодетные", pearson.Mnogodet);
            command.Parameters.AddWithValue("Доп_Дни", pearson.DopDni);
            if (pearson.Vacation != null)
            {
                String vacationDays = String.Join(",", pearson.Vacation.ToArray()); // Преобразование отпуска в формат БД
                command.Parameters.AddWithValue("Отпуск", vacationDays);
            }
            else
            {
                command.Parameters.AddWithValue("Отпуск", "");
            }
            command.Parameters.AddWithValue("Не_отгул_за_пГод", pearson.PrevYearDays);
            command.Parameters.AddWithValue("Коэффициент", pearson.Koeff);

            command.ExecuteNonQuery();
        }

        static public List<Pearson> ReadAllFromDB()
        {
            List<Pearson> result = new List<Pearson>();
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            SqlCommand cmd = new SqlCommand("SELECT * FROM [Sotr]", sqlConnection);

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    Pearson pearson = ParseStr(rdr);
                    result.Add(pearson);
                }
            }
            return result;
        }

        static public Pearson ReadFromDB(string TableNum)
        {
            Pearson result = new Pearson();
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            String str = "SELECT * FROM [Sotr] WHERE Табельный_номер = '" + TableNum+"'";

            SqlCommand cmd = new SqlCommand(str, sqlConnection);
            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    result = ParseStr(rdr);
                }
            }

            return result;
        }

        static public Pearson ReadFromDB(int id)
        {
            Pearson result = new Pearson();
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            String str = "SELECT * FROM [Sotr] WHERE id = " + id.ToString();

            SqlCommand cmd = new SqlCommand(str, sqlConnection);
            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    result = ParseStr(rdr);
                }
            }

            return result;
        }

        static private Pearson ParseStr(SqlDataReader rdr)
        {
            Pearson pearson = new Pearson();
            pearson.FIO = rdr["ФИО"] as String;
            pearson.Id = (int)rdr["Id"];

            if (rdr["Возраст"].GetType() != typeof(DBNull))
            {
                pearson.Age = Convert.ToInt32(rdr["Возраст"]);
            }

            pearson.Proffession = rdr["Должность"] as String;

            if (rdr["Табельный_номер"].GetType() != typeof(DBNull))
            {
                pearson.TableNum = rdr["Табельный_номер"] as String;
            }

            pearson.Otdel = rdr["Отдел"] as String;
            pearson.AdditionalPearsonId = new List<string>();
            if (rdr["Доп_табельный_номер"].GetType() != typeof(DBNull))
            {
                string tmp = rdr["Доп_табельный_номер"] as String;
                string[] array = tmp.Split(',');
                foreach (string str in array)
                {
                    pearson.AdditionalPearsonId.Add(str);
                }
            }

            if (rdr["Декрет"].GetType() != typeof(DBNull))
            {
                pearson.Dikret = Convert.ToInt32(rdr["Декрет"]);
            }

            if (rdr["Жена_Отпуск"].GetType() != typeof(DBNull))
            {
                pearson.Zhena_otpusk = Convert.ToInt32(rdr["Жена_Отпуск"]);
            }

            if (rdr["Жена_Муж_военнослужащие"].GetType() != typeof(DBNull))
            {
                pearson.Zhena_much_voenn = Convert.ToInt32(rdr["Жена_Муж_военнослужащие"]);
            }

            if (rdr["Ветеран_боевых_дейcтвий"].GetType() != typeof(DBNull))
            {
                pearson.Veteran = Convert.ToInt32(rdr["Ветеран_боевых_дейcтвий"]);
            }

            if (rdr["Ликвидатор"].GetType() != typeof(DBNull))
            {
                pearson.Likvidator = Convert.ToInt32(rdr["Ликвидатор"]);
            }

            if (rdr["ребенка_2_меньше_12_лет"].GetType() != typeof(DBNull))
            {
                pearson.Zhena_2detei_menee12let = Convert.ToInt32(rdr["ребенка_2_меньше_12_лет"]);
            }

            if (rdr["Многодетные"].GetType() != typeof(DBNull))
            {
                pearson.Mnogodet = Convert.ToInt32(rdr["Многодетные"]);
            }

            if (rdr["Доп_Дни"].GetType() != typeof(DBNull))
            {
                pearson.DopDni = Convert.ToInt32(rdr["Доп_Дни"]);
            }

            if (rdr["Не_отгул_за_пГод"].GetType() != typeof(DBNull))
            {
                pearson.PrevYearDays = Convert.ToInt32(rdr["Не_отгул_за_пГод"]);
            }

            if (rdr["Коэффициент"].GetType() != typeof(DBNull))
            {
                pearson.Koeff = Convert.ToInt32(rdr["Коэффициент"]);
            }

            if (rdr["Отпуск"].GetType() != typeof(DBNull))
            {
                String str = rdr["Отпуск"] as String;
                String[] a = str.Split(',');
                foreach (String t in a)
                {
                    try
                    {
                        pearson.Vacation.Add(Convert.ToDateTime(t));
                    }
                    catch (Exception e)
                    {

                    }
                }
            }

            return pearson;
        }

        static public void EditAfterExcelImport(Pearson pearson, string type)
        {
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            if (type == "dates")
            {
                SqlCommand command = new SqlCommand("UPDATE [Sotr] SET [Не_отгул_за_пГод] = @Не_отгул_за_пГод, [Коэффициент] = @Коэффициент WHERE [Табельный_номер] = @Табельный_номер");
                command.Connection = sqlConnection;
                command.Parameters.AddWithValue("Табельный_номер", pearson.TableNum);
                command.Parameters.AddWithValue("Не_отгул_за_пГод", pearson.PrevYearDays);
                command.Parameters.AddWithValue("Коэффициент", pearson.Koeff);
                command.ExecuteNonQuery();
            }
            else if (type == "worker")
            {
                SqlCommand command = new SqlCommand("INSERT INTO[Sotr](ФИО, Должность, Табельный_номер ) VALUES(@ФИО, @Должность, @Табельный_номер)");
                command.Connection = sqlConnection;
                command.Parameters.AddWithValue("Табельный_номер", pearson.TableNum);
                command.Parameters.AddWithValue("ФИО", pearson.FIO);
                command.Parameters.AddWithValue("Должность", pearson.Proffession);
                command.ExecuteNonQuery();
            }
        }

        static public void ReadKoeffs(int year)//Загрузить коэффиценты по месяцам
        {
            String str = "SELECT * FROM [Koeffs] WHERE Year = " + year.ToString();
            SqlCommand cmd = new SqlCommand(str, sqlConnection);

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    for (int i = 0; i < 12; i++)
                    {
                        Constant.month_koeffs[i] = Convert.ToSingle(rdr[i]);
                    }
                }
            }

            year += 10000;
            str = "SELECT * FROM [Koeffs] WHERE Year = " + year.ToString();
            cmd = new SqlCommand(str, sqlConnection);

            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    for (int i = 0; i < 12; i++)
                    {
                        Constant.max_month_koeffs[i] = Convert.ToSingle(rdr[i]);
                    }
                }
            }

        }

        static public void WriteKoeffs(int year)
        {
            //SqlCommand command = new SqlCommand("UPDATE [Sotr] SET [ФИО] = @ФИО, [Возраст] = @Возраст, [Должность] = @Должность, [Отдел] = @Отдел," +
            //    " [Доп_табельный_номер] = @Доп_табельный_номер, [Декрет] = @Декрет, [Жена_Отпуск] = @Жена_Отпуск, [Жена_Муж_военнослужащие] = @Жена_Муж_военнослужащие," +
            //    "[Ветеран_боевых_дейтвий] = @Ветеран_боевых_дейтвий, [Ветеран_труда] = @Ветеран_труда ,[Ликвидатор] = @Ликвидатор," +
            //    "[ребенка_2_меньше_12_лет] = @ребенка_2_меньше_12_лет, [Многодетные] = @Многодетные, [Доп_Дни] = @Доп_Дни, [Отпуск] = @Отпуск," +
            //    "[Не_отгул_за_пГод] = @Не_отгул_за_пГод, [Коэффициент] = @Коэффициент WHERE [Табельный_номер] = @Табельный_номер", sqlConnection);

            String str = "UPDATE [Koeffs] SET [Январь] = @Январь, [Февраль] = @Февраль, [Март] = @Март, [Апрель] = @Апрель, [Май] = @Май,[Июнь] = @Июнь, [Июль] = @Июль,[Август] = @Август,[Сентябрь] = @Сентябрь,[Октябрь] = @Октябрь,[Ноябрь] = @Ноябрь,[Декабрь] = @Декабрь WHERE [Year] = @Year";
            SqlCommand cmd = new SqlCommand(str, sqlConnection);

            cmd.Parameters.AddWithValue("Январь",Constant.month_koeffs[0]);
            cmd.Parameters.AddWithValue("Февраль", Constant.month_koeffs[1]);
            cmd.Parameters.AddWithValue("Март", Constant.month_koeffs[2]);
            cmd.Parameters.AddWithValue("Апрель", Constant.month_koeffs[3]);
            cmd.Parameters.AddWithValue("Май", Constant.month_koeffs[4]);
            cmd.Parameters.AddWithValue("Июнь", Constant.month_koeffs[5]);
            cmd.Parameters.AddWithValue("Июль", Constant.month_koeffs[6]);
            cmd.Parameters.AddWithValue("Август", Constant.month_koeffs[7]);
            cmd.Parameters.AddWithValue("Сентябрь", Constant.month_koeffs[8]);
            cmd.Parameters.AddWithValue("Октябрь", Constant.month_koeffs[9]);
            cmd.Parameters.AddWithValue("Ноябрь", Constant.month_koeffs[10]);
            cmd.Parameters.AddWithValue("Декабрь", Constant.month_koeffs[11]);
            cmd.Parameters.AddWithValue("Year", year.ToString());

            cmd.ExecuteNonQuery();
        }

        static public List<Otdel> ReadAllOtdels()
        {
            List<Otdel> result = new List<Otdel>();
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            SqlCommand cmd = new SqlCommand("SELECT * FROM [Otdels]", sqlConnection);

            SqlDataReader sqlReader = cmd.ExecuteReader();
            while (sqlReader.Read())
            {
                Otdel otdel = new Otdel();
                otdel.Id = Convert.ToInt32(sqlReader["Id"]);
                otdel.OtdelName = Convert.ToString(sqlReader["Отдел"]);
                otdel.OtdelShortName = Convert.ToString(sqlReader["СокрНазв"]);
                result.Add(otdel);
            }
            sqlReader.Close();
            sqlConnection.Close();

            return result;
        }

        static public void AddOtdel(string name, string short_name)
        {
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();

            SqlCommand command = new SqlCommand("INSERT INTO [Otdels] (Отдел,Связь,СокрНазв) VALUES (@otdel_name,@link,@short_name)");
            command.Connection = sqlConnection;
            command.Parameters.AddWithValue("otdel_name", name);
            command.Parameters.AddWithValue("link", "0");
            command.Parameters.AddWithValue("short_name", short_name);
            command.ExecuteNonQuery();
        }

        //Удаление отдела
        static public void deleteOtdel(string name)
        {
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM [Otdels] WHERE [Отдел] = @Отдел");
            command.Connection = sqlConnection;
            command.Parameters.AddWithValue("Отдел", name);
            command.ExecuteNonQuery();
            
        }

        static public void deletePerson(string name)
        {
            sqlConnection = new SqlConnection(connectionString); //Объект соединения с базой данных
            sqlConnection.Open();
            SqlCommand command = new SqlCommand("DELETE FROM [Sotr] WHERE [Табельный_номер] = @Табельный_номер");
            command.Connection = sqlConnection;
            command.Parameters.AddWithValue("Табельный_номер", name);
            command.ExecuteNonQuery();
        }

    }
}
