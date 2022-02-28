using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Otpuska
{
    class Constant
    {
        /*Таблица Отделы*/
        public const string strOtdel = "\t" + "\t" + "\t";

        /*Таблица Сотрудники*/
        public const string ID = "\t";
        public const string FIO = "\t" + "\t";
        public const string Age = "\t" + "\t";
        public const string Dol = "\t" + "\t";
        public const string Tabel = "\t" + "\t" + "\t";
        public const string Buro_Otdel = "\t" + "\t" + "\t";
        public const string Odna_Dol = "\t" + "\t";
        public const string Dekr = "\t" + "\t";
        public const string F_Ot = "\t" + "\t";
        public const string Воен = "\t" + "\t" + "\t";
        public const string Ветеран = "\t" + "\t";
        public const string Otpusk_1 = "\t";
        public const string Otpusk_2 = "\t";
        public const string АЭС = "\t" + "\t";
        public const string Жен_2_Ребен = "\t" + "\t" + "\t";
        public const string Многодет = "\t" + "\t";


        public const int numberMassiv = 5;

        public int[] munth = { 8, 12, 11, 10, 5, 4, 1, 2, 3, 6, 9, 7 };
        public int[] monthDay = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        public const int age = 800;
        public const int dekr = 800;
        public const int f_ot = 800;
        public const int воен = 800;
        public const int ветеран_1 = 800;
        public const int ветеран_2 = 1600;
        public const int аэс = 800;
        public const int жен_2_ребен = 800;
        public const int многодет = 800;

        public const int countMasivRow = 14;
        public const int countMassivColumn = 3;

        public static float[] month_koeffs = new float[12];//Текущие значения заполненности по месяцам(строка год:2021)
        public static float[] max_month_koeffs = new float[12];//Максимальные значения запоненности по месяцам (строка 1год:12021)
        /* Коэффициенты отдыхающего в месяц на 1 день всего 14
         *  8 - январь      (01)    [0] 
         *  12 - февраль    (02)    [1] 
         *  11 - март       (03)    [2]
         *  10 - апрель     (04)    [3] 
         *  5 - май         (05)    [4] 
         *  4 - июнь        (06)    [5] 
         *  1 - июль        (07)    [6] 
         *  2 - август      (08)    [7] 
         *  3 - сентябрь    (09)    [8] 
         *  6 - октябрь     (10)    [9] 
         *  9 - ноябрь      (11)    [10] 
         *  7 - декабрь     (12)    [11]
         *  
         * Возраст < 18 + 800 к коеффициенту
         * 
         * Женщина по беременности + 800 к коэффициенту
         * 
         * Муж когда жена беремена и в отпуске + 800 к коэффициенту
         * 
         * Супруг(а) военослужащего когда тот в отпуске + 800 к коэфициенту
         * 
         * Ветеран труда - 1 + 800
         * Ветеран труда и военных действий - 2 + 1600
         * 
         * АЭС + 200 к коэфициенту
         * 
         * Женщина имеющая от 2 детей меньше 12 лет + 800 коэффициента
         * 
         * Многодетная семья(от 3 детей) + 800 коэффициент
         */

    }
}
