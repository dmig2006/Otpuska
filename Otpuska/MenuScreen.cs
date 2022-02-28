using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Xml;
using System.Threading;


namespace Otpuska
{
    public partial class MenuScreen : MetroFramework.Forms.MetroForm
    {
        MainScreen screen;
        AddPearsonScreen addPearsonScreen;
        ModifyPearsonScreen modifyPearsonScreen;
        AddEditOtdelForm addOtdelScreen;

        List<Otdel> otdels = new List<Otdel>();
        List<Pearson> pearsonsList = new List<Pearson>();

        //WaitForm ctf = new WaitForm();

        public MenuScreen()
        {
            InitializeComponent();
            pearsonsList = SQLClient.ReadAllFromDB();
            otdels = SQLClient.ReadAllOtdels();
            foreach (Otdel otdel in otdels)
            {
                otdelsListView.Items.Add(otdel.OtdelName);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addOtdelScreen = new AddEditOtdelForm();
            addOtdelScreen.Text = "Добавить отдел";
            addOtdelScreen.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            addPearsonScreen = new AddPearsonScreen();
            addPearsonScreen.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            screen = new MainScreen();
            screen.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            modifyPearsonScreen = new ModifyPearsonScreen();
            modifyPearsonScreen.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "C:\\Users\\v.k.koscheev\\Desktop\\отпуска2019";
                openFileDialog.Filter = "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Создаём приложение.
                    Excel.Application ObjExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook;
                    try
                    {
                        //Открываем книгу.                                                                                                                                                        
                        ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog.FileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    }
                    catch
                    {
                        return;
                    }

                    //Выбираем таблицу(лист).
                    Excel.Worksheet ObjWorkSheet;
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];


                    int i = 28;
                    string tabNumValue;
                    string sFIO;
                    string[] sHolidaysCount;
                    DateTime firstHolidayDate;
                    Dictionary<string, Pearson> pearsonDictionary = new Dictionary<string, Pearson>();
                    //Dictionary<string, int> pearsonDictionary = new Dictionary<string, int>();
                    //List<Pearson> workersList = new List<Pearson>();
                    List<Pearson> workersList = SQLClient.ReadAllFromDB();

                    while (ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text != "")
                    {
                        if (ObjWorkSheet.get_Range("I" + i.ToString(), "I" + i.ToString()).Text != "РАСЧЕТ" && workersList.Exists(x => x.TableNum == ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text) == true)
                        {
                            //ФИО сотрудника
                            sFIO = ObjWorkSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Text;
                            //получение табельного номера
                            tabNumValue = ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text;

                            //получение должности
                            //sWork = ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text;

                            //расчёт количества дней отпуска
                            sHolidaysCount = ObjWorkSheet.get_Range("F" + i.ToString(), "F" + i.ToString()).Text.Split('+');
                            int iHolidaysCount = 0;
                            foreach (var u in sHolidaysCount)
                                iHolidaysCount += Int32.Parse(u);

                            //оставшиеся с прошлого года неотгулянные дни и дополнительные дни отпуска
                            int ostDni = workersList.Find(x => x.TableNum == tabNumValue).PrevYearDays + workersList.Find(x => x.TableNum == tabNumValue).DopDni;
                            DateTime firstWorkDay = workersList.Find(x => x.TableNum == tabNumValue).FirstWorkDay;

                            //получение даты начала отпуска
                            try
                            {
                                firstHolidayDate = (ObjWorkSheet.get_Range("H" + i.ToString(), "H" + i.ToString()).Text != "  .  .") ? Convert.ToDateTime(ObjWorkSheet.get_Range("H" + i.ToString(), "H" + i.ToString()).Text) : Convert.ToDateTime(ObjWorkSheet.get_Range("G" + i.ToString(), "G" + i.ToString()).Text);
                            }
                            catch
                            {
                                MessageBox.Show("Не удаётся считать дату начала отпуска у " + sFIO, "Ошибка", MessageBoxButtons.OK);
                                return;
                            }

                            //если устроился в этом году и ходил в отпуск, то было 14 дней, если устроился ранее, то 28
                            if (firstHolidayDate.Year == firstWorkDay.Year)
                            {
                                ostDni += 14;
                            }
                            else
                            {
                                ostDni += 28;
                            }

                            //проверка по табельному номеру есть ли сотрудник в pearsonDictionary
                            if (!pearsonDictionary.ContainsKey(tabNumValue))
                            {
                                Pearson pearson = new Pearson();
                                pearson.FIO = sFIO;
                                pearson.TableNum = tabNumValue;
                                pearson.PrevYearDays = ostDni - iHolidaysCount;

                                //запись всех дней отпуска в pearson.vacation
                                pearson.Vacation = new List<DateTime>();
                                for (int day = 0; day < iHolidaysCount; day++)
                                {
                                    pearson.Vacation.Add(firstHolidayDate.AddDays(day));
                                }
                                //добавление класса сотрудника в массив, где ключ - его табельный номер
                                pearsonDictionary.Add(tabNumValue, pearson);
                            }
                            else
                            {
                                pearsonDictionary[tabNumValue].PrevYearDays -= iHolidaysCount;

                                //запись всех дней отпуска в pearson.vacation
                                for (int day = 0; day < iHolidaysCount; day++)
                                {
                                    pearsonDictionary[tabNumValue].Vacation.Add(firstHolidayDate.AddDays(day));
                                }
                                //pearsonDictionary[tabNumValue] += iHolidaysCount;
                            }
                        }
                        //count++;
                        i++;
                    }

                    //вычисляем коэффициент для каждого сотрудника из pearsonDictionary и записываем его и количество неотгулянных дней в базу
                    foreach (KeyValuePair<string, Pearson> pearson in pearsonDictionary)
                    {
                        pearson.Value.Koeff = pearson.Value.koefficient();
                        if (pearson.Value.PrevYearDays < 0)
                        {
                            pearson.Value.PrevYearDays = 0;
                        }

                        SQLClient.EditAfterExcelImport(pearson.Value,"dates");
                    }
                    //выход из экселя
                    ObjExcel.Quit();
                    MessageBox.Show("Импорт данных об отпусках и расчет коэффицентов окончен");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "C:\\Users\\v.k.koscheev\\Desktop\\отпуска2019";
                openFileDialog.Filter = "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Создаём приложение.
                    Excel.Application ObjExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook;
                    try
                    {
                        //Открываем книгу.                                                                                                                                                        
                        ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog.FileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    }
                    catch
                    {
                        return;
                    }

                    //Выбираем таблицу(лист).
                    Excel.Worksheet ObjWorkSheet;
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    int i = 28;
                    Dictionary<string, Pearson> pearsonDictionary = new Dictionary<string, Pearson>();
                    List<Pearson> workersList = SQLClient.ReadAllFromDB();

                    while (ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text != "")
                    {
                        if (!pearsonDictionary.ContainsKey(ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text) && workersList.Exists(x => x.TableNum == ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text) == false)
                        {
                            Pearson pearson = new Pearson();
                            pearson.FIO = ObjWorkSheet.get_Range("D" + i.ToString(), "D" + i.ToString()).Text;
                            pearson.TableNum = ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text;
                            pearson.Proffession = ObjWorkSheet.get_Range("C" + i.ToString(), "C" + i.ToString()).Text;
                            pearsonDictionary.Add(ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString()).Text, pearson);
                            //pearson.SaveToDB();
                            //SQLClient.EditAfterExcelImport(pearson, "worker");
                            pearson.Age = 99;
                            pearson.Dikret = 0;
                            pearson.FirstWorkDay = new DateTime(2000, 1, 1);
                            pearson.Likvidator = 0;
                            pearson.Otdel = "";
                            pearson.PrevYearDays = 0;
                            pearson.Zhena_otpusk = 0;
                            pearson.Zhena_much_voenn = 0;
                            pearson.Zhena_2detei_menee12let = 0;
                            pearson.Veteran = 0;
                            pearson.Mnogodet = 0;
                            pearson.AdditionalPearsonId.Add("0");
                            pearson.SaveToDB();
                        }
                        i++;
                    }
                    //выход из экселя
                    ObjExcel.Quit();
                }
                MessageBox.Show("Импорт сотрудников завершён", "Уведомление", MessageBoxButtons.OK);
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application ObjExcel = new Excel.Application();
                    ObjExcel.SheetsInNewWorkbook = 1;
                    Excel.Workbook workBook = ObjExcel.Workbooks.Add(Type.Missing);

                    Excel.Worksheet workSheet = workBook.Worksheets.Add();

                    //Excel.Worksheet workSheet = ObjExcel.Worksheets.Add();

                    //workSheet.Cells[24, 1] = "№";
                    //workSheet.Cells[24, 2] = "Структурное подразделение";
                    //workSheet.Cells[24, 3] = "Должность(по штатному расписанию)";
                    //workSheet.Cells[24, 4] = "Фамилия, имя, отчество";
                    //workSheet.Cells[24, 5] = "Табельный номер";
                    //workSheet.Cells[24, 6] = "Отпуск";

                    Excel.Range excelCell;

                    excelCell = workSheet.Cells[1, 11];
                    excelCell.Value2 = "Положение 4";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 21.71;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[2, 11], (Excel.Range)workSheet.Cells[2, 12]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "к приказу от ________ №____";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.WrapText = false;

                    excelCell = workSheet.Cells[4, 9];
                    excelCell.Value2 = "Унифицированная форма № Т-7";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 36.71;

                    excelCell = workSheet.Cells[5, 9];
                    excelCell.Value2 = "Утверждена постановлением";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[6, 9];
                    excelCell.Value2 = "Госкомстата России от 05.01.04  №1";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[8, 10];
                    excelCell.Value2 = "Форма по ОКУД";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 19.86;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[9, 10];
                    excelCell.Value2 = "по ОКПО";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[8, 11];
                    excelCell.Value2 = "Код";
                    excelCell.Font.Size = 10;
                    excelCell.ColumnWidth = 21.71;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[9, 11];
                    excelCell.Value2 = "0301020";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[9, 11];
                    excelCell.Value2 = " ";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[10, 2], (Excel.Range)workSheet.Cells[10, 6]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Акционерное общество "+'"'+"Концерн "+'"'+"Калашников"+'"';
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.Cells.Style.WrapText = true;
                    excelCell.Font.Bold = true;
                    excelCell.Font.Italic = true;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[11, 2], (Excel.Range)workSheet.Cells[11, 6]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "наименование организации";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Style.WrapText = true;
                    excelCell.Font.Bold = false;
                    excelCell.Font.Italic = false;

                    excelCell = workSheet.Cells[12, 8];
                    excelCell.Value2 = "УТВЕРЖДАЮ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 21.29;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[12, 9], (Excel.Range)workSheet.Cells[13, 11]);
                    excelCell.Merge(Type.Missing);
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[14, 10];
                    excelCell.Value2 = "должность";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[13, 1], (Excel.Range)workSheet.Cells[13, 4]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Мнение выборного профсоюзного органа   УТВЕРЖДАЮ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.Cells.Style.WrapText = true;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[14, 1], (Excel.Range)workSheet.Cells[14, 4]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "от  "+'"'+"_____ "+'"'+" ________________ 2018 г.  №             учтено";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.Cells.Style.WrapText = true;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[16, 8], (Excel.Range)workSheet.Cells[16, 9]);
                    excelCell.Merge(Type.Missing);
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[17, 8], (Excel.Range)workSheet.Cells[17, 9]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "личная подпись";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Style.WrapText = true;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[16, 11];
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[17, 11];
                    excelCell.Value2 = "расшифровка подписи";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[21, 3];
                    excelCell.Value2 = "ГРАФИК ОТПУСКОВ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    excelCell.Font.Bold = true;

                    excelCell = workSheet.Cells[20, 4];
                    excelCell.Value2 = "Номер \nдокумента";
                    excelCell.Font.Size = 7;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Font.Bold = false;

                    excelCell = workSheet.Cells[20, 5];
                    excelCell.Value2 = "Дата \nсоставления";
                    excelCell.Font.Size = 7;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[20, 6];
                    excelCell.Value2 = "На год";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[21, 4];
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[21, 5];
                    excelCell.Value2 = DateTime.Now.ToShortDateString();
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[21, 6];
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    for(int i = 1; i<=5; i++)
                    {
                        excelCell = workSheet.Cells[25, i];
                        //excelCell.Value2 = " ";
                        excelCell.Font.Size = 10;
                        excelCell.Font.Name = "Arial";
                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        excelCell = workSheet.Cells[26, i];
                        //excelCell.Value2 = " ";
                        excelCell.Font.Size = 10;
                        excelCell.Font.Name = "Arial";
                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }

                    excelCell = workSheet.Cells[25, 6];
                    excelCell.Value2 = "Количество календ. дней осн+вр + кд + многодет";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[26, 6];
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[25, 7], (Excel.Range)workSheet.Cells[25, 8]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Дата";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.Cells.Style.WrapText = true;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[26, 7];
                    excelCell.Value2 = "Запланированная";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[26, 8];
                    excelCell.Value2 = "Фактическая";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[25, 9], (Excel.Range)workSheet.Cells[25, 10]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Перенесения отпуска";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    //excelCell.Cells.Style.WrapText = true;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[26, 9];
                    excelCell.Value2 = "Основание \nдокумент";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[26, 10];
                    excelCell.Value2 = "Дата предполагаемого отпуска";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[25, 11];
                    excelCell.Value2 = "Подпись";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[26, 11];
                    excelCell.Value2 = "Ознакомлен, уведомлен о дате начала отпуска, согласен с разделением отпуска на части";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    excelCell = workSheet.Cells[27, 1];
                    //excelCell.Value2 = " ";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    for(int i = 1; i <= 10; i++)
                    {
                        excelCell = workSheet.Cells[27, i+1];
                        excelCell.Value2 = i;
                        excelCell.Font.Size = 10;
                        excelCell.Font.Name = "Arial";
                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }

                    excelCell = workSheet.Cells[24, 1];
                    excelCell.Value2 = "№";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 5.29;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCell.RowHeight = 63.75;

                    excelCell = workSheet.Cells[24, 2];
                    excelCell.Value2 = "Структурное подразделение";
                    excelCell.Font.Size = 10;
                    excelCell.ColumnWidth = 5.29;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[24, 3];
                    excelCell.Value2 = "Должность (специальность, профессия) по штатному расписанию";
                    excelCell.ColumnWidth = 36.57;
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[24, 4];
                    excelCell.Value2 = "Фамилия, имя, отчество";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.ColumnWidth = 35.14;
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[24, 5];
                    excelCell.Value2 = "Табельный номер";
                    excelCell.ColumnWidth = 13;
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[24, 6], (Excel.Range)workSheet.Cells[24, 10]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Отпуск";
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[24, 11];
                    //excelCell.Value2 = " ";
                    excelCell.ColumnWidth = 21.71;
                    excelCell.Font.Size = 10;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                    List<DateTime> holidays = MainScreen.getHolidaysList("2021");
                    List<Pearson> pearsons = SQLClient.ReadAllFromDB();
                    int currentLine = 28;
                    int start = currentLine;
                    for (int i = 0; i < pearsons.Count; i++)
                    {
                        //workSheet.Cells[i + 2, 1] = pearsons[i].TableNum;
                        //workSheet.Cells[i + 2, 2] = pearsons[i].FIO;
                        //workSheet.Cells[i + 2, 3] = pearsons[i].Proffession;
                        if (pearsons[i].Vacation != null)
                        {
                            List<DateTime> lst = pearsons[i].Vacation;
                            lst.Sort();
                            if (lst.Count > 1)
                            {
                                int t = 1;
                                DateTime tmp_start = lst[0];
                                for (int j = 1; j < lst.Count; j++)
                                {
                                    if (lst[j].Subtract(lst[j - 1]) > TimeSpan.FromDays(1.0f))
                                    {
                                        start++;

                                        excelCell = workSheet.Cells[currentLine, 1];
                                        excelCell.Value2 = currentLine - 27;
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 5.29;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                                        excelCell = workSheet.Cells[currentLine, 2];
                                        excelCell.Value2 = 840;
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 6.29;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                        excelCell = workSheet.Cells[currentLine, 3];
                                        excelCell.Value2 = pearsons[i].Proffession;
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 36.57;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        //excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                        excelCell = workSheet.Cells[currentLine, 4];
                                        excelCell.Value2 = pearsons[i].FIO;
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 35.14;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        //excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                        excelCell = workSheet.Cells[currentLine, 5];
                                        excelCell.NumberFormat = "@";
                                        excelCell.Value2 = (string)pearsons[i].TableNum;
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 13;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                        excelCell = workSheet.Cells[currentLine, 6];
                                        excelCell.Value2 = t.ToString() + "+0+0+0";
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 10.14;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                        excelCell = workSheet.Cells[currentLine, 7];
                                        excelCell.Value2 = tmp_start.ToString("dd.MM.yyyy");
                                        excelCell.Font.Size = 11;
                                        excelCell.Font.Name = "Arial";
                                        excelCell.ColumnWidth = 11.14;
                                        excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                        excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                        for(int s = 8; s <= 11; s++)
                                        {
                                            excelCell = workSheet.Cells[currentLine, s];
                                            //excelCell.Value2 = " ";
                                            excelCell.Font.Size = 11;
                                            excelCell.Font.Name = "Arial";
                                            //excelCell.ColumnWidth = 35.14;
                                            excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        }

                                        //workSheet.Cells[currentLine, 3] = pearsons[i].Proffession;
                                        //workSheet.Cells[currentLine, 4] = pearsons[i].FIO;
                                        //workSheet.Cells[currentLine, 5] = pearsons[i].TableNum;
                                        //workSheet.Cells[currentLine, 6] = t.ToString()+"+0+0+0";
                                        //workSheet.Cells[currentLine, 7] = tmp_start.ToString("dd.MM.yyyy");

                                        tmp_start = lst[j];
                                        t = 1;
                                        currentLine++;
                                    }
                                    else
                                    {
                                        if (!holidays.Contains(lst[j]))
                                        {
                                            t++;
                                        }
                                    }
                                }

                                start++;

                                excelCell = workSheet.Cells[currentLine, 1];
                                excelCell.Value2 = currentLine - 27;
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 5.29;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                                excelCell = workSheet.Cells[currentLine, 2];
                                excelCell.Value2 = 840;
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 6.29;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                                excelCell = workSheet.Cells[currentLine, 3];
                                excelCell.Value2 = pearsons[i].Proffession;
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 36.57;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                //excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                //excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                excelCell = workSheet.Cells[currentLine, 4];
                                excelCell.Value2 = pearsons[i].FIO;
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 35.14;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                               // excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                //excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                excelCell = workSheet.Cells[currentLine, 5];
                                excelCell.NumberFormat = "@";
                                excelCell.Value2 = pearsons[i].TableNum;
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 13;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                excelCell = workSheet.Cells[currentLine, 6];
                                excelCell.Value2 = t.ToString() + "+0+0+0";
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 10.14;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                excelCell = workSheet.Cells[currentLine, 7];
                                excelCell.Value2 = tmp_start.ToString("dd.MM.yyyy");
                                excelCell.Font.Size = 11;
                                excelCell.Font.Name = "Arial";
                                excelCell.ColumnWidth = 11.14;
                                excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                excelCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                                for (int s = 8; s <= 11; s++)
                                {
                                    excelCell = workSheet.Cells[currentLine, s];
                                    //excelCell.Value2 = " ";
                                    excelCell.Font.Size = 11;
                                    excelCell.Font.Name = "Arial";
                                    //excelCell.ColumnWidth = 35.14;
                                    excelCell.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                }

                                //workSheet.Cells[currentLine, 3] = pearsons[i].Proffession;
                                //workSheet.Cells[currentLine, 4] = pearsons[i].FIO;
                                //workSheet.Cells[currentLine, 5] = pearsons[i].TableNum;
                                //workSheet.Cells[currentLine, 6] = t.ToString() + "+0+0+0";
                                //workSheet.Cells[currentLine, 7] = tmp_start.ToString("dd.MM.yyyy");
                                currentLine++;
                            }
                        }
                    }

                    //Pearson.Count

                    //int start = pearsons.Count + currentLine;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[(start + 3), 2], (Excel.Range)workSheet.Cells[(start + 3), 3]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Руководитель направления:";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[(start + 3), 5];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[(start + 3), 6];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 3, 8], (Excel.Range)workSheet.Cells[start + 3, 11]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 5, 2], (Excel.Range)workSheet.Cells[start + 5, 4]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Конструкторский центр";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 6, 2], (Excel.Range)workSheet.Cells[start + 6, 4]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Руководитель структурного подразделения:";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[start + 6, 5];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[start + 6, 6];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 6, 8], (Excel.Range)workSheet.Cells[start + 6, 11]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 9, 2], (Excel.Range)workSheet.Cells[start + 9, 3]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Согласовано:";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 11, 2], (Excel.Range)workSheet.Cells[start + 11, 3]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Начальник управления компенсации и льгот";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[start + 11, 5];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[start + 11, 6];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 11, 8], (Excel.Range)workSheet.Cells[start + 11, 11]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 13, 2], (Excel.Range)workSheet.Cells[start + 13, 3]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Начальник отдела кадрового";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[start + 13, 5];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[start + 13, 6];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 13, 8], (Excel.Range)workSheet.Cells[start + 13, 11]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 15, 2], (Excel.Range)workSheet.Cells[start + 15, 3]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Value2 = "Начальник управления оплаты труда и кадрового документооборота";
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";

                    excelCell = workSheet.Cells[start + 15, 5];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = workSheet.Cells[start + 15, 6];
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    excelCell = (Excel.Range)workSheet.get_Range((Excel.Range)workSheet.Cells[start + 15, 8], (Excel.Range)workSheet.Cells[start + 15, 11]);
                    excelCell.Merge(Type.Missing);
                    excelCell.Font.Size = 11;
                    excelCell.Font.Name = "Arial";
                    excelCell.Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;


                    //workSheet.Columns.AutoFit();
                    //workBook.SaveAs(saveFileDialog.FileName);

                    workSheet.SaveAs(saveFileDialog.FileName);
                    ObjExcel.Quit();

                    MessageBox.Show("Экспорт завершен");

                }
            }
        }

        private void createTable_Click(object sender, EventArgs e)
        {
            
            int inputYear = 0;
            try
            {
                inputYear = Convert.ToInt32(inputYearTextBox.Text);
                //int inputHumans = Convert.ToInt32(inputHumansTextBox.Text);

            }
            catch (FormatException)
            {
                MetroFramework.MetroMessageBox.Show(this, "Введите корректные данные!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Question);
                //MessageBox.Show("Введите корректные данные!", "Предупреждение");
            }
            if (inputYear != 0)
            {
                XmlDocument xmlDoc = new XmlDocument();
                string sRequestData = "http://xmlcalendar.ru/data/ru/" + inputYear + "/calendar.xml";
                try
                {
                    xmlDoc.Load(sRequestData);
                }
                catch (System.Net.WebException)
                {
                    MetroFramework.MetroMessageBox.Show(this, "На сервере нет данных о праздничных днях этого года. Таблицу создать не удалось.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    //MessageBox.Show("На сервере нет данных о праздничных днях этого года. Таблицу создать не удалось.", "Предупреждение");
                }
                if(xmlDoc.ChildNodes.Count != 0)
                {
                    //CancellationTokenSource tokenSource = new CancellationTokenSource();
                    //CancellationToken token = tokenSource.Token;
                    //await Task.Run(() => ctf.ShowDialog(), token);
                    //await task;

                    //WaitForm ctf = new WaitForm();
                    //ctf.ShowDialog();

                    string folderpath = "";
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    DialogResult dr = fbd.ShowDialog();
                    int numberOtdInList = 1;

                    if (dr == DialogResult.OK)
                    {
                        folderpath = fbd.SelectedPath;
                    }

                    if(folderpath == "")
                    {
                        MetroFramework.MetroMessageBox.Show(this, "Необходимо выбрать папку для сохранения файлов с табелями", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    }
                    else
                    {
                        //MessageBox.Show(folderpath, "Message");

                        Thread thread = new Thread(openWindow);
                        thread.Start();

                        foreach (Otdel otdel in otdels)
                        {
                            createExcel(inputYear, otdel, xmlDoc,numberOtdInList, folderpath);
                            numberOtdInList++;
                        }
                        //ctf.Close();
                        //tokenSource.Cancel();
                        thread.Abort();
                    }
                    
                }
            }
 
        }

        private void openWindow()
        {
            WaitForm ctf = new WaitForm();
            ctf.ShowDialog();
        }

        //private void createExcel(Object input)
        private void createExcel(int inputYear, Otdel otd, XmlDocument xmlDoc, int numberOtdInList, string folderpath)
        {
            Excel.Application excelapp;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Range excelcells;

            // Otdel tOtdel = new Otdel();
            //tOtdel = (Otdel)input;
            //int inputHumans = 2;

            List<Pearson> otdelsPearson = new List<Pearson>();
            foreach(Pearson p in pearsonsList)
            {
                if(p.Otdel == otd.Id.ToString())
                {
                    otdelsPearson.Add(p);
                }
            }

            int chiefOtdIndex = otdelsPearson.FindIndex(x => x.Proffession == "начальник отдела");
            int chiefBuroIndex = otdelsPearson.FindIndex(x => x.Proffession == "начальник бюро");
            if (chiefOtdIndex != -1)
            {
                //otdelsPearson.Add(otdelsPearson[0]);
                otdelsPearson.Insert(0, otdelsPearson[chiefOtdIndex]);
                otdelsPearson.RemoveAt(chiefOtdIndex + 1);
            }
            else if(chiefBuroIndex != -1)
            {
                //otdelsPearson.Add(otdelsPearson[0]);
                otdelsPearson.Insert(0, otdelsPearson[chiefBuroIndex]);
                otdelsPearson.RemoveAt(chiefBuroIndex + 1);
            }


            int inputHumans = otdelsPearson.Count();

            string[] sMonths = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
            string[] sDays = { "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс" };
            string sDayCounter, sMonthsCounter;
            string adder;
            int hoursInMonth, daysInMonth;
            int dayInWeek;
            Dictionary<string, string> HolidaysData = new Dictionary<string, string>();
            List<string> HolidaysList = new List<string>();
           // Excel.Range formulaRange;

            //Получаем все узлы, имеющие имя day
            XmlNodeList nodeList = xmlDoc.GetElementsByTagName("day");
            //Выводим значения атрибутов d и t у всех найденных узлов day 
            foreach (XmlNode xmlnode in nodeList)
            {
                HolidaysData.Add(xmlnode.Attributes["d"].InnerText, xmlnode.Attributes["t"].InnerText);
            }

            excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 12;
            excelapp.Workbooks.Add(Type.Missing);
            

            for (int month = 1; month <= 12; month++)
            {
                daysInMonth = 0;
                hoursInMonth = 0;
                HolidaysList.Clear();
                //записываем месяца в название листов
                excelapp.Sheets[month].Name = sMonths[month - 1];

                excelsheets = excelapp.Worksheets;
                //Получаем ссылку на лист 1
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(month);


                // Выделяем диапазон ячеек А1-А3, записываем номер
                Excel.Range c1 = excelworksheet.Cells[1, 1];
                Excel.Range c2 = excelworksheet.Cells[3, 1];
                Excel.Range excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "№";
                excelCells1.ColumnWidth = 5.86;
                excelCells1.Font.Size = 11;
                excelCells1.Font.Name = "Calibri";
                excelCells1.EntireRow.Font.Bold = true;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, 2];
                c2 = excelworksheet.Cells[3, 2];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "ФИО";
                excelCells1.EntireRow.Font.Bold = true;
                excelCells1.ColumnWidth = 36;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //statika
                //excelworksheet.Activate();
                //excelworksheet.Application.ActiveWindow.SplitRow = 3;
                //excelworksheet.Application.ActiveWindow.SplitColumn = 2;
                //excelworksheet.Application.ActiveWindow.FreezePanes = true;


                for (int date = 1; date <= 31; date++)
                {

                    dayInWeek = dateCheck(inputYear, month, date);
                    string sDateForCompare;
                    if (dayInWeek != 0)
                    {

                        if ((int)Math.Log10(month) + 1 == 2)
                        {
                            sDateForCompare = month.ToString() + ".";
                        }
                        else
                        {
                            sDateForCompare = "0" + month.ToString() + ".";
                        }

                        if ((int)Math.Log10(date) + 1 == 2)
                        {
                            sDateForCompare = sDateForCompare + date.ToString();
                        }
                        else
                        {
                            sDateForCompare = sDateForCompare + "0" + date.ToString();
                        }

                        int dayVariable = 0;
                        foreach (KeyValuePair<string, string> kvp in HolidaysData)
                        {
                            if (kvp.Key == sDateForCompare)
                            {
                                if (kvp.Value == "1")
                                {
                                    dayVariable = 1;//выходной
                                    break;
                                }
                                else if (kvp.Value == "2")
                                {
                                    dayVariable = 2;//предпраздничный
                                    break;
                                }
                                else
                                {
                                    dayVariable = 0;//рабочий
                                    break;
                                }

                            }
                            else
                            {
                                if (dayInWeek == 6 || dayInWeek == 7)
                                {
                                    dayVariable = 1;//выходной
                                }
                                else
                                {
                                    dayVariable = 0;//рабочий
                                }
                            }
                        }
                        switch (dayVariable)
                        {
                            case 0:
                                daysInMonth = 2 + date;
                                excelCells1 = (Excel.Range)excelworksheet.Cells[3, daysInMonth];
                                excelCells1.Value2 = date;
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                //DateTime dateValue = new DateTime(iInputYear, month, date);
                                excelCells1 = (Excel.Range)excelworksheet.Cells[2, daysInMonth];
                                excelCells1.Value2 = sDays[dayInWeek - 1];
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                hoursInMonth = hoursInMonth + 8;
                                break;
                            case 1:
                                daysInMonth = 2 + date;
                                excelCells1 = (Excel.Range)excelworksheet.Cells[3, daysInMonth];
                                excelCells1.Value2 = date;
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                excelCells1.Interior.Color = 12632256; //серый
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                //DateTime dateValue = new DateTime(iInputYear, month, date);
                                excelCells1 = (Excel.Range)excelworksheet.Cells[2, daysInMonth];
                                excelCells1.Value2 = sDays[dayInWeek - 1];
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                // excelCells1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                                excelCells1.Interior.Color = 12632256;
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                HolidaysList.Add(date.ToString());
                                break;
                            case 2:
                                daysInMonth = 2 + date;
                                excelCells1 = (Excel.Range)excelworksheet.Cells[3, daysInMonth];
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Value2 = date;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                //DateTime dateValue = new DateTime(iInputYear, month, date);
                                excelCells1 = (Excel.Range)excelworksheet.Cells[2, daysInMonth];
                                excelCells1.Value2 = sDays[dayInWeek - 1];
                                excelCells1.ColumnWidth = 5.29;
                                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                hoursInMonth = hoursInMonth + 7;
                                break;
                        }


                    }
                    else
                    {
                        //date = 32;
                        break;
                    }
                }
                c1 = excelworksheet.Cells[1, 3];
                c2 = excelworksheet.Cells[1, daysInMonth];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = sMonths[month - 1] + ' ' + inputYear;
                excelCells1.Font.Size = 16;
                excelCells1.Font.Name = "Arial Black";
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //DateTime myDate = DateTime.ParseExact(hoursInMonth.ToString() + ":00", "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
                excelCells1 = (Excel.Range)excelworksheet.Cells[1, daysInMonth + 1];
                //string time = hoursInMonth.ToString() + ":00";
                //double hours = TimeSpan.Parse(time).TotalHours;
                //excelCells1.Value2 = hours;
                excelCells1.Value2 = hoursInMonth.ToString() + ":00";
                excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                excelCells1.Font.Size = 14;
                excelCells1.Font.Name = "Calibri";
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, daysInMonth + 2];
                c2 = excelworksheet.Cells[3, daysInMonth + 2];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "Отработано";
                excelCells1.ColumnWidth = 14.86;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, daysInMonth + 3];
                c2 = excelworksheet.Cells[3, daysInMonth + 3];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "План";
                excelCells1.ColumnWidth = 14.86;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, daysInMonth + 4];
                c2 = excelworksheet.Cells[3, daysInMonth + 4];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "Выходные";
                excelCells1.ColumnWidth = 14.86;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, daysInMonth + 5];
                c2 = excelworksheet.Cells[3, daysInMonth + 5];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "Переработка";
                excelCells1.ColumnWidth = 14.86;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[2, daysInMonth + 1];
                c2 = excelworksheet.Cells[inputHumans + 3, daysInMonth + 1];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек

                c1 = excelworksheet.Cells[inputHumans + 5, 3];
                c2 = excelworksheet.Cells[inputHumans + 5, 15];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                //excelCells1.Cells.Style.WrapText = true;
                excelCells1.Value2 = "А - административный отпуск, Б - больничный, К - командировка, З - Отгул";
                excelCells1.Cells.Style.WrapText = true;
                excelCells1.EntireRow.Font.Bold = false;
                //excelCells1.ColumnWidth = 14.86;
                //excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                //excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //excelCells1.Orientation = 90;
                //excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                for (int i = 1; i <= inputHumans; i++)
                {
                    //номер
                    excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, 1];
                    excelCells1.Value2 = i;
                    excelCells1.Font.Size = 11;
                    excelCells1.Font.Name = "Calibri";
                    excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    if (i % 2 != 0)
                    {
                        excelCells1.Interior.Color = 13434828;//зелёный

                        excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, 2];
                        excelCells1.Interior.Color = 13434828;
                        excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelCells1.Value2 = otdelsPearson[i - 1].FIO;
                    }
                    else
                    {
                        excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, 2];
                        excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        excelCells1.Value2 = otdelsPearson[i - 1].FIO;
                    }

                    for (int y = 1; y <= (daysInMonth - 2); y++)
                    {
                        if (HolidaysList.IndexOf(y.ToString()) == -1)
                        {
                            if (i % 2 != 0)
                            {
                                excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, y + 2];
                                excelCells1.Interior.Color = 13434828; // зелёный
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            }
                            else
                            {
                                excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, y + 2];
                                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            }
                        }
                        else
                        {
                            excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, y + 2];
                            excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            excelCells1.Interior.Color = 12632256; //серый
                            excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            //excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            //excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }

                    //расчет общих часов
                    //c1 = excelworksheet.Cells[3 + i, 3];
                    // c2 = excelworksheet.Cells[3 + i, daysInMonth];
                    // formulaRange = (Excel.Range)excelworksheet.get_Range(c1, c2);
                    //adder = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, daysInMonth + 2];
                    excelCells1.Formula = String.Format("=IF(SUM(" + excelCellNumber(3, 3 + i) + ":" + excelCellNumber(daysInMonth, 3 + i) + ")<=0," + '"' + "0:00" + '"' + "," + "SUM(" + excelCellNumber(3, 3 + i) + ":" + excelCellNumber(daysInMonth, 3 + i) + "))");
                    //excelCells1.Formula = String.Format("=SUMM({0}", adder);

                    excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                    excelCells1.Font.Size = 11;
                    excelCells1.Font.Name = "Calibri";
                    excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    if (i % 2 != 0)
                    {
                        excelCells1.Interior.Color = 13434828;
                    }

                    //расчет выходных дней
                    // c1 = excelworksheet.Cells[3 + i, 2 + Convert.ToInt32(HolidaysList[0])];
                    string holidaysCellsSum = "";
                    for (int y = 0; y < HolidaysList.Count(); y++)
                    {
                        //c2 = excelworksheet.Cells[3 + i, 2 + Convert.ToInt32(HolidaysList[y])];
                        //c1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                        if (holidaysCellsSum != "") { holidaysCellsSum += "+"; }
                        holidaysCellsSum += excelCellNumber(Convert.ToInt32(HolidaysList[y]) + 2, 3 + i);
                    }
                    //c2 = excelworksheet.Cells[3 + i, 2 + Convert.ToInt32(HolidaysList.Count()-1)];
                    //formulaRange = (Excel.Range)excelworksheet.get_Range(c1, c2);
                    // adder = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, daysInMonth + 4];
                    // excelCells1.Formula = String.Format("=СУММ({0}", adder);
                    excelCells1.Formula = String.Format("=IF(" + holidaysCellsSum + "=0," + '"' + '"' + "," + holidaysCellsSum + ")");
                    excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                    excelCells1.Font.Size = 11;
                    excelCells1.Font.Name = "Calibri";
                    excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelCells1.Interior.Color = 12632256; //серый


                    //мес.норма
                    excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, daysInMonth + 3];
                    excelCells1.Value2 = hoursInMonth.ToString() + ":00";
                    excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                    excelCells1.Font.Size = 11;
                    excelCells1.Font.Name = "Calibri";
                    excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    if (i % 2 != 0)
                    {
                        excelCells1.Interior.Color = 13434828;
                    }

                    //переработка
                    //c1 = excelworksheet.Cells[3 + i, daysInMonth + 2];
                    // c2 = excelworksheet.Cells[3 + i, daysInMonth + 3];
                    //formulaRange = (Excel.Range)excelworksheet.get_Range(c1, c2);

                    //adder = formulaRange.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    excelCells1 = (Excel.Range)excelworksheet.Cells[3 + i, daysInMonth + 5];
                    //excelCells1.Formula = String.Format("=СУММ({0}", adder);

                    excelCells1.Formula = String.Format("=IF((" + excelCellNumber(daysInMonth + 2, 3 + i) + "-" + excelCellNumber(daysInMonth + 3, 3 + i) + ")<0," + '"' + '"' + "," + excelCellNumber(daysInMonth + 2, 3 + i) + "-" + excelCellNumber(daysInMonth + 3, 3 + i) + ")");

                    excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                    excelCells1.Font.Size = 11;
                    excelCells1.Font.Name = "Calibri";
                    excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    if (i % 2 != 0)
                    {
                        excelCells1.Interior.Color = 13434828;
                    }

                }

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 4, 2];
                excelCells1.Value2 = "Норма времени рабочего дня студента";
                excelCells1.Font.Size = 11;
                excelCells1.Font.Name = "Calibri";
                excelCells1.RowHeight = 15;

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 4, daysInMonth + 3];
                excelCells1.Value2 = (hoursInMonth % 2 == 0) ? ((hoursInMonth / 2).ToString() + ":00") : (hoursInMonth / 2).ToString() + ":30";
                excelCells1.EntireColumn.NumberFormat = "[H]:mm";
                excelCells1.Font.Size = 11;
                excelCells1.Font.Name = "Calibri";
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 8, 2];
                excelCells1.Value2 = "Заявление";
                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //excelCells1.Interior.Color = 39423;
                excelCells1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 9, 2];
                excelCells1.Value2 = "Время для отгула";
                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelCells1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 10, 2];
                excelCells1.Value2 = "Ночные";
                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                excelCells1.Interior.Color = 8388736;

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 11, 2];
                excelCells1.Value2 = "Время ставить с учетом обеда!!!!!!!!\nРаботал 6 пишем 5:30";
                excelCells1.Cells.Style.WrapText = true;
                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 12, 2];
                excelCells1.Value2 = "Если работал менее 4 часов, обед не фиксируется, и в табель заносится фактическое время!";
                excelCells1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                excelCells1 = (Excel.Range)excelworksheet.Cells[inputHumans + 13, 2];
                excelCells1.Value2 = "Пример: Работал с 8:00 до 16:30, фактическое время 8:30, заносим сюда лишь 8:30-0:30=8:00";
            }

            folderpath = folderpath + "/" + numberOtdInList + ". " + otd.OtdelShortName + ".xlsx";
            excelapp.ActiveWorkbook.SaveAs(@folderpath);

            excelapp.Visible = true;



            //return true;
        }

        private int dateCheck(int year, int month, int date)
        {
            // int result;
            try
            {
                DateTime dateValue = new DateTime(year, month, date);
                if ((int)dateValue.DayOfWeek == 0)
                {
                    return 7;
                }
                else
                {
                    return (int)dateValue.DayOfWeek;
                }

                // result;
            }
            catch
            {
                //result = 0;
                return 0;
            }

        }

        private string excelCellNumber(int column, int str)
        {
            string cellNumber = null;
            string sColumn = null;
            string[] alphabet = new string[26] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int count = 0;

            if (column <= 26)
            {
                sColumn = alphabet[column - 1];
            }
            else
            {
                count = column - 26;
                sColumn = "A" + alphabet[column - 26 - 1];
            }


            cellNumber = sColumn + str.ToString();
            return cellNumber;
        }

        private void CreateGeneralExcelFile(XmlDocument xmlDoc)
        {
            Excel.Application excelapp;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;

            string[] sMonths = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };
            string[] sDays = { "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс" };

            int hoursInMonth, daysInMonth;
            int dayInWeek;
            Dictionary<string, string> HolidaysData = new Dictionary<string, string>();
            List<string> HolidaysList = new List<string>();
            // Excel.Range formulaRange;

            //Получаем все узлы, имеющие имя day
            XmlNodeList nodeList = xmlDoc.GetElementsByTagName("day");
            //Выводим значения атрибутов d и t у всех найденных узлов day 
            foreach (XmlNode xmlnode in nodeList)
            {
                HolidaysData.Add(xmlnode.Attributes["d"].InnerText, xmlnode.Attributes["t"].InnerText);
            }

            excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 12;
            excelapp.Workbooks.Add(Type.Missing);
            for (int month = 1; month <= 12; month++)
            {
                daysInMonth = 0;
                hoursInMonth = 0;
                HolidaysList.Clear();
                //записываем месяца в название листов
                excelapp.Sheets[month].Name = sMonths[month - 1];

                excelsheets = excelapp.Worksheets;
                //Получаем ссылку на лист 1
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(month);


                // Выделяем диапазон ячеек А1-А3, записываем номер
                Excel.Range c1 = excelworksheet.Cells[1, 1];
                Excel.Range c2 = excelworksheet.Cells[3, 1];
                Excel.Range excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "№";
                excelCells1.ColumnWidth = 5.86;
                excelCells1.Font.Size = 11;
                excelCells1.Font.Name = "Calibri";
                excelCells1.EntireRow.Font.Bold = true;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                c1 = excelworksheet.Cells[1, 2];
                c2 = excelworksheet.Cells[3, 2];
                excelCells1 = (Excel.Range)excelworksheet.get_Range(c1, c2);
                excelCells1.Merge(Type.Missing);//объединение ячеек
                excelCells1.Value2 = "ФИО";
                excelCells1.EntireRow.Font.Bold = true;
                excelCells1.ColumnWidth = 36;
                excelCells1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                excelCells1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                excelCells1.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                for (int date = 1; date <= 31; date++)
                {

                }

            }

        }


        private void otdelsListView_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (otdelsListView.SelectedItems.Count == 1)
            {
                deleteOtdelButton.Enabled = true;
                FIOListView.Items.Clear();
                for (int i = 0; i < otdels.Count; i++)
                {
                    if (otdelsListView.SelectedIndices[0] == i)
                    {
                        foreach(Pearson p in pearsonsList)
                        {
                            if(p.Otdel == (i+1).ToString())
                            {
                                FIOListView.Items.Add("(" + p.TableNum + ") " + p.FIO + "  -  " + p.Proffession);
                            }
                        }
                        break;
                    }
                }
                //FIOListView.Clear();
            }
                
        }

        private void deleteOtdelButton_Click(object sender, EventArgs e)
        {
            // MessageBox.Show(otdelsListView.SelectedIndices[0].ToString(), "", MessageBoxButtons.YesNo);
            var result = MetroFramework.MetroMessageBox.Show(this, "Вы действительно хотите удалить " + otdels[otdelsListView.SelectedIndices[0]].OtdelName, "Требуется подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                string str1 = otdels[otdelsListView.SelectedIndices[0]].OtdelName;

                SQLClient.deleteOtdel(str1);
            }

            //string str = otdels[otdelsListView.SelectedIndices[0]].OtdelName;

        }

        private void deleteSotrButton_Click(object sender, EventArgs e)
        {
            string str1 = FIOListView.SelectedItems[0].Text;

            char[] fioCharArray;
            List<char> fioBaseCharArray = new List<char>();

            fioCharArray = str1.ToCharArray();
            int i = 0;
            while (i < fioCharArray.Length)
            {
                if (fioCharArray[i] == '(')
                {
                    i++;
                    continue;
                }

                else
                    if (fioCharArray[i] == ')')
                    break;
                else
                    fioBaseCharArray.Add(fioCharArray[i]);
                i++;        
                    
            }

            string strTest = "";

            for (i = 0; i < fioBaseCharArray.Count; i++)
            {
                strTest = strTest + fioBaseCharArray[i].ToString();
            }

            SQLClient.deletePerson(strTest);
        }

    }
}
