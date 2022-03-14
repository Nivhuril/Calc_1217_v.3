using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    class StringOfTableZh4
    {
        public int ID_CS;//Столбец 0, кодовый номер объекта КС
        public string gcs_name;//Столбец 1, наименование КС
        public string kc_number;//Столбец 2, номер КЦ
        public string pipeline_name;//Столбец 3, Наименование МГ
        public double OperatingLifeOfTheMGsection;//Столбец 4, Срок эксплуатации участка МГ 
        public int a0;//Столбец 5, Количество аварий и инцидентов На территории КЦ a0
        public int a1;//Столбец 6, Количество аварий и инцидентов На расстоянии от КЦ:  от 0 до 3 км a1
        public int a2;//Столбец 7, Количество аварий и инцидентов На расстоянии от КЦ:  от 3 до 5 км a2
        public int a3;//Столбец 8, Количество аварий и инцидентов На расстоянии от КЦ: от 5 до 10 км a3
        public int a4;//Столбец 9, Количество аварий и инцидентов На расстоянии от КЦ: от 10 до 50 км a4
        public double Expiration_year_of_the_safe_operation_of_the_gas_pipeline_section;//Столбец 10, Год истечения срока безопасной эксплуатации участка ЛЧ МГ, в том числе назначенный по результатам ЭПБ
        public string note;//Столбец 11, Прочие сведения
    }
    partial class Form1 : Form
    {
        private List<StringOfTableZh4> get_Zh4_data_from_file(string fileName)//метод для чтения из файла таблицы Ж.2
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, false, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkBook.Unprotect();
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = "Ж.4";//получаем название вкладки
            int a = Convert.ToInt32(textBox7.Text) + 13;
            progressBar3.Invoke(new Action(() => progressBar3.Minimum = 0));
            progressBar3.Invoke(new Action(() => progressBar3.Maximum = Convert.ToInt32(textBox7.Text)));
            progressBar3.Invoke(new Action(() => progressBar3.Step = 1));
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];
            List<StringOfTableZh4> tableZh4 = new List<StringOfTableZh4>();//создаём список строк
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.4...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.4...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 14;
            bool mark = true;
            while (mark)
            {
                StringOfTableZh4 stringOfTable_Zh4 = new StringOfTableZh4();//создаём экземпляр класса строки журнала аномалий
                stringOfTable_Zh4.ID_CS = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 1].Text);//Столбец 0, кодовый номер объекта КС
                stringOfTable_Zh4.gcs_name = ObjWorkSheet.Cells[i, 1].Text;//Столбец 1, наименование КС
                stringOfTable_Zh4.kc_number = ObjWorkSheet.Cells[i, 1].Text;//Столбец 2, номер КЦ
                stringOfTable_Zh4.pipeline_name = ObjWorkSheet.Cells[i, 1].Text;//Столбец 3, Наименование МГ
                stringOfTable_Zh4.OperatingLifeOfTheMGsection = Converting.ConvertToDouble(ObjWorkSheet.Cells[i, 1].Text);//Столбец 4, Срок эксплуатации участка МГ 
                stringOfTable_Zh4.a0 = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 6].Text);//Столбец 5, Количество аварий и инцидентов На территории КЦ
                stringOfTable_Zh4.a1 = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 7].Text);//Столбец 6, Количество аварий и инцидентов На расстоянии от КЦ:  от 0 до 3 км
                stringOfTable_Zh4.a2 = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 8].Text);//Столбец 7, Количество аварий и инцидентов На расстоянии от КЦ:  от 3 до 5 км 
                stringOfTable_Zh4.a3 = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 9].Text);//Столбец 8, Количество аварий и инцидентов На расстоянии от КЦ: от 5 до 10 км
                stringOfTable_Zh4.a4 = Converting.ConvertToInt(ObjWorkSheet.Cells[i, 10].Text);//Столбец 9, Количество аварий и инцидентов На расстоянии от КЦ: от 10 до 50 км
                stringOfTable_Zh4.Expiration_year_of_the_safe_operation_of_the_gas_pipeline_section = Converting.ConvertToDouble(ObjWorkSheet.Cells[i, 1].Text);//Столбец 10, Год истечения срока безопасной эксплуатации участка ЛЧ МГ, в том числе назначенный по результатам ЭПБ
                stringOfTable_Zh4.note = ObjWorkSheet.Cells[i, 1].Text;//Столбец 11, Прочие сведения

                if (String.IsNullOrWhiteSpace(stringOfTable_Zh4.gcs_name))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    tableZh4.Add(stringOfTable_Zh4);//добавляем заполненный экземпляр класса к списку
                    textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(tableZh4.Count)));
                    progressBar3.Invoke(new Action(() => progressBar3.Value++));
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }
            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.4 прочитан, количество строк:" + tableZh4.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.4 прочитан, количество строк: " + tableZh4.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjExcel.Quit();
            return tableZh4;
        }

    }

}
