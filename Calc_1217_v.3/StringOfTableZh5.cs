using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    class StringOfTableZh5
    {
        public int ID_CS;//Столбец 0, кодовый номер объекта КС
        public string gcs_name;//Столбец 1, наименование КС
        public string kc_number;//Столбец 2, номер КЦ
        public string pipeline_name;//Столбец 3, Наименование МГ
        public string number_of_diagnosed_element;//Столбец 4, номер продиагностированного элемента
        public int number_of_defect_KRN;//Столбец 5, номер дефекта, п/п
        public double Distance_of_the_element_with_SCC_defects_from_the_CC_km;//Столбец 6,Расстояние элемента с дефектами КРН от КЦ, км
    }
    partial class Form1 : Form
    {
        private List<StringOfTableZh5> get_Zh5_data_from_file(string fileName)//метод для чтения из файла таблицы Ж.2
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, false, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkBook.Unprotect();
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = "Ж.5";//получаем название вкладки
            progressBar4.Invoke(new Action(() => progressBar4.Minimum = 0));
            progressBar4.Invoke(new Action(() => progressBar4.Maximum = 0));
            progressBar4.Invoke(new Action(() => progressBar4.Step = 1));
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];
            List<StringOfTableZh5> tableZh5 = new List<StringOfTableZh5>();//создаём список строк
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.5...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.5...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));

            int j = 14;
            bool mark1 = true;
            while (mark1)//проводим цикл, чтобы узнать число строк таблицы
            {
                if (String.IsNullOrWhiteSpace(ObjWorkSheet.Cells[j, 2].Text))
                {
                    mark1 = false;//дошли до конца трубного журлала
                }
                else
                {
                    //progressBar1.Maximum++;
                    progressBar4.Invoke(new Action(() => progressBar4.Maximum++));
                    textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(progressBar4.Maximum)));
                }
                j++;
            }



            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 14;
            bool mark = true;
            while (mark)
            {
                StringOfTableZh5 stringOfTable_Zh5 = new StringOfTableZh5();//создаём экземпляр класса строки журнала аномалий
                stringOfTable_Zh5.ID_CS=Converting.ConvertToInt(ObjWorkSheet.Cells[i, 1].Text);//Столбец 0, кодовый номер объекта
                stringOfTable_Zh5.gcs_name= ObjWorkSheet.Cells[i, 2].Text;//Столбец 1, наименование КС
                stringOfTable_Zh5.kc_number= ObjWorkSheet.Cells[i, 3].Text;//Столбец 2, номер КЦ
                stringOfTable_Zh5.pipeline_name= ObjWorkSheet.Cells[i, 4].Text;//Столбец 3, Наименование МГ
                stringOfTable_Zh5.number_of_diagnosed_element= ObjWorkSheet.Cells[i, 5].Text;//Столбец 4, номер продиагностированного элемента
                stringOfTable_Zh5.number_of_defect_KRN= Converting.ConvertToInt(ObjWorkSheet.Cells[i, 6].Text);//Столбец 5, номер дефекта, п/п
                stringOfTable_Zh5.Distance_of_the_element_with_SCC_defects_from_the_CC_km= Converting.ConvertToDouble(ObjWorkSheet.Cells[i, 7].Text);//Столбец 6,Расстояние элемента с дефектами КРН от КЦ, км


                if (String.IsNullOrWhiteSpace(stringOfTable_Zh5.gcs_name))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    tableZh5.Add(stringOfTable_Zh5);//добавляем заполненный экземпляр класса к списку
                    textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(tableZh5.Count)+" из "+ Convert.ToString(progressBar4.Maximum)));
                    progressBar4.Invoke(new Action(() => progressBar4.Value++));
                }
                //textBox1.Text = Convert.ToString(tableZh5.Count);
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }
            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.5 прочитан, количество строк: " + tableZh5.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.5 прочитан, количество строк: " + tableZh5.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));
            textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(tableZh5.Count)));
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjExcel.Quit();
            return tableZh5;
        }
      
    }
}
