using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    internal class StringOfTableZh2_string
    {
        public string ID_TP;//Столбец 1, кодовый номер объекта
        public string lpu_name;//Столбец 2, наименование ЛПУ
        public string gcs_name;//Столбец 3, наименование КС
        public string kc_name;//Столбец 4, наименование КЦ
        public string tt_gcs_type;//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
        public string technological_number_tt_gks;//Столбец 6, технологический номер (П,Н,7,8,7А,8А)
        public string diameter_exterlal;//Столбец 7, наружный диаметр
        public string is_underground;//Столбец 8, "истина", если подземный, "ложь", если надземный
        public DateTime date_of_examination;//Столбец 9, дата завершения обследования
        public string is_insulation_removed;//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали
        public string number_of_diagnosed_areal;//Столбец 11, номер продиагностированного участка (уникальное имя диагностированного участка)
        public string length_of_diagnosed_areal;//Столбец 12, протяженность диагностированного участка, м.
        public string diagnostician;//Столбец 13, исполнитель работ по диагностике
        public string is_external_diagnostics;//Столбец 14, "истина", если проводилось наружное обследование, "ложь", если не проводилось
        public string is_VTD_diagnostics;//Столбец 15, "истина", если проводилось ВТД, "ложь", если не проводилось
        public string methods_of_diagnostics;//Столбец 16, методы диагностики
        public string number_of_diagnosed_element;//Столбец 17, номер продиагностированного элемента
        public string type_of_element;//Столбец 18, тип обследованного элемента (КСС, отвод, переход, тройник, труба);
        public string element_design_in_Zh3;// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3
        public string conclusion_of_diagnostics;//Cтолбец 20, заключение по презультатам обследования (Ремонт, Замена, Эксплуатация)
        public string binding_object;//Cтолбец 21, объект привязки
        public string distance_to_binding_object;//Cтолбец 22, расстояние до объекта привязки
        public string angular_orientation_of_weld_1;//Cтолбец 23, угловая ориентация продольного шва №1
        public string angular_orientation_of_weld_2;//Cтолбец 24, угловая ориентация продольного шва №2
        public string element_thikness;//Cтолбец 25, толщина стенки элемента, мм
        public string element_length;//Столбец 26, длина элемента, мм
        public string element_ovality;//Столбец 27, величина овальности, %
        public string length_of_coating_peeling;//Столбец 28, протяженность отслаивания защитного покрытия
        public string beginning_of_coating_peeling;//Столбец 29, координата начала отслаивания покрытия
        public string length_of_uncontrolled_zones;//Столбец 30, протяженность непроконтролированныз зон
        public string coordinates_of_uncontrolled_zones;//Столбец 31, координаты непроконтролированныз зон
        public string number_of_defect;//Столбец 32, номер дефекта, п/п
        public string type_of_defect;//Столбец 33, вид дефекта (Смещение кромок, подрез, закат и т.д.)
        public string defect_begin_coordinates;//Столбец 34, координаты начала дефекта
        public string defect_end_coordinates;//Столбец 35, координаты конца дефекта
        public string defect_begin_orientation;//Столбец 36, ориентация начала дефекта
        public string defect_end_orientation;//Столбец 37, ориентация конца дефекта
        public string defect_length;//Столбец 38, длина дефекта, мм
        public string defect_width;//Столбец 39, ширина дефекта, мм
        public string defect_depth;//Столбец 40, глубина дефекта, мм
        public string defect_depth_in_percent;//Столбец 41, глубина дефекта в процентах
        public string degree_of_danger;//Столбец 42, оценка степени опасности
        public string recommendations_for_defect_elimination;//Столбец 43, рекомендации по устранению дефекта
        public string repair_method;//Столбец 44,  метод ремонта
        public DateTime date_of_repair;//Столбец 45, дата ремонта
        public string repair_contractor;//Столбец 46, исполнитель ремонта
       
    }
    partial class Form1 : Form
    {
        private List<StringOfTableZh2_string> get_Zh2_data_from_file_to_string(string fileName, Form1 form1)//метод для чтения из файла таблицы Ж.2
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkBook.Unprotect();
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = "Ж.2";//получаем название вкладки

            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];
            List<StringOfTableZh2_string> tableZh2 = new List<StringOfTableZh2_string>();//создаём список строк
            form1.richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.2...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 14;
            bool mark = true;
            while (mark)
            {
                StringOfTableZh2_string stringOfTable_Zh2 = new StringOfTableZh2_string();//создаём экземпляр класса строки журнала аномалий

                stringOfTable_Zh2.ID_TP = ObjWorkSheet.Cells[i, 1].Text;//Столбец 1, кодовый номер объекта
                stringOfTable_Zh2.lpu_name = ObjWorkSheet.Cells[i, 2].Text;//Столбец 2, наименование ЛПУ
                stringOfTable_Zh2.gcs_name = ObjWorkSheet.Cells[i, 3].Text;//Столбец 3, наименование КС
                stringOfTable_Zh2.kc_name = ObjWorkSheet.Cells[i, 4].Text;//Столбец 4, наименование КЦ
                stringOfTable_Zh2.tt_gcs_type = ObjWorkSheet.Cells[i, 5].Text;//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                stringOfTable_Zh2.technological_number_tt_gks = ObjWorkSheet.Cells[i, 6].Text;//Столбец 6, технологический номер (П,Н,7,8,7А,8А)
                stringOfTable_Zh2.diameter_exterlal = ObjWorkSheet.Cells[i, 7].Text;//Столбец 7, наружный диаметр
                stringOfTable_Zh2.is_underground = ObjWorkSheet.Cells[i, 8].Text;//Столбец 8, "истина", если подземный, "ложь", если надземный
                stringOfTable_Zh2.date_of_examination = ObjWorkSheet.Cells[i, 9].Text;//Столбец 9, дата завершения обследования
                stringOfTable_Zh2.is_insulation_removed = ObjWorkSheet.Cells[i, 10].Text;//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали
                stringOfTable_Zh2.number_of_diagnosed_areal = ObjWorkSheet.Cells[i, 11].Text;//Столбец 11, номер продиагностированного участка (уникальное имя диагностированного участка)
                stringOfTable_Zh2.length_of_diagnosed_areal = ObjWorkSheet.Cells[i, 12].Text;//Столбец 12, протяженность диагностированного участка, м.
                stringOfTable_Zh2.diagnostician = ObjWorkSheet.Cells[i, 13].Text;//Столбец 13, исполнитель работ по диагностике
                stringOfTable_Zh2.is_external_diagnostics = ObjWorkSheet.Cells[i, 14].Text;//Столбец 14, "истина", если проводилось наружное обследование, "ложь", если не проводилось
                stringOfTable_Zh2.is_VTD_diagnostics = ObjWorkSheet.Cells[i, 15].Text;//Столбец 15, "истина", если проводилось ВТД, "ложь", если не проводилось
                stringOfTable_Zh2.methods_of_diagnostics = ObjWorkSheet.Cells[i, 16].Text;//Столбец 16, методы диагностики
                stringOfTable_Zh2.number_of_diagnosed_element = ObjWorkSheet.Cells[i, 17].Text;//Столбец 17, номер продиагностированного элемента
                stringOfTable_Zh2.type_of_element = ObjWorkSheet.Cells[i, 18].Text;//Столбец 18, тип обследованного элемента (КСС, отвод, переход, тройник, труба);
                stringOfTable_Zh2.element_design_in_Zh3 = ObjWorkSheet.Cells[i, 19].Text;// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3
                stringOfTable_Zh2.conclusion_of_diagnostics = ObjWorkSheet.Cells[i, 20].Text;//Cтолбец 20, заключение по презультатам обследования (Ремонт, Замена, Эксплуатация)
                stringOfTable_Zh2.binding_object = ObjWorkSheet.Cells[i, 21].Text;//Cтолбец 21, объект привязки
                stringOfTable_Zh2.distance_to_binding_object = ObjWorkSheet.Cells[i, 22].Text;//Cтолбец 22, расстояние до объекта привязки
                stringOfTable_Zh2.angular_orientation_of_weld_1 = ObjWorkSheet.Cells[i, 23].Text;//Cтолбец 23, угловая ориентация продольного шва №1
                stringOfTable_Zh2.angular_orientation_of_weld_2 = ObjWorkSheet.Cells[i, 24].Text;//Cтолбец 24, угловая ориентация продольного шва №2
                stringOfTable_Zh2.element_thikness = ObjWorkSheet.Cells[i, 25].Text;//Cтолбец 25, толщина стенки элемента, мм
                stringOfTable_Zh2.element_length = ObjWorkSheet.Cells[i, 26].Text;//Столбец 26, длина элемента, мм
                stringOfTable_Zh2.element_ovality = ObjWorkSheet.Cells[i, 27].Text;//Столбец 27, величина овальности, %
                stringOfTable_Zh2.length_of_coating_peeling = ObjWorkSheet.Cells[i, 28].Text;//Столбец 28, протяженность отслаивания защитного покрытия
                stringOfTable_Zh2.beginning_of_coating_peeling = ObjWorkSheet.Cells[i, 29].Text;//Столбец 29, координата начала отслаивания покрытия
                stringOfTable_Zh2.length_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 30].Text;//Столбец 30, протяженность непроконтролированныз зон
                stringOfTable_Zh2.coordinates_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 31].Text;//Столбец 31, координаты непроконтролированныз зон
                stringOfTable_Zh2.number_of_defect = ObjWorkSheet.Cells[i, 32].Text;//Столбец 32, номер дефекта, п/п
                stringOfTable_Zh2.type_of_defect = ObjWorkSheet.Cells[i, 33].Text;//Столбец 33, вид дефекта (Смещение кромок, подрез, закат и т.д.)
                stringOfTable_Zh2.defect_begin_coordinates = ObjWorkSheet.Cells[i, 34].Text;//Столбец 34, координаты начала дефекта
                stringOfTable_Zh2.defect_end_coordinates = ObjWorkSheet.Cells[i, 35].Text;//Столбец 35, координаты конца дефекта
                stringOfTable_Zh2.defect_begin_orientation = ObjWorkSheet.Cells[i, 36].Text;//Столбец 36, ориентация начала дефекта
                stringOfTable_Zh2.defect_end_orientation = ObjWorkSheet.Cells[i, 37].Text;//Столбец 37, ориентация конца дефекта
                stringOfTable_Zh2.defect_length = ObjWorkSheet.Cells[i, 38].Text;//Столбец 38, длина дефекта, мм
                stringOfTable_Zh2.defect_width = ObjWorkSheet.Cells[i, 39].Text;//Столбец 39, ширина дефекта, мм
                stringOfTable_Zh2.defect_depth = ObjWorkSheet.Cells[i, 40].Text;//Столбец 40, глубина дефекта, мм
                stringOfTable_Zh2.defect_depth_in_percent = ObjWorkSheet.Cells[i, 41].Text;//Столбец 41, глубина дефекта в процентах
                stringOfTable_Zh2.degree_of_danger = ObjWorkSheet.Cells[i, 42].Text;//Столбец 42, оценка степени опасности
                stringOfTable_Zh2.recommendations_for_defect_elimination = ObjWorkSheet.Cells[i, 43].Text;//Столбец 43, рекомендации по устранению дефекта
                stringOfTable_Zh2.repair_method = ObjWorkSheet.Cells[i, 44].Text;//Столбец 44,  метод ремонта
                stringOfTable_Zh2.date_of_repair = DateTime.FromOADate((double)ObjWorkSheet.Cells[i, 45].value2);//Столбец 45, дата ремонта
                stringOfTable_Zh2.repair_contractor = ObjWorkSheet.Cells[i, 46].Text;//Столбец 46, исполнитель ремонта

                if (String.IsNullOrWhiteSpace(stringOfTable_Zh2.lpu_name))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    tableZh2.Add(stringOfTable_Zh2);//добавляем заполненный экземпляр класса к списку
                    //richTextBox1.AppendText(FurnishingsLog.characterFeatures);
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
                i++;
            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 прочитан, количество строк:" + tableZh2.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return tableZh2;
        }
        private List<StringOfTableZh2_string> get_Zh2_data_from_file_to_string_short(string fileName)//короткий метод для чтения из файла таблицы Ж.2
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkBook.Unprotect();
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = "Ж.2";//получаем название вкладки


            //progressBar1.Minimum = 0;
            //progressBar1.Maximum = 0;
            //progressBar1.Step = 1;
            progressBar1.Invoke(new Action(() => progressBar1.Minimum = 0));
            progressBar1.Invoke(new Action(() => progressBar1.Maximum = 0));
            progressBar1.Invoke(new Action(() => progressBar1.Step = 1));

            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];
            List<StringOfTableZh2_string> tableZh2 = new List<StringOfTableZh2_string>();//создаём список строк
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.2...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.2...")));
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
                    progressBar1.Invoke(new Action(() => progressBar1.Maximum++));
                }
                j++;
            }
            //textBox1.Text = Convert.ToString(progressBar1.Maximum);
            textBox1.Invoke(new Action(() => textBox1.Text = Convert.ToString(progressBar1.Maximum)));
            
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 14;
            bool mark = true;
            while (mark)
            {
                StringOfTableZh2_string stringOfTable_Zh2 = new StringOfTableZh2_string();//создаём экземпляр класса строки журнала аномалий

                stringOfTable_Zh2.ID_TP = ObjWorkSheet.Cells[i, 1].Text;//Столбец 1, кодовый номер объекта
                //stringOfTable_Zh2.lpu_name = ObjWorkSheet.Cells[i, 2].Text;//Столбец 2, наименование ЛПУ
                //stringOfTable_Zh2.gcs_name = ObjWorkSheet.Cells[i, 3].Text;//Столбец 3, наименование КС
                //stringOfTable_Zh2.kc_name = ObjWorkSheet.Cells[i, 4].Text;//Столбец 4, наименование КЦ
                //stringOfTable_Zh2.tt_gcs_type = ObjWorkSheet.Cells[i, 5].Text;//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                //stringOfTable_Zh2.technological_number_tt_gks = ObjWorkSheet.Cells[i, 6].Text;//Столбец 6, технологический номер (П,Н,7,8,7А,8А)


                stringOfTable_Zh2.lpu_name = "0";//Столбец 2, наименование ЛПУ
                stringOfTable_Zh2.gcs_name = "0";//Столбец 3, наименование КС
                stringOfTable_Zh2.kc_name = "0";//Столбец 4, наименование КЦ
                stringOfTable_Zh2.tt_gcs_type = "0";//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                stringOfTable_Zh2.technological_number_tt_gks = "0";//Столбец 6, технологический номер (П,Н,7,8,7А,8А)



                stringOfTable_Zh2.diameter_exterlal = ObjWorkSheet.Cells[i, 7].Text;//Столбец 7, наружный диаметр
                //stringOfTable_Zh2.is_underground = ObjWorkSheet.Cells[i, 8].Text;//Столбец 8, "истина", если подземный, "ложь", если надземный
                //stringOfTable_Zh2.date_of_examination = ObjWorkSheet.Cells[i, 9].Text;//Столбец 9, дата завершения обследования
                //stringOfTable_Zh2.is_insulation_removed = ObjWorkSheet.Cells[i, 10].Text;//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали


                stringOfTable_Zh2.is_underground = "0";//Столбец 8, "истина", если подземный, "ложь", если надземный
                //stringOfTable_Zh2.date_of_examination = "0";//Столбец 9, дата завершения обследования
                stringOfTable_Zh2.is_insulation_removed = "0";//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали


                stringOfTable_Zh2.number_of_diagnosed_areal = ObjWorkSheet.Cells[i, 11].Text;//Столбец 11, номер продиагностированного участка (уникальное имя диагностированного участка)
                stringOfTable_Zh2.length_of_diagnosed_areal = ObjWorkSheet.Cells[i, 12].Text;//Столбец 12, протяженность диагностированного участка, м.
                //stringOfTable_Zh2.diagnostician = ObjWorkSheet.Cells[i, 13].Text;//Столбец 13, исполнитель работ по диагностике

                stringOfTable_Zh2.diagnostician = "0";//Столбец 13, исполнитель работ по диагностике

                stringOfTable_Zh2.is_external_diagnostics = ObjWorkSheet.Cells[i, 14].Text;//Столбец 14, "истина", если проводилось наружное обследование, "ложь", если не проводилось
                stringOfTable_Zh2.is_VTD_diagnostics = ObjWorkSheet.Cells[i, 15].Text;//Столбец 15, "истина", если проводилось ВТД, "ложь", если не проводилось
                //stringOfTable_Zh2.methods_of_diagnostics = ObjWorkSheet.Cells[i, 16].Text;//Столбец 16, методы диагностики

                stringOfTable_Zh2.methods_of_diagnostics = "0";//Столбец 16, методы диагностики

                stringOfTable_Zh2.number_of_diagnosed_element = ObjWorkSheet.Cells[i, 17].Text;//Столбец 17, номер продиагностированного элемента
                stringOfTable_Zh2.type_of_element = ObjWorkSheet.Cells[i, 18].Text;//Столбец 18, тип обследованного элемента (КСС, отвод, переход, тройник, труба);
                //stringOfTable_Zh2.element_design_in_Zh3 = ObjWorkSheet.Cells[i, 19].Text;// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3

                stringOfTable_Zh2.element_design_in_Zh3 = "0";// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3

                stringOfTable_Zh2.conclusion_of_diagnostics = ObjWorkSheet.Cells[i, 20].Text;//Cтолбец 20, заключение по презультатам обследования (Ремонт, Замена, Эксплуатация)
                //stringOfTable_Zh2.binding_object = ObjWorkSheet.Cells[i, 21].Text;//Cтолбец 21, объект привязки
                //stringOfTable_Zh2.distance_to_binding_object = ObjWorkSheet.Cells[i, 22].Text;//Cтолбец 22, расстояние до объекта привязки
                //stringOfTable_Zh2.angular_orientation_of_weld_1 = ObjWorkSheet.Cells[i, 23].Text;//Cтолбец 23, угловая ориентация продольного шва №1
                //stringOfTable_Zh2.angular_orientation_of_weld_2 = ObjWorkSheet.Cells[i, 24].Text;//Cтолбец 24, угловая ориентация продольного шва №2

                stringOfTable_Zh2.binding_object = "0";//Cтолбец 21, объект привязки
                stringOfTable_Zh2.distance_to_binding_object = "0";//Cтолбец 22, расстояние до объекта привязки
                stringOfTable_Zh2.angular_orientation_of_weld_1 = "0";//Cтолбец 23, угловая ориентация продольного шва №1
                stringOfTable_Zh2.angular_orientation_of_weld_2 = "0";//Cтолбец 24, угловая ориентация продольного шва №2


                stringOfTable_Zh2.element_thikness = ObjWorkSheet.Cells[i, 25].Text;//Cтолбец 25, толщина стенки элемента, мм
                stringOfTable_Zh2.element_length = ObjWorkSheet.Cells[i, 26].Text;//Столбец 26, длина элемента, мм
                //stringOfTable_Zh2.element_ovality = ObjWorkSheet.Cells[i, 27].Text;//Столбец 27, величина овальности, %
                //stringOfTable_Zh2.length_of_coating_peeling = ObjWorkSheet.Cells[i, 28].Text;//Столбец 28, протяженность отслаивания защитного покрытия
                //stringOfTable_Zh2.beginning_of_coating_peeling = ObjWorkSheet.Cells[i, 29].Text;//Столбец 29, координата начала отслаивания покрытия

                stringOfTable_Zh2.element_ovality = "0";//Столбец 27, величина овальности, %
                stringOfTable_Zh2.length_of_coating_peeling = "0";//Столбец 28, протяженность отслаивания защитного покрытия
                stringOfTable_Zh2.beginning_of_coating_peeling = "0";//Столбец 29, координата начала отслаивания покрытия


                stringOfTable_Zh2.length_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 30].Text;//Столбец 30, протяженность непроконтролированныз зон
                //stringOfTable_Zh2.coordinates_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 31].Text;//Столбец 31, координаты непроконтролированныз зон

                stringOfTable_Zh2.coordinates_of_uncontrolled_zones = "0";//Столбец 31, координаты непроконтролированныз зон


                stringOfTable_Zh2.number_of_defect = ObjWorkSheet.Cells[i, 32].Text;//Столбец 32, номер дефекта, п/п
                stringOfTable_Zh2.type_of_defect = ObjWorkSheet.Cells[i, 33].Text;//Столбец 33, вид дефекта (Смещение кромок, подрез, закат и т.д.)
                //stringOfTable_Zh2.defect_begin_coordinates = ObjWorkSheet.Cells[i, 34].Text;//Столбец 34, координаты начала дефекта
                //stringOfTable_Zh2.defect_end_coordinates = ObjWorkSheet.Cells[i, 35].Text;//Столбец 35, координаты конца дефекта
                //stringOfTable_Zh2.defect_begin_orientation = ObjWorkSheet.Cells[i, 36].Text;//Столбец 36, ориентация начала дефекта
                //stringOfTable_Zh2.defect_end_orientation = ObjWorkSheet.Cells[i, 37].Text;//Столбец 37, ориентация конца дефекта

                stringOfTable_Zh2.defect_begin_coordinates = "0";//Столбец 34, координаты начала дефекта
                stringOfTable_Zh2.defect_end_coordinates = "0";//Столбец 35, координаты конца дефекта
                stringOfTable_Zh2.defect_begin_orientation = "0";//Столбец 36, ориентация начала дефекта
                stringOfTable_Zh2.defect_end_orientation = "0";//Столбец 37, ориентация конца дефекта


                stringOfTable_Zh2.defect_length = ObjWorkSheet.Cells[i, 38].Text;//Столбец 38, длина дефекта, мм
                stringOfTable_Zh2.defect_width = ObjWorkSheet.Cells[i, 39].Text;//Столбец 39, ширина дефекта, мм
                stringOfTable_Zh2.defect_depth = ObjWorkSheet.Cells[i, 40].Text;//Столбец 40, глубина дефекта, мм
                stringOfTable_Zh2.defect_depth_in_percent = ObjWorkSheet.Cells[i, 41].Text;//Столбец 41, глубина дефекта в процентах
                stringOfTable_Zh2.degree_of_danger = ObjWorkSheet.Cells[i, 42].Text;//Столбец 42, оценка степени опасности
                stringOfTable_Zh2.recommendations_for_defect_elimination = ObjWorkSheet.Cells[i, 43].Text;//Столбец 43, рекомендации по устранению дефекта
                stringOfTable_Zh2.repair_method = ObjWorkSheet.Cells[i, 44].Text;//Столбец 44,  метод ремонта
                stringOfTable_Zh2.date_of_repair = ObjWorkSheet.Cells[i, 45].Text;//Столбец 45, дата ремонта
                                                                                  //stringOfTable_Zh2.repair_contractor = ObjWorkSheet.Cells[i, 46].Text;//Столбец 46, исполнитель ремонта


                stringOfTable_Zh2.repair_contractor = "0";//Столбец 46, исполнитель ремонта
                if (String.IsNullOrWhiteSpace(stringOfTable_Zh2.ID_TP))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    tableZh2.Add(stringOfTable_Zh2);//добавляем заполненный экземпляр класса к списку
                    //textBox2.Text = Convert.ToString(tableZh2.Count);
                    //progressBar1.Value++;
                    textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(tableZh2.Count)));
                    progressBar1.Invoke(new Action(() => progressBar1.Value++));
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
            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 прочитан, количество строк:" + tableZh2.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 прочитан, количество строк:")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            ObjExcel.Quit();
            return tableZh2;
        }
    }

}
