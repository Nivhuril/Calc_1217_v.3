using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace CalcGCS_1217
{
    internal class StringOfTableZh2//Класс для хранения одной строки таблицы Ж2
    {
        public int ID_TP;//Столбец 1, кодовый номер объекта
        public string lpu_name;//Столбец 2, наименование ЛПУ
        public string gcs_name;//Столбец 3, наименование КС
        public string kc_name;//Столбец 4, наименование КЦ
        public string tt_gcs_type;//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
        public string technological_number_tt_gks;//Столбец 6, технологический номер (П,Н,7,8,7А,8А)
        public double diameter_exterlal;//Столбец 7, наружный диаметр
        public bool is_underground;//Столбец 8, "истина", если подземный, "ложь", если надземный
        public DateTime date_of_examination;//Столбец 9, дата завершения обследования
        public bool is_insulation_removed;//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали
        public string number_of_diagnosed_areal;//Столбец 11, номер продиагностированного участка (уникальное имя диагностированного участка)
        public double length_of_diagnosed_areal;//Столбец 12, протяженность диагностированного участка, м.
        public string diagnostician;//Столбец 13, исполнитель работ по диагностике
        public bool is_external_diagnostics;//Столбец 14, "истина", если проводилось наружное обследование, "ложь", если не проводилось
        public bool is_VTD_diagnostics;//Столбец 15, "истина", если проводилось ВТД, "ложь", если не проводилось
        public string methods_of_diagnostics;//Столбец 16, методы диагностики
        public string number_of_diagnosed_element;//Столбец 17, номер продиагностированного элемента
        public string type_of_element;//Столбец 18, тип обследованного элемента (КСС, отвод, переход, тройник, труба);
        public string element_design_in_Zh3;// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3
        public string conclusion_of_diagnostics;//Cтолбец 20, заключение по презультатам обследования (Ремонт, Замена, Эксплуатация)
        public string binding_object;//Cтолбец 21, объект привязки
        public string distance_to_binding_object;//Cтолбец 22, расстояние до объекта привязки
        public double angular_orientation_of_weld_1;//Cтолбец 23, угловая ориентация продольного шва №1
        public double angular_orientation_of_weld_2;//Cтолбец 24, угловая ориентация продольного шва №2
        public double element_thikness;//Cтолбец 25, толщина стенки элемента, мм
        public double element_length;//Столбец 26, длина элемента, мм
        public double element_ovality;//Столбец 27, величина овальности, %
        public double length_of_coating_peeling;//Столбец 28, протяженность отслаивания защитного покрытия
        public double beginning_of_coating_peeling;//Столбец 29, координата начала отслаивания покрытия
        public double length_of_uncontrolled_zones;//Столбец 30, протяженность непроконтролированныз зон
        public double coordinates_of_uncontrolled_zones;//Столбец 31, координаты непроконтролированныз зон
        public int number_of_defect;//Столбец 32, номер дефекта, п/п
        public string type_of_defect;//Столбец 33, вид дефекта (Смещение кромок, подрез, закат и т.д.)
        public double defect_begin_coordinates;//Столбец 34, координаты начала дефекта
        public double defect_end_coordinates;//Столбец 35, координаты конца дефекта
        public double defect_begin_orientation;//Столбец 36, ориентация начала дефекта
        public double defect_end_orientation;//Столбец 37, ориентация конца дефекта
        public double defect_length;//Столбец 38, длина дефекта, мм
        public double defect_width;//Столбец 39, ширина дефекта, мм
        public double defect_depth;//Столбец 40, глубина дефекта, мм
        public double defect_depth_in_percent;//Столбец 41, глубина дефекта в процентах
        public string degree_of_danger;//Столбец 42, оценка степени опасности
        public string recommendations_for_defect_elimination;//Столбец 43, рекомендации по устранению дефекта
        public string repair_method;//Столбец 44,  метод ремонта
        public DateTime date_of_repair;//Столбец 45, дата ремонта
        public string repair_contractor;//Столбец 46, исполнитель ремонта
                                        //формируется отдельно
        public string type_for_sort;//тип элемента+диаметр, для сортировки
        public double rang_of_danger;//ранг опасности
               
    }
    public class repare_inform//хранение информации о способе и результатах ремонта
    {
        public string recommendations_for_defect_elimination;//Столбец 43, рекомендации по устранению дефекта
        public string repair_method;//Столбец 44,  метод ремонта
        public double rang_of_danger;//ранг опасности
        public repare_inform(string recommendations_for_defect_elimination, string repair_method, double rang_of_danger)
        {
            this.recommendations_for_defect_elimination = recommendations_for_defect_elimination;
            this.repair_method = repair_method;
            this.rang_of_danger = rang_of_danger;
        }
    }
    partial class Form1 : Form
    {
        private List<StringOfTableZh2> getSortMassiveZh2(List<StringOfTableZh2> input)
        {            
            input = input.OrderBy(i => i.ID_TP).ToList();
            input = input.OrderBy(i => i.type_for_sort).ToList();
            input = input.OrderBy(i => i.number_of_diagnosed_element).ToList();
            return input;
        }
        private List<StringOfTableZh2> getSortMassiveZh2_for_rangs(List<StringOfTableZh2> input)
        {           
            input = input.OrderBy(i => i.type_for_sort).ToList();
            input = input.OrderBy(i => i.number_of_diagnosed_element).ToList();
            input = input.OrderBy(i => i.ID_TP).ToList();
            return input;
        }
        private List<StringOfTableZh2> Convert_Zh2_to_current_type(List<StringOfTableZh2_string> input)//конвертация сырой таблицы в таблицу с нужными типами данных
        {
            List<StringOfTableZh2> result = new List<StringOfTableZh2>();
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала Ж.2...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала Ж.2...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 0; i < input.Count; i++)
            {
                StringOfTableZh2 current = new StringOfTableZh2();
                current.ID_TP = Converting.ConvertToInt(input[i].ID_TP);
                //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала Ж.2..."+ i +" "+ input[i].ID_TP);
                current.lpu_name = input[i].lpu_name;
                current.gcs_name = input[i].gcs_name;
                current.kc_name = input[i].kc_name;
                current.tt_gcs_type = input[i].tt_gcs_type;
                current.technological_number_tt_gks = input[i].technological_number_tt_gks;
                current.diameter_exterlal = Converting.ConvertToDouble(input[i].diameter_exterlal);
                current.is_underground = Converting.GetBoolean(input[i].is_underground, "одзем");
                //current.date_of_examination = input[i].date_of_examination;
                current.is_insulation_removed = Converting.GetBoolean(input[i].is_insulation_removed, "снятием");
                current.number_of_diagnosed_areal = input[i].number_of_diagnosed_areal;
                current.length_of_diagnosed_areal = Converting.ConvertToDouble(input[i].length_of_diagnosed_areal);
                current.diagnostician = input[i].diagnostician;
                current.is_external_diagnostics = Converting.GetBoolean(input[i].is_external_diagnostics, "Да");
                current.is_VTD_diagnostics = Converting.GetBoolean(input[i].is_VTD_diagnostics, "Да");
                current.methods_of_diagnostics = input[i].methods_of_diagnostics;
                current.number_of_diagnosed_element = input[i].number_of_diagnosed_element;
                current.type_of_element = input[i].type_of_element;
                current.element_design_in_Zh3 = input[i].element_design_in_Zh3;
                current.conclusion_of_diagnostics = input[i].conclusion_of_diagnostics;
                current.binding_object = input[i].binding_object;
                current.distance_to_binding_object = input[i].distance_to_binding_object;
                current.angular_orientation_of_weld_1 = Converting.ConvertToDouble(input[i].angular_orientation_of_weld_1);
                current.angular_orientation_of_weld_2 = Converting.ConvertToDouble(input[i].angular_orientation_of_weld_2);
                current.element_thikness = Converting.ConvertToDouble(input[i].element_thikness);
                current.element_length = Converting.ConvertToDouble(input[i].element_length);
                current.element_ovality = Converting.ConvertToDouble(input[i].element_ovality);
                current.length_of_coating_peeling = Converting.ConvertToDouble(input[i].length_of_coating_peeling);
                current.beginning_of_coating_peeling = Converting.ConvertToDouble(input[i].beginning_of_coating_peeling);
                current.length_of_uncontrolled_zones = Converting.ConvertToDouble(input[i].length_of_uncontrolled_zones);
                current.coordinates_of_uncontrolled_zones = Converting.ConvertToDouble(input[i].coordinates_of_uncontrolled_zones);
                current.number_of_defect = Converting.ConvertToInt(input[i].number_of_defect);
                current.type_of_defect = input[i].type_of_defect;
                current.defect_begin_coordinates = Converting.ConvertToDouble(input[i].defect_begin_coordinates);
                current.defect_end_coordinates = Converting.ConvertToDouble(input[i].defect_end_coordinates);
                current.defect_begin_orientation = Converting.ConvertToDouble(input[i].defect_begin_orientation);
                current.defect_end_orientation = Converting.ConvertToDouble(input[i].defect_end_orientation);
                current.defect_length = Converting.ConvertToDouble(input[i].defect_length);
                current.defect_width = Converting.ConvertToDouble(input[i].defect_width);
                current.defect_depth = Converting.ConvertToDouble(input[i].defect_depth);
                current.defect_depth_in_percent = Converting.ConvertToDouble(input[i].defect_depth_in_percent);
                current.degree_of_danger = input[i].degree_of_danger;
                current.recommendations_for_defect_elimination = input[i].recommendations_for_defect_elimination;
                current.repair_method = input[i].repair_method;
                current.date_of_repair = input[i].date_of_repair;
                current.repair_contractor = input[i].repair_contractor;
                current.type_for_sort = String.Concat(input[i].type_of_element, input[i].diameter_exterlal);
                //richTextBox1.AppendText(Environment.NewLine + "type_for_sort:" + current.type_for_sort);
                result.Add(current);
                //textBox2.Text = Convert.ToString(result.Count);
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.                
                if (incrementor > 29)
                {
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;

                }
            }
            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 обработан, количество строк:" + result.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 обработан, количество строк:" + result.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));
            return result;
        }
        private double get_rang_local(StringOfTableZh2 input)//определение ранга опасности для одного (любого) дефекта
        {
            double result = 0;
            repare_inform var1 = new repare_inform("Замена", "Замена", 1);
            repare_inform var2 = new repare_inform("Замена", "Не выполнен", 1);
            repare_inform var3 = new repare_inform("Замена", "Замена катушки", 0.75);
            repare_inform var4 = new repare_inform("Замена", "Контролируемая шлифовка", 0.25);
            repare_inform var5 = new repare_inform("Замена", "Сварка", 0.25);
            repare_inform var6 = new repare_inform("Ремонт", "Замена", 1);
            repare_inform var7 = new repare_inform("Ремонт", "Замена катушки", 0.75);
            repare_inform var8 = new repare_inform("Ремонт", "Контролируемая шлифовка", 0.25);
            repare_inform var9 = new repare_inform("Ремонт", "Сварка", 0.25);
            repare_inform var10 = new repare_inform("Ремонт", "Не выполнен", 0.75);
            repare_inform var11 = new repare_inform("Эксплуатация", "Замена катушки", 0.75);
            repare_inform var12 = new repare_inform("Эксплуатация", "Контролируемая шлифовка", 0.25);
            repare_inform var13 = new repare_inform("Эксплуатация", "Не выполнен", 0.1);
            repare_inform var14 = new repare_inform("-", "Не выполнен", 0.1);
            repare_inform var15 = new repare_inform("Замена/муфта/катушка", "Не выполнен", 1);//Значения написаны наугад, требуется уточнение.
            repare_inform var16 = new repare_inform("Эксплуатация", "Замена", 0.25);//Значения написаны наугад, требуется уточнение.

            if (input.recommendations_for_defect_elimination == var1.recommendations_for_defect_elimination & input.repair_method == var1.repair_method)
            {
                result = var1.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var2.recommendations_for_defect_elimination & input.repair_method == var2.repair_method)
            {
                result = var2.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var3.recommendations_for_defect_elimination & input.repair_method == var3.repair_method)
            {
                result = var3.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var4.recommendations_for_defect_elimination & input.repair_method == var4.repair_method)
            {
                result = var4.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var5.recommendations_for_defect_elimination & input.repair_method == var5.repair_method)
            {
                result = var5.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var6.recommendations_for_defect_elimination & input.repair_method == var6.repair_method)
            {
                result = var6.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var7.recommendations_for_defect_elimination & input.repair_method == var7.repair_method)
            {
                result = var7.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var8.recommendations_for_defect_elimination & input.repair_method == var8.repair_method)
            {
                result = var8.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var9.recommendations_for_defect_elimination & input.repair_method == var9.repair_method)
            {
                result = var9.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var10.recommendations_for_defect_elimination & input.repair_method == var10.repair_method)
            {
                result = var10.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var11.recommendations_for_defect_elimination & input.repair_method == var11.repair_method)
            {
                result = var11.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var12.recommendations_for_defect_elimination & input.repair_method == var12.repair_method)
            {
                result = var12.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var13.recommendations_for_defect_elimination & input.repair_method == var13.repair_method)
            {
                result = var13.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var14.recommendations_for_defect_elimination & input.repair_method == var14.repair_method)
            {
                result = var14.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var15.recommendations_for_defect_elimination & input.repair_method == var15.repair_method)
            {
                result = var15.rang_of_danger;
            }
            else if (input.recommendations_for_defect_elimination == var16.recommendations_for_defect_elimination & input.repair_method == var16.repair_method)
            {
                result = var16.rang_of_danger;
            }

            return result;
        }
        private List<StringOfTableZh2> get_rang_of_danger(List<StringOfTableZh2> input)//добавить к строкам таблицы Ж.2 значение ранга опасности
        {
            List<StringOfTableZh2> result = new List<StringOfTableZh2>();
            result = input;
            int incrementor = 0;
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            richTextBox1.AppendText(Environment.NewLine + "Выполняется определение рангов опасности для дефектов таблицы Ж.2");
            for (int i = 0; i < result.Count; i++)
            {
                double rang = get_rang_local(result[i]);
                result[i].rang_of_danger = rang;
                //richTextBox1.AppendText(Environment.NewLine + result[i].rang_of_danger);
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
            }
            for (int i = 0; i < result.Count; i++)
            {
                if (result[i].rang_of_danger==0)
                {
                    result[i].rang_of_danger = 0;
                }
            }
            richTextBox1.AppendText(Environment.NewLine + "Определение рангов опасности для дефектов таблицы Ж.2 выполнено, количество строк:" + result.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            return result;
        }
        private List<StringOfTableZh2> get_Zh2_data_from_file_short(string fileName)//короткий метод для чтения из файла таблицы Ж.2
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, false, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
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
            //List<StringOfTableZh2_string> tableZh2 = new List<StringOfTableZh2_string>();//создаём список строк
            List<StringOfTableZh2> tableZh2_converted = new List<StringOfTableZh2>();//создаём список строк
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
                    textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(progressBar1.Maximum)));
                }
                j++;
            }
            //textBox1.Text = Convert.ToString(progressBar1.Maximum);
            //textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(progressBar1.Maximum)));

            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 14;
            bool mark = true;
            while (mark)
            {
                DateTime dtfv = new DateTime();
                StringOfTableZh2_string stringOfTable_Zh2 = new StringOfTableZh2_string();//создаём экземпляр класса строки журнала аномалий
                stringOfTable_Zh2.ID_TP = ObjWorkSheet.Cells[i, 1].Text;//Столбец 1, кодовый номер объекта
                stringOfTable_Zh2.lpu_name = ObjWorkSheet.Cells[i, 2].Text;//Столбец 2, наименование ЛПУ
                stringOfTable_Zh2.gcs_name = ObjWorkSheet.Cells[i, 3].Text;//Столбец 3, наименование КС
                stringOfTable_Zh2.kc_name = ObjWorkSheet.Cells[i, 4].Text;//Столбец 4, наименование КЦ
                stringOfTable_Zh2.tt_gcs_type = ObjWorkSheet.Cells[i, 5].Text;//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                stringOfTable_Zh2.technological_number_tt_gks = ObjWorkSheet.Cells[i, 6].Text;//Столбец 6, технологический номер (П,Н,7,8,7А,8А)
                {/*stringOfTable_Zh2.lpu_name = "0";//Столбец 2, наименование ЛПУ
                stringOfTable_Zh2.gcs_name = "0";//Столбец 3, наименование КС
                stringOfTable_Zh2.kc_name = "0";//Столбец 4, наименование КЦ
                stringOfTable_Zh2.tt_gcs_type = "0";//Столбец 5, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                stringOfTable_Zh2.technological_number_tt_gks = "0";//Столбец 6, технологический номер (П,Н,7,8,7А,8А)*/
                }
                stringOfTable_Zh2.diameter_exterlal = ObjWorkSheet.Cells[i, 7].Text;//Столбец 7, наружный диаметр
                stringOfTable_Zh2.is_underground = ObjWorkSheet.Cells[i, 8].Text;//Столбец 8, "истина", если подземный, "ложь", если надземный

                try
                {
                    dtfv = DateTime.FromOADate(Convert.ToDouble(ObjWorkSheet.Cells[i, 9].Value2));
                    if (Convert.ToDouble(ObjWorkSheet.Cells[i, 9].Value2) > 40000)
                    {
                        stringOfTable_Zh2.date_of_examination = dtfv;//Столбец 9, дата завершения обследования
                    }
                }
                catch (Exception){}
               

                stringOfTable_Zh2.is_insulation_removed = ObjWorkSheet.Cells[i, 10].Text;//Столбец 10, "истина", если изоляцию снимали, "ложь", если изоляцию не снимали
                stringOfTable_Zh2.number_of_diagnosed_areal = ObjWorkSheet.Cells[i, 11].Text;//Столбец 11, номер продиагностированного участка (уникальное имя диагностированного участка)
                stringOfTable_Zh2.length_of_diagnosed_areal = ObjWorkSheet.Cells[i, 12].Text;//Столбец 12, протяженность диагностированного участка, м.
                stringOfTable_Zh2.diagnostician = ObjWorkSheet.Cells[i, 13].Text;//Столбец 13, исполнитель работ по диагностике
                //stringOfTable_Zh2.diagnostician = "0";//Столбец 13, исполнитель работ по диагностике
                stringOfTable_Zh2.is_external_diagnostics = ObjWorkSheet.Cells[i, 14].Text;//Столбец 14, "истина", если проводилось наружное обследование, "ложь", если не проводилось
                stringOfTable_Zh2.is_VTD_diagnostics = ObjWorkSheet.Cells[i, 15].Text;//Столбец 15, "истина", если проводилось ВТД, "ложь", если не проводилось
                stringOfTable_Zh2.methods_of_diagnostics = ObjWorkSheet.Cells[i, 16].Text;//Столбец 16, методы диагностики
                //stringOfTable_Zh2.methods_of_diagnostics = "0";//Столбец 16, методы диагностики
                stringOfTable_Zh2.number_of_diagnosed_element = ObjWorkSheet.Cells[i, 17].Text;//Столбец 17, номер продиагностированного элемента
                stringOfTable_Zh2.type_of_element = ObjWorkSheet.Cells[i, 18].Text;//Столбец 18, тип обследованного элемента (КСС, отвод, переход, тройник, труба);
                stringOfTable_Zh2.element_design_in_Zh3 = ObjWorkSheet.Cells[i, 19].Text;// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3
                //stringOfTable_Zh2.element_design_in_Zh3 = "0";// Cтолбец 19, конструкция элемента в соответствии с таблицей Ж3
                stringOfTable_Zh2.conclusion_of_diagnostics = ObjWorkSheet.Cells[i, 20].Text;//Cтолбец 20, заключение по презультатам обследования (Ремонт, Замена, Эксплуатация)
                {
                    //stringOfTable_Zh2.binding_object = ObjWorkSheet.Cells[i, 21].Text;//Cтолбец 21, объект привязки
                    //stringOfTable_Zh2.distance_to_binding_object = ObjWorkSheet.Cells[i, 22].Text;//Cтолбец 22, расстояние до объекта привязки
                    //stringOfTable_Zh2.angular_orientation_of_weld_1 = ObjWorkSheet.Cells[i, 23].Text;//Cтолбец 23, угловая ориентация продольного шва №1
                    //stringOfTable_Zh2.angular_orientation_of_weld_2 = ObjWorkSheet.Cells[i, 24].Text;//Cтолбец 24, угловая ориентация продольного шва №2
                }
                stringOfTable_Zh2.binding_object = ObjWorkSheet.Cells[i, 21].Text;//Cтолбец 21, объект привязки
                stringOfTable_Zh2.distance_to_binding_object = ObjWorkSheet.Cells[i, 22].Text;//Cтолбец 22, расстояние до объекта привязки
                stringOfTable_Zh2.angular_orientation_of_weld_1 = ObjWorkSheet.Cells[i, 23].Text;//Cтолбец 23, угловая ориентация продольного шва №1
                stringOfTable_Zh2.angular_orientation_of_weld_2 = ObjWorkSheet.Cells[i, 24].Text;//Cтолбец 24, угловая ориентация продольного шва №2
                stringOfTable_Zh2.element_thikness = ObjWorkSheet.Cells[i, 25].Text;//Cтолбец 25, толщина стенки элемента, мм
                stringOfTable_Zh2.element_length = ObjWorkSheet.Cells[i, 26].Text;//Столбец 26, длина элемента, мм
                {
                    //stringOfTable_Zh2.element_ovality = ObjWorkSheet.Cells[i, 27].Text;//Столбец 27, величина овальности, %
                    //stringOfTable_Zh2.length_of_coating_peeling = ObjWorkSheet.Cells[i, 28].Text;//Столбец 28, протяженность отслаивания защитного покрытия
                    //stringOfTable_Zh2.beginning_of_coating_peeling = ObjWorkSheet.Cells[i, 29].Text;//Столбец 29, координата начала отслаивания покрытия
                }
                stringOfTable_Zh2.element_ovality = ObjWorkSheet.Cells[i, 27].Text;//Столбец 27, величина овальности, %
                stringOfTable_Zh2.length_of_coating_peeling = ObjWorkSheet.Cells[i, 28].Text;//Столбец 28, протяженность отслаивания защитного покрытия
                stringOfTable_Zh2.beginning_of_coating_peeling = ObjWorkSheet.Cells[i, 29].Text;//Столбец 29, координата начала отслаивания покрытия
                stringOfTable_Zh2.length_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 30].Text;//Столбец 30, протяженность непроконтролированных зон
                {
                    //stringOfTable_Zh2.coordinates_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 31].Text;//Столбец 31, координаты непроконтролированныз зон
                }
                stringOfTable_Zh2.coordinates_of_uncontrolled_zones = ObjWorkSheet.Cells[i, 31].Text;//Столбец 31, координаты непроконтролированныз зон
                stringOfTable_Zh2.number_of_defect = ObjWorkSheet.Cells[i, 32].Text;//Столбец 32, номер дефекта, п/п
                stringOfTable_Zh2.type_of_defect = ObjWorkSheet.Cells[i, 33].Text;//Столбец 33, вид дефекта (Смещение кромок, подрез, закат и т.д.)
                stringOfTable_Zh2.defect_begin_coordinates = ObjWorkSheet.Cells[i, 34].Text;//Столбец 34, координаты начала дефекта
                stringOfTable_Zh2.defect_end_coordinates = ObjWorkSheet.Cells[i, 35].Text;//Столбец 35, координаты конца дефекта
                stringOfTable_Zh2.defect_begin_orientation = ObjWorkSheet.Cells[i, 36].Text;//Столбец 36, ориентация начала дефекта
                stringOfTable_Zh2.defect_end_orientation = ObjWorkSheet.Cells[i, 37].Text;//Столбец 37, ориентация конца дефекта
                {
                 /*stringOfTable_Zh2.defect_begin_coordinates = "0";//Столбец 34, координаты начала дефекта
                stringOfTable_Zh2.defect_end_coordinates = "0";//Столбец 35, координаты конца дефекта
                stringOfTable_Zh2.defect_begin_orientation = "0";//Столбец 36, ориентация начала дефекта
                stringOfTable_Zh2.defect_end_orientation = "0";//Столбец 37, ориентация конца дефекта*/
                }
                stringOfTable_Zh2.defect_length = ObjWorkSheet.Cells[i, 38].Text;//Столбец 38, длина дефекта, мм
                stringOfTable_Zh2.defect_width = ObjWorkSheet.Cells[i, 39].Text;//Столбец 39, ширина дефекта, мм
                stringOfTable_Zh2.defect_depth = ObjWorkSheet.Cells[i, 40].Text;//Столбец 40, глубина дефекта, мм
                stringOfTable_Zh2.defect_depth_in_percent = ObjWorkSheet.Cells[i, 41].Text;//Столбец 41, глубина дефекта в процентах
                stringOfTable_Zh2.degree_of_danger = ObjWorkSheet.Cells[i, 42].Text;//Столбец 42, оценка степени опасности
                stringOfTable_Zh2.recommendations_for_defect_elimination = ObjWorkSheet.Cells[i, 43].Text;//Столбец 43, рекомендации по устранению дефекта
                stringOfTable_Zh2.repair_method = ObjWorkSheet.Cells[i, 44].Text;//Столбец 44,  метод ремонта
                try
                {
                    dtfv = DateTime.FromOADate(Convert.ToDouble(ObjWorkSheet.Cells[i, 45].Value2));
                    if (Convert.ToDouble(ObjWorkSheet.Cells[i, 45].Value2) > 40000)
                    {
                        stringOfTable_Zh2.date_of_repair = dtfv;//Столбец 45, дата ремонта
                    }
                }
                catch (Exception) { }

                stringOfTable_Zh2.date_of_repair = dtfv;//Столбец 45, дата ремонта
                {
                    //stringOfTable_Zh2.repair_contractor = ObjWorkSheet.Cells[i, 46].Text;//Столбец 46, исполнитель ремонта
                }
                stringOfTable_Zh2.repair_contractor = ObjWorkSheet.Cells[i, 46].Text;//Столбец 46, исполнитель ремонта
                if (String.IsNullOrWhiteSpace(stringOfTable_Zh2.ID_TP))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    tableZh2_converted.Add(ConvertStringZh2ToCurrentType(stringOfTable_Zh2));
                    textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(tableZh2_converted.Count)+" из "+ Convert.ToString(progressBar1.Maximum)));
                    progressBar1.Invoke(new Action(() => progressBar1.Value++));
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.2 прочитан, количество строк: "+ tableZh2_converted.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));
            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(tableZh2_converted.Count)));
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjExcel.Quit();
            return tableZh2_converted;
        }
        private StringOfTableZh2 ConvertStringZh2ToCurrentType(StringOfTableZh2_string stringOfTable_Zh2)//конвертируем строку Ж2 к нужным типам
        {
            StringOfTableZh2 current = new StringOfTableZh2
            {
                ID_TP = Converting.ConvertToInt(stringOfTable_Zh2.ID_TP),
                lpu_name = stringOfTable_Zh2.lpu_name,
                gcs_name = stringOfTable_Zh2.gcs_name,
                kc_name = stringOfTable_Zh2.kc_name,
                tt_gcs_type = stringOfTable_Zh2.tt_gcs_type,
                technological_number_tt_gks = stringOfTable_Zh2.technological_number_tt_gks,
                diameter_exterlal = Converting.ConvertToDouble(stringOfTable_Zh2.diameter_exterlal),
                is_underground = Converting.GetBoolean(stringOfTable_Zh2.is_underground, "одзем"),
                date_of_examination = stringOfTable_Zh2.date_of_examination,
                is_insulation_removed = Converting.GetBoolean(stringOfTable_Zh2.is_insulation_removed, "снятием"),
                number_of_diagnosed_areal = stringOfTable_Zh2.number_of_diagnosed_areal,
                length_of_diagnosed_areal = Converting.ConvertToDouble(stringOfTable_Zh2.length_of_diagnosed_areal),
                diagnostician = stringOfTable_Zh2.diagnostician,
                is_external_diagnostics = Converting.GetBoolean(stringOfTable_Zh2.is_external_diagnostics, "Да"),
                is_VTD_diagnostics = Converting.GetBoolean(stringOfTable_Zh2.is_VTD_diagnostics, "Да"),
                methods_of_diagnostics = stringOfTable_Zh2.methods_of_diagnostics,
                number_of_diagnosed_element = stringOfTable_Zh2.number_of_diagnosed_element,
                type_of_element = stringOfTable_Zh2.type_of_element,
                element_design_in_Zh3 = stringOfTable_Zh2.element_design_in_Zh3,
                conclusion_of_diagnostics = stringOfTable_Zh2.conclusion_of_diagnostics,
                binding_object = stringOfTable_Zh2.binding_object,
                distance_to_binding_object = stringOfTable_Zh2.distance_to_binding_object,
                angular_orientation_of_weld_1 = Converting.ConvertToDouble(stringOfTable_Zh2.angular_orientation_of_weld_1),
                angular_orientation_of_weld_2 = Converting.ConvertToDouble(stringOfTable_Zh2.angular_orientation_of_weld_2),
                element_thikness = Converting.ConvertToDouble(stringOfTable_Zh2.element_thikness),
                element_length = Converting.ConvertToDouble(stringOfTable_Zh2.element_length),
                element_ovality = Converting.ConvertToDouble(stringOfTable_Zh2.element_ovality),
                length_of_coating_peeling = Converting.ConvertToDouble(stringOfTable_Zh2.length_of_coating_peeling),
                beginning_of_coating_peeling = Converting.ConvertToDouble(stringOfTable_Zh2.beginning_of_coating_peeling),
                length_of_uncontrolled_zones = Converting.ConvertToDouble(stringOfTable_Zh2.length_of_uncontrolled_zones),
                coordinates_of_uncontrolled_zones = Converting.ConvertToDouble(stringOfTable_Zh2.coordinates_of_uncontrolled_zones),
                number_of_defect = Converting.ConvertToInt(stringOfTable_Zh2.number_of_defect),
                type_of_defect = stringOfTable_Zh2.type_of_defect,
                defect_begin_coordinates = Converting.ConvertToDouble(stringOfTable_Zh2.defect_begin_coordinates),
                defect_end_coordinates = Converting.ConvertToDouble(stringOfTable_Zh2.defect_end_coordinates),
                defect_begin_orientation = Converting.ConvertToDouble(stringOfTable_Zh2.defect_begin_orientation),
                defect_end_orientation = Converting.ConvertToDouble(stringOfTable_Zh2.defect_end_orientation),
                defect_length = Converting.ConvertToDouble(stringOfTable_Zh2.defect_length),
                defect_width = Converting.ConvertToDouble(stringOfTable_Zh2.defect_width),
                defect_depth = Converting.ConvertToDouble(stringOfTable_Zh2.defect_depth),
                defect_depth_in_percent = Converting.ConvertToDouble(stringOfTable_Zh2.defect_depth_in_percent),
                degree_of_danger = stringOfTable_Zh2.degree_of_danger,
                recommendations_for_defect_elimination = stringOfTable_Zh2.recommendations_for_defect_elimination,
                repair_method = stringOfTable_Zh2.repair_method,
                date_of_repair = stringOfTable_Zh2.date_of_repair,
                repair_contractor = stringOfTable_Zh2.repair_contractor,

                type_for_sort = String.Concat(stringOfTable_Zh2.ID_TP, "_", stringOfTable_Zh2.number_of_diagnosed_areal, "_", stringOfTable_Zh2.type_of_element, "_", stringOfTable_Zh2.diameter_exterlal)

            };
            return current;
        }
    }


}
