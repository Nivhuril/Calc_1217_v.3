using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace CalcGCS_1217
{
    internal class StringOfTableZh1_string//Класс для хранения одной строки таблицы Ж1 В СТРОКОВОМ ФОРМАТЕ
    {        
        public string ID_TP;//Столбец 1, кодовый номер объекта
        public string lpu_name;//Столбец 2, наименование ЛПУ
        public string gcs_name;//Столбец 3, наименование КС
        public string kc_number;//Столбец 4, номер КЦ
        public string pipeline_name;//Столбец 5, Наименование МГ
        public string coordinate_of_connecting_TT_to_MG;//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
        public string region;//Столбец 7, область
        public string tt_gcs_type;//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
        public string technological_number_tt_gks;//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
        public string inventory_number;//Столбец 10, Инвентарный номер ТТ КС
        public string plot_category;//Столбец 11, Категория участка
        public string maximum_outer_diameter;//Столбец 12, Максимальный наружный диаметр, мм
        public string maximum_wall_thickness;//Столбец 13, Максимальная толщина стенки, мм
        public string total_length_of_pipelines_DN_150_400;//Столбец 14, Общая протяженность трубопроводов 150_400, м
        public string total_length_of_pipelines_DN_500_1400;//Столбец 15, Общая протяженность трубопроводов 500_1400, м
        public string total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 16, Общая протяженность трубопроводов DN 150-1400 в пересчете на DN 1000
        public string list_of_designations_of_pipes_DN_500_1400_as_table_G3;//Столбец 17, Перечень обозначений труб DN 500-1400 в составе ТТ согласно таблице Ж.3
        public string Insulation_cover_type;//Столбец 18, Тип изоляционного покрытия
        public string commissioning_year_of_the_gcs;//Столбец 19, Год ввода в эксплуатацию КЦ
        public string year_of_commissioning_of_th_LPG_section;//Столбец 20, Год ввода в эксплуатацию участка ЛЧ МГ
        public string design_pressure;//Столбец 21, Проектное давление, МПа
        public string maximum_allowed_working_pressure;//Столбец 22, Максимальное разрешенное рабочее давление, МПа
        public string average_number_of_cycles_less_than_10proc;//Столбец 23, Среднее количество циклов в год с различной амплитудой менее 10% от проектного давления
        public string average_number_of_cycles_more_than_10proc;//Столбец 24, Среднее количество циклов в год с различной амплитудой 10% и более от проектного давления
        public string operating_time_in_design_mode;//Столбец 25, Время эксплуатации в различных режимах, лет (в проектном)
        public string operating_time_in_throughput_mode;//Столбец 26, Время эксплуатации в различных режимах, лет (в пропускном)
        public string operating_time_in_disconnected_state_under_pressure;//Столбец 27, Время эксплуатации в различных режимах, лет (в отключенном состоянии под давлением)
        public string operating_time_in_shutdown_state_without_pressure;//Столбец 28, Время эксплуатации в различных режимах, лет (в отключенном состоянии без давления)
        public string if_there_were_non_design_loads;//Столбец 29, Информация о непроектных нагрузках и воздействиях ("Да", если были)
        public string total_length_of_surveyed_pipelines_DN_150_400;//Столбец 30, Общая протяженность обследованных трубопроводов 150_400, м
        public string total_length_of_surveyed_pipelines_DN_500_1400;//Столбец 31, Общая протяженность обследованных трубопроводов 500_1400, м
        public string total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 32, Общая протяженность обследованных трубопроводов DN 150-1400 в пересчете на DN 1000
        public string total_number_of_inspected_joints_500_1400;//Столбец 33, Количество проконтролированных монтажных стыков DN 500-1400 (Всего)
        public string number_of_inspected_joints_500_1400_with_ext_examination;//Столбец 34, Количество проконтролированных монтажных стыков DN 500-1400 (Из них c наружным обследованием методами НК)
        public string number_of_circular_joints_500_1400_defects_requiring_repair;//Столбец 35, Количество кольцевых сварных соединений  DN 500-1400 с дефектами, требующими ремонта, шт.
        public string total_number_of_inspected_pipes_500_1400;//Столбец 36, Количество проконтролированных труб  DN 500-1400 (Всего)
        public string number_of_inspected_pipes_500_1400_with_ext_examination;//Столбец 37, Количество проконтролированных труб  DN 500-1400 (Из них c наружным обследованием методами НК)
        public string number_of_pipes_500_1400_defects_requiring_repair;//Столбец 38,  Количество труб DN 500-1400 с дефектами, требующими ремонта, шт.
        public string total_number_of_inspected_details_500_1400;//Столбец 39, Количество проконтролированных СДТ  DN 500-1400 (Всего)
        public string number_of_inspected_details_500_1400_with_ext_examination;//Столбец 40, Количество проконтролированных СДТ  DN 500-1400 (Из них c наружным обследованием методами НК)
        public string number_of_details_500_1400_defects_requiring_repair;//Столбец 41, Количество СДТ  DN 500-1400 с дефектами, требующими ремонта, шт.
        public string number_of_monitored_valves;//Столбец 42, Суммарное количество проконтролированных единиц ТПА  DN 500-1400, шт.
        public string number_of_valves_requiring_repair;//Столбец 43, Количество единиц ТПА  DN 500-1400, требующих ремонта, шт.
        public string soil_resistivity;//Столбец 44, Удельное сопротивление грунта в районе пролегания трубопровода (минимальное), Ом*м
        public string length_peeling_insulating_coating;//Столбец 45, Суммарная протяженность отслаивания изоляционного покрытия на обследованных участках DN 500-1400, м.
        public string area_of_zones_not_controlled_VTD;//Столбец 46, Суммарная площадь зон, не проконтролированных ВТД, м2
        public string year_of_KRTT;//Столбец 47, Год проведения КРТТ
        public string year_of_VD;//Столбец 48, Год проведения ВТД
        public string total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 49, Протяженность отремонтированных ТТ DN 150-1400, м в пересчете на Ду1000
        public string planned_date_of_EPB;//Столбец 50, Планируемый срок проведения ЭПБ, согласно НД и результатам предыдущей, год
        public string note;//Столбец 51, Примечание
        public int ID_CS;//ID компрессорной станции

        public List<int> IdenticalObjectsInsideTheCC = new List<int>();//Идентичные объекты внутри КЦ
        public List<int> OtherObjectsInsideTheCC = new List<int>();//Другие объекты внутри КЦ
        public List<int> IdenticalObjectsOfTheNeighboringCC = new List<int>();//Идентичные объекты соседнего КЦ
        public List<int> OtherObjectsOfTheNeighboringCC = new List<int>();//Другие объекты соседнего КЦ
        public List<int> ObjectsOfTheCС = new List<int>();//Объекты текуцего КЦ

    }
        partial class Form1 : Form
        {
        private List<StringOfTableZh1_string> get_Zh1_data_from_file_to_string(string fileName)//метод для чтения из файла таблицы Ж.1
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, false, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkBook.Unprotect();
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = "Ж.1";//получаем название вкладки
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];
            List<StringOfTableZh1_string> tableZh1 = new List<StringOfTableZh1_string>();//создаём список строк
            int a = Convert.ToInt32(textBox8.Text) + 14;
            progressBar2.Invoke(new Action(() => progressBar2.Minimum = 0));
            progressBar2.Invoke(new Action(() => progressBar2.Maximum = Convert.ToInt32(textBox8.Text)));
            progressBar2.Invoke(new Action(() => progressBar2.Step = 1));

            //richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных журнала Ж.1...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            int incrementor = 0;//переменная для прогресс - индикатора
            
            for (int i = 14; i < a; i++)
            {
                StringOfTableZh1_string stringOfTable_Zh1 = new StringOfTableZh1_string();//создаём экземпляр класса строки журнала аномалий

                stringOfTable_Zh1.ID_TP = ObjWorkSheet.Cells[i, 1].Text;//Столбец 1, кодовый номер объекта
                stringOfTable_Zh1.lpu_name = ObjWorkSheet.Cells[i, 2].Text;//Столбец 2, наименование ЛПУ
                stringOfTable_Zh1.gcs_name = ObjWorkSheet.Cells[i, 3].Text;//Столбец 3, наименование КС
                stringOfTable_Zh1.kc_number = ObjWorkSheet.Cells[i, 4].Text;//Столбец 4, номер КЦ
                stringOfTable_Zh1.pipeline_name = ObjWorkSheet.Cells[i, 5].Text;//Столбец 5, Наименование МГ
                stringOfTable_Zh1.coordinate_of_connecting_TT_to_MG = ObjWorkSheet.Cells[i, 6].Text;//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
                stringOfTable_Zh1.region = ObjWorkSheet.Cells[i, 7].Text;//Столбец 7, область
                stringOfTable_Zh1.tt_gcs_type = ObjWorkSheet.Cells[i, 8].Text;//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                stringOfTable_Zh1.technological_number_tt_gks = ObjWorkSheet.Cells[i, 9].Text;//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
                stringOfTable_Zh1.inventory_number = ObjWorkSheet.Cells[i, 10].Text;//Столбец 10, Инвентарный номер ТТ КС
                stringOfTable_Zh1.plot_category = ObjWorkSheet.Cells[i, 11].Text;//Столбец 11, Категория участка
                stringOfTable_Zh1.maximum_outer_diameter = ObjWorkSheet.Cells[i, 12].Text;//Столбец 12, Максимальный наружный диаметр, мм
                stringOfTable_Zh1.maximum_wall_thickness = ObjWorkSheet.Cells[i, 13].Text;//Столбец 13, Максимальная толщина стенки, мм
                stringOfTable_Zh1.total_length_of_pipelines_DN_150_400 = ObjWorkSheet.Cells[i, 14].Text;//Столбец 14, Общая протяженность трубопроводов 150_400, м
                stringOfTable_Zh1.total_length_of_pipelines_DN_500_1400 = ObjWorkSheet.Cells[i, 15].Text;//Столбец 15, Общая протяженность трубопроводов 500_1400, м
                stringOfTable_Zh1.total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000 = ObjWorkSheet.Cells[i, 16].Text;//Столбец 16, Общая протяженность трубопроводов DN 150-1400 в пересчете на DN 1000
                stringOfTable_Zh1.list_of_designations_of_pipes_DN_500_1400_as_table_G3 = ObjWorkSheet.Cells[i, 17].Text;//Столбец 17, Перечень обозначений труб DN 500-1400 в составе ТТ согласно таблице Ж.3
                stringOfTable_Zh1.Insulation_cover_type = ObjWorkSheet.Cells[i, 18].Text;//Столбец 18, Тип изоляционного покрытия
                stringOfTable_Zh1.commissioning_year_of_the_gcs = ObjWorkSheet.Cells[i, 19].Text;//Столбец 19, Год ввода в эксплуатацию КЦ
                stringOfTable_Zh1.year_of_commissioning_of_th_LPG_section = ObjWorkSheet.Cells[i, 20].Text;//Столбец 20, Год ввода в эксплуатацию участка ЛЧ МГ
                stringOfTable_Zh1.design_pressure = ObjWorkSheet.Cells[i, 21].Text;//Столбец 21, Проектное давление, МПа
                stringOfTable_Zh1.maximum_allowed_working_pressure = ObjWorkSheet.Cells[i, 22].Text;//Столбец 22, Максимальное разрешенное рабочее давление, МПа
                stringOfTable_Zh1.average_number_of_cycles_less_than_10proc = ObjWorkSheet.Cells[i, 23].Text;//Столбец 23, Среднее количество циклов в год с различной амплитудой менее 10% от проектного давления
                stringOfTable_Zh1.average_number_of_cycles_more_than_10proc = ObjWorkSheet.Cells[i, 24].Text;//Столбец 24, Среднее количество циклов в год с различной амплитудой 10% и более от проектного давления
                stringOfTable_Zh1.operating_time_in_design_mode = ObjWorkSheet.Cells[i, 25].Text;//Столбец 25, Время эксплуатации в различных режимах, лет (в проектном)
                stringOfTable_Zh1.operating_time_in_throughput_mode = ObjWorkSheet.Cells[i, 26].Text;//Столбец 26, Время эксплуатации в различных режимах, лет (в пропускном)
                stringOfTable_Zh1.operating_time_in_disconnected_state_under_pressure = ObjWorkSheet.Cells[i, 27].Text;//Столбец 27, Время эксплуатации в различных режимах, лет (в отключенном состоянии под давлением)
                stringOfTable_Zh1.operating_time_in_shutdown_state_without_pressure = ObjWorkSheet.Cells[i, 28].Text;//Столбец 28, Время эксплуатации в различных режимах, лет (в отключенном состоянии без давления)
                stringOfTable_Zh1.if_there_were_non_design_loads = ObjWorkSheet.Cells[i, 29].Text;//Столбец 29, Информация о непроектных нагрузках и воздействиях ("Да", если были)
                stringOfTable_Zh1.total_length_of_surveyed_pipelines_DN_150_400 = ObjWorkSheet.Cells[i, 30].Text;//Столбец 30, Общая протяженность обследованных трубопроводов 150_400, м
                stringOfTable_Zh1.total_length_of_surveyed_pipelines_DN_500_1400 = ObjWorkSheet.Cells[i, 31].Text;//Столбец 31, Общая протяженность обследованных трубопроводов 500_1400, м
                stringOfTable_Zh1.total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000 = ObjWorkSheet.Cells[i, 32].Text;//Столбец 32, Общая протяженность обследованных трубопроводов DN 150-1400 в пересчете на DN 1000
                stringOfTable_Zh1.total_number_of_inspected_joints_500_1400 = ObjWorkSheet.Cells[i, 33].Text;//Столбец 33, Количество проконтролированных монтажных стыков DN 500-1400 (Всего)
                stringOfTable_Zh1.number_of_inspected_joints_500_1400_with_ext_examination = ObjWorkSheet.Cells[i, 34].Text;//Столбец 34, Количество проконтролированных монтажных стыков DN 500-1400 (Из них c наружным обследованием методами НК)
                stringOfTable_Zh1.number_of_circular_joints_500_1400_defects_requiring_repair = ObjWorkSheet.Cells[i, 35].Text;//Столбец 35, Количество кольцевых сварных соединений  DN 500-1400 с дефектами, требующими ремонта, шт.
                stringOfTable_Zh1.total_number_of_inspected_pipes_500_1400 = ObjWorkSheet.Cells[i, 36].Text;//Столбец 36, Количество проконтролированных труб  DN 500-1400 (Всего)
                stringOfTable_Zh1.number_of_inspected_pipes_500_1400_with_ext_examination = ObjWorkSheet.Cells[i, 37].Text;//Столбец 37, Количество проконтролированных труб  DN 500-1400 (Из них c наружным обследованием методами НК)
                stringOfTable_Zh1.number_of_pipes_500_1400_defects_requiring_repair = ObjWorkSheet.Cells[i, 38].Text;//Столбец 38,  Количество труб DN 500-1400 с дефектами, требующими ремонта, шт.
                stringOfTable_Zh1.total_number_of_inspected_details_500_1400 = ObjWorkSheet.Cells[i, 39].Text;//Столбец 39, Количество проконтролированных СДТ  DN 500-1400 (Всего)
                stringOfTable_Zh1.number_of_inspected_details_500_1400_with_ext_examination = ObjWorkSheet.Cells[i, 40].Text;//Столбец 40, Количество проконтролированных СДТ  DN 500-1400 (Из них c наружным обследованием методами НК)
                stringOfTable_Zh1.number_of_details_500_1400_defects_requiring_repair = ObjWorkSheet.Cells[i, 41].Text;//Столбец 41, Количество СДТ  DN 500-1400 с дефектами, требующими ремонта, шт.
                stringOfTable_Zh1.number_of_monitored_valves = ObjWorkSheet.Cells[i, 42].Text;//Столбец 42, Суммарное количество проконтролированных единиц ТПА  DN 500-1400, шт.
                stringOfTable_Zh1.number_of_valves_requiring_repair = ObjWorkSheet.Cells[i, 43].Text;//Столбец 43, Количество единиц ТПА  DN 500-1400, требующих ремонта, шт.
                stringOfTable_Zh1.soil_resistivity = ObjWorkSheet.Cells[i, 44].Text;//Столбец 44, Удельное сопротивление грунта в районе пролегания трубопровода (минимальное), Ом*м
                stringOfTable_Zh1.length_peeling_insulating_coating = ObjWorkSheet.Cells[i, 45].Text;//Столбец 45, Суммарная протяженность отслаивания изоляционного покрытия на обследованных участках DN 500-1400, м.
                stringOfTable_Zh1.area_of_zones_not_controlled_VTD = ObjWorkSheet.Cells[i, 46].Text;//Столбец 46, Суммарная площадь зон, не проконтролированных ВТД, м2
                stringOfTable_Zh1.year_of_KRTT = ObjWorkSheet.Cells[i, 47].Text;//Столбец 47, Год проведения КРТТ
                stringOfTable_Zh1.year_of_VD = ObjWorkSheet.Cells[i, 48].Text;//Столбец 48, Год проведения ВТД
                stringOfTable_Zh1.total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000 = ObjWorkSheet.Cells[i, 49].Text;//Столбец 49, Протяженность отремонтированных ТТ DN 150-1400, м в пересчете на Ду1000
                stringOfTable_Zh1.planned_date_of_EPB = ObjWorkSheet.Cells[i, 50].Text;//Столбец 50, Планируемый срок проведения ЭПБ, согласно НД и результатам предыдущей, год
                stringOfTable_Zh1.note = ObjWorkSheet.Cells[i, 51].Text;//Столбец 51, Примечание
                stringOfTable_Zh1.ID_CS= Converting.ConvertToInt(ObjWorkSheet.Cells[i, 98].Text);//Считываем ID компрессорных станций, расставленных согласно таблице Ж.4

                stringOfTable_Zh1.IdenticalObjectsInsideTheCC = Converting.GetNumbersOfTT(ObjWorkSheet.Cells[i, 72].Text);

               stringOfTable_Zh1.OtherObjectsInsideTheCC = Converting.GetNumbersOfTT(ObjWorkSheet.Cells[i, 75].Text);

               stringOfTable_Zh1.IdenticalObjectsOfTheNeighboringCC = Converting.GetNumbersOfTT(ObjWorkSheet.Cells[i, 80].Text);

               stringOfTable_Zh1.OtherObjectsOfTheNeighboringCC = Converting.GetNumbersOfTT(ObjWorkSheet.Cells[i, 85].Text);

               stringOfTable_Zh1.ObjectsOfTheCС = Converting.GetNumbersOfTT(ObjWorkSheet.Cells[i, 97].Text);



                tableZh1.Add(stringOfTable_Zh1);//добавляем заполненный экземпляр класса к списку
                textBox1.Invoke(new Action(() => textBox1.Text = Convert.ToString(tableZh1.Count)));
                progressBar2.Invoke(new Action(() => progressBar2.Value++));
                //textBox1.Text = Convert.ToString(tableZh1.Count);

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    //richTextBox1.AppendText("*");
                    incrementor = 0;

                }
            }
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.1 прочитан, количество строк:" + tableZh1.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));
            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.1 прочитан, количество строк:" + tableZh1.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
           
            ObjExcel.Quit();
            return tableZh1;
        }

    }
    }
