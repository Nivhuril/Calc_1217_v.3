using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
namespace CalcGCS_1217
{
   internal class StringOfTableZh1 //Класс для хранения одной строки таблицы Ж1 В СТРОКОВОМ ФОРМАТЕ
    {
        
        public int ID_TP;//Столбец 1, кодовый номер объекта
        public string lpu_name;//Столбец 2, наименование ЛПУ
        public string gcs_name;//Столбец 3, наименование КС
        public string kc_number;//Столбец 4, номер КЦ
        public string pipeline_name;//Столбец 5, Наименование МГ
        public double coordinate_of_connecting_TT_to_MG;//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
        public string region;//Столбец 7, область
        public string tt_gcs_type;//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
        public string technological_number_tt_gks;//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
        public string inventory_number;//Столбец 10, Инвентарный номер ТТ КС
        public string plot_category;//Столбец 11, Категория участка
        public double maximum_outer_diameter;//Столбец 12, Максимальный наружный диаметр, мм
        public double maximum_wall_thickness;//Столбец 13, Максимальная толщина стенки, мм
        public double total_length_of_pipelines_DN_150_400;//Столбец 14, Общая протяженность трубопроводов 150_400, м
        public double total_length_of_pipelines_DN_500_1400;//Столбец 15, Общая протяженность трубопроводов 500_1400, м
        public double total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 16, Общая протяженность трубопроводов DN 150-1400 в пересчете на DN 1000
        public string list_of_designations_of_pipes_DN_500_1400_as_table_G3;//Столбец 17, Перечень обозначений труб DN 500-1400 в составе ТТ согласно таблице Ж.3
        public string Insulation_cover_type;//Столбец 18, Тип изоляционного покрытия
        public double commissioning_year_of_the_gcs;//Столбец 19, Год ввода в эксплуатацию КЦ
        public double year_of_commissioning_of_th_LPG_section;//Столбец 20, Год ввода в эксплуатацию участка ЛЧ МГ
        public double design_pressure;//Столбец 21, Проектное давление, МПа
        public double maximum_allowed_working_pressure;//Столбец 22, Максимальное разрешенное рабочее давление, МПа
        public int average_number_of_cycles_less_than_10proc;//Столбец 23, Среднее количество циклов в год с различной амплитудой менее 10% от проектного давления
        public int average_number_of_cycles_more_than_10proc;//Столбец 24, Среднее количество циклов в год с различной амплитудой 10% и более от проектного давления
        public double operating_time_in_design_mode;//Столбец 25, Время эксплуатации в различных режимах, лет (в проектном)
        public double operating_time_in_throughput_mode;//Столбец 26, Время эксплуатации в различных режимах, лет (в пропускном)
        public double operating_time_in_disconnected_state_under_pressure;//Столбец 27, Время эксплуатации в различных режимах, лет (в отключенном состоянии под давлением)
        public double operating_time_in_shutdown_state_without_pressure;//Столбец 28, Время эксплуатации в различных режимах, лет (в отключенном состоянии без давления)
        public bool if_there_were_non_design_loads;//Столбец 29, Информация о непроектных нагрузках и воздействиях ("Да", если были)
        public double total_length_of_surveyed_pipelines_DN_150_400;//Столбец 30, Общая протяженность обследованных трубопроводов 150_400, м
        public double total_length_of_surveyed_pipelines_DN_500_1400;//Столбец 31, Общая протяженность обследованных трубопроводов 500_1400, м
        public double total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 32, Общая протяженность обследованных трубопроводов DN 150-1400 в пересчете на DN 1000
        public int total_number_of_inspected_joints_500_1400;//Столбец 33, Количество проконтролированных монтажных стыков DN 500-1400 (Всего)
        public int number_of_inspected_joints_500_1400_with_ext_examination;//Столбец 34, Количество проконтролированных монтажных стыков DN 500-1400 (Из них c наружным обследованием методами НК)
        public int number_of_circular_joints_500_1400_defects_requiring_repair;//Столбец 35, Количество кольцевых сварных соединений  DN 500-1400 с дефектами, требующими ремонта, шт.
        public int total_number_of_inspected_pipes_500_1400;//Столбец 36, Количество проконтролированных труб  DN 500-1400 (Всего)
        public int number_of_inspected_pipes_500_1400_with_ext_examination;//Столбец 37, Количество проконтролированных труб  DN 500-1400 (Из них c наружным обследованием методами НК)
        public int number_of_pipes_500_1400_defects_requiring_repair;//Столбец 38,  Количество труб DN 500-1400 с дефектами, требующими ремонта, шт.
        public int total_number_of_inspected_details_500_1400;//Столбец 39, Количество проконтролированных СДТ  DN 500-1400 (Всего)
        public int number_of_inspected_details_500_1400_with_ext_examination;//Столбец 40, Количество проконтролированных СДТ  DN 500-1400 (Из них c наружным обследованием методами НК)
        public int number_of_details_500_1400_defects_requiring_repair;//Столбец 41, Количество СДТ  DN 500-1400 с дефектами, требующими ремонта, шт.
        public int number_of_monitored_valves;//Столбец 42, Суммарное количество проконтролированных единиц ТПА  DN 500-1400, шт.
        public int number_of_valves_requiring_repair;//Столбец 43, Количество единиц ТПА  DN 500-1400, требующих ремонта, шт.
        public double soil_resistivity;//Столбец 44, Удельное сопротивление грунта в районе пролегания трубопровода (минимальное), Ом*м
        public double length_peeling_insulating_coating;//Столбец 45, Суммарная протяженность отслаивания изоляционного покрытия на обследованных участках DN 500-1400, м.
        public double area_of_zones_not_controlled_VTD;//Столбец 46, Суммарная площадь зон, не проконтролированных ВТД, м2
        public double year_of_KRTT;//Столбец 47, Год проведения КРТТ
        public double year_of_VD;//Столбец 48, Год проведения ВТД
        public double total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000;//Столбец 49, Протяженность отремонтированных ТТ DN 150-1400, м в пересчете на Ду1000
        public double planned_date_of_EPB;//Столбец 50, Планируемый срок проведения ЭПБ, согласно НД и результатам предыдущей, год
        public string note;//Столбец 51, Примечание
        public int ID_CS;//ID компрессорной станции

        public List<int> IdenticalObjectsInsideTheCC = new List<int>();//Идентичные объекты внутри КЦ
        public List<int> OtherObjectsInsideTheCC = new List<int>();//Другие объекты внутри КЦ
        public List<int> IdenticalObjectsOfTheNeighboringCC = new List<int>();//Идентичные объекты соседнего КЦ
        public List<int> OtherObjectsOfTheNeighboringCC = new List<int>();//Другие объекты соседнего КЦ
        public List<int> ObjectsOfTheCС = new List<int>();//Объекты текущего КЦ
        public List<int> NeighboringTheCС = new List<int>();//ID соседних КЦ

        //расчетные значения для таблицы Ж1
        public double sum_of_monitored_areas_NK_and_VTD;//Сумма обследованных труб СДТ и ТПА в составе ТТ методами НК и/или ВТД
        public double sum_of_monitored_areas_NK_only;//Сумма обследованных труб СДТ и ТПА в составе ТТ только методами НК
        public double reduced_length_uncontrolled_NK_and_VTD;//Приведённая длина непроконтролированных зон НК и ВТД
        public double reduced_length_uncontrolled_NK_only;//Приведённая длина непроконтролированных зон НК
        public double frequency_indicator_of_completeness_o_initial_data;//Р5//Частотный показатель полноты исходных данных
        public double summ_of_rangs;//Р10//сумма рангов опасности для ТТ
        public double summ_of_rangs_joint;//Р13//сумма рангов опасности сварных соединений для ТТ
        public double summ_of_rangs_KRN;//сумма рангов опасности сварных соединений для ТТ для дефектов КРН
        //public double summ_of_rangs_KRN_with_MT;//сумма рангов опасности сварных соединений для ТТ для дефектов КРН


        public double summ_of_rangs_KRN_A4;//результаты расчета по формуле А4-столбец 70 таблицы Ж1
        public double summ_of_rangs_KRN_A5;//результаты расчета по формуле А5
        public double sum_of_length_Identical_Objects_Inside_The_CC;//общая протяженность идентичных ТТ внутри КЦ
        public double sum_of_length_other_Objects_Inside_The_CC;//общая протяженность других ТТ внутри КЦ
        public double sum_of_length_Identical_Objects_Neighboring_The_CC;//общая протяженность идентичных ТТ соседних КЦ
        public double sum_of_length_other_Objects_Neighboring_The_CC;//общая протяженность идентичных ТТ соседних КЦ


        public double parameter_according_to_formula_6_5;//Р9//параметр по формуле 6.5
        public double wanted_number_of_joints;//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
        public double base_number_of_joints;//P14//Базовое число КСС, формула 6.7
        public double parameter_taking_into_account_the_technical_condition_of_KSS_TT;//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
        public double parameter_according_to_formula_6_2_1;//Р9//параметр по формуле 6.2.1
        public double service_life_of_the_workshop_of_string_Zh1;//срок эксплуатации текущего объекта таблицы Ж.1 с момента ввода или капремонта, лет.
        
        public double service_life_of_the_workshop_of_CC;//срок эксплуатации текущего КЦ с момента ввода или капремонта, лет. (максимальное время из всех объектов доанного КЦ)
        public double service_life_of_the_workshop_Neighboring_The_CC;//Максимальный из сроков эксплуатации соседнего КЦ

        public double service_life_of_the_pipeline_of_CC;//срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет.
        public double service_life_of_the_pipeline_Neighboring_The_CC;//срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет.


        public double parameter_according_formula_A7;//Величина коэффициента дефектности идентичных ТТ соседних КЦ (По формуле А7)
        public double parameter_according_formula_A8;//Величина коэффициента дефектности других ТТ соседних КЦ (По формуле А8)
        public double parameter_according_formula_A10;//Коэффициент дефектности прилегающих участков МГ текущего КЦ (По формуле А10)
        public double parameter_according_formula_A11;//Коэффициент дефектности прилегающих участков МГ соседних КЦ (По формуле А11)
        public double parameter_according_formula_A9;//Величина показателя фактора дефектности прилегающих участков МГ(По формуле А9)
        public double parameter_according_formula_A13;//Коэффициент аварийности ЛЧ МГ, примыкающей к текущему КЦ (По формуле А13)
        public double parameter_according_formula_A12;//Количественная оценка фактора аварийности Ra (По формуле А12)
        public double parameter_according_formula_A14;//Коэффициент аварийности соседних КЦ (По формуле А14)
        public double parameter_according_formula_A15;//Количественная оценка фактора наработки Rн (По формуле А15)


        public int number_of_defects_KRN;//количество дефектов КРН на объекте
    }
    partial class Form1 : Form
    {
        private List<StringOfTableZh1> convert_Zh1_to_current_type(List<StringOfTableZh1_string> input)//конвертация сырой таблицы в таблицу с нужными типами данных
        {
            List<StringOfTableZh1> result = new List<StringOfTableZh1>();
            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала Ж.1...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 0; i < input.Count; i++)
            {
                StringOfTableZh1 current = new StringOfTableZh1();
                current.ID_TP = Converting.ConvertToInt(input[i].ID_TP);
                current.lpu_name = input[i].lpu_name;//Столбец 2, наименование ЛПУ
                current.gcs_name = input[i].gcs_name;//Столбец 3, наименование КС
                current.kc_number = input[i].kc_number;//Столбец 4, номер КЦ
                current.pipeline_name = input[i].pipeline_name;//Столбец 5, Наименование МГ
                current.coordinate_of_connecting_TT_to_MG = Converting.ConvertToDouble(input[i].coordinate_of_connecting_TT_to_MG);//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
                current.region = input[i].region;//Столбец 7, область
                current.tt_gcs_type = input[i].tt_gcs_type;//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                current.technological_number_tt_gks = input[i].technological_number_tt_gks;//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
                current.inventory_number = input[i].inventory_number;//Столбец 10, Инвентарный номер ТТ КС
                current.plot_category = input[i].plot_category;//Столбец 11, Категория участка
                current.maximum_outer_diameter = Converting.ConvertToDouble(input[i].maximum_outer_diameter);//Столбец 12, Максимальный наружный диаметр, мм
                current.maximum_wall_thickness = Converting.ConvertToDouble(input[i].maximum_wall_thickness);//Столбец 13, Максимальная толщина стенки, мм
                current.total_length_of_pipelines_DN_150_400 = Converting.ConvertToDouble(input[i].total_length_of_pipelines_DN_150_400);//Столбец 14, Общая протяженность трубопроводов 150_400, м
                current.total_length_of_pipelines_DN_500_1400 = Converting.ConvertToDouble(input[i].total_length_of_pipelines_DN_500_1400);//Столбец 15, Общая протяженность трубопроводов 500_1400, м
                current.total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000 = Converting.ConvertToDouble(input[i].total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000);//Столбец 16, Общая протяженность трубопроводов DN 150-1400 в пересчете на DN 1000
                current.list_of_designations_of_pipes_DN_500_1400_as_table_G3 = input[i].list_of_designations_of_pipes_DN_500_1400_as_table_G3;//Столбец 17, Перечень обозначений труб DN 500-1400 в составе ТТ согласно таблице Ж.3
                current.Insulation_cover_type = input[i].Insulation_cover_type;//Столбец 18, Тип изоляционного покрытия
                current.commissioning_year_of_the_gcs = Converting.ConvertToDouble(input[i].commissioning_year_of_the_gcs);//Столбец 19, Год ввода в эксплуатацию КЦ
                current.year_of_commissioning_of_th_LPG_section = Converting.ConvertToDouble(input[i].year_of_commissioning_of_th_LPG_section);//Столбец 20, Год ввода в эксплуатацию участка ЛЧ МГ
                current.design_pressure = Converting.ConvertToDouble(input[i].design_pressure);//Столбец 21, Проектное давление, МПа
                current.maximum_allowed_working_pressure = Converting.ConvertToDouble(input[i].maximum_allowed_working_pressure);//Столбец 22, Максимальное разрешенное рабочее давление, МПа
                current.average_number_of_cycles_less_than_10proc = Converting.ConvertToInt(input[i].average_number_of_cycles_less_than_10proc);//Столбец 23, Среднее количество циклов в год с различной амплитудой менее 10% от проектного давления
                current.average_number_of_cycles_more_than_10proc = Converting.ConvertToInt(input[i].average_number_of_cycles_more_than_10proc);//Столбец 24, Среднее количество циклов в год с различной амплитудой 10% и более от проектного давления
                current.operating_time_in_design_mode = Converting.ConvertToDouble(input[i].operating_time_in_design_mode);//Столбец 25, Время эксплуатации в различных режимах, лет (в проектном)
                current.operating_time_in_throughput_mode = Converting.ConvertToDouble(input[i].operating_time_in_throughput_mode);//Столбец 26, Время эксплуатации в различных режимах, лет (в пропускном)
                current.operating_time_in_disconnected_state_under_pressure = Converting.ConvertToDouble(input[i].operating_time_in_disconnected_state_under_pressure);//Столбец 27, Время эксплуатации в различных режимах, лет (в отключенном состоянии под давлением)
                current.operating_time_in_shutdown_state_without_pressure = Converting.ConvertToDouble(input[i].operating_time_in_shutdown_state_without_pressure);//Столбец 28, Время эксплуатации в различных режимах, лет (в отключенном состоянии без давления)
                current.if_there_were_non_design_loads = Converting.GetBoolean(input[i].if_there_were_non_design_loads, "не было");//Столбец 29, Информация о непроектных нагрузках и воздействиях ("Да", если были)
                current.total_length_of_surveyed_pipelines_DN_150_400 = Converting.ConvertToDouble(input[i].total_length_of_surveyed_pipelines_DN_150_400);//Столбец 30, Общая протяженность обследованных трубопроводов 150_400, м
                current.total_length_of_surveyed_pipelines_DN_500_1400 = Converting.ConvertToDouble(input[i].total_length_of_surveyed_pipelines_DN_500_1400);//Столбец 31, Общая протяженность обследованных трубопроводов 500_1400, м
                current.total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000 = Converting.ConvertToDouble(input[i].total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000);//Столбец 32, Общая протяженность обследованных трубопроводов DN 150-1400 в пересчете на DN 1000
                current.total_number_of_inspected_joints_500_1400 = Converting.ConvertToInt(input[i].total_number_of_inspected_joints_500_1400);//Столбец 33, Количество проконтролированных монтажных стыков DN 500-1400 (Всего)
                current.number_of_inspected_joints_500_1400_with_ext_examination = Converting.ConvertToInt(input[i].number_of_inspected_joints_500_1400_with_ext_examination);//Столбец 34, Количество проконтролированных монтажных стыков DN 500-1400 (Из них c наружным обследованием методами НК)
                current.number_of_circular_joints_500_1400_defects_requiring_repair = Converting.ConvertToInt(input[i].number_of_circular_joints_500_1400_defects_requiring_repair);//Столбец 35, Количество кольцевых сварных соединений  DN 500-1400 с дефектами, требующими ремонта, шт.
                current.total_number_of_inspected_pipes_500_1400 = Converting.ConvertToInt(input[i].total_number_of_inspected_pipes_500_1400);//Столбец 36, Количество проконтролированных труб  DN 500-1400 (Всего)
                current.number_of_inspected_pipes_500_1400_with_ext_examination = Converting.ConvertToInt(input[i].number_of_inspected_pipes_500_1400_with_ext_examination);//Столбец 37, Количество проконтролированных труб  DN 500-1400 (Из них c наружным обследованием методами НК)
                current.number_of_pipes_500_1400_defects_requiring_repair = Converting.ConvertToInt(input[i].number_of_pipes_500_1400_defects_requiring_repair);//Столбец 38,  Количество труб DN 500-1400 с дефектами, требующими ремонта, шт.
                current.total_number_of_inspected_details_500_1400 = Converting.ConvertToInt(input[i].total_number_of_inspected_details_500_1400);//Столбец 39, Количество проконтролированных СДТ  DN 500-1400 (Всего)
                current.number_of_inspected_details_500_1400_with_ext_examination = Converting.ConvertToInt(input[i].number_of_inspected_details_500_1400_with_ext_examination);//Столбец 40, Количество проконтролированных СДТ  DN 500-1400 (Из них c наружным обследованием методами НК)
                current.number_of_details_500_1400_defects_requiring_repair = Converting.ConvertToInt(input[i].number_of_details_500_1400_defects_requiring_repair);//Столбец 41, Количество СДТ  DN 500-1400 с дефектами, требующими ремонта, шт.
                current.number_of_monitored_valves = Converting.ConvertToInt(input[i].number_of_monitored_valves);//Столбец 42, Суммарное количество проконтролированных единиц ТПА  DN 500-1400, шт.
                current.number_of_valves_requiring_repair = Converting.ConvertToInt(input[i].number_of_valves_requiring_repair);//Столбец 43, Количество единиц ТПА  DN 500-1400, требующих ремонта, шт.
                current.soil_resistivity = Converting.ConvertToDouble(input[i].soil_resistivity);//Столбец 44, Удельное сопротивление грунта в районе пролегания трубопровода (минимальное), Ом*м
                current.length_peeling_insulating_coating = Converting.ConvertToDouble(input[i].length_peeling_insulating_coating);//Столбец 45, Суммарная протяженность отслаивания изоляционного покрытия на обследованных участках DN 500-1400, м.
                current.area_of_zones_not_controlled_VTD = Converting.ConvertToDouble(input[i].area_of_zones_not_controlled_VTD);//Столбец 46, Суммарная площадь зон, не проконтролированных ВТД, м2
                current.year_of_KRTT = Converting.ConvertToDouble(input[i].year_of_KRTT);//Столбец 47, Год проведения КРТТ
                current.year_of_VD = Converting.ConvertToDouble(input[i].year_of_VD);//Столбец 48, Год проведения ВТД
                current.total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000 = Converting.ConvertToDouble(input[i].total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000);//Столбец 49, Протяженность отремонтированных ТТ DN 150-1400, м в пересчете на Ду1000
                current.planned_date_of_EPB = Converting.ConvertToDouble(input[i].planned_date_of_EPB);//Столбец 50, Планируемый срок проведения ЭПБ, согласно НД и результатам предыдущей, год
                current.note = input[i].note;//Столбец 51, Примечание
                current.ID_CS = input[i].ID_CS;//Считываем ID компрессорных станций, расставленных согласно таблице Ж.4

                current.IdenticalObjectsInsideTheCC = input[i].IdenticalObjectsInsideTheCC;
                current.OtherObjectsInsideTheCC = input[i].OtherObjectsInsideTheCC;
                current.IdenticalObjectsOfTheNeighboringCC = input[i].IdenticalObjectsOfTheNeighboringCC;
                current.OtherObjectsOfTheNeighboringCC = input[i].OtherObjectsOfTheNeighboringCC;
                current.ObjectsOfTheCС = input[i].ObjectsOfTheCС;

                result.Add(current);
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 29)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала Ж.1 обработан, количество строк:" + result.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_monitored_areas_NK_and_or_VTD(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> TableZh1)//Р1//Сумма проконтролированных зон НК и ВТД
        {
            double localLength = 0;

            if (sortMassive[0].is_external_diagnostics | sortMassive[0].is_VTD_diagnostics)
            {
                if (sortMassive[0].type_of_element.Contains("КСС") == false)
                {
                    localLength = sortMassive[0].element_length;
                }
            }

            richTextBox1.AppendText(Environment.NewLine + "Номер элемента: 0, "+"тип: "+sortMassive[0].type_of_element+" Длина: "+ sortMassive[0].element_length);

            for (int i = 1; i < sortMassive.Count - 1; i++)
            {
                richTextBox1.AppendText(Environment.NewLine + "Номер элемента: "+ i + ", " + "тип: " + sortMassive[i].type_of_element + " Длина: " + sortMassive[i].element_length);
                
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }

                bool isNewID_TP = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].ID_TP, sortMassive[i - 1].ID_TP) == false)
                {
                    isNewID_TP = true;
                }

                bool isJoint = false;//истина, если текущий элемент является КСС
                if (sortMassive[i].type_of_element.Contains("КСС"))
                {
                    isJoint = true;
                }

                if (isNewElement & isNewID_TP == false & isJoint == false)
                {
                    if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                    {
                        localLength += sortMassive[i].element_length;
                    }
                }

                if (isNewElement & isNewID_TP)
                {
                    int strNumber = Converting.getStringNumber(sortMassive[i - 1].ID_TP);
                    TableZh1[strNumber].sum_of_monitored_areas_NK_and_VTD += localLength / 1000;
                    localLength = 0;
                    if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                    {
                        if (isJoint == false)
                        {
                            localLength = sortMassive[i].element_length;
                        }

                    }
                }

                if (isNewElement == false & isNewID_TP)
                {
                    int strNumber = Converting.getStringNumber(sortMassive[i - 1].ID_TP);
                    TableZh1[strNumber].sum_of_monitored_areas_NK_and_VTD += localLength / 1000;
                    localLength = 0;
                    if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                    {
                        if (isJoint == false)
                        {
                            localLength = sortMassive[i].element_length;
                        }

                    }
                }

            }
            //richTextBox1.AppendText(Environment.NewLine + "get_sum_of_monitored_areas_NK_and_or_VTD=" + TableZh1.Count);
            return TableZh1;
        }
        private List<StringOfTableZh1> get_sum_of_monitored_areas_NK_and_or_VTD_new(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> TableZh1)//Р1//Сумма проконтролированных зон НК и ВТД
        {
            double localLength = 0;
            List<StringOfTableZh2> Massive = new List<StringOfTableZh2>();

            for (int i = 0; i < sortMassive.Count; i++)
            {
                if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                {
                    if (sortMassive[i].type_of_element.Contains("КСС") == false)
                    {
                        Massive.Add(sortMassive[i]);
                    }
                }
                else if (sortMassive[i].is_external_diagnostics && sortMassive[i].is_VTD_diagnostics)
                {
                    if (sortMassive[i].type_of_element.Contains("КСС") == false)
                    {
                        Massive.Add(sortMassive[i]);
                    }
                }
            }
            richTextBox1.AppendText(Environment.NewLine + "Длина массива после отсеивания: "+ Massive.Count);
            Massive = getSortMassiveZh2(Massive);
            richTextBox1.AppendText(Environment.NewLine + "Длина массива после сортировки: " + Massive.Count);
            localLength = Massive[0].element_length;

            for (int i = 1; i < Massive.Count; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(Massive[i].number_of_diagnosed_element, Massive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                bool isNewID_TP = false;//истина, если текущий элемент не равен предыдущему
                
                if (String.Equals(Massive[i].ID_TP, Massive[i - 1].ID_TP) == false)
                {
                    isNewID_TP = true;
                }
                
                if (isNewElement & isNewID_TP == false)//новый элепмент в составе текущего ТТ. Увеличиваем сумму.
                {
                    localLength += Massive[i].element_length;
                }

                if (isNewID_TP && i != Massive.Count - 1)
                {
                    TableZh1[Converting.getStringNumber(Massive[i - 1].ID_TP)].sum_of_monitored_areas_NK_and_VTD += localLength / 1000;
                    
                    localLength = Massive[i].element_length;
                }
                
                if (i== Massive.Count-1 && isNewID_TP==false)
                {
                    TableZh1[Converting.getStringNumber(Massive[i].ID_TP)].sum_of_monitored_areas_NK_and_VTD += localLength / 1000;
                }

            }
            
            return TableZh1;
        }
        private List<StringOfTableZh1> get_sum_of_monitored_areas_NK_only_new(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> TableZh1)//Р1//Сумма проконтролированных зон НК и ВТД
        {
            double localLength = 0;
            List<StringOfTableZh2> Massive = new List<StringOfTableZh2>();

            for (int i = 0; i < sortMassive.Count; i++)
            {
                if (sortMassive[i].is_external_diagnostics && sortMassive[i].is_VTD_diagnostics==false)
                {
                    if (sortMassive[i].type_of_element.Contains("КСС") == false)
                    {
                        Massive.Add(sortMassive[i]);
                    }
                }               
            }

            Massive = getSortMassiveZh2(Massive);
            if (Massive.Count>0)
            {
                localLength = Massive[0].element_length;
                for (int i = 1; i < Massive.Count; i++)
                {
                    bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                    if (String.Equals(Massive[i].number_of_diagnosed_element, Massive[i - 1].number_of_diagnosed_element) == false)
                    {
                        isNewElement = true;
                    }
                    bool isNewID_TP = false;//истина, если текущий трубопровод не равен предыдущему

                    if (String.Equals(Massive[i].ID_TP, Massive[i - 1].ID_TP) == false)
                    {
                        isNewID_TP = true;
                    }

                    if (isNewElement & isNewID_TP == false)//новый элепмент в составе текущего ТТ. Увеличиваем сумму.
                    {
                        localLength += Massive[i].element_length;
                    }

                    if (isNewID_TP && i != Massive.Count - 1)
                    {
                        TableZh1[Converting.getStringNumber(Massive[i - 1].ID_TP)].sum_of_monitored_areas_NK_only += localLength / 1000;

                        localLength = Massive[i].element_length;
                    }

                    if (i == Massive.Count - 1 && isNewID_TP == false)
                    {
                        TableZh1[Converting.getStringNumber(Massive[i].ID_TP)].sum_of_monitored_areas_NK_only += localLength / 1000;
                    }
                }
            }
            
            return TableZh1;
        }
        private List<StringOfTableZh1> get_sum_of_monitored_areas_NK_only(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)//Р3//Сумма проконтролированных зон НК и ВТД
        {
            double localLength = 0;
            if (sortMassive[0].is_external_diagnostics & sortMassive[0].is_VTD_diagnostics == false)
            {
                if (sortMassive[0].type_of_element.Contains("КСС") == false)
                {
                    localLength = sortMassive[0].element_length;
                }
            }

            for (int i = 1; i < sortMassive.Count - 1; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                bool isNewID_TP = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].ID_TP, sortMassive[i - 1].ID_TP) == false)
                {
                    isNewID_TP = true;
                }
                bool isJoint = false;//истина, если текущий элемент является КСС
                if (sortMassive[i].type_of_element.Contains("КСС"))
                {
                    isJoint = true;
                }
                if (isNewElement & isNewID_TP == false & isJoint == false)
                {
                    if (sortMassive[i].is_external_diagnostics & sortMassive[i].is_VTD_diagnostics == false)
                    {
                        localLength += sortMassive[i].element_length;
                    }
                }

                if (isNewElement & isNewID_TP)
                {
                    int strNumber = Converting.getStringNumber(sortMassive[i - 1].ID_TP);
                    result[strNumber].sum_of_monitored_areas_NK_only += localLength / 1000;
                    localLength = 0;
                    if (sortMassive[i].is_external_diagnostics & sortMassive[i].is_VTD_diagnostics == false)
                    {
                        if (isJoint == false)
                        {
                            localLength = sortMassive[i].element_length;
                        }

                    }
                }

                if (isNewElement == false & isNewID_TP)
                {
                    int strNumber = Converting.getStringNumber(sortMassive[i - 1].ID_TP);
                    result[strNumber].sum_of_monitored_areas_NK_only += localLength / 1000;
                    localLength = 0;
                    if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                    {
                        if (isJoint == false)
                        {
                            localLength = sortMassive[i].element_length;
                        }

                    }
                }

            }
            //richTextBox1.AppendText(Environment.NewLine + "get_sum_of_monitored_areas_NK_only=" + result.Count);
            return result;
        }
        private List<StringOfTableZh1> get_reduced_length_uncontrolled_NK_and_or_VTD_zones(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)//Р2//!!!!!!!!!!!!!!!!
        {
            int strNumber = 0;
            if (sortMassive[0].is_external_diagnostics | sortMassive[0].is_VTD_diagnostics)
            {
                strNumber = Converting.getStringNumber(sortMassive[0].ID_TP);
                double q = sortMassive[0].length_of_uncontrolled_zones;
                double w = Convert.ToDouble(sortMassive[0].diameter_exterlal);
                double e = (1000 * q) / (Math.PI * w);
                result[strNumber].reduced_length_uncontrolled_NK_and_VTD += e;
            }


            for (int i = 1; i < sortMassive.Count; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                if (sortMassive[i].is_external_diagnostics | sortMassive[i].is_VTD_diagnostics)
                {
                    if (isNewElement)
                    {
                        strNumber = Converting.getStringNumber(sortMassive[i].ID_TP);
                        double q = sortMassive[i].length_of_uncontrolled_zones;
                        double w = Convert.ToDouble(sortMassive[i].diameter_exterlal);
                        double e = (1000 * q) / (Math.PI * w);
                        result[strNumber].reduced_length_uncontrolled_NK_and_VTD += e;
                    }

                }
            }
            //richTextBox1.AppendText(Environment.NewLine + "get_reduced_length_uncontrolled_NK_and_or_VTD_zones=" + result.Count);
            return result;
        }
        private List<StringOfTableZh1> get_reduced_length_uncontrolled_NK_only(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)//Р4//!!!!!!!!!!!!!!!!
        {
            int strNumber = 0;
            if (sortMassive[0].is_external_diagnostics & sortMassive[0].is_VTD_diagnostics == false)
            {
                strNumber = Converting.getStringNumber(sortMassive[0].ID_TP);
                double q = sortMassive[0].length_of_uncontrolled_zones;
                double w = Convert.ToDouble(sortMassive[0].diameter_exterlal);
                double e = (1000 * q) / (Math.PI * w);
                result[strNumber].reduced_length_uncontrolled_NK_only += e;
            }
            for (int i = 1; i < sortMassive.Count; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                if (sortMassive[i].is_external_diagnostics & sortMassive[i].is_VTD_diagnostics == false & isNewElement)
                {
                    strNumber = Converting.getStringNumber(sortMassive[i].ID_TP);
                    double q = sortMassive[i].length_of_uncontrolled_zones;
                    double w = Convert.ToDouble(sortMassive[i].diameter_exterlal);
                    double e = (1000 * q) / (Math.PI * w);
                    result[strNumber].reduced_length_uncontrolled_NK_only += e;
                }
            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_danger_rangs(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result, List<int> NumbersOfObjects)//Р10//Подсчет суммы максимальных рангов опасности для каждого дефектного элемента в каждом ТТ
        {
            double localsumm = 0;
            double summOfElement = 0;
            localsumm = sortMassive[0].rang_of_danger;
            
            for (int i = 1; i < sortMassive.Count; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                bool isNewID_TP = false;//истина, если текущий объект не равен предыдущему
                if (String.Equals(sortMassive[i].ID_TP, sortMassive[i - 1].ID_TP) == false)
                {
                    isNewID_TP = true;
                }
                bool isJoint = false;//истина, если текущий элемент является КСС
                if (sortMassive[i].number_of_diagnosed_element.Contains("КСС"))
                {
                    isJoint = true;
                }
                if (isNewElement == false & isJoint == false)
                {
                    if (localsumm < sortMassive[i].rang_of_danger)
                    {
                        localsumm = sortMassive[i].rang_of_danger;
                    }

                }
                if (isNewElement & isNewID_TP == false)
                {
                    summOfElement = summOfElement + localsumm;
                    localsumm = sortMassive[i].rang_of_danger;
                }
                if (i == sortMassive.Count - 1)//великая поправка к расчету
                {
                    summOfElement = summOfElement + localsumm;
                }

                if (isNewID_TP)
                {
                    result[Converting.getStringNumber(sortMassive[i - 1].ID_TP)].summ_of_rangs = summOfElement;
                    summOfElement = 0;
                    localsumm = 0;
                    //richTextBox1.AppendText(Environment.NewLine + "get_sum_of_danger_rangs=" + summOfElement);
                }

            }

            return result;
        }
        private List<StringOfTableZh1> get_sum_of_max_danger_rangs_new(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result, List<int> NumbersOfObjects)//Р10//Подсчет суммы максимальных рангов опасности для каждого дефектного элемента в каждом ТТ
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].summ_of_rangs=0;                
            }


            for (int k = 0; k < NumbersOfObjects.Count; k++)
            {
                List<StringOfTableZh2> listZh2OnlyForSomeObject = new List<StringOfTableZh2>();
                
                
                double summOfMaxRang = 0;


                for (int r = 0; r < sortMassive.Count; r++)
                {
                    if (NumbersOfObjects[k] == sortMassive[r].ID_TP)//NumbersOfObjects[k]==sortMassive[r].ID_TP
                    {

                        string kcc = "КСС";
                        if (String.Equals(sortMassive[r].type_of_element, kcc))//sortMassive[r].type_of_element.Contains("КСС") == false
                        {
                            //listZh2OnlyForSomeObject.Add(sortMassive[r]);
                            //richTextBox1.AppendText(Environment.NewLine + "listZh2OnlyForSomeObject.Add(sortMassive[r] )" + listZh2OnlyForSomeObject.Count);
                        }
                        else
                        {
                            listZh2OnlyForSomeObject.Add(sortMassive[r]);
                        }
                        
                    }
  
                }
                listZh2OnlyForSomeObject = getSortMassiveZh2_for_rangs(listZh2OnlyForSomeObject);//сортировка  для подсчета сумм рангов опасности
                richTextBox1.AppendText(Environment.NewLine + "Lits of NO! joints " + NumbersOfObjects[k] + "-" + listZh2OnlyForSomeObject.Count);


                if (listZh2OnlyForSomeObject.Count>0)
                {
                    double localRang = listZh2OnlyForSomeObject[0].rang_of_danger;

                    for (int t = 1; t < listZh2OnlyForSomeObject.Count; t++)
                    {
                        if (String.Equals(Convert.ToString(listZh2OnlyForSomeObject[t].number_of_diagnosed_element), Convert.ToString(listZh2OnlyForSomeObject[t - 1].number_of_diagnosed_element)))
                        {
                            if (listZh2OnlyForSomeObject[t].rang_of_danger > localRang)
                            {
                                localRang = listZh2OnlyForSomeObject[t].rang_of_danger;
                            }
                        }
                        else
                        {
                            summOfMaxRang += localRang;
                            localRang = 0;
                        }
                        if (t == listZh2OnlyForSomeObject.Count - 1)//великая поправка к расчету
                        {
                            summOfMaxRang += localRang;
                        }

                    }
                }
                if (listZh2OnlyForSomeObject.Count > 0)//костыль
                {
                    if (listZh2OnlyForSomeObject.Count < 2)
                    {
                        summOfMaxRang = listZh2OnlyForSomeObject[0].rang_of_danger;
                    }
                }
                result[Converting.getStringNumber(NumbersOfObjects[k])].summ_of_rangs += summOfMaxRang;
                

            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_max_danger_rangs_new_joint(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result, List<int> NumbersOfObjects)//Р10//Подсчет суммы максимальных рангов опасности для каждого дефектного элемента в каждом ТТ
        {
            for (int m = 0; m < result.Count; m++)
            {               
                result[m].summ_of_rangs_joint = 0;
            }


            for (int k = 0; k < NumbersOfObjects.Count; k++)
            {
                List<StringOfTableZh2> listZh2OnlyForSomeObject = new List<StringOfTableZh2>();


                double summOfMaxRang = 0;


                for (int r = 0; r < sortMassive.Count; r++)
                {
                    if (NumbersOfObjects[k] == sortMassive[r].ID_TP)//NumbersOfObjects[k]==sortMassive[r].ID_TP
                    {

                        string kcc = "КСС";
                        if (String.Equals(sortMassive[r].type_of_element, kcc))//sortMassive[r].type_of_element.Contains("КСС") == false
                        {
                            listZh2OnlyForSomeObject.Add(sortMassive[r]);
                            //richTextBox1.AppendText(Environment.NewLine + "listZh2OnlyForSomeObject.Add(sortMassive[r] )" + listZh2OnlyForSomeObject.Count);
                        }
                        else
                        {
                            //listZh2OnlyForSomeObject.Add(sortMassive[r]);
                        }

                    }

                }
                listZh2OnlyForSomeObject = getSortMassiveZh2_for_rangs(listZh2OnlyForSomeObject);//сортировка  для подсчета сумм рангов опасности
                richTextBox1.AppendText(Environment.NewLine + "Lits of joints " + NumbersOfObjects[k] + "-" + listZh2OnlyForSomeObject.Count);


                if (listZh2OnlyForSomeObject.Count > 0)
                {
                    double localRang = listZh2OnlyForSomeObject[0].rang_of_danger;

                    for (int t = 1; t < listZh2OnlyForSomeObject.Count; t++)
                    {
                        if (String.Equals(Convert.ToString(listZh2OnlyForSomeObject[t].number_of_diagnosed_element), Convert.ToString(listZh2OnlyForSomeObject[t - 1].number_of_diagnosed_element)))
                        {
                            if (listZh2OnlyForSomeObject[t].rang_of_danger > localRang)
                            {
                                localRang = listZh2OnlyForSomeObject[t].rang_of_danger;
                            }
                        }
                        else
                        {
                            summOfMaxRang += localRang;
                            localRang = 0;
                        }
                        if (t == listZh2OnlyForSomeObject.Count - 1)//великая поправка к расчету
                        {
                            summOfMaxRang += localRang;
                        }


                    }
                }
                if (listZh2OnlyForSomeObject.Count > 0)//костыль
                {
                    if (listZh2OnlyForSomeObject.Count < 2)
                    {
                        summOfMaxRang = listZh2OnlyForSomeObject[0].rang_of_danger;
                    }
                }
                result[Converting.getStringNumber(NumbersOfObjects[k])].summ_of_rangs_joint += summOfMaxRang;


            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_max_danger_rangs_new_KRN(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result, List<int> NumbersOfObjects)//подсчет суммы максимальных рангов опасности для каждого дефектного элемента в каждом ТТ для трещиноподобных дефектов
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].summ_of_rangs_KRN = 0;
            }
            for (int k = 0; k < NumbersOfObjects.Count; k++)
            {
                List<StringOfTableZh2> listZh2OnlyForSomeObject = new List<StringOfTableZh2>();
                double summOfMaxRang = 0;
                for (int r = 0; r < sortMassive.Count; r++)
                {
                    if (NumbersOfObjects[k] == sortMassive[r].ID_TP)//NumbersOfObjects[k]==sortMassive[r].ID_TP
                    {
                        string def1 = "РН";
                        string def2 = "щин";
                        if (sortMassive[r].type_of_defect.Contains(def1))//sortMassive[r].type_of_element.Contains("КСС") == false
                        {
                            listZh2OnlyForSomeObject.Add(sortMassive[r]);
                            //richTextBox1.AppendText(Environment.NewLine + "listZh2OnlyForSomeObject.Add(sortMassive[r] )" + listZh2OnlyForSomeObject.Count);
                        }
                        else if (sortMassive[r].type_of_defect.Contains(def2))
                        {
                            listZh2OnlyForSomeObject.Add(sortMassive[r]);
                        }
                        else
                        {
                            //listZh2OnlyForSomeObject.Add(sortMassive[r]);
                        }
                    }
                }


                if (listZh2OnlyForSomeObject.Count >2)
                {
                    listZh2OnlyForSomeObject = getSortMassiveZh2_for_rangs(listZh2OnlyForSomeObject);//сортировка  для подсчета сумм рангов опасности
                }

                if (listZh2OnlyForSomeObject.Count > 0)
                {
                    double localRang = listZh2OnlyForSomeObject[0].rang_of_danger;

                    for (int t = 1; t < listZh2OnlyForSomeObject.Count; t++)
                    {
                        if (String.Equals(Convert.ToString(listZh2OnlyForSomeObject[t].number_of_diagnosed_element), Convert.ToString(listZh2OnlyForSomeObject[t - 1].number_of_diagnosed_element)))
                        {
                            if (listZh2OnlyForSomeObject[t].rang_of_danger > localRang)
                            {
                                localRang = listZh2OnlyForSomeObject[t].rang_of_danger;
                            }
                        }
                        else
                        {
                            summOfMaxRang += localRang;
                            localRang = 0;
                        }

                        if (t== listZh2OnlyForSomeObject.Count-1)//великая поправка к расчету
                        {
                            summOfMaxRang += localRang;
                        }
                    }
                }
                if (listZh2OnlyForSomeObject.Count>0)//костыль
                {
                    if (listZh2OnlyForSomeObject.Count < 2)
                    {
                        summOfMaxRang = listZh2OnlyForSomeObject[0].rang_of_danger;
                    }
                }
                
                result[Converting.getStringNumber(NumbersOfObjects[k])].summ_of_rangs_KRN += summOfMaxRang;
                result[Converting.getStringNumber(NumbersOfObjects[k])].number_of_defects_KRN = listZh2OnlyForSomeObject.Count;
            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_length_other_Objects_Inside_The_CC(List<StringOfTableZh1> result, List<int> NumbersOfObjects)//подсчет суммы длин других ТТ внутри КЦ//OtherObjectsInsideTheCC
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].sum_of_length_other_Objects_Inside_The_CC = 0;
            }
            for (int i = 0; i < result.Count; i++)
            {
                double summ_length = 0;
                for (int j = 0; j < result[i].OtherObjectsInsideTheCC.Count; j++)
                {
                    //richTextBox1.AppendText(Environment.NewLine + "Вычисляем сумму...сравниваем" + result[i].OtherObjectsInsideTheCC[j] + "___" + result[i].ID_TP);
                    for (int k = 0; k < result.Count; k++)
                    {
                        if (result[k].ID_TP == result[i].OtherObjectsInsideTheCC[j])
                        {
                            summ_length += result[k].total_length_of_pipelines_DN_500_1400;
                            //richTextBox1.AppendText(Environment.NewLine + "Вычисляем сумму..." + result[i].OtherObjectsInsideTheCC[j] + "___" + summ_length);
                        }
                    }                    
                    
                }
                result[i].sum_of_length_other_Objects_Inside_The_CC = summ_length;
            }
            return result;           
        }
        private List<StringOfTableZh1> get_sum_of_length_Identical_Objects_Inside_The_CC(List<StringOfTableZh1> result, List<int> NumbersOfObjects)//подсчет суммы длин других ТТ внутри КЦ//OtherObjectsInsideTheCC
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].sum_of_length_Identical_Objects_Inside_The_CC = 0;
            }
            for (int i = 0; i < result.Count; i++)
            {
                double summ_length = 0;
                for (int j = 0; j < result[i].IdenticalObjectsInsideTheCC.Count; j++)
                {
                    //richTextBox1.AppendText(Environment.NewLine + "Вычисляем сумму...сравниваем" + result[i].OtherObjectsInsideTheCC[j] + "___" + result[i].ID_TP);
                    for (int k = 0; k < result.Count; k++)
                    {
                        if (result[k].ID_TP == result[i].IdenticalObjectsInsideTheCC[j])
                        {
                            summ_length += result[k].total_length_of_pipelines_DN_500_1400;
                            //richTextBox1.AppendText(Environment.NewLine + "Вычисляем сумму..." + result[i].IdenticalObjectsInsideTheCC[j] + "___" + summ_length);
                        }
                    }

                }
                result[i].sum_of_length_Identical_Objects_Inside_The_CC = summ_length;
            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_length_Identical_Objects_Neighboring_The_CC(List<StringOfTableZh1> result, List<int> NumbersOfObjects)//подсчет суммы длин идентичных ТТ соседних КЦ//OtherObjectsInsideTheCC
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].sum_of_length_Identical_Objects_Neighboring_The_CC = 0;
            }
            for (int i = 0; i < result.Count; i++)
            {
                double summ_length = 0;
                for (int j = 0; j < result[i].IdenticalObjectsOfTheNeighboringCC.Count; j++)
                {                   
                    for (int k = 0; k < result.Count; k++)
                    {
                        if (result[k].ID_TP == result[i].IdenticalObjectsOfTheNeighboringCC[j])
                        {
                            summ_length += result[k].total_length_of_pipelines_DN_500_1400;
                        }
                    }
                }
                result[i].sum_of_length_Identical_Objects_Neighboring_The_CC = summ_length;
            }
            return result;
        }
        private List<StringOfTableZh1> get_sum_of_length_other_Objects_Neighboring_The_CC(List<StringOfTableZh1> result, List<int> NumbersOfObjects)//подсчет суммы длин других ТТ соседних КЦ//OtherObjectsInsideTheCC
        {
            for (int m = 0; m < result.Count; m++)
            {
                result[m].sum_of_length_other_Objects_Neighboring_The_CC = 0;
            }
            for (int i = 0; i < result.Count; i++)
            {
                double summ_length = 0;
                for (int j = 0; j < result[i].OtherObjectsOfTheNeighboringCC.Count; j++)
                {
                    for (int k = 0; k < result.Count; k++)
                    {
                        if (result[k].ID_TP == result[i].OtherObjectsOfTheNeighboringCC[j])
                        {
                            summ_length += result[k].total_length_of_pipelines_DN_500_1400;
                        }
                    }
                }
                result[i].sum_of_length_other_Objects_Neighboring_The_CC = summ_length;
            }
            return result;
        }
        private double get_sum_of_length_Objects_ONE(List<StringOfTableZh1> result,  List<int> Objects)//подсчет суммы длин любых объектов по списку
        {
            double summ = 0;
            for (int i = 0; i < Objects.Count; i++)
            {
                for (int j = 0; j < result.Count; j++)
                {
                    if (result[j].ID_TP == Objects[i])
                    {
                        summ += result[j].total_length_of_pipelines_DN_500_1400;
                    }
                }

            }
            return summ;   
        }
        private List<StringOfTableZh1> get_summ_KRN_formula_A4(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ_KRN_TT = 0;
                double summ_length = 0;
                    for (int j = 0; j < result[i].IdenticalObjectsInsideTheCC.Count; j++)
                        {
                            for (int k = 0; k < result.Count; k++)
                            {
                                    if (result[i].IdenticalObjectsInsideTheCC[j]== result[k].ID_TP)
                                    {
                                        summ_KRN_TT += result[k].summ_of_rangs_KRN;
                                        summ_length += result[k].total_length_of_pipelines_DN_500_1400;
                                    }
                            }
                        }
                
                result[i].summ_of_rangs_KRN_A4 = 1 - Math.Exp(-(4.63 / (Math.Pow(summ_length, 0.17))) * summ_KRN_TT);
            }

            return result;
        }
        private List<StringOfTableZh1> get_summ_KRN_formula_A5(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ_KRN_TT = 0;

                if (result[i].OtherObjectsInsideTheCC.Count>0)
                {                    
                    for (int j = 0; j < result[i].OtherObjectsInsideTheCC.Count; j++)
                    {
                        for (int g = 0; g < result.Count; g++)
                        {
                            if (result[g].ID_TP == result[i].OtherObjectsInsideTheCC[j])
                            {
                                double MT_A1 = Converting.get_MT_from_table_A1(result[i].tt_gcs_type, result[g].tt_gcs_type);
                                summ_KRN_TT += MT_A1 * result[g].summ_of_rangs_KRN;
                            }
                        }      
                    }
                }               
                
                result[i].summ_of_rangs_KRN_A5 = 1 - Math.Exp(-(4.63 / (Math.Pow(result[i].sum_of_length_other_Objects_Inside_The_CC, 0.17))) * summ_KRN_TT);
            }
            return result;
        }       
        private List<StringOfTableZh1> get_sum_of_danger_rangs_for_joint(List<StringOfTableZh2> sortMassive, List<StringOfTableZh1> result)//Р13//Подсчет суммы максимальных рангов опасности для каждого дефектного СТЫКА в каждом ТТ
        {
            double localsumm = 0;
            double summOfElement = 0;
            localsumm = sortMassive[0].rang_of_danger;
            for (int i = 1; i < sortMassive.Count; i++)
            {
                bool isNewElement = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].number_of_diagnosed_element, sortMassive[i - 1].number_of_diagnosed_element) == false)
                {
                    isNewElement = true;
                }
                bool isNewID_TP = false;//истина, если текущий элемент не равен предыдущему
                if (String.Equals(sortMassive[i].ID_TP, sortMassive[i - 1].ID_TP) == false)
                {
                    isNewID_TP = true;
                }
                bool isJoint = false;//истина, если текущий элемент является КСС
                if (sortMassive[i].number_of_diagnosed_element.Contains("КСС"))
                {
                    isJoint = true;
                }
                if (isNewElement == false & isJoint)
                {
                    if (localsumm < sortMassive[i].rang_of_danger)
                    {
                        localsumm = sortMassive[i].rang_of_danger;
                    }

                }
                if (isNewElement & isNewID_TP == false)
                {
                    summOfElement = summOfElement + localsumm;
                    localsumm = sortMassive[i].rang_of_danger;
                }
                if (isNewID_TP)
                {
                    result[Converting.getStringNumber(sortMassive[i - 1].ID_TP)].summ_of_rangs_joint = summOfElement;
                    summOfElement = 0;
                    localsumm = 0;
                }

            }
            //richTextBox1.AppendText(Environment.NewLine + "get_sum_of_danger_rangs_for_joint=" + result.Count);
            return result;
        }
        private List<StringOfTableZh1> get_wanted_number_of_joints(List<StringOfTableZh1> massive)//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
        {
            for (int i = 0; i < massive.Count; i++)
            {
                massive[i].wanted_number_of_joints = 0.1 * (massive[i].number_of_inspected_joints_500_1400_with_ext_examination / 11);
                //richTextBox1.AppendText(Environment.NewLine + "get_wanted_number_of_joints" + massive[i].number_of_inspected_joints_500_1400_with_ext_examination + " ##" + i + " Number: " + result[i].wanted_number_of_joints);
            }

            return massive;
        }
        private List<StringOfTableZh1> get_base_number_of_joints(List<StringOfTableZh1> massive)//P14//Базовое число КСС, формула 6.7
        {
            for (int i = 0; i < massive.Count; i++)
            {
                massive[i].base_number_of_joints = Math.Max(massive[i].wanted_number_of_joints, Convert.ToDouble(massive[i].number_of_inspected_joints_500_1400_with_ext_examination));
                if (massive[i].wanted_number_of_joints == 0 & massive[i].number_of_inspected_joints_500_1400_with_ext_examination == 0)
                {
                    massive[i].base_number_of_joints = 0;
                }

                //richTextBox1.AppendText(Environment.NewLine + "get_base_number_of_joints" + result[i].wanted_number_of_joints + " ##" + i + " Number: " + massive[i].number_of_inspected_joints_500_1400_with_ext_examination + "***" + result[i].wanted_number_of_joints);
            }

            return massive;
        }
        private List<StringOfTableZh1> get_parameter_taking_into_account_the_technical_condition_of_KSS_TT(List<StringOfTableZh1> table_Zh1)//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
        {
            for (int i = 0; i < table_Zh1.Count; i++)
            {
                table_Zh1[i].parameter_taking_into_account_the_technical_condition_of_KSS_TT = table_Zh1[i].summ_of_rangs_joint / table_Zh1[i].base_number_of_joints;
            }
            return table_Zh1;
        }
        private List<StringOfTableZh1> get_parameter_according_to_formula_6_2_1(List<StringOfTableZh1> table_Zh1)//P12//формула 6.2.1//метод недоработан
        {
            for (int i = 0; i < table_Zh1.Count; i++)
            {
                table_Zh1[i].parameter_according_to_formula_6_2_1 = table_Zh1[i].summ_of_rangs_joint / table_Zh1[i].base_number_of_joints;
            }
            return table_Zh1;
        }
        private List<StringOfTableZh1> get_frequency_indicator_of_ompleteness_o_initial_data(List<StringOfTableZh1> input)//Р5//Частотный показатель полноты исходных данных
        {
            for (int i = 0; i < input.Count; i++)
            {
                double K = 0.22756;
                double k1 = 80;
                double k2 = 15;
                if (input[i].total_length_of_pipelines_DN_500_1400 > 0)
                {
                    double a = 1 + k1 * ((input[i].sum_of_monitored_areas_NK_only - input[i].reduced_length_uncontrolled_NK_only) / input[i].total_length_of_pipelines_DN_500_1400);
                    double b = 1 + k2 * ((input[i].sum_of_monitored_areas_NK_and_VTD - input[i].reduced_length_uncontrolled_NK_and_VTD) / input[i].total_length_of_pipelines_DN_500_1400);
                    double c = 1 + k2 * ((input[i].sum_of_monitored_areas_NK_only - input[i].reduced_length_uncontrolled_NK_only) / input[i].total_length_of_pipelines_DN_500_1400);
                    input[i].frequency_indicator_of_completeness_o_initial_data = K * Math.Log((a * b) / c);
                    //richTextBox1.AppendText(Environment.NewLine /*+ a + "=" + b + "=" + c + "="*/ + result[i].frequency_indicator_of_completeness_o_initial_data /*+ "=" + input[i].total_length_of_pipelines_DN_500_1400*/);
                    if (input[i].technological_number_tt_gks.Contains("Н"))//стр. 12. "расчет не проводится для надземной части ТТ промплощадки"
                    {
                        input[i].frequency_indicator_of_completeness_o_initial_data = 0;
                    }
                    DateTime dateToday = new DateTime();
                    dateToday = DateTime.Today;
                    double totalYears = 1;
                    if (isOtherYear.Checked==true)
                    {
                        totalYears = Convert.ToDouble(textBox9.Text) - input[i].year_of_KRTT;
                    }
                    else
                    {
                        totalYears = dateToday.Year - input[i].year_of_KRTT;
                    }
                    
                    //richTextBox1.AppendText(Environment.NewLine + dateToday.Year + "-" + input[i].year_of_KRTT.Year + "=" + totalYears);
                    if (totalYears < 10 & Math.Abs(input[i].total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000 - input[i].total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000) < 1)
                    {
                        input[i].frequency_indicator_of_completeness_o_initial_data = 1;
                    }
                }
                else
                {
                    input[i].frequency_indicator_of_completeness_o_initial_data = 0;
                }
            }
            return input;
        }
        private List<StringOfTableZh1> get_parameter_according_to_formula_6_5(List<StringOfTableZh1> input)//Р9//параметр по формуле 6.5
        {
            for (int i = 0; i < input.Count; i++)
            {
                double R = input[i].summ_of_rangs;
                double N1 = input[i].number_of_valves_requiring_repair;
                double N2 = input[i].total_number_of_inspected_pipes_500_1400 + input[i].total_number_of_inspected_details_500_1400 + input[i].number_of_monitored_valves;
                richTextBox1.AppendText(Environment.NewLine + "get_parameter_according_to_formula_6_5" + R);
                if (N2 > 0)
                {
                    input[i].parameter_according_to_formula_6_5 = (R + N1) / N2;
                    //richTextBox1.AppendText(Environment.NewLine /*+ a + "=" + b + "=" + c + "="*/ + result[i].parameter_according_to_formula_6_5 /*+ "=" + input[i].total_length_of_pipelines_DN_500_1400*/);
                }
                else
                {
                    input[i].parameter_according_to_formula_6_5 = 0;
                }
            }
            return input;
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
        private void save_to_excel(List<StringOfTableZh1> inputZh1)//Сохранение результатов в таблицу EXCEL
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
            worksheet.Name = "CalculationResult";


            worksheet.Cells[1, 1] = "ID_TP";//Столбец 1, кодовый номер объекта
            worksheet.Cells[1, 2] = "lpu_name";//Столбец 2, наименование ЛПУ
            worksheet.Cells[1, 3] = "gcs_name";//Столбец 3, наименование КС
            worksheet.Cells[1, 4] = "kc_number";//Столбец 4, номер КЦ
            worksheet.Cells[1, 5] = "pipeline_name";//Столбец 5, Наименование МГ
            worksheet.Cells[1, 6] = "coordinate_of_connecting_TT_to_MG";//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
            worksheet.Cells[1, 7] = "region";//Столбец 7, область
            worksheet.Cells[1, 8] = "tt_gcs_type";//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
            worksheet.Cells[1, 9] = "technological_number_tt_gks";//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
            worksheet.Cells[1, 10] = "inventory_number";//Столбец 10, Инвентарный номер ТТ КС
            worksheet.Cells[1, 11] = "plot_category";//Столбец 11, Категория участка
            worksheet.Cells[1, 12] = "maximum_outer_diameter";//Столбец 12, Максимальный наружный диаметр, мм
            worksheet.Cells[1, 13] = "maximum_wall_thickness";//Столбец 13, Максимальная толщина стенки, мм

            worksheet.Cells[1, 14] = "sum_of_monitored_areas_NK_and_VTD";//Сумма обследованных труб СДТ и ТПА в составе ТТ методами НК и/или ВТД
            worksheet.Cells[1, 15] = "sum_of_monitored_areas_NK_only";//Сумма обследованных труб СДТ и ТПА в составе ТТ только методами НК
            worksheet.Cells[1, 16] = "reduced_length_uncontrolled_NK_and_VTD";//Приведённая длина непроконтролированных зон НК и ВТД
            worksheet.Cells[1, 17] = "reduced_length_uncontrolled_NK_only";//Приведённая длина непроконтролированных зон НК
            worksheet.Cells[1, 18] = "frequency_indicator_of_completeness_o_initial_data";//Р5//Частотный показатель полноты исходных данных
            worksheet.Cells[1, 19] = "summ_of_rangs";//Р10//сумма рангов опасности для ТТ
            worksheet.Cells[1, 20] = "summ_of_rangs_joint";//Р13//сумма рангов опасности сварных соединений для ТТ
            worksheet.Cells[1, 21] = "summ_of_rangs_KRN_A4";//Р9//параметр по формуле 6.5
            
            
            worksheet.Cells[1, 22] = "summ_of_rangs_KRN_A5";//результаты расчета по формуле А5
            worksheet.Cells[1, 23] = "sum_of_length_Identical_Objects_Inside_The_CC";//общая протяженность идентичных ТТ внутри КЦ
            worksheet.Cells[1, 24] = "sum_of_length_other_Objects_Inside_The_CC";//общая протяженность других ТТ внутри КЦ
            worksheet.Cells[1, 25] = "sum_of_length_Identical_Objects_Neighboring_The_CC";//общая протяженность идентичных ТТ соседних КЦ
            worksheet.Cells[1, 26] = "sum_of_length_other_Objects_Neighboring_The_CC";// общая протяженность других ТТ соседних КЦ
            worksheet.Cells[1, 27] = "service_life_of_the_workshop_of_CC";// срок эксплуатации текущего КЦ с момента ввода или капремонта, лет. (максимальное время из всех объектов доанного КЦ)
            worksheet.Cells[1, 28] = "service_life_of_the_workshop_Neighboring_The_CC";// Максимальный из сроков эксплуатации соседнего КЦ
            
            worksheet.Cells[1, 29] = "service_life_of_the_workshop_of_string_Zh1";// service_life_of_the_workshop_of_string_Zh1;//срок эксплуатации текущего объекта таблицы Ж.1 с момента ввода или капремонта, лет.
            worksheet.Cells[1, 30] = "parameter_according_formula_A7";// parameter_according_formula_A7;//Величина коэффициента дефектности идентичных ТТ соседних КЦ (По формуле А7)
            worksheet.Cells[1, 31] = "parameter_according_formula_A8";// parameter_according_formula_A8;//Величина коэффициента дефектности других ТТ соседних КЦ (По формуле А8)
            worksheet.Cells[1, 32] = "parameter_according_formula_A10";// parameter_according_formula_A10;//Коэффициент дефектности прилегающих участков МГ текущего КЦ (По формуле А10)
            worksheet.Cells[1, 33] = "service_life_of_the_pipeline_of_CC";//срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет.
            worksheet.Cells[1, 34] = "service_life_of_the_pipeline_Neighboring_The_CC";//срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет.
            
            worksheet.Cells[1, 35] = "parameter_according_formula_A9";//Величина показателя фактора дефектности прилегающих участков МГ(По формуле А9)
            worksheet.Cells[1, 36] = "parameter_according_formula_A13";//Коэффициент аварийности ЛЧ МГ, примыкающей к текущему КЦ (По формуле А13)
            worksheet.Cells[1, 37] = "parameter_according_formula_A12";//Количественная оценка фактора аварийности Ra (По формуле А12)
            
            //worksheet.Cells[1, 38] = "parameter_according_formula_A15";//Количественная оценка фактора наработки Rн (По формуле А15)

        //worksheet.Cells[1, 22] = "wanted_number_of_joints";//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
        //worksheet.Cells[1, 23] = "base_number_of_joints";//P14//Базовое число КСС, формула 6.7
        //worksheet.Cells[1, 24] = "parameter_taking_into_account_the_technical_condition_of_KSS_TT";//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
        //worksheet.Cells[1, 25] = "parameter_according_to_formula_6_2_1";//Р9//параметр по формуле 6.2.1

            // worksheet.Cells[1, 26] = "Идентичные объекты внутри КЦ";
            // worksheet.Cells[1, 27] = "Другие объекты внутри КЦ";
            // worksheet.Cells[1, 28] = "Идентичные объекты соседнего КЦ";
            // worksheet.Cells[1, 29] = "Другие объекты соседнего КЦ";

            worksheet.Cells[2, 1] = "Кодовый номер объекта";//Столбец 1, кодовый номер объекта
            worksheet.Cells[2, 2] = "Наименование ЛПУ";//Столбец 2, наименование ЛПУ
            worksheet.Cells[2, 3] = "Наименование КС";//Столбец 3, наименование КС
            worksheet.Cells[2, 4] = "Номер КЦ";//Столбец 4, номер КЦ
            worksheet.Cells[2, 5] = "Наименование МГ";//Столбец 5, Наименование МГ
            worksheet.Cells[2, 6] = "Координата подключения ТТ КС к ЛЧ МГ, км";//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
            worksheet.Cells[2, 7] = "Область";//Столбец 7, область
            worksheet.Cells[2, 8] = "Тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)";//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
            worksheet.Cells[2, 9] = "Технологический номер (П,Н,7,8,7А,8А)";//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
            worksheet.Cells[2, 10] = "Инвентарный номер ТТ КС";//Столбец 10, Инвентарный номер ТТ КС
            worksheet.Cells[2, 11] = "Категория участка";//Столбец 11, Категория участка
            worksheet.Cells[2, 12] = "Максимальный наружный диаметр, мм";//Столбец 12, Максимальный наружный диаметр, мм
            worksheet.Cells[2, 13] = "Максимальная толщина стенки, мм";//Столбец 13, Максимальная толщина стенки, мм

            worksheet.Cells[2, 14] = "Сумма обследованных труб СДТ и ТПА в составе ТТ методами НК и/или ВТД";//Сумма обследованных труб СДТ и ТПА в составе ТТ методами НК и/или ВТД
            worksheet.Cells[2, 15] = "Сумма обследованных труб СДТ и ТПА в составе ТТ только методами НК";//Сумма обследованных труб СДТ и ТПА в составе ТТ только методами НК
            worksheet.Cells[2, 16] = "Приведённая длина непроконтролированных зон НК и ВТД";//Приведённая длина непроконтролированных зон НК и ВТД
            worksheet.Cells[2, 17] = "Приведённая длина непроконтролированных зон НК";//Приведённая длина непроконтролированных зон НК
            worksheet.Cells[2, 18] = "Показатель полноты исходных данных";//Р5//Частотный показатель полноты исходных данных
            worksheet.Cells[2, 19] = "Сумма рангов опасности для ТТ";//Р10//сумма рангов опасности для ТТ
            worksheet.Cells[2, 20] = "Сумма рангов опасности сварных соединений для ТТ";//Р13//сумма рангов опасности сварных соединений для ТТ
            worksheet.Cells[2, 21] = "Сумма рангов КРН по формуле А4";//Сумма рангов КРН по формуле А4

            worksheet.Cells[2, 22] = "результаты расчета по формуле А5";//результаты расчета по формуле А5
            worksheet.Cells[2, 23] = "общая протяженность идентичных ТТ внутри КЦ";//общая протяженность идентичных ТТ внутри КЦ
            worksheet.Cells[2, 24] = "общая протяженность других ТТ внутри КЦ";//общая протяженность других ТТ внутри КЦ
            worksheet.Cells[2, 25] = "общая протяженность идентичных ТТ соседних КЦ";//общая протяженность идентичных ТТ соседних КЦ
            worksheet.Cells[2, 26] = "общая протяженность других ТТ соседних КЦ";// sum_of_length_other_Objects_Neighboring_The_CC
            worksheet.Cells[2, 27] = "срок эксплуатации текущего КЦ с момента ввода или капремонта, лет. (максимальное время из всех объектов доанного КЦ) (t0)";//service_life_of_the_workshop_of_CC 
            worksheet.Cells[2, 28] = "Максимальный из сроков эксплуатации соседнего КЦ (t)";//  service_life_of_the_workshop_Neighboring_The_CC
            worksheet.Cells[2, 29] = "срок эксплуатации текущего объекта таблицы Ж.1 с момента ввода или капремонта, лет.";// service_life_of_the_workshop_of_string_Zh1;//срок эксплуатации текущего объекта таблицы Ж.1 с момента ввода или капремонта, лет.
            worksheet.Cells[2, 30] = "Величина коэффициента дефектности идентичных ТТ соседних КЦ (По формуле А7)";// parameter_according_formula_A7;//Величина коэффициента дефектности идентичных ТТ соседних КЦ (По формуле А7)
            worksheet.Cells[2, 31] = "Величина коэффициента дефектности других ТТ соседних КЦ (По формуле А8)";// parameter_according_formula_A8;//Величина коэффициента дефектности других ТТ соседних КЦ (По формуле А8)
            worksheet.Cells[2, 32] = "Коэффициент дефектности прилегающих участков МГ текущего КЦ (По формуле А10)";// parameter_according_formula_A10
            worksheet.Cells[2, 33] = "срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет. ";//срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет. service_life_of_the_pipeline_of_CC
            worksheet.Cells[2, 34] = "срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет.";//срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет. service_life_of_the_pipeline_Neighboring_The_CC

            worksheet.Cells[2, 35] = "Величина показателя фактора дефектности прилегающих участков МГ(По формуле А9)";//Величина показателя фактора дефектности прилегающих участков МГ(По формуле А9) parameter_according_formula_A9 

            worksheet.Cells[2, 36] = "Коэффициент аварийности ЛЧ МГ, примыкающей к текущему КЦ (По формуле А13)";//Коэффициент аварийности ЛЧ МГ, примыкающей к текущему КЦ (По формуле А13) parameter_according_formula_A13
            worksheet.Cells[2, 37] = "Количественная оценка фактора аварийности Ra (По формуле А12)";//Количественная оценка фактора аварийности Ra (По формуле А12) parameter_according_formula_A12
            //worksheet.Cells[2, 38] = "Количественная оценка фактора наработки Rн (По формуле А15)";//Количественная оценка фактора наработки Rн (По формуле А15) parameter_according_formula_A15


            //worksheet.Cells[1, 22] = "wanted_number_of_joints";//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
            //worksheet.Cells[1, 23] = "base_number_of_joints";//P14//Базовое число КСС, формула 6.7
            //worksheet.Cells[1, 24] = "parameter_taking_into_account_the_technical_condition_of_KSS_TT";//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
            //worksheet.Cells[1, 25] = "parameter_according_to_formula_6_2_1";//Р9//параметр по формуле 6.2.1

            worksheet.Cells[2, 50] = "Идентичные объекты внутри КЦ";
            worksheet.Cells[2, 51] = "Другие объекты внутри КЦ";
            worksheet.Cells[2, 52] = "Идентичные объекты соседнего КЦ";
            worksheet.Cells[2, 53] = "Другие объекты соседнего КЦ";
            worksheet.Cells[2, 54] = "Объекты текущего КЦ";
            worksheet.Cells[2, 55] = "Номера соседних КЦ КЦ";
            int strNunber = 3;//номер строки в формируемой таблице
            for (int i = 0; i < 92; i++)
            {
                worksheet.Cells[strNunber, 1] = inputZh1[i].ID_TP;//Столбец 1, кодовый номер объекта
                worksheet.Cells[strNunber, 2] = inputZh1[i].lpu_name;//Столбец 2, наименование ЛПУ
                worksheet.Cells[strNunber, 3] = inputZh1[i].gcs_name;//Столбец 3, наименование КС
                worksheet.Cells[strNunber, 4] = inputZh1[i].kc_number;//Столбец 4, номер КЦ
                worksheet.Cells[strNunber, 5] = inputZh1[i].pipeline_name;//Столбец 5, Наименование МГ
                worksheet.Cells[strNunber, 6] = inputZh1[i].coordinate_of_connecting_TT_to_MG;//Столбец 6, Координата подключения ТТ КС к ЛЧ МГ, км
                worksheet.Cells[strNunber, 7] = inputZh1[i].region;//Столбец 7, область
                worksheet.Cells[strNunber, 8] = inputZh1[i].tt_gcs_type;//Столбец 8, тип ТТ КС (входной шлейф, выходной шлейфы, промплощадка)
                worksheet.Cells[strNunber, 9] = inputZh1[i].technological_number_tt_gks;//Столбец 9, технологический номер (П,Н,7,8,7А,8А)
                worksheet.Cells[strNunber, 10] = inputZh1[i].inventory_number;//Столбец 10, Инвентарный номер ТТ КС
                worksheet.Cells[strNunber, 11] = inputZh1[i].plot_category;//Столбец 11, Категория участка
                worksheet.Cells[strNunber, 12] = inputZh1[i].maximum_outer_diameter;//Столбец 12, Максимальный наружный диаметр, мм
                worksheet.Cells[strNunber, 13] = inputZh1[i].maximum_wall_thickness;//Столбец 13, Максимальная толщина стенки, мм

                worksheet.Cells[strNunber, 14] = inputZh1[i].sum_of_monitored_areas_NK_and_VTD;//Сумма обследованных труб СДТ и ТПА в составе ТТ методами НК и/или ВТД
                worksheet.Cells[strNunber, 15] = inputZh1[i].sum_of_monitored_areas_NK_only;//Сумма обследованных труб СДТ и ТПА в составе ТТ только методами НК
                worksheet.Cells[strNunber, 16] = inputZh1[i].reduced_length_uncontrolled_NK_and_VTD;//Приведённая длина непроконтролированных зон НК и ВТД
                worksheet.Cells[strNunber, 17] = inputZh1[i].reduced_length_uncontrolled_NK_only;//Приведённая длина непроконтролированных зон НК
                worksheet.Cells[strNunber, 18] = inputZh1[i].frequency_indicator_of_completeness_o_initial_data;//Р5//Частотный показатель полноты исходных данных
                worksheet.Cells[strNunber, 19] = inputZh1[i].summ_of_rangs;//Р10//сумма рангов опасности для ТТ
                worksheet.Cells[strNunber, 20] = inputZh1[i].summ_of_rangs_joint;//Р13//сумма рангов опасности сварных соединений для ТТ
                worksheet.Cells[strNunber, 21] = inputZh1[i].summ_of_rangs_KRN_A4;//результаты расчета по формуле А4
                worksheet.Cells[strNunber, 22] = inputZh1[i].summ_of_rangs_KRN_A5;//результаты расчета по формуле А5
                worksheet.Cells[strNunber, 23] = inputZh1[i].sum_of_length_Identical_Objects_Inside_The_CC;//общая протяженность идентичных ТТ внутри КЦ
                worksheet.Cells[strNunber, 24] = inputZh1[i].sum_of_length_other_Objects_Inside_The_CC;//общая протяженность других ТТ внутри КЦ
                worksheet.Cells[strNunber, 25] = inputZh1[i].sum_of_length_Identical_Objects_Neighboring_The_CC;//общая протяженность идентичных ТТ соседних КЦ
                worksheet.Cells[strNunber, 26] = inputZh1[i].sum_of_length_other_Objects_Neighboring_The_CC;//общая протяженность других ТТ соседних КЦ
                worksheet.Cells[strNunber, 27] = inputZh1[i].service_life_of_the_workshop_of_CC;//срок эксплуатации текущего КЦ с момента ввода или капремонта, лет. (максимальное время из всех объектов доанного КЦ)
                worksheet.Cells[strNunber, 28] = inputZh1[i].service_life_of_the_workshop_Neighboring_The_CC;// Максимальный из сроков эксплуатации соседнего КЦ
                worksheet.Cells[strNunber, 29] = inputZh1[i].service_life_of_the_workshop_of_string_Zh1;//срок эксплуатации текущего объекта таблицы Ж.1 с момента ввода или капремонта, лет.
                worksheet.Cells[strNunber, 30] = inputZh1[i].parameter_according_formula_A7;//Величина коэффициента дефектности идентичных ТТ соседних КЦ (По формуле А7)
                worksheet.Cells[strNunber, 31] = inputZh1[i].parameter_according_formula_A8;//Величина коэффициента дефектности других ТТ соседних КЦ (По формуле А8)
                worksheet.Cells[strNunber, 32] = inputZh1[i].parameter_according_formula_A10;// "Коэффициент дефектности прилегающих участков МГ текущего КЦ (По формуле А10)"
                worksheet.Cells[strNunber, 33] = inputZh1[i].service_life_of_the_pipeline_of_CC;//срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет.
                worksheet.Cells[strNunber, 34] = inputZh1[i].service_life_of_the_pipeline_Neighboring_The_CC;//срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет.
                worksheet.Cells[strNunber, 35] = inputZh1[i].parameter_according_formula_A9;//Величина показателя фактора дефектности прилегающих участков МГ(По формуле А9)
                worksheet.Cells[strNunber, 36] = inputZh1[i].parameter_according_formula_A13;//Коэффициент аварийности ЛЧ МГ, примыкающей к текущему КЦ (По формуле А13)
                worksheet.Cells[strNunber, 37] = inputZh1[i].parameter_according_formula_A12;//Количественная оценка фактора аварийности Ra (По формуле А12)
                //worksheet.Cells[strNunber, 38] = inputZh1[i].parameter_according_formula_A15;//Количественная оценка фактора наработки Rн (По формуле А15)


                //worksheet.Cells[strNunber, 21] = inputZh1[i].parameter_according_to_formula_6_5;//Р9//параметр по формуле 6.5
                //worksheet.Cells[strNunber, 22] = inputZh1[i].wanted_number_of_joints;//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
                //worksheet.Cells[strNunber, 23] = inputZh1[i].base_number_of_joints;//P14//Базовое число КСС, формула 6.7
                //worksheet.Cells[strNunber, 24] = inputZh1[i].parameter_taking_into_account_the_technical_condition_of_KSS_TT;//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
                //worksheet.Cells[strNunber, 25] = inputZh1[i].parameter_according_to_formula_6_2_1;//Р9//параметр по формуле 6.2.1
                //result[i].summ_of_rangs_KRN_A4

                //выводим списки объектов для проверки качаства их чтения

                worksheet.Cells[strNunber, 50] = Converting.concatNumbersOfObjects(inputZh1[i].IdenticalObjectsInsideTheCC);
                worksheet.Cells[strNunber, 51] = Converting.concatNumbersOfObjects(inputZh1[i].OtherObjectsInsideTheCC);
                worksheet.Cells[strNunber, 52] = Converting.concatNumbersOfObjects(inputZh1[i].IdenticalObjectsOfTheNeighboringCC);
                worksheet.Cells[strNunber, 53] = Converting.concatNumbersOfObjects(inputZh1[i].OtherObjectsOfTheNeighboringCC);
                worksheet.Cells[strNunber, 54] = Converting.concatNumbersOfObjects(inputZh1[i].ObjectsOfTheCС);
                worksheet.Cells[strNunber, 55] = Converting.concatNumbersOfObjects(inputZh1[i].NeighboringTheCС);
                strNunber++;

            }

            // Show save file dialog
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //richTextBox1.AppendText(Environment.NewLine + saveFileDialog1.FileName);
                worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                //workBook.Close(false, Type.Missing, Type.Missing);
                //app.Quit();
            }
        }
        private List<StringOfTableZh1> get_service_life_of_the_workshop(List<StringOfTableZh1> input)//вычисляем срок эксплуатации текущего КЦ с момента ввода или капремонта, лет.
        {
            double currentYear =Converting.ConvertToDouble(Convert.ToString(DateTime.Now.Year));//получили текущий год
            for (int i = 0; i < input.Count; i++)
            {
                //currentYear- input[i].commissioning_year_of_the_gcs==0
                if (input[i].year_of_KRTT>0)
                {
                    input[i].service_life_of_the_workshop_of_string_Zh1 = currentYear-input[i].year_of_KRTT;
                }
                else if (true)
                {
                    input[i].service_life_of_the_workshop_of_string_Zh1 = currentYear-input[i].commissioning_year_of_the_gcs;
                }

            }
            for (int i = 0; i < input.Count; i++)//Здесь мы ищем максимальное время эксплуатации текущего КЦ
            {
                double max_service_life = 0;
                for (int j = 0; j < input[i].ObjectsOfTheCС.Count; j++)
                {
                    for (int k = 0; k < input.Count; k++)
                    {
                        if (input[i].ObjectsOfTheCС[j]== input[k].ID_TP)
                        {
                            if (max_service_life < input[k].service_life_of_the_workshop_of_string_Zh1)
                            {
                                max_service_life = input[k].service_life_of_the_workshop_of_string_Zh1;
                            }
                            
                        }
                    }

                }
                input[i].service_life_of_the_workshop_of_CC = max_service_life;
            }
            
            
            for (int i = 0; i < input.Count; i++)//здесь мы ищем максимальное время эксплуатации соседнего КЦ
            {
                double max_service_life = 0;
                for (int j = 0; j < input[i].IdenticalObjectsOfTheNeighboringCC.Count; j++)
                {
                    for (int k = 0; k < input.Count; k++)
                    {
                        if (input[i].IdenticalObjectsOfTheNeighboringCC[j] == input[k].ID_TP)
                        {
                            if (max_service_life < input[k].service_life_of_the_workshop_of_string_Zh1)
                            {
                                max_service_life = input[k].service_life_of_the_workshop_of_string_Zh1;
                            }

                        }
                    }

                }
                for (int j = 0; j < input[i].OtherObjectsOfTheNeighboringCC.Count; j++)
                {
                    for (int k = 0; k < input.Count; k++)
                    {
                        if (input[i].OtherObjectsOfTheNeighboringCC[j] == input[k].ID_TP)
                        {
                            if (max_service_life < input[k].service_life_of_the_workshop_of_string_Zh1)
                            {
                                max_service_life = input[k].service_life_of_the_workshop_of_string_Zh1;
                            }

                        }
                    }

                }
                input[i].service_life_of_the_workshop_Neighboring_The_CC = max_service_life;
            }

           
            return input;
        }
        private List<StringOfTableZh1> get_service_life_of_the_pipeline_near_the_station(List<StringOfTableZh1> input)//вычисляем срок эксплуатации прилегающего участка МГ с момента ввода, лет.
        {
            /*
             public double service_life_of_the_pipeline_of_CC;//срок эксплуатации участка ЛЧМГ примыкающего к текущему объекту таблицы Ж.1 с момента ввода, лет.
             public double service_life_of_the_pipeline_Neighboring_The_CC;//срок эксплуатации участка ЛЧМГ примыкающего к соседнему объекту с момента ввода, лет.             
             */


            double currentYear = Converting.ConvertToDouble(Convert.ToString(DateTime.Now.Year));//получили текущий год
            for (int i = 0; i < input.Count; i++)
            {                
                    input[i].service_life_of_the_pipeline_of_CC = currentYear - input[i].year_of_commissioning_of_th_LPG_section;
            }
            

            for (int i = 0; i < input.Count; i++)//здесь мы ищем максимальное время эксплуатации соседнего КЦ
            {
                double max_service_life = 0;
                for (int j = 0; j < input[i].IdenticalObjectsOfTheNeighboringCC.Count; j++)
                {
                    for (int k = 0; k < input.Count; k++)
                    {
                        if (input[i].IdenticalObjectsOfTheNeighboringCC[j] == input[k].ID_TP)
                        {
                            if (max_service_life < input[k].service_life_of_the_pipeline_of_CC)
                            {
                                max_service_life = input[k].service_life_of_the_pipeline_of_CC;
                            }

                        }
                    }

                }
                for (int j = 0; j < input[i].OtherObjectsOfTheNeighboringCC.Count; j++)
                {
                    for (int k = 0; k < input.Count; k++)
                    {
                        if (input[i].OtherObjectsOfTheNeighboringCC[j] == input[k].ID_TP)
                        {
                            if (max_service_life < input[k].service_life_of_the_pipeline_of_CC)
                            {
                                max_service_life = input[k].service_life_of_the_pipeline_of_CC;
                            }

                        }
                    }

                }
                input[i].service_life_of_the_pipeline_Neighboring_The_CC = max_service_life;
            }


            return input;
        }
        private List<StringOfTableZh1> get_parameter_according_formula_A7(List<StringOfTableZh1> result)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ_KRN_TT = 0;

                if (result[i].IdenticalObjectsOfTheNeighboringCC.Count > 0)
                {
                    for (int j = 0; j < result[i].IdenticalObjectsOfTheNeighboringCC.Count; j++)
                    {
                        for (int g = 0; g < result.Count; g++)
                        {
                            if (result[g].ID_TP == result[i].IdenticalObjectsOfTheNeighboringCC[j])
                            {                                
                                summ_KRN_TT += result[g].summ_of_rangs_KRN;
                            }
                        }
                    }
                }
                
                if (result[i].service_life_of_the_workshop_Neighboring_The_CC>0)//проверяем, чтобы в знаменатель не попал ноль
                {
                    result[i].parameter_according_formula_A7 = result[i].service_life_of_the_workshop_of_CC / result[i].service_life_of_the_workshop_Neighboring_The_CC * (1 - Math.Exp(-(4.63 / Math.Pow(result[i].sum_of_length_Identical_Objects_Neighboring_The_CC, 0.17)) * summ_KRN_TT));
                }

                else
                {
                    result[i].parameter_according_formula_A7 = 0;
                }
            }
            return result;
        }
        private List<StringOfTableZh1> get_parameter_according_formula_A8(List<StringOfTableZh1> result)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ_KRN_TT = 0;

                if (result[i].OtherObjectsOfTheNeighboringCC.Count > 0)
                {
                    for (int j = 0; j < result[i].OtherObjectsOfTheNeighboringCC.Count; j++)
                    {
                        for (int g = 0; g < result.Count; g++)
                        {
                            if (result[g].ID_TP == result[i].OtherObjectsOfTheNeighboringCC[j])
                            {
                                double MT_A1 = Converting.get_MT_from_table_A1(result[i].tt_gcs_type, result[g].tt_gcs_type);
                                summ_KRN_TT += MT_A1*result[g].summ_of_rangs_KRN;
                            }
                        }
                    }
                }
                if (result[i].service_life_of_the_workshop_Neighboring_The_CC > 0)//проверяем, чтобы в знаменатель не попал ноль
                {
                    result[i].parameter_according_formula_A8 = result[i].service_life_of_the_workshop_of_CC / result[i].service_life_of_the_workshop_Neighboring_The_CC * (1 - Math.Exp(-(4.63 / Math.Pow(result[i].sum_of_length_other_Objects_Neighboring_The_CC, 0.17)) * summ_KRN_TT));
                }
                else
                {
                    result[i].parameter_according_formula_A8 = 0;
                }

            }
            return result;
        }
        private List<StringOfTableZh1> get_parameter_according_formula_A10(List<StringOfTableZh1> result, List<StringOfTableZh5> table_Zh5)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ = 0;
                for (int j = 0; j < table_Zh5.Count; j++)
                {
                    if (result[i].ID_CS == table_Zh5[j].ID_CS)
                    {
                        //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10" + result[i].ID_CS);
                        summ += Math.Exp(-(table_Zh5[j].Distance_of_the_element_with_SCC_defects_from_the_CC_km* table_Zh5[j].Distance_of_the_element_with_SCC_defects_from_the_CC_km) /(7*20));
                        //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10 Результат: " + summ);
                    }
                }
                double res= (summ * 20) / 300;
                if (res>1)
                {
                    result[i].parameter_according_formula_A10 = 1;
                }
                else
                {
                    result[i].parameter_according_formula_A10 = res;
                }

            }
            return result;
        }
        private double get_parameter_according_formula_A11_only_one(List<StringOfTableZh1> result, List<StringOfTableZh5> table_Zh5, int ID_CS_current, int ID_CS_other)
        {
            double tl = 0;
            double tl0 = 0;
            double summ = 0;
            double resultat = 0;

            for (int i = 0; i < result.Count; i++)
            {
                if (ID_CS_current== result[i].ID_CS)
                {
                    tl = result[i].service_life_of_the_pipeline_of_CC;
                }
            }
            for (int i = 0; i < result.Count; i++)
            {
                if (ID_CS_other == result[i].ID_CS)
                {
                    tl0 = result[i].service_life_of_the_pipeline_of_CC;
                }
            }

            for (int i = 0; i < table_Zh5.Count; i++)
            {
                if (ID_CS_current == table_Zh5[i].ID_CS)
                {
                    //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10" + result[i].ID_CS);
                    summ += Math.Exp(-(table_Zh5[i].Distance_of_the_element_with_SCC_defects_from_the_CC_km * table_Zh5[i].Distance_of_the_element_with_SCC_defects_from_the_CC_km) / (7 * 20));
                    //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10 Результат: " + summ);
                }
            }
            double a = 20;
            double b = 300;
            double res = (tl0 / tl)*(a/b)*summ;
            
            if (res > 1)
            {
                resultat = 1;
            }
            else
            {
                resultat = res;
            }

            return resultat;
        }
        private List<StringOfTableZh1> get_parameter_according_formula_A9_new(List<StringOfTableZh1> result)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ = 0;

                if (result[i].NeighboringTheCС.Count > 0)
                {
                    for (int j = 0; j < result[i].NeighboringTheCС.Count; j++)
                    {
                        for (int g = 0; g < result.Count; g++)
                        {
                            if (result[g].ID_CS == result[i].NeighboringTheCС[j])
                            {
                                summ += get_parameter_according_formula_A11_only_one(result, table_Zh5, result[i].ID_CS, result[i].NeighboringTheCС[j]);
                            }
                        }
                    }
                }
                double a = 0.6;
                double b = 0.4;
                result[i].parameter_according_formula_A9 = a * result[i].parameter_according_formula_A10 + b * summ;

            }
            return result;
        }
        private double get_parameter_according_formula_A14_only_one(List<StringOfTableZh1> result, List<StringOfTableZh4> table_Zh4, int ID_CS_current, int ID_CS_other)
        {
            double tl = 0;
            double tl0 = 0;
            double summ = 0;
            double resultat = 0;

            for (int i = 0; i < result.Count; i++)
            {
                if (ID_CS_current == result[i].ID_CS)
                {
                    tl = result[i].service_life_of_the_pipeline_of_CC;
                }
            }
            for (int i = 0; i < result.Count; i++)
            {
                if (ID_CS_other == result[i].ID_CS)
                {
                    tl0 = result[i].service_life_of_the_pipeline_of_CC;
                }
            }

            for (int i = 0; i < table_Zh4.Count; i++)
            {
                if (ID_CS_current == table_Zh4[i].ID_CS)
                {
                    //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10" + result[i].ID_CS);
                    summ = (tl0 / tl) *(0.5 * Math.Sqrt(table_Zh4[i].a0) + 0.26 * Math.Sqrt(table_Zh4[i].a1) + 0.18 * Math.Sqrt(table_Zh4[i].a2) + 0.06 * Math.Sqrt(table_Zh4[i].a3) + 0.03 * Math.Sqrt(table_Zh4[i].a4));
                    //richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А10 Результат: " + summ);
                }
            }

            double res = summ;

            if (res > 1)
            {
                resultat = 1;
            }
            else
            {
                resultat = res;
            }

            return resultat;
        }
        private List<StringOfTableZh1> get_parameter_according_formula_A13(List<StringOfTableZh1> result, List<StringOfTableZh4> table_Zh4)
        {
            for (int i = 0; i < result.Count; i++)
            {
                for (int j = 0; j < table_Zh4.Count; j++)
                {
                    richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].ID_TP + "Расчет по формуле А13 Результат: " + table_Zh4[j].ID_CS+"_"+ result[i].ID_CS);
                    if (table_Zh4[j].ID_CS== result[i].ID_CS)
                    {
                        result[i].parameter_according_formula_A13 = 0.5 * Math.Sqrt(Convert.ToDouble(table_Zh4[j].a0)) + 0.26 * Math.Sqrt(table_Zh4[j].a1) + 0.18 * Math.Sqrt(Convert.ToDouble(table_Zh4[j].a2)) + 0.06 * Math.Sqrt(Convert.ToDouble(table_Zh4[j].a3)) + 0.03 * Math.Sqrt(Convert.ToDouble(table_Zh4[j].a4));
                    }
                }
            }
            return result;//parameter_according_formula_A13
        }
        private List<StringOfTableZh1> get_paraqmeer_according_formula_A12(List<StringOfTableZh1> result, List<StringOfTableZh4> table_Zh4)
        {
            for (int i = 0; i < result.Count; i++)
            {
                double summ = 0;

                if (result[i].NeighboringTheCС.Count > 0)
                {
                    for (int j = 0; j < result[i].NeighboringTheCС.Count; j++)
                    {
                        for (int g = 0; g < result.Count; g++)
                        {
                            if (result[g].ID_CS == result[i].NeighboringTheCС[j])
                            {
                                summ += get_parameter_according_formula_A14_only_one(result, table_Zh4, result[i].ID_CS, result[i].NeighboringTheCС[j]);
                            }
                        }
                    }
                }
                double a = 0.7;
                double b = 0.3;
                result[i].parameter_according_formula_A12 = a * result[i].parameter_according_formula_A13 + b * summ;

            }
            return result;
        }
        private List<StringOfTableZh1> get_neighboring_CS_ID(List<StringOfTableZh1> result)//создание списка ID_CS (номеров соседних КЦ)
        {
            for (int i = 0; i < result.Count; i++)
            {
                if (result[i].IdenticalObjectsOfTheNeighboringCC.Count>0)
                {
                    for (int j = 0; j < result[i].IdenticalObjectsOfTheNeighboringCC.Count; j++)
                    {
                        for (int k = 0; k < result.Count; k++)
                        {
                            if (result[i].IdenticalObjectsOfTheNeighboringCC[j]== result[k].ID_TP)
                            {
                                result[i].NeighboringTheCС.Add(result[k].ID_CS);             
                            }
                        }
                    }
                    result[i].NeighboringTheCС= result[i].NeighboringTheCС.Distinct().ToList();
                }

            }
            return result;
        }
    }
}
