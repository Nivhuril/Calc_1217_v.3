using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //public string Connect = "Database=sql11475221;Data Source=sql11.freemysqlhosting.net;User Id=sql11475221;Password=HjT9hIIX8J;charset=utf8; charset = utf8";
        public string Connect = "Database=u1611387_bdforwork;Data Source = server179.hosting.reg.ru; User Id = u1611387_bduser; Password=Mirinda22+; charset = utf8";
        List<StringOfTableZh1_string> table_Zh1_string = new List<StringOfTableZh1_string>();
        List<StringOfTableZh1> table_Zh1 = new List<StringOfTableZh1>();
        List<StringOfTableZh2_string> table_Zh2_string = new List<StringOfTableZh2_string>();
        List<StringOfTableZh2> table_Zh2 = new List<StringOfTableZh2>();
        List<StringOfTableZh4> table_Zh4 = new List<StringOfTableZh4>();
        List<StringOfTableZh5> table_Zh5 = new List<StringOfTableZh5>();
        public string Invert(string input)
        {
            string result = "";
            int strLength = input.Length;
            if (strLength == 0)
            {
                result = "";
            }
            else if (strLength == 1)
            {
                result = input;
            }

            else if (strLength > 1)
            {
                for (int i = 1; i < strLength + 1; i++)
                {
                    result = String.Concat(result, input.Substring(strLength - i, 1));
                }
            }
            return result;
        }
        private async void openFileButton_Click(object sender, EventArgs e)
        {
            string fileName;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                this.Text = String.Concat("Расчет Р Газпром 2-2.3-1217-2020 - ", fileName);
                richTextBox1.AppendText(Environment.NewLine + "Открыт файл: " + fileName);
                await Task.Run(() => table_Zh1_string = get_Zh1_data_from_file_to_string(fileName));//Читаем таблицу Ж.1
                table_Zh1 = convert_Zh1_to_current_type(table_Zh1_string);//Обрабатываем таблицу Ж.1
                richTextBox1.AppendText(Environment.NewLine + "Таблица Ж1 прочитана");
                await Task.Run(() => table_Zh2 = get_Zh2_data_from_file_short(fileName));//Читаем и обрабатываем таблицу Ж.2
                await Task.Run(() => table_Zh4 = get_Zh4_data_from_file(fileName));
                await Task.Run(() => table_Zh5 = get_Zh5_data_from_file(fileName));
            }
        }
        private void calculation_button_Click(object sender, EventArgs e)
        {
            List<int> NumbersOfObjects = new List<int>();
            for (int i = 0; i < table_Zh1.Count; i++)
            {
                NumbersOfObjects.Add(table_Zh1[i].ID_TP);
            }
            textBox1.Text = Convert.ToString(NumbersOfObjects.Count);
            table_Zh2 = get_rang_of_danger(table_Zh2);//Добавляем в табльицу Ж.2 значения рангов опасности для каждого дефекта
            List<StringOfTableZh2> sortMassive_Zh2 = getSortMassiveZh2(table_Zh2);//сортируем массив последовательно по трём полям
            table_Zh1 = get_sum_of_monitored_areas_NK_and_or_VTD_new(table_Zh2, table_Zh1);//get_sum_of_monitored_areas_NK_and_VTD
            table_Zh1 = get_sum_of_monitored_areas_NK_only_new(table_Zh2, table_Zh1);//get_sum_of_monitored_areas_NK_only
            table_Zh1 = get_reduced_length_uncontrolled_NK_and_or_VTD_zones(sortMassive_Zh2, table_Zh1);//непроконтролированные зоны ВТД и/или НК
            table_Zh1 = get_reduced_length_uncontrolled_NK_only(sortMassive_Zh2, table_Zh1);//непроконтролированные зоны  НК
            table_Zh1 = get_frequency_indicator_of_ompleteness_o_initial_data(table_Zh1);//Р5//Частотный показатель полноты исходных данных
            table_Zh1 = get_parameter_according_to_formula_6_5(table_Zh1);//Р9//параметр по формуле 6.5
            table_Zh1 = get_wanted_number_of_joints(table_Zh1);//P15//формула 6.8. Целевое количество КСС, подлежащее наружному обследованию в составе j-го ТТ
            table_Zh1 = get_base_number_of_joints(table_Zh1);//P14//Базовое число КСС, формула 6.7
            table_Zh1 = get_parameter_taking_into_account_the_technical_condition_of_KSS_TT(table_Zh1);//P12//параметр учитывающий техническое состояние КСС ТТ, формула 6.6
            List<StringOfTableZh2> sortMassiveZh2_for_rangs = getSortMassiveZh2_for_rangs(table_Zh2);//сортировка в другом порядке для подсчета сумм рангов опасности
            table_Zh1 = get_sum_of_max_danger_rangs_new(sortMassiveZh2_for_rangs, table_Zh1, NumbersOfObjects);//Р10//выполнение подсчета сумм рангов опасности
            table_Zh1 = get_sum_of_max_danger_rangs_new_joint(sortMassiveZh2_for_rangs, table_Zh1, NumbersOfObjects);//Р13//выполнение подсчета сумм рангов опасности дефектов КСС
            table_Zh1 = get_sum_of_max_danger_rangs_new_KRN(sortMassiveZh2_for_rangs, table_Zh1, NumbersOfObjects);//выполнение подсчета сумм рангов опасности дефектов для трещиноподобных дефектов каждого ТТ в отдельности
            table_Zh1 = get_sum_of_length_Identical_Objects_Inside_The_CC(table_Zh1, NumbersOfObjects);//подсчет суммы длин идентичных ТТ внутри КЦ
            table_Zh1 = get_sum_of_length_other_Objects_Inside_The_CC(table_Zh1, NumbersOfObjects);//подсчет суммы длин других ТТ внутри КЦ
            table_Zh1 = get_sum_of_length_Identical_Objects_Neighboring_The_CC(table_Zh1, NumbersOfObjects);//подсчет суммы длин идентичных ТТ соседних КЦ
            table_Zh1 = get_sum_of_length_other_Objects_Neighboring_The_CC(table_Zh1, NumbersOfObjects);//подсчет суммы длин других ТТ соседних КЦ
            table_Zh1 = get_neighboring_CS_ID(table_Zh1);//Получение номеров соседних КЦ
            table_Zh1 = get_summ_KRN_formula_A4(sortMassiveZh2_for_rangs, table_Zh1);//Расчет по формуле А4
            table_Zh1 = get_summ_KRN_formula_A5(sortMassiveZh2_for_rangs, table_Zh1);//Расчет по формуле А5
            table_Zh1 = get_service_life_of_the_workshop(table_Zh1);//Расчет t и t0
            table_Zh1 = get_parameter_according_formula_A7(table_Zh1);//Расчет по формуле А7
            table_Zh1 = get_parameter_according_formula_A8(table_Zh1);//Расчет по формуле А8
            table_Zh1 = get_parameter_according_formula_A10(table_Zh1, table_Zh5);//Расчет по формуле А10
            table_Zh1 = get_service_life_of_the_pipeline_near_the_station(table_Zh1);//Расчет tл и tл0
            //table_Zh1 = get_parameter_according_formula_A11(table_Zh1, table_Zh5);//Расчет по формуле А11
            table_Zh1 = get_parameter_according_formula_A9_new(table_Zh1);//Расчет по формуле А9
            table_Zh1 = get_parameter_according_formula_A13(table_Zh1, table_Zh4);//Расчет по формуле А9
            table_Zh1 = get_paraqmeer_according_formula_A12(table_Zh1, table_Zh4);//Расчет по формуле А12
            richTextBox1.AppendText(Environment.NewLine + "======================================");
            textBox5.Text = "Расчет выполнен";
            textBox6.Text = "Данные готовы к сохранению";
        }
        private void save_to_Excel_Click(object sender, EventArgs e)
        {
            save_to_excel(table_Zh1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < table_Zh1.Count; i++)
            {
                richTextBox1.AppendText(Environment.NewLine + table_Zh1[i].summ_of_rangs_KRN);
            }

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            registerForm frm = new registerForm();
            frm.Show();

        }
        private async void ReadZh1_Click(object sender, EventArgs e)
        {
            string fileName;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                this.Text = String.Concat("Расчет Р Газпром 2-2.3-1217-2020 - ", fileName);
                richTextBox1.AppendText(Environment.NewLine + "Открыт файл: " + fileName);
                await Task.Run(() => table_Zh1_string = get_Zh1_data_from_file_to_string(fileName));//Читаем таблицу Ж.1
                table_Zh1 = convert_Zh1_to_current_type(table_Zh1_string);//Обрабатываем таблицу Ж.1
                richTextBox1.AppendText(Environment.NewLine + "Таблица Ж1 прочитана");
            }
        }
        private void SendZh1ToDB_Click(object sender, EventArgs e)
        {
            SendZh1ToDBMeth();
        }
        private async void sendZh2ToDB_Click(object sender, EventArgs e)
        {
            await Task.Run(() => SendZh2ToDBMeth());
        }
        private async void SendZh1ToDB_Click_1(object sender, EventArgs e)
        {
            await Task.Run(() => SendZh1ToDBMeth());
        }
        private void addNewStringToGh2_Click(object sender, EventArgs e)
        {
            String loginUser = loginField.Text;
            String passUser = passField.Text;

            string CommandText = "INSERT INTO `users` (`login`, `pass`) VALUES (@login, @pass);";
            string Connect = "Database=sql11475221;Data Source=sql11.freemysqlhosting.net;User Id=sql11475221;Password=HjT9hIIX8J";

            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter();

            MySqlConnection myConnection = new MySqlConnection(Connect);

            MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);


            myCommand.Parameters.Add("@login", MySqlDbType.VarChar).Value = loginUser;
            myCommand.Parameters.Add("@pass", MySqlDbType.VarChar).Value = passUser;

            myConnection.Open(); //Устанавливаем соединение с базой данных.
            if (myCommand.ExecuteNonQuery() == 1)
            {
                MessageBox.Show("Account is created");
            }
            else
            {
                MessageBox.Show("Account is not created");
            }


            myConnection.Close();

            //adapter.SelectCommand = myCommand;
            //adapter.Fill(table);
        }
    }
}