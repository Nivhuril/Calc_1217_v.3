using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    public class SqlConnecting
    {
        
    }
    partial class Form1 : Form
    {
        private void CheckConnectionEdication()
        {
            String loginUser = loginField.Text;
            String passUser = passField.Text;

            string CommandText = "SELECT * FROM `users` WHERE `login` = @ul AND `pass` = @up";
            //string Connect = ;

            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter();

            MySqlConnection myConnection = new MySqlConnection(Connect);

            MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);

            myConnection.Open(); //Устанавливаем соединение с базой данных.
            myCommand.Parameters.Add("@ul", MySqlDbType.VarChar).Value = loginUser;
            myCommand.Parameters.Add("@up", MySqlDbType.VarChar).Value = passUser;
            adapter.SelectCommand = myCommand;
            adapter.Fill(table);


            if (table.Rows.Count > 0)
            {
                MessageBox.Show("Yes");
            }
            else
            {
                MessageBox.Show("No");
            }

            myConnection.Close();//Закрываем соединение с базой данных

        }
        private void ReadFromDBEdication()
        {
            String loginUser = loginField.Text;
            String passUser = passField.Text;

            string CommandText = "SELECT * FROM `users` WHERE `login` = @ul";//"SELECT * FROM `users` WHERE `login` = @ul AND `pass` = @up"
            //string Connect = "Database=sql11475221;Data Source=sql11.freemysqlhosting.net;User Id=sql11475221;Password=HjT9hIIX8J";
            //string Connect = "Database=u1611387_bdforwork;Data Source=server179.hosting.reg.ru;User Id=u1611387_bduser;Password=Mirinda22+";

            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter();

            MySqlConnection myConnection = new MySqlConnection(Connect);

            MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);

            myConnection.Open(); //Устанавливаем соединение с базой данных.
            myCommand.Parameters.Add("@ul", MySqlDbType.VarChar).Value = loginUser;
            myCommand.Parameters.Add("@up", MySqlDbType.VarChar).Value = passUser;
            adapter.SelectCommand = myCommand;
            adapter.Fill(table);
            string commentAuthor = table.Rows[0].Field<string>("pass");
            MessageBox.Show(commentAuthor);
            myConnection.Close();//Закрываем соединение с базой данных
        }
        private void SendZh1ToDBMeth()
        {
            //string CommandText = "SELECT * FROM `users` WHERE `login` = @ul";//"SELECT * FROM `users` WHERE `login` = @ul AND `pass` = @up"
            //string Connect = "Database=sql11475221;Data Source=sql11.freemysqlhosting.net;User Id=sql11475221;Password=HjT9hIIX8J";
            //string Connect = "Database=u1611387_bdforwork;Data Source=server179.hosting.reg.ru;User Id=u1611387_bduser;Password=Mirinda22+; charset = utf8";
            string CommandText = "INSERT INTO `Zh1` (`ID_TP`, `x2`, `x3`, `x4`, `x5`, `x6`, `x7`, `x8`, `x9`, `x10`, `x11`, `x12`, `x13`, `x14`, `x15`, " +
                "`x16`, `x17`, `x18`, `x19`, `x20`, `x21`, `x22`, `x23`, `x24`, `x25`, `x26`, `x27`, `x28`, `x29`, `x30`, `x31`, `x32`, `x33`, `x34`, `x35`," +
                " `x36`, `x37`, `x38`, `x39`, `x40`, `x41`, `x42`, `x43`, `x44`, `x45`, `x46`, `x47`, `x48`, `x49`, `x50`, `x51`, `x52`, `x53`, `x54`, `x55`," +
                " `x56`, `x57`, `x58`) VALUES (@x1, @x2, @x3, @x4, @x5, @x6, @x7, @x8, @x9, @x10, @x11, @x12, @x13, @x14, @x15, @x16, @x17, @x18, @x19," +
                " @x20, @x21, @x22, @x23, @x24, @x25, @x26, @x27, @x28, @x29, @x30, @x31, @x32, @x33, @x34, @x35, @x36, @x37, @x38, @x39, @x40, @x41, @x42," +
                " @x43, @x44, @x45, @x46, @x47, @x48, @x49, @x50, @x51, @x52, @x53, @x54, @x55, @x56, @x57, @x58);";

            MySqlConnection myConnection = new MySqlConnection(Connect);

            myConnection.Open(); //Устанавливаем соединение с базой данных.
            for (int i = 0; i < table_Zh1.Count; i++)
            {

                MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);

                myCommand.Parameters.Add("@x1", MySqlDbType.VarChar).Value = table_Zh1[i].ID_TP;
                myCommand.Parameters.Add("@x2", MySqlDbType.VarChar).Value = table_Zh1[i].lpu_name;
                myCommand.Parameters.Add("@x3", MySqlDbType.VarChar).Value = table_Zh1[i].gcs_name;
                myCommand.Parameters.Add("@x4", MySqlDbType.VarChar).Value = table_Zh1[i].kc_number;
                myCommand.Parameters.Add("@x5", MySqlDbType.VarChar).Value = table_Zh1[i].pipeline_name;
                myCommand.Parameters.Add("@x6", MySqlDbType.VarChar).Value = table_Zh1[i].coordinate_of_connecting_TT_to_MG;
                myCommand.Parameters.Add("@x7", MySqlDbType.VarChar).Value = table_Zh1[i].region;
                myCommand.Parameters.Add("@x8", MySqlDbType.VarChar).Value = table_Zh1[i].tt_gcs_type;
                myCommand.Parameters.Add("@x9", MySqlDbType.VarChar).Value = table_Zh1[i].technological_number_tt_gks;
                myCommand.Parameters.Add("@x10", MySqlDbType.VarChar).Value = table_Zh1[i].inventory_number;
                myCommand.Parameters.Add("@x11", MySqlDbType.VarChar).Value = table_Zh1[i].plot_category;
                myCommand.Parameters.Add("@x12", MySqlDbType.VarChar).Value = table_Zh1[i].maximum_outer_diameter;
                myCommand.Parameters.Add("@x13", MySqlDbType.VarChar).Value = table_Zh1[i].maximum_wall_thickness;
                myCommand.Parameters.Add("@x14", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_pipelines_DN_150_400;
                myCommand.Parameters.Add("@x15", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_pipelines_DN_500_1400;
                myCommand.Parameters.Add("@x16", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_pipelines_DN_150_1400_in_terms_of_DN_1000;
                myCommand.Parameters.Add("@x17", MySqlDbType.VarChar).Value = table_Zh1[i].list_of_designations_of_pipes_DN_500_1400_as_table_G3;
                myCommand.Parameters.Add("@x18", MySqlDbType.VarChar).Value = table_Zh1[i].Insulation_cover_type;
                myCommand.Parameters.Add("@x19", MySqlDbType.VarChar).Value = table_Zh1[i].commissioning_year_of_the_gcs;
                myCommand.Parameters.Add("@x20", MySqlDbType.VarChar).Value = table_Zh1[i].year_of_commissioning_of_th_LPG_section;
                myCommand.Parameters.Add("@x21", MySqlDbType.VarChar).Value = table_Zh1[i].design_pressure;
                myCommand.Parameters.Add("@x22", MySqlDbType.VarChar).Value = table_Zh1[i].maximum_allowed_working_pressure;
                myCommand.Parameters.Add("@x23", MySqlDbType.VarChar).Value = table_Zh1[i].average_number_of_cycles_less_than_10proc;
                myCommand.Parameters.Add("@x24", MySqlDbType.VarChar).Value = table_Zh1[i].average_number_of_cycles_more_than_10proc;
                myCommand.Parameters.Add("@x25", MySqlDbType.VarChar).Value = table_Zh1[i].operating_time_in_design_mode;
                myCommand.Parameters.Add("@x26", MySqlDbType.VarChar).Value = table_Zh1[i].operating_time_in_throughput_mode;
                myCommand.Parameters.Add("@x27", MySqlDbType.VarChar).Value = table_Zh1[i].operating_time_in_disconnected_state_under_pressure;
                myCommand.Parameters.Add("@x28", MySqlDbType.VarChar).Value = table_Zh1[i].operating_time_in_shutdown_state_without_pressure;
                myCommand.Parameters.Add("@x29", MySqlDbType.VarChar).Value = table_Zh1[i].if_there_were_non_design_loads;
                myCommand.Parameters.Add("@x30", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_surveyed_pipelines_DN_150_400;
                myCommand.Parameters.Add("@x31", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_surveyed_pipelines_DN_500_1400;
                myCommand.Parameters.Add("@x32", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_surveyed_pipelines_DN_150_1400_in_terms_of_DN_1000;
                myCommand.Parameters.Add("@x33", MySqlDbType.VarChar).Value = table_Zh1[i].total_number_of_inspected_joints_500_1400;
                myCommand.Parameters.Add("@x34", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_inspected_joints_500_1400_with_ext_examination;
                myCommand.Parameters.Add("@x35", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_circular_joints_500_1400_defects_requiring_repair;
                myCommand.Parameters.Add("@x36", MySqlDbType.VarChar).Value = table_Zh1[i].total_number_of_inspected_pipes_500_1400;
                myCommand.Parameters.Add("@x37", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_inspected_pipes_500_1400_with_ext_examination;
                myCommand.Parameters.Add("@x38", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_pipes_500_1400_defects_requiring_repair;
                myCommand.Parameters.Add("@x39", MySqlDbType.VarChar).Value = table_Zh1[i].total_number_of_inspected_details_500_1400;
                myCommand.Parameters.Add("@x40", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_inspected_details_500_1400_with_ext_examination;
                myCommand.Parameters.Add("@x41", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_details_500_1400_defects_requiring_repair;
                myCommand.Parameters.Add("@x42", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_monitored_valves;
                myCommand.Parameters.Add("@x43", MySqlDbType.VarChar).Value = table_Zh1[i].number_of_valves_requiring_repair;
                myCommand.Parameters.Add("@x44", MySqlDbType.VarChar).Value = table_Zh1[i].soil_resistivity;
                myCommand.Parameters.Add("@x45", MySqlDbType.VarChar).Value = table_Zh1[i].length_peeling_insulating_coating;
                myCommand.Parameters.Add("@x46", MySqlDbType.VarChar).Value = table_Zh1[i].area_of_zones_not_controlled_VTD;
                myCommand.Parameters.Add("@x47", MySqlDbType.VarChar).Value = table_Zh1[i].year_of_KRTT;
                myCommand.Parameters.Add("@x48", MySqlDbType.VarChar).Value = table_Zh1[i].year_of_VD;
                myCommand.Parameters.Add("@x49", MySqlDbType.VarChar).Value = table_Zh1[i].total_length_of_repared_pipelines_DN_150_1400_in_terms_of_DN_1000;
                myCommand.Parameters.Add("@x50", MySqlDbType.VarChar).Value = table_Zh1[i].planned_date_of_EPB;
                myCommand.Parameters.Add("@x51", MySqlDbType.VarChar).Value = table_Zh1[i].note;
                myCommand.Parameters.Add("@x52", MySqlDbType.VarChar).Value = table_Zh1[i].ID_CS;
                myCommand.Parameters.Add("@x53", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].IdenticalObjectsInsideTheCC);
                myCommand.Parameters.Add("@x54", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].OtherObjectsInsideTheCC);
                myCommand.Parameters.Add("@x55", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].IdenticalObjectsOfTheNeighboringCC);
                myCommand.Parameters.Add("@x56", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].OtherObjectsOfTheNeighboringCC);
                myCommand.Parameters.Add("@x57", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].ObjectsOfTheCС);
                myCommand.Parameters.Add("@x58", MySqlDbType.VarChar).Value = Converting.concatNumbersOfObjects(table_Zh1[i].NeighboringTheCС);
                if (myCommand.ExecuteNonQuery() == 1)
                {
                    //MessageBox.Show("Account is created");
                }
                else
                {
                    MessageBox.Show("Account is not created");
                }
            }
            myConnection.Close();//Закрываем соединение с базой данных

        }
        private void SendZh2ToDBMeth()
        {
            //string CommandText = "SELECT * FROM `users` WHERE `login` = @ul";//"SELECT * FROM `users` WHERE `login` = @ul AND `pass` = @up"
            //string Connect = "Database=sql11475221;Data Source=sql11.freemysqlhosting.net;User Id=sql11475221;Password=HjT9hIIX8J";
            //string Connect = "Database=u1611387_bdforwork;Data Source=server179.hosting.reg.ru;User Id=u1611387_bduser;Password=Mirinda22+; charset = utf8";
            string CommandText = "INSERT INTO `Zh2` (`x1`, `x2`, `x3`, `x4`, `x5`, `x6`, `x7`, `x8`, `x9`, `x10`, `x11`, `x12`, `x13`, `x14`, `x15`, " +
                "`x16`, `x17`, `x18`, `x19`, `x20`, `x21`, `x22`, `x23`, `x24`, `x25`, `x26`, `x27`, `x28`, `x29`, `x30`, `x31`, `x32`, `x33`, `x34`, `x35`," +
                " `x36`, `x37`, `x38`, `x39`, `x40`, `x41`, `x42`, `x43`, `x44`, `x45`, `x46`, `x47`, `x48`) VALUES (@x1, @x2, @x3, @x4, @x5, @x6, @x7, @x8, @x9, @x10, @x11, @x12, @x13, @x14, @x15, @x16, @x17, @x18, @x19," +
                " @x20, @x21, @x22, @x23, @x24, @x25, @x26, @x27, @x28, @x29, @x30, @x31, @x32, @x33, @x34, @x35, @x36, @x37, @x38, @x39, @x40, @x41, @x42," +
                " @x43, @x44, @x45, @x46, @x47, @x48);";

            MySqlConnection myConnection = new MySqlConnection(Connect);

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Идет загрузка в БД данных таблицы Ж.2")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "/->*")));
            int incrementor = 0;
            myConnection.Open(); //Устанавливаем соединение с базой данных.
            for (int i = 0; i < table_Zh2.Count; i++)
            {

                MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);
                myCommand.Parameters.Add("@x1", MySqlDbType.VarChar).Value = table_Zh2[i].ID_TP;
                myCommand.Parameters.Add("@x2", MySqlDbType.VarChar).Value = table_Zh2[i].lpu_name;
                myCommand.Parameters.Add("@x3", MySqlDbType.VarChar).Value = table_Zh2[i].gcs_name;
                myCommand.Parameters.Add("@x4", MySqlDbType.VarChar).Value = table_Zh2[i].kc_name;
                myCommand.Parameters.Add("@x5", MySqlDbType.VarChar).Value = table_Zh2[i].tt_gcs_type;
                myCommand.Parameters.Add("@x6", MySqlDbType.VarChar).Value = table_Zh2[i].technological_number_tt_gks;
                myCommand.Parameters.Add("@x7", MySqlDbType.VarChar).Value = table_Zh2[i].diameter_exterlal;

                if (table_Zh2[i].is_underground)
                {
                    myCommand.Parameters.Add("@x8", MySqlDbType.VarChar).Value = "1";
                }
                else
                {
                    myCommand.Parameters.Add("@x8", MySqlDbType.VarChar).Value = "0";
                }



                myCommand.Parameters.Add("@x9", MySqlDbType.VarChar).Value = String.Concat(Convert.ToInt32(table_Zh2[i].date_of_examination.Year), "-12-31T01:01:01");
                //richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine +Convert.ToString(table_Zh2[i].date_of_examination.Year))));
                if (table_Zh2[i].is_insulation_removed)
                {
                    myCommand.Parameters.Add("@x10", MySqlDbType.VarChar).Value = "1";
                }
                else
                {
                    myCommand.Parameters.Add("@x10", MySqlDbType.VarChar).Value = "0";
                }

                myCommand.Parameters.Add("@x11", MySqlDbType.VarChar).Value = table_Zh2[i].number_of_diagnosed_areal;
                myCommand.Parameters.Add("@x12", MySqlDbType.VarChar).Value = table_Zh2[i].length_of_diagnosed_areal;
                myCommand.Parameters.Add("@x13", MySqlDbType.VarChar).Value = table_Zh2[i].diagnostician;
                if (table_Zh2[i].is_external_diagnostics)
                {
                    myCommand.Parameters.Add("@x14", MySqlDbType.VarChar).Value = "1";
                }
                else
                {
                    myCommand.Parameters.Add("@x14", MySqlDbType.VarChar).Value = "0";
                }

                if (table_Zh2[i].is_VTD_diagnostics)
                {
                    myCommand.Parameters.Add("@x15", MySqlDbType.VarChar).Value = "1";
                }
                else
                {
                    myCommand.Parameters.Add("@x15", MySqlDbType.VarChar).Value = "0";
                }
                myCommand.Parameters.Add("@x16", MySqlDbType.VarChar).Value = table_Zh2[i].methods_of_diagnostics;
                myCommand.Parameters.Add("@x17", MySqlDbType.VarChar).Value = table_Zh2[i].number_of_diagnosed_element;
                myCommand.Parameters.Add("@x18", MySqlDbType.VarChar).Value = table_Zh2[i].type_of_element;
                myCommand.Parameters.Add("@x19", MySqlDbType.VarChar).Value = table_Zh2[i].element_design_in_Zh3;
                myCommand.Parameters.Add("@x20", MySqlDbType.VarChar).Value = table_Zh2[i].conclusion_of_diagnostics;
                myCommand.Parameters.Add("@x21", MySqlDbType.VarChar).Value = table_Zh2[i].binding_object;
                myCommand.Parameters.Add("@x22", MySqlDbType.VarChar).Value = table_Zh2[i].distance_to_binding_object;
                myCommand.Parameters.Add("@x23", MySqlDbType.VarChar).Value = table_Zh2[i].angular_orientation_of_weld_1;
                myCommand.Parameters.Add("@x24", MySqlDbType.VarChar).Value = table_Zh2[i].angular_orientation_of_weld_2;
                myCommand.Parameters.Add("@x25", MySqlDbType.VarChar).Value = table_Zh2[i].element_thikness;
                myCommand.Parameters.Add("@x26", MySqlDbType.VarChar).Value = table_Zh2[i].element_length;
                myCommand.Parameters.Add("@x27", MySqlDbType.VarChar).Value = table_Zh2[i].element_ovality;
                myCommand.Parameters.Add("@x28", MySqlDbType.VarChar).Value = table_Zh2[i].length_of_coating_peeling;
                myCommand.Parameters.Add("@x29", MySqlDbType.VarChar).Value = table_Zh2[i].beginning_of_coating_peeling;
                myCommand.Parameters.Add("@x30", MySqlDbType.VarChar).Value = table_Zh2[i].length_of_uncontrolled_zones;
                myCommand.Parameters.Add("@x31", MySqlDbType.VarChar).Value = table_Zh2[i].coordinates_of_uncontrolled_zones;
                myCommand.Parameters.Add("@x32", MySqlDbType.VarChar).Value = table_Zh2[i].number_of_defect;
                myCommand.Parameters.Add("@x33", MySqlDbType.VarChar).Value = table_Zh2[i].type_of_defect;
                myCommand.Parameters.Add("@x34", MySqlDbType.VarChar).Value = table_Zh2[i].defect_begin_coordinates;
                myCommand.Parameters.Add("@x35", MySqlDbType.VarChar).Value = table_Zh2[i].defect_end_coordinates;
                myCommand.Parameters.Add("@x36", MySqlDbType.VarChar).Value = table_Zh2[i].defect_begin_orientation;
                myCommand.Parameters.Add("@x37", MySqlDbType.VarChar).Value = table_Zh2[i].defect_end_orientation;
                myCommand.Parameters.Add("@x38", MySqlDbType.VarChar).Value = table_Zh2[i].defect_length;
                myCommand.Parameters.Add("@x39", MySqlDbType.VarChar).Value = table_Zh2[i].defect_width;
                myCommand.Parameters.Add("@x40", MySqlDbType.VarChar).Value = table_Zh2[i].defect_depth;
                myCommand.Parameters.Add("@x41", MySqlDbType.VarChar).Value = table_Zh2[i].defect_depth_in_percent;
                myCommand.Parameters.Add("@x42", MySqlDbType.VarChar).Value = table_Zh2[i].degree_of_danger;
                myCommand.Parameters.Add("@x43", MySqlDbType.VarChar).Value = table_Zh2[i].recommendations_for_defect_elimination;
                myCommand.Parameters.Add("@x44", MySqlDbType.VarChar).Value = table_Zh2[i].repair_method;
                myCommand.Parameters.Add("@x45", MySqlDbType.VarChar).Value = String.Concat(Convert.ToInt32(table_Zh2[i].date_of_repair.Year), "-12-31T01:01:01");
                myCommand.Parameters.Add("@x46", MySqlDbType.VarChar).Value = table_Zh2[i].repair_contractor;
                myCommand.Parameters.Add("@x47", MySqlDbType.VarChar).Value = table_Zh2[i].type_for_sort;
                myCommand.Parameters.Add("@x48", MySqlDbType.VarChar).Value = table_Zh2[i].rang_of_danger;

                if (myCommand.ExecuteNonQuery() == 1)
                {
                    incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.                
                    if (incrementor > 29)
                    {
                        richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                        incrementor = 0;
                    }
                }
                else
                {
                    MessageBox.Show("Upload error");
                }
            }
            myConnection.Close();//Закрываем соединение с базой данных

        }
    }
}
