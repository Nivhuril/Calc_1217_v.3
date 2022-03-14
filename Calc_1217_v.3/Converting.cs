using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalcGCS_1217
{
    partial class Form1 : Form
    {
        private DateTime getDataTime(string date)//конвертируем строку в настоящую дату
        {
            DateTime someDate = new DateTime();
            int length = date.Length;
            if (length == 4)
            {
                int day = 31;
                int mount = 12;
                int year = Convert.ToInt32(date);
                someDate = new DateTime(year, mount, day);
                richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("короткое имя")));
            }
            else
            {
                try
                {
                    int day = Convert.ToInt32(date.Substring(0, 2));
                    int mount = Convert.ToInt32(date.Substring(3, 2));
                    int year = Convert.ToInt32(date.Substring(6, 4));
                    someDate = new DateTime(year, mount, day);
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("длинное имя")));
                }
                catch (Exception)
                {
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("не получилось")));
                }
            }
            return someDate;
        }
    }
    internal class Converting
        {
            public static int getStringNumber(int ID_TP)//По номеру ID_TP узнаём номер строки списка, куда добавлять результат расчета
            {
                int result = 2;
                if (ID_TP == 10001402) { result = 0; }
                if (ID_TP == 10001401) { result = 1; }
                if (ID_TP == 10001202) { result = 2; }
                if (ID_TP == 10001201) { result = 3; }
                if (ID_TP == 10001301) { result = 4; }
                if (ID_TP == 10001302) { result = 5; }
                if (ID_TP == 10002402) { result = 6; }
                if (ID_TP == 10002401) { result = 7; }
                if (ID_TP == 10002202) { result = 8; }
                if (ID_TP == 10002201) { result = 9; }
                if (ID_TP == 10002301) { result = 10; }
                if (ID_TP == 10002302) { result = 11; }
                if (ID_TP == 10003402) { result = 12; }
                if (ID_TP == 10003401) { result = 13; }
                if (ID_TP == 10003202) { result = 14; }
                if (ID_TP == 10003201) { result = 15; }
                if (ID_TP == 10003301) { result = 16; }
                if (ID_TP == 10003302) { result = 17; }
                if (ID_TP == 10004402) { result = 18; }
                if (ID_TP == 10004401) { result = 19; }
                if (ID_TP == 10004202) { result = 20; }
                if (ID_TP == 10004201) { result = 21; }
                if (ID_TP == 10004301) { result = 22; }
                if (ID_TP == 10004302) { result = 23; }
                if (ID_TP == 10005402) { result = 24; }
                if (ID_TP == 10005401) { result = 25; }
                if (ID_TP == 10005202) { result = 26; }
                if (ID_TP == 10005201) { result = 27; }
                if (ID_TP == 10005301) { result = 28; }
                if (ID_TP == 10005302) { result = 29; }
                if (ID_TP == 10006402) { result = 30; }
                if (ID_TP == 10006401) { result = 31; }
                if (ID_TP == 10006202) { result = 32; }
                if (ID_TP == 10006201) { result = 33; }
                if (ID_TP == 10006301) { result = 34; }
                if (ID_TP == 10006302) { result = 35; }
                if (ID_TP == 10007402) { result = 36; }
                if (ID_TP == 10007401) { result = 37; }
                if (ID_TP == 10007201) { result = 38; }
                if (ID_TP == 10007202) { result = 39; }
                if (ID_TP == 10007301) { result = 40; }
                if (ID_TP == 10007302) { result = 41; }
                if (ID_TP == 10008402) { result = 42; }
                if (ID_TP == 10008401) { result = 43; }
                if (ID_TP == 10008201) { result = 44; }
                if (ID_TP == 10008301) { result = 45; }
                if (ID_TP == 10009402) { result = 46; }
                if (ID_TP == 10009401) { result = 47; }
                if (ID_TP == 10009201) { result = 48; }
                if (ID_TP == 10009202) { result = 49; }
                if (ID_TP == 10009301) { result = 50; }
                if (ID_TP == 10009302) { result = 51; }
                if (ID_TP == 10010402) { result = 52; }
                if (ID_TP == 10010401) { result = 53; }
                if (ID_TP == 10010201) { result = 54; }
                if (ID_TP == 10010301) { result = 55; }
                if (ID_TP == 10011402) { result = 56; }
                if (ID_TP == 10011401) { result = 57; }
                if (ID_TP == 10011201) { result = 58; }
                if (ID_TP == 10011202) { result = 59; }
                if (ID_TP == 10011301) { result = 60; }
                if (ID_TP == 10011302) { result = 61; }
                if (ID_TP == 10012402) { result = 62; }
                if (ID_TP == 10012401) { result = 63; }
                if (ID_TP == 10012201) { result = 64; }
                if (ID_TP == 10012301) { result = 65; }
                if (ID_TP == 10013402) { result = 66; }
                if (ID_TP == 10013401) { result = 67; }
                if (ID_TP == 10013201) { result = 68; }
                if (ID_TP == 10013202) { result = 69; }
                if (ID_TP == 10013301) { result = 70; }
                if (ID_TP == 10013302) { result = 71; }
                if (ID_TP == 10014402) { result = 72; }
                if (ID_TP == 10014401) { result = 73; }
                if (ID_TP == 10014201) { result = 74; }
                if (ID_TP == 10014202) { result = 75; }
                if (ID_TP == 10014301) { result = 76; }
                if (ID_TP == 10014302) { result = 77; }
                if (ID_TP == 10015402) { result = 78; }
                if (ID_TP == 10015401) { result = 79; }
                if (ID_TP == 10015201) { result = 80; }
                if (ID_TP == 10015301) { result = 81; }
                if (ID_TP == 10016402) { result = 82; }
                if (ID_TP == 10016401) { result = 83; }
                if (ID_TP == 10016201) { result = 84; }
                if (ID_TP == 10016202) { result = 85; }
                if (ID_TP == 10016301) { result = 86; }
                if (ID_TP == 10016302) { result = 87; }
                if (ID_TP == 10017402) { result = 88; }
                if (ID_TP == 10017401) { result = 89; }
                if (ID_TP == 10017201) { result = 90; }
                if (ID_TP == 10017301) { result = 91; }
                return result;
            }
            public static double ConvertToDouble(string inputString)//конвертируем строку в число double
            {
                double result = 0;

                if (Double.TryParse(inputString.Trim().Replace(".", ","), out result))
                {
                    result = Double.Parse(inputString.Trim().Replace(".", ","));
                }

                return result;
            }
            public static int ConvertToInt(string inputString)//конвертируем строку в число int
            {
                int result = 0;
                try
                {
                    result = Int32.Parse(inputString.Trim());
                }
                catch (Exception)
                {
                    result = 0;
                }
                return result;
            }
            public static bool GetBoolean(string input, string is_true)//проверяем условие "да"/"нет"
            {
                bool result = false;

                if (String.IsNullOrEmpty(input) == false)
                {
                    if (input.Contains(is_true))
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }


                return result;
            }
            public static List<int> GetNumbersOfTT(string input)//превращаем перечисленные через запятую натуральные числа в список
            {
                List<int> result = new List<int>();
                input = input.Replace(" ", "");
                int n0 = 0;
                bool mark = true;
                while (mark)
                {
                    int num = input.IndexOf(",");
                    int currentNumber = 0;
                    if (num >= 0)
                    {
                        currentNumber = Converting.ConvertToInt(input.Substring(n0, num));
                        input = input.Substring(num + 1, input.Length - (num + 1));
                    }
                    else if (num == -1 & input.Length > 2)
                    {
                        currentNumber = Converting.ConvertToInt(input.Substring(0, input.Length));
                        input = "";
                    }

                    if (currentNumber == 0)
                    {
                        mark = false;
                    }
                    else
                    {
                        //richTextBox1.AppendText("/" + currentNumber);
                        result.Add(currentNumber);
                    }
                }
                return result;
            }
            public static string concatNumbersOfObjects(List<int> input)
            {
                string result = "";
                for (int i = 0; i < input.Count; i++)
                {
                    result = String.Concat(result, ", ", input[i]);
                }

                return result;
            }
            public static double get_MT_from_table_A1(string curremt, string other)//метод для оопределания коэффициента Мт по таблице А.1
            {
                double result = 0;
                string prompl = "площ";
                string vhod = "входн";
                string vihod = "ыходн";
                if (curremt.Contains(vihod))
                {
                    if (other.Contains(vhod))
                    {
                        result = 2.4;
                    }
                    else if (other.Contains(prompl))
                    {
                        result = 18.2;
                    }
                }
                if (curremt.Contains(vhod))
                {
                    if (other.Contains(vihod))
                    {
                        result = 0.4;
                    }
                    else if (other.Contains(prompl))
                    {
                        result = 7.5;
                    }
                }
                if (curremt.Contains(prompl))
                {
                    if (other.Contains(vhod))
                    {
                        result = 0.1;
                    }
                    else if (other.Contains(vihod))
                    {
                        result = 0.05;
                    }
                }
                if (result == 0)
                {
                    result = 0.05;
                }
                return result;
            }
        }
}
    
