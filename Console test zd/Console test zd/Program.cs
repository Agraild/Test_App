using Aspose.Cells;
using ConsoleTables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Console_test_zd
{
    class Program
    {
        //ВЫВОД ЗАДАЧ
        static void print()
            {
            // Загрузить файл Excel
            Workbook wb = new Workbook("tasks.xlsx");

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[0];

            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            int i = rows + 1;





            var table = new ConsoleTable("№", "task number", "task name", "description", "status");
            for (int u = 0; u < i; u++)
            {
                table.AddRow(u, worksheet.Cells[u, 0].Value, worksheet.Cells[u, 1].Value, worksheet.Cells[u, 2].Value, worksheet.Cells[u, 3].Value);
                
            }
            

            table.Write();
            
            }
        //ДОБАВИТЬ ЗАДАЧУ
        static void add_task(string number, string name, string description)
        {
            // Загрузить файл Excel
            Workbook wb = new Workbook("tasks.xlsx");

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[0];

            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            int i = rows+1;

            worksheet.Cells[i, 0].PutValue(number);
            worksheet.Cells[i, 1].PutValue(name);
            worksheet.Cells[i, 2].PutValue(description);


            //Save to Excel file (XLSX)
            wb.Worksheets.RemoveAt("Evaluation Warning");
                wb.Save("tasks.xlsx");
            
        }
        //ИЗМЕНИТЬ СТАТУС
        static void change_status(int change, int change2 )
        {
            Console.Clear();
            // Загрузить файл Excel
            Workbook wb = new Workbook("tasks.xlsx");

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[0];

            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            int i = rows + 1;

            string s1 = "To do";
            string s2 = "In Progress";
            string s3 = "Done";

            switch (change2)
            {

                case 1:
                    worksheet.Cells[change, 3].PutValue(s1);
                    break;
                case 2:
                    worksheet.Cells[change, 3].PutValue(s2);
                    break;
                case 3:
                    worksheet.Cells[change, 3].PutValue(s3);
                    break;
            }


            
            


            //Save to Excel file (XLSX)
            wb.Worksheets.RemoveAt("Evaluation Warning");
            wb.Save("tasks.xlsx");

        }

        //ВЫВОД ПОЛЬЗОВАТЕЛЕЙ
        static void print_users()
        {
            var table = new ConsoleTable("№", "user", "password" );

            for (int u = 0; u < intArray.Length; u++)
            {
                table.AddRow(u, Convert.ToString(intArray[u]), Convert.ToString(intArray2[u]));

            }
            
            
            table.Write();
        }

        static bool Enter = false;

        static bool sotrudnik = false;
        
        //ФУНКЦИЯ АВТОРИЗАЦИИ
        static void Authorization(string login1, string password1 )
            {
            Console.Clear();
            for (int u = 0; u < intArray.Length - 1; u++)
            {
                if (login1 == Convert.ToString(intArray[u]) && password1 == Convert.ToString(intArray2[u])) { Enter = true; if (login1 == "admin") { sotrudnik = true; }  }

            }
            if (Enter == false) { read_users();  Console.WriteLine("Данные неверны, повторите вход" + "\n");   }
            }

        static void delete_user(int num_user)
        {

                    // Загрузить файл Excel
                    Workbook wb = new Workbook("users.xlsx");

                    // Removing a worksheet using its index

                    //wb.Worksheets.RemoveAt(1);
                    //wb.Worksheets.RemoveAt("Evaluation Warning (2)");

                    // Получить все рабочие листы
                    WorksheetCollection collection = wb.Worksheets;

                    // Получить рабочий лист, используя его индекс
                    Worksheet worksheet = collection[0];

                    worksheet.Cells.ClearContents(num_user, 0, num_user, 1);

                    

                    //Save to Excel file (XLSX)
                    wb.Worksheets.RemoveAt("Evaluation Warning");
                    wb.Save("users.xlsx");
            
                
        }
        
        //ДОБАВИТЬ ПОЛЬЗОВАТЕЛЯ В БАЗУ
        static void Add_User(string login2, string password2)
        {
            // Загрузить файл Excel
            Workbook wb = new Workbook("users.xlsx");

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[0];

            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;

            string user = "";
            string password = "";

            

            user = login2;
            password = password2;

            int i = 0;
            bool stat = true;
            bool login = true;

            while (i < rows + 1)
            {
                if (user == Convert.ToString(worksheet.Cells[i, 0].Value)) { Console.WriteLine("номер уже занят"); login = false; }
                i++;

            }
            



            if (login == true)
            {
                i = 0;
                //bool isLetter = true;
                bool result = false;
                stat = false;

                while (result == false)
                {
                    string check;
                    i++;
                    check = Convert.ToString(worksheet.Cells[i, 0].Value);

                    

                    bool isLetter = Regex.IsMatch(check, @"^[a-zA-Z0-9_]+$");



                    if (isLetter == false) { result = true; stat = true; break; }


                }
                
                if (user == "") { Console.WriteLine("введите логин"); stat = false; }
                if (password == "") { Console.WriteLine("введите пароль"); stat = false; }

                if (stat == true)
                {
                    worksheet.Cells[i, 0].PutValue(user);
                    worksheet.Cells[i, 1].PutValue(password);
                }
                

                //Save to Excel file (XLSX)
                wb.Worksheets.RemoveAt("Evaluation Warning");
                wb.Save("users.xlsx");

                
            }


        }
        static string[] intArray;
        static string[] intArray2;

        static void read_users()
        {
            // Загрузить файл Excel
            Workbook wb = new Workbook("users.xlsx");
            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;
            // Получить рабочий лист, используя его индекс
            Worksheet worksheet = collection[0];

            // Получить количество строк и столбцов
            int rows = worksheet.Cells.MaxDataRow + 1;
            int cols = worksheet.Cells.MaxDataColumn;
            //массив первого столбца (серийный)
            intArray = new string[rows];
            intArray2 = new string[rows];


            int y;
            int x;
            //присваивает массиву значения из таблицы
            for (int d = 0; d < rows; d++)
            {

                string temp = (worksheet.Cells[d, 0].Value) + "";
                intArray[d] = temp;
            }
            for (int d = 0; d < rows; d++)
            {

                string temp = (worksheet.Cells[d, 1].Value) + "";
                intArray2[d] = temp;
            }


            print_users();

            
        }




        static void Main(string[] args)
        {
            
            
            read_users();
            while (true) {

                if (Enter == false)
                {
                    Console.WriteLine("Необходима авторизация");
                    Console.WriteLine("Введите логин");
                    string login = Console.ReadLine();
                    Console.WriteLine("Введите пароль");
                    string password = Console.ReadLine();
                    Authorization(login, password);
                }
            int choice;
            //вход сотр
            if(Enter == true && sotrudnik == false)
                {
                    print();
                    Console.WriteLine(">1-change status");
                    Console.WriteLine(">2-exit");
                    int choose = Convert.ToInt32(Console.ReadLine());
                    switch (choose)
                    {
                        case 1:
                            Console.Clear();
                            print();
                            Console.WriteLine("number of project to change status");
                            int project_number = Convert.ToInt32(Console.ReadLine());
                            Console.WriteLine(">1-To do");
                            Console.WriteLine(">2-In Progress");
                            Console.WriteLine(">3-Done");
                            int project_stat = Convert.ToInt32(Console.ReadLine());

                            change_status(project_number, project_stat);
                            break;
                        case 2:
                            //выход
                            Environment.Exit(0);
                            break;

                    }
                }
                
            //вход админ
            if(Enter == true && sotrudnik == true)
                {

                    

            Console.WriteLine(">1-Add task");
            Console.WriteLine(">2-Add User");
            Console.WriteLine(">3-View Tasks");
            Console.WriteLine(">4-Delete user");

            choice = Convert.ToInt32(Console.ReadLine());

            switch (choice)
            {
                
                case 1:
                            Console.Clear();
                            Console.WriteLine("number of project");
                    string project_number = Console.ReadLine();
                    Console.WriteLine("task name");
                    string task_name = Console.ReadLine();
                    Console.WriteLine("description of task");
                    string task_description = Console.ReadLine();
                            Console.Clear();

                            add_task(project_number, task_name, task_description);
                            print();

                            break;
                case 2:
                    Console.WriteLine("enter new username");
                    string login_add = Console.ReadLine();
                    Console.WriteLine("enter password");
                    string password_add = Console.ReadLine();
                    Add_User(login_add, password_add);
                            Console.Clear();
                            read_users();
                            break;
                case 3:
                    Console.Clear();
                    print();
                    break;
                case 4:
                    read_users();
                    Console.WriteLine("Введите номер пользователя для удаления");
                    int delete_us = Convert.ToInt32(Console.ReadLine());

                    delete_user(delete_us);
                            Console.Clear();
                            read_users();
                            break;
            }
            }


            }



            Console.ReadLine();
        }
    }
}
