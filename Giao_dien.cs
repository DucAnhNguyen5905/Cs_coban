using DataProcessing;
using System;

namespace QLHS
{
    class Giao_dien
    {
        static void Main(string[] args)
        {
            Student_Management Stu_Mana = new Student_Management();
            string defaultFilePath = @"C:\Users\LENOVO\Desktop\Quan_ly_hoc_sinh.xlsx";

            while (true)
            {
                Console.WriteLine("|----- CHUONG TRINH QUAN LY HOC SINH -----|");
                Console.WriteLine("Vui long chon: ");
                Console.WriteLine("1. Them hoc sinh.");
                Console.WriteLine("2. Xuat danh sach ra file Excel");
                Console.WriteLine("3. Tim kiem hoc sinh theo hoc luc");
                Console.WriteLine("4. Thoat chuong trinh");
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        Console.Write("Nhap ten hoc sinh: ");
                        string name = Console.ReadLine();

                        Console.Write("Nhap ngay sinh (dd/MM/yyyy): ");
                        string date = Console.ReadLine();

                        Console.Write("Nhap diem trung binh: ");
                        string gpa = Console.ReadLine();

                        string result = Stu_Mana.Student_Insert(name, date, gpa);
                        Console.WriteLine(result);
                        break;  

                    case "2":
                        string exportResult = Stu_Mana.ExportToExcel(defaultFilePath);
                        Console.WriteLine(exportResult);
                        break;

                    case "3":
                        Console.Write("Nhap hoc luc (Hoc Lai, Yeu, Trung Binh, Kha, Gioi): ");
                        string classificationInput = Console.ReadLine();

                        
                        var studentsByPerformance = Stu_Mana.GetStudentsByPerformance(classificationInput);

                        if (studentsByPerformance.Count > 0)
                        {
                            Console.WriteLine($"Danh sach hoc sinh xep loai {classificationInput}:");
                            foreach (var student in studentsByPerformance)
                            {
                                Console.WriteLine($"{student.Name} - {student.Date.ToString("dd/MM/yyyy")} - {student.GPA}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Khong tim thay hoc sinh nao phu hop voi hoc luc nay.");
                        }
                        break;

                    case "4":
                        Console.WriteLine("Thoat chuong trinh");
                        return;

                    default:
                        Console.WriteLine("Tuy chon khong hop le. Vui long chon lai");
                        break;
                }
            }
        }
    }
}
