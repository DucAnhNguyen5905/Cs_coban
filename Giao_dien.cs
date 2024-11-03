using DataProcessing;
using DataProcessing.STRUCT;
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
                Console.WriteLine("4. Tim kiem hoc sinh theo ten");
                Console.WriteLine("5. Tim kiem hoc sinh theo ngay sinh");
                Console.WriteLine("6. Thoat chuong trinh");
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        Console.Write("Ban muon them bao nhieu hoc sinh? ");
                        if (!int.TryParse(Console.ReadLine(), out int quantity) || quantity < 1 || quantity >= 50)
                        {
                            Console.WriteLine("So luong khong hop le!");
                            break;
                        }

                        for (int i = 0; i < quantity; i++)
                        {
                            Console.WriteLine($"Nhap thong tin cho hoc sinh thu {i + 1}:");
                            Console.Write("Nhap ten hoc sinh: ");
                            string name = Console.ReadLine();

                            Console.Write("Nhap ngay sinh (dd/MM/yyyy): ");
                            string date = Console.ReadLine();

                            Console.Write("Nhap diem trung binh: ");
                            string gpa = Console.ReadLine();

                            string result = Stu_Mana.Student_Insert(name, date, gpa, "1");
                            Console.WriteLine(result);
                        }
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
                        Console.Write("Nhap ten hoc sinh can tim: ");
                        string searchName = Console.ReadLine();

                        var studentsByName = Stu_Mana.FindStudentsByName(searchName);

                        if (studentsByName.Count > 0 && studentsByName.Count <= 50)
                        {
                            Console.WriteLine($"Danh sach hoc sinh co ten '{searchName}':");
                            foreach (var student in studentsByName)
                            {
                                Console.WriteLine($"{student.Name} - {student.Date.ToString("dd/MM/yyyy")} - {student.GPA}");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Khong tim thay hoc sinh nao phu hop voi ten nay.");
                        }
                        break;

                    case "5":
                        Console.Write("Nhap ngay sinh cua hoc sinh can tim: ");
                        string searchDate = Console.ReadLine();

                        var studentsByDate = Stu_Mana.FindStudentsByDate(searchDate);
                        
                        if (Common.Validate.Check_DateTime(searchDate))
                        {
                            Console.WriteLine($"Danh sach hoc sinh co ten '{searchDate}':");
                            foreach (var student in studentsByDate)
                            {
                                Console.WriteLine($"{student.Name} - {student.Date.ToString("dd/MM/yyyy")} - {student.GPA}");
                            }
                        } else
                        {
                            Console.WriteLine("Khong tim thay hoc sinh nao phu hop voi ten nay.");
                        }

                        Console.Write("Bạn có muốn xóa thông tin của học sinh này không? (Y/N): ");
                        string deleteConfirmation = Console.ReadLine();

                        if (deleteConfirmation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                        {
                            Stu_Mana.Student_Remove(searchDate); 
                            Console.WriteLine("Da xoa hoc sinh thanh cong.");
                        }
                        else
                        {
                            Console.WriteLine("Khong the xoa học sinh.");
                        }

                        break;

                    case "6":
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

