using DataProcessing.STRUCT;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace DataProcessing
{
    public class Student_Management
    {
        private List<Student> list = new List<Student>();

        
        public string Student_Insert(string name, string date, string gpa)
        {
            var check_name = Common.Validate.Check_Name(name);
            if (check_name == false)
            {
                return "Ten khong hop le";
            }

            var check_date = Common.Validate.Check_DateTime(date);
            if (check_date == false)
            {
                return "Ngay sinh khong hop le";
            }

            var check_gpa = Common.Validate.Check_GPA(gpa);
            if (check_gpa == false)
            {
                return "Diem trung binh khong hop le";
            }

            var stu = new Student
            {
                Name = name,
                Date = DateTime.ParseExact(date, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture),
                GPA = (int)float.Parse(gpa) // Sử dụng float cho GPA
            };

            list.Add(stu);
            return "Them thanh cong !";
        }

        
        public List<Student> Student_GetList()
        {
            return list;
        }

        
        public string ExportToExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Danh sách học sinh");

                    worksheet.Cells[1, 1].Value = "Ho va Ten";
                    worksheet.Cells[1, 2].Value = "Ngay sinh";
                    worksheet.Cells[1, 3].Value = "Diem trung binh";
                    worksheet.Cells[1, 4].Value = "Xep loai";

                    for (int i = 0; i < list.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = list[i].Name;
                        worksheet.Cells[i + 2, 2].Value = list[i].Date.ToString("dd/MM/yyyy");
                        worksheet.Cells[i + 2, 3].Value = list[i].GPA;
                        worksheet.Cells[i + 2, 4].Value = GetClassification(list[i].GPA);
                    }

                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }

                return "Xuat thanh cong ra file Excel!";
            }
            catch (Exception ex)
            {
                return $"Loi xuat du lieu: {ex.Message}";
            }
        }

        
        public List<Student> GetStudentsByPerformance(string classification)
        {
            List<Student> filteredStudents = new List<Student>();

            foreach (var student in list)
            {
                if (GetClassification(student.GPA).Equals(classification, StringComparison.OrdinalIgnoreCase))
                {
                    filteredStudents.Add(student);
                }
            }

            return filteredStudents;
        }


        
        private string GetClassification(float gpa)
        {
            if (gpa >= 0.0 && gpa <= 3.0)
                return "Hoc Lai";
            else if (gpa >= 3.0 && gpa <= 4.9)
                return "Yeu";
            else if (gpa >= 5.0 && gpa <= 6.5)
                return "Trung Binh";
            else if (gpa >= 6.5 && gpa <= 7.9)
                return "Kha";
            else if (gpa >= 8.0)
                return "Gioi";
            else
                return "Khong hop le";

       

        }
    }
}
