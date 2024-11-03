using DataProcessing.STRUCT;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace DataProcessing
{
    public class Student_Management
    {
        private List<Student> list = new List<Student>();
        private readonly object student;

        public string Student_Insert(string name, string date, string gpa, string quantity)
        {
            
            var check_name = Common.Validate.Check_Name(name);
            if (!check_name)
            {
                return "Ten khong hop le";
            }

            
            var check_date = Common.Validate.Check_DateTime(date);
            if (!check_date)
            {
                return "Ngay sinh khong hop le";
            }

            
            var check_gpa = Common.Validate.Check_GPA(gpa);
            if (!check_gpa)
            {
                return "Diem trung binh khong hop le";
            }

            

            var stu = new Student
            {
                Name = name,
                Date = DateTime.ParseExact(date, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture),
                GPA = (int)float.Parse(gpa),
            };

            list.Add(stu);
            return "Them thanh cong!";
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

                    int hien_thi_gio = list.Count + 2;
                    worksheet.Cells[hien_thi_gio, 1].Value = "Du lieu duoc cap nhat vao luc:";
                    worksheet.Cells[hien_thi_gio, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    s
                    var fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);
                }

                return "Xuat thanh cong ra file Excel!";
            }
            catch (Exception ex)
            {
                return $"loi xuat du lieu: {ex.Message}";
            }
        }

        public List<Student> GetStudentsByPerformance(string classification)
        {
            List<Student> find_gpa = new List<Student>();

            foreach (var student in list)
            {
                if (GetClassification(student.GPA).Equals(classification, StringComparison.OrdinalIgnoreCase))
                {
                    find_gpa.Add(student);
                }
            }

            return find_gpa;
        }

        private string GetClassification(float gpa)
        {
            if (gpa >= 0.0 && gpa <= 3.0)
                return "Hoc Lai";
            else if (gpa > 3.0 && gpa <= 4.9)
                return "Yeu";
            else if (gpa >= 5.0 && gpa <= 6.5)
                return "Trung Binh";
            else if (gpa > 6.5 && gpa <= 7.9)
                return "Kha";
            else if (gpa >= 8.0)
                return "Gioi";
            else
                return "Khong hop le";
        }

        public List<Student> FindStudentsByName(string name)
        {
            
            List<Student> find_name = new List<Student>();

            foreach (var student in list)
            { 
                if (student.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    find_name.Add(student);
                }
            }

            return find_name;
        }

        public List<Student> FindStudentsByDate(string inputDate)
        {

            DateTime targetDate = DateTime.ParseExact(inputDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            List<Student> find_date = new List<Student>();

            foreach (var student in list)
            {
                if (student.Date.Date == targetDate)
                {
                    find_date.Add(student);
                }
            }

            return find_date;
        }

        public bool Student_Remove(string DOB)
        {
            
            if (!Common.Validate.Check_DateTime(DOB))
            {
                return false; 
            }

            
            DateTime dateOfBirth = DateTime.ParseExact(DOB, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            
            var studentToRemove = list.FirstOrDefault(student => student.Date == dateOfBirth);

            if (studentToRemove != null)
            {
                
                list.Remove(studentToRemove);
                return true; 
            }
            else
            {
                
                return false; 
            }
        }

    }
}
