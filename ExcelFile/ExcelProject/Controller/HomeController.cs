using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelProject.Models;
using FastMember;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExcelProject.Controller
{
    public class HomeController : Microsoft.AspNetCore.Mvc.Controller
    {
        public IActionResult GetExcelFile()
        {
            ExcelPackage excelPackage = new ExcelPackage();
            var excelBlank = excelPackage.Workbook.Worksheets.Add("DOSYA");
            // excelBlank.Cells[1, 1].Value = "NAME";
            // excelBlank.Cells[1, 2].Value = "SURNAME";
            //
            // excelBlank.Cells[2, 1].Value = "Esat";
            // excelBlank.Cells[2, 2].Value = "YILMAZ";

            excelBlank.Cells.LoadFromCollection(new List<Customer>
            {
                new Customer
                {
                    Id = 1,
                    CustomerName = "Esat",
                    CustomerSurName = "Yılmaz",
                    CustomerEmail = "esatyy@outlook.com",
                    CustomerAdress = "Sivas"
                },
                new Customer
                {
                    Id = 2,
                    CustomerName = "Onur",
                    CustomerSurName = "Yılmaz",
                    CustomerAdress = "Ankara",
                    CustomerEmail = "onnur36@gmail.com"
                }
            }, true, OfficeOpenXml.Table.TableStyles.Light10);
            var bytes = excelPackage.GetAsByteArray();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\n",
                Guid.NewGuid() + "" + ".xlsx");
        }

        public IActionResult GetPdfFile()
        {
            DataTable dataTable = new DataTable();
            dataTable.Load(ObjectReader.Create(new List<Customer>
            {
                new Customer
                {
                    Id = 1,
                    CustomerName = "Esat",
                    CustomerSurName = "Yılmaz",
                    CustomerEmail = "esatyy@outlook.com",
                    CustomerAdress = "Sivas"
                },
                new Customer
                {
                    Id = 2,
                    CustomerName = "Onur",
                    CustomerSurName = "Yılmaz",
                    CustomerAdress = "Ankara",
                    CustomerEmail = "onnur36@gmail.com"
                }
            }));
            string fileName = Guid.NewGuid() + ".pdf";
            string path = Path.Combine(Directory.GetCurrentDirectory() + "www.root/documents/" + fileName);
            var stream = new FileStream(path, FileMode.Create);
            Document document = new Document(PageSize.A4, 25f, 25f, 25f, 25f);
            PdfWriter.GetInstance(document, stream);
            document.Open();
            // Paragraph paragraph = new Paragraph("PDF DOSYASI'NIN İÇİNDE BU YAZI YAZACAK");
            PdfTable pdfTable = new PdfTable(dataTable.Columns.Count);
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                pdfTable.AddCell(dataTable.Columns[i].ColumnName);
            }

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    pdfTable.AddCell(dataTable.Rows[i][j].ToString());
                }
            }
            document.Add(pdfTable);
            return File("/documents/" + fileName, "application/pdf", fileName);
        }
    }
}