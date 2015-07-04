using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using Memorandum;

namespace GeneratedCode
{
    class Program
    {
        static void Main(string[] args)
        {
            Memo.semester = "2015Spring";
            Memo.instructor = "Horst Hohberger";
            Memo.studentCount = 2;
            Memo.students = new List<Student>();
            Memo.students.Add(new Student()
            {
                name = "A B",
                id = "23423532",
                decision = Decision.Guilty,
                count = 1
            });
            Memo.students.Add(new Student()
            {
                name = "C D",
                id = "2342432",
                decision = Decision.Innocent,
                count = 0
            });
            Memo.reportDate = DateTime.Now;
            Memo.code = "VV186";
            Memo.assignment = "assignment 1";
            Memo.hearingDate = DateTime.Now;
            Memo.description = "asodfijweofjiwoej";
            Memo.decision = "both of them are guilty";

            // Create a document by path, and write some text in it.
            //string fileName = @"C:\Users\Jia\Desktop\demo.docx";

            // Create the Word document. 
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Create(@"C:\Users\Jia\Desktop\" + Memo.getFileName(),
                    WordprocessingDocumentType.Document);

            // Create a MainDocumentPart instance.
            MainDocumentPart mainDocumentPart =
                wordprocessingDocument.AddMainDocumentPart();

            Memo.CreateMemo(mainDocumentPart);

            // Close the document handle
            wordprocessingDocument.Close();
        }
    }
}
