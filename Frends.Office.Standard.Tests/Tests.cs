using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Data;

namespace Frends.Office.Standard.Tests
{
    [TestFixture]
    class TestClass
    {
        [TestFixture]
        public class Tests
        {
            [Test]
            public void TestCsvToDataTable()
            {
                string input = "a;b;c\r\n1;2;3";
                string lineDelimiter = "\r\n";
                char cellDelimiter = ';';

                DataTable dt = WriteExcelFile.CsvToDataTable(input, lineDelimiter, cellDelimiter);

                Assert.That(dt.Columns.Count, Is.EqualTo(3));
                Assert.That(dt.Rows.Count, Is.EqualTo(1));

                Assert.That(dt.Columns[0].ColumnName, Is.EqualTo("a"));
                Assert.That(dt.Columns[2].ColumnName, Is.EqualTo("c"));

                Assert.That(dt.Rows[0][0], Is.EqualTo("1"));
                Assert.That(dt.Rows[0][2], Is.EqualTo("3"));
            }

            [Test]
            public void TestCreateWorkbookObject()
            {
                WriteExcelFileInput i = new WriteExcelFileInput();
                i.StringInput = "a;b;c\r\n1;2;3";
                i.CellDelimiter = ';';
                i.LineDelimiter = "\r\n";

                XLWorkbook w = WriteExcelFile.CreateWorkbookObject(i);

                IXLWorksheet s = w.Worksheet(1);

                Assert.That(s.Row(1).Cell(1).Value, Is.EqualTo("a"));
                Assert.That(s.Row(2).Cell(3).Value, Is.EqualTo("3"));

                Assert.That(s.Column(1).Cell(1).Value, Is.EqualTo("a"));
                Assert.That(s.Column(1).Cell(2).Value, Is.EqualTo("1"));
            }
        }
    }
}
