﻿//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Zulu.Table.CachedWorkSheetNamespace;
using Zulu.Table.SpreadSheet;
using Zulu.Table.Table;

namespace Zulu.Table.NunitTests
{
  #region Tests
  [TestFixture]
  [Category("SpreadSheetHelpersTests")]
  public class SpreadSheetHelpersTests
  {
    /// <summary>
    /// In Excel/OpenOffice, the colums are numbered A..Z, AA..AZ, BA..BZ, ....
    /// This functionality is implemented in the method 'Excel.ExcelColumn.intToAZ(i)'.
    ///
    /// This tests verifies the method 'Excel.ExcelColumn.intToAZ(i)'.
    /// </summary>
    [TestCase(0, ExpectedResult = "A")]
    [TestCase(1, ExpectedResult = "B")]
    [TestCase(25, ExpectedResult = "Z")]
    [TestCase(26, ExpectedResult = "AA")]
    [TestCase(27, ExpectedResult = "AB")]
    public string Test_AZ(int i)
    {
      return SpreadSheetReaderFactory.IntToAZ(i);
    }

    /// <summary>
    /// Given a cell reference, for example AB5, this will return the zero-based row and the column index.
    /// </summary>
    [TestCase("A1", ExpectedResult = "0 0")]
    [TestCase("A2", ExpectedResult = "0 1")]
    [TestCase("B2", ExpectedResult = "1 1")]
    [TestCase("AA22", ExpectedResult = "26 21")]
    [TestCase("AB22", ExpectedResult = "27 21")]
    [TestCase("BD7", ExpectedResult = "55 6")]
    public string Test_AZ_reverse(string addressString)
    {
      CellAddress address = SpreadSheetReaderFactory.AZtoAddress(addressString);
      return $"{address.Column0} {address.Row0}";
    }
  }

  public abstract class TestsBase
  {
    protected string PathToTestSheet
    {
      get
      {
        string directoryAssembly = System.Reflection.Assembly.GetAssembly(typeof(TestsBase)).Location;
        // string directoryTrunk = Path.GetFullPath(Path.Combine(directoryAssembly, "..", "..", "..", "..", "..", "zuluspreadsheet_example", "bin", "Debug", "net45"));
        string directoryTrunk = Path.GetFullPath(Path.Combine(directoryAssembly, "..", "..", "..", "..", "..", "zuluspreadsheet_example"));
        return directoryTrunk;
      }
    }

    // In cell 'G3' in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods
    protected const string TableReference = "Reference: row 3 in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'";
    protected const string CellReferenceE13 = "Reference: cell 'E13' in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'";
    protected const string CellReferenceG3 = "Reference: cell 'G3' in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'";

    protected object fixOdtXls(string msg)
    {
      foreach (string ext in SpreadSheetReaderFactory.EXTENSIONS_EXCEL)
      {
        msg = msg.Replace(ext, SpreadSheetReaderFactory.EXTENSION_ODS);
      }
      return msg;
    }
  }

  [TestFixture]
  [Category("SpreadSheetTests")]
  public class SpreadSheetTests : TestsBase
  {
    private CachedSpreadSheet GetCachedSpreadSheet(string extension)
    {
      return new CachedSpreadSheet(PathToTestSheet + "/zuluspreadsheet_test." + extension);
    }

    [Test]
    public void TestFailToOpenDocument([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      FileNotFoundException ex = Assert.Throws<FileNotFoundException>(
        delegate
       {
         SpreadSheetReaderFactory.factory(PathToTestSheet + "/wrong_filename." + extension);
       });
    }

    [Test]
    public void TestWorksheetReference([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      IWorksheet worksheet = GetCachedSpreadSheet(extension)["SheetA"].Worksheet;
      Assert.AreEqual("worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'", fixOdtXls(worksheet.Reference));
    }

    [Test]
    public void TestRowReference([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      IRow row = GetCachedSpreadSheet(extension)["SheetA"][1];
      Assert.AreEqual("row 2 in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'", fixOdtXls(row.Reference));
    }

    [Test]
    public void TestCellReference([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"]["G3"];
      Assert.AreEqual("cell 'G3' in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods'", fixOdtXls(cell.Reference));
    }

    [Test]
    public void TestFailToReadInteger([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"]["E13"];
      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
       delegate
       {
         cell.Parse(out int i);
       });

      string msg = "'male' is not a valid integer!";
      Assert.AreEqual(msg, ex.Msg);
      Assert.AreEqual($"{msg} {CellReferenceE13}", fixOdtXls(ex.Message));
    }

    enum EnumGender { EnumA, EnumB, EnumC };
    [Test]
    public void TestFailToReadEnum([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"]["E13"];
      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
       delegate
       {
         cell.Parse(out EnumGender enumGender);
       });
      string msg = "'male' is not valid. Use one of EnumA, EnumB, EnumC!";
      Assert.AreEqual(msg, ex.Msg);
      Assert.AreEqual($"{msg} {CellReferenceE13}", fixOdtXls(ex.Message));
    }

    [Test]
    public void TestReadFloat([Values(SpreadSheetReaderFactory.EXTENSION_ODS)] string extension)
    {
      // TODO: SpreadSheetReaderFactory.EXTENSION_XLSX
      //   This test failes for float in case of Excel
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"]["F16"];

      cell.Parse(out double size);

      Assert.AreEqual(165.35, size, delta: 0.00000001);
    }

    [Test]
    public void TestReadDate([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      // TODO: SpreadSheetReaderFactory.EXTENSION_XLSX
      //   This test failes for DateTime in case of Excel
      CachedSpreadSheet css = GetCachedSpreadSheet(extension);
      ICell cell = css["SheetA"]["G16"];

      {
        string dateTimeString = cell.String;
        string expectedTimeString = css.SpreadSheetReader.IsExcel ? "33646.043090277803" : "1992-02-12";
        Assert.AreEqual(expectedTimeString, dateTimeString);
      }

      {
        cell.Parse(out DateTime dateTime);
        // On Excel, we have some minutes and seconds: This is a rounding error of the float stored by excel.
        // We format the 'dateTime' as a string to get rid of minutes and seconds.
        string dateTimeString = dateTime.ToString("yyyy-MM-dd");
        Assert.AreEqual("1992-02-12", dateTimeString);
      }
    }

    [Test]
    public void TestAccessByAB([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"]["C5"];
      Assert.AreEqual("Spalte4Zeile7", cell.String);
    }

    [Test]
    public void TestAccessByColumnCell([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      // Note: Index is 0based. [row][column].
      ICell cell = GetCachedSpreadSheet(extension)["SheetA"][4][2];
      Assert.AreEqual("Spalte4Zeile7", cell.String);
    }

    [Test]
    public void TestNamedCells([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      foreach (var list in new List<Tuple<string, string>> {
             Tuple.Create( "NamedCellA", "cell 'B2' in worksheet 'Named Cells' in file 'zuluspreadsheet_test.<EXTENSION>'" ),
             Tuple.Create( "NamedCellB", "cell 'B4' in worksheet 'Named Cells' in file 'zuluspreadsheet_test.<EXTENSION>'" ),
          })
      {
        string namedCellName = list.Item1;
        string expectedReference = list.Item2.Replace("<EXTENSION>", extension);

        CachedSpreadSheet spreadSheet = GetCachedSpreadSheet(extension);
        INamedCells namedCells = spreadSheet.NamedCells;

        Assert.AreEqual(2, namedCells.Names.Length);

        // Access a named cell and get it's value
        string value = namedCells[namedCellName];
        string expectedValue = $"This is: {namedCellName}";
        Assert.AreEqual(expectedValue, value);

        // Access a named cell and get it's cell
        ICell cell = namedCells.GetCell(namedCellName);
        Assert.AreEqual(expectedReference, cell.Reference);
      }
    }
  }

  [TestFixture]
  [Category("TableTests")]
  public class TableTests : TestsBase
  {

    /// <summary>
    /// This class represents 'TableA' in our Excel/OpenOffice-Sheet
    /// </summary>
    [TableName("TableA")]
    private class TableARow : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }
      public readonly string Spalte4 = null;
      public readonly string Spalte5 = null;
      public readonly string Spalte6 = null;
      public readonly string Spalte7 = null;
      public readonly int Spalte8 = 0;
    }

    /// <summary>
    /// This test accesses the cell H6, which contains the value 1008
    /// </summary>
    /// <param name="extension">Test Excel and OpenOffice</param>
    [Test]
    public void TestOk([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ITableCollection tableCollection = TableCollection.factory(PathToTestSheet + "/zuluspreadsheet_test." + extension);
      const int cellH6 = 1008;

      //
      // Access 'TableA' using names
      //
      foreach (ITableRow row in tableCollection["TableA"])
      {
        ICell cell = row["Spalte5"];
        if (cell.String == "D")
        {
          // Verify contents of cell H6
          Assert.AreEqual("1008", row["Spalte8"].String);
          // Verify contents of cell H6
          Assert.AreEqual(cellH6, row["Spalte8"].Parse<int>());
        }
      }

      //
      // Access 'TableA' using a class 'TableA' representing one row
      //
      foreach (TableARow row in tableCollection.TypedRows<TableARow>())
      {
        if (row.Spalte5 == "D")
        {
          // Verify contents of cell H6
          Assert.AreEqual(cellH6, row.Spalte8);
        }
      }

      //
      // Same as above but using Linq
      //
      int value = tableCollection.TypedRows<TableARow>().Where(row => row.Spalte5 == "D").Select(row => row.Spalte8).First();
      // Verify contents of cell H6
      Assert.AreEqual(1008, value);
    }

    /// <summary>
    /// This class represents 'TableA' in our Excel/OpenOffice-Sheet.
    /// </summary>
    [TableName("TableA")]
    private class TableARow_WrongDatatype : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }
      /// <summary>'int' is a wrong datatype for 'Spalte7'.</summary>
      public readonly int Spalte7 = 0;
    }

    /// <summary>
    /// Test that a exception is thrown when a cell contains a string 'Spalte7Zeile5' but an integer is expected
    /// </summary>
    [Test]
    public void Test_WrongDatatype([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ITableCollection tableCollection = TableCollection.factory(PathToTestSheet + "/zuluspreadsheet_test." + extension);
      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
         delegate
         {
           foreach (TableARow_WrongDatatype row in tableCollection.TypedRows<TableARow_WrongDatatype>()) { }
         });
      string msg = "'Spalte7Zeile5' is not a valid integer!";
      Assert.AreEqual(msg, ex.Msg);
      Assert.AreEqual(msg + " " + CellReferenceG3, fixOdtXls(ex.Message));
    }

    public class UserTypeOnlyHttpsUri : Uri
    {
      private UserTypeOnlyHttpsUri(string uri) : base(uri) { }

      private const string ALLOWED_SCHEME = "https";
      public static bool TryParse(string s, out UserTypeOnlyHttpsUri uri)
      {
        uri = new UserTypeOnlyHttpsUri(s);
        if (uri.Scheme != ALLOWED_SCHEME)
        {
          throw new FormatException($"Only {ALLOWED_SCHEME} is allowed");
        }
        return true;
      }
    }

    [Test]
    public void TestUserType([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ITableCollection tableCollection = TableCollection.factory(PathToTestSheet + "/zuluspreadsheet_test." + extension);
      string[] expectedValues = new string[] {
        "github.com", // https://github.com/hmaerki/ZuluSpreadSheet
        "www.nuget.org", // https://www.nuget.org/packages/zuluspreadsheet
      };
      ITableRow[] rows = tableCollection["TableC"].ToArray();
      for (int row = 0; row < expectedValues.Length; row++)
      {
        ICell cell = rows[row]["OnlyHttpsUri"];
        cell.Parse(out UserTypeOnlyHttpsUri uri);
        Assert.AreEqual(expectedValues[row], uri.Host);
      }

      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
        delegate
        {
          // http://www.url.ch
          ICell cell = rows[2]["OnlyHttpsUri"];
          // The following line will throw an exception
          cell.Parse(out UserTypeOnlyHttpsUri uri);
        });
      string msg = "'http://www.url.ch/uv' is not a valid UserTypeOnlyHttpsUri (Only https is allowed)!";
      Assert.AreEqual(msg, ex.Msg);
    }

    /// <summary>
    /// This class represents 'TableA' in our Excel/OpenOffice-Sheet.
    /// </summary>
    [TableName("TableA")]
    private class TableARow_WrongColumnName : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }
      /// <summary>Column 'Spalte77' is invalid.</summary>
      public readonly string Spalte77 = null;
    }

    /// <summary>
    /// Test that a exception is thrown when a column, in this case 'Spalte77', does not exist. 
    /// </summary>
    [Test]
    public void Test_WrongColumnName([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ITableCollection tableCollection = TableCollection.factory(PathToTestSheet + "/zuluspreadsheet_test." + extension);

      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
        delegate
        {
          foreach (TableARow_WrongColumnName row in tableCollection.TypedRows<TableARow_WrongColumnName>()) { }
        });

      // No column 'Spalte77'! Existing columns are (Spalte4|Spalte5|Spalte6|Spalte7|Spalte8). Reference: row 3 in worksheet 'SheetA' in file 'zuluspreadsheet_test.ods''
      string msg = "No column 'Spalte77'! Existing columns are (Spalte4|Spalte5|Spalte6|Spalte7|Spalte8).";
      Assert.AreEqual(msg, ex.Msg);
      Assert.AreEqual(msg + " " + TableReference, fixOdtXls(ex.Message));
    }

    /// <summary>
    /// This class represents 'TableA' in our Excel/OpenOffice-Sheet.
    /// </summary>
    [TableName("TableA_WrongTableName")]
    private class TableARow_WrongTableName : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }
      public readonly string Spalte7 = null;
    }

    /// <summary>
    /// Test that a exception is thrown when a table, in this case 'TableA_WrongTableName', does not exist. 
    /// </summary>
    [Test]
    public void Test_WrongTableName([Values(SpreadSheetReaderFactory.EXTENSION_ODS, SpreadSheetReaderFactory.EXTENSION_XLSX)] string extension)
    {
      ITableCollection tableCollection = TableCollection.factory(PathToTestSheet + "/zuluspreadsheet_test." + extension);

      SpreadSheetException ex = Assert.Throws<SpreadSheetException>(
        delegate
        {
          foreach (TableARow_WrongTableName row in tableCollection.TypedRows<TableARow_WrongTableName>()) { }
        });
      Assert.AreEqual("Table 'TableA_WrongTableName' does not exist! Existing tables are (Equipment|Measurement|TableA|TableC).", ex.Msg);
    }
  }
  #endregion
}
