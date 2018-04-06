//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;

namespace Zulu.Table.SpreadSheet
{
  #region Interfaces
  /// <summary>
  /// Most of the classes support the <code>IReference</code> interface.
  /// The interface provides a human readable form of the objects origin in Excel- or OpenOffice.
  /// This eases programming a lot.
  /// All these classes may be passed to the <code>SpreadSheetException</code>.
  /// </summary>
  public interface IReference
  {
    /// <summary>
    /// A textual representation where the object comes from.
    /// <![CDATA[
    ///    TYPE            EXAMPLE REFERENCE
    ///    SpreadSheet     "xy.odt"
    ///    Worksheet       Worksheet "test" in "xy.odt"
    ///    Cell            Cell AZ in Worksheet "test" in "xy.odt"
    /// ]]>
    /// </summary>
    string Reference { get; }

    /// <summary>
    /// For example
    /// <![CDATA[
    ///   xy.odt          // For a ISpreadSheetReader
    ///   Configuration   // For a IWorksheet
    ///   Members         // For a Table
    ///   A5              // For a ICell
    /// ]]>
    /// </summary>
    string Name { get; }

    /// <summary>
    /// For example
    /// <![CDATA[
    ///   file 'xy.od'               // For a ISpreadSheetReader
    ///   worksheet 'Configuration'   // For a IWorksheet
    ///   table 'Members'             // For a Table
    ///   cell 'A5'                   // For a ICell
    /// ]]>
    /// </summary>
    string Description { get; }
  }

  /// <summary>
  /// The address of a cell of a Excel- or OpenOffice Calc-document.
  /// </summary>
  public struct CellAddress
  {
    /// <summary>
    /// For example <code>A5</code>
    /// </summary>
    public readonly string Text;

    /// <summary>
    /// The 0-based column. Note that the numbering in Excel- or Calc is 1-based.
    /// </summary>
    public readonly int Column0;

    /// <summary>
    /// The 0-based row. Note that the numbering in Excel- or Calc is 1-based.
    /// </summary>
    public readonly int Row0;

    public CellAddress(string text, int column0, int row0)
    {
      Text = text;
      Column0 = column0;
      Row0 = row0;
    }
  }

  /// <summary>
  /// Access to named cells
  /// </summary>
  public interface INamedCells
  {
    string this[string name] { get; }
    ICell GetCell(string named);
    string[] Names { get; }
  }

  /// <summary>
  /// This is the interface to a Excel- or OpenOffice Calc-document.
  /// To instantiate it use <code>SpreadSheetReaderFactory.factory(<filename)</code>
  /// 
  /// The spreadsheet is implemented using XML-DOM. However, the interface to access is implemented sequencial.
  /// This would allow to refactor the implementation to use XML-SAX which would be much faster.
  /// </summary>
  public interface ISpreadSheetReader : IReference
  {
    string Filename { get; }
    List<IWorksheet> Worksheets { get; }
    INamedCells NamedCells { get; }
    bool IsExcel { get; }
  }

  /// <summary>
  /// This is the interface to a worksheet of a Excel- or OpenOffice Calc-document.
  /// In the GUI, a worksheet is a tab.
  /// The name of the worksheet is the name of the tab
  /// </summary> 
  public interface IWorksheet : IReference
  {
    IEnumerable<IRow> Rows { get; }
  }

  /// <summary>
  /// This is the interface to one row of a Excel- or OpenOffice Calc-document.
  /// </summary>
  public interface IRow : IReference
  {
    /// <summary>
    /// The last cell which is not empty.
    /// </summary>
    int Columns { get; }

    /// <summary>
    /// The row number 0-based. Note that the numbering in Excel- or Calc is 1-based.
    /// </summary>
    int RowNumber0 { get; }

    /// <summary>
    /// Accessor for one cell.
    /// It is allowed to access beyond <code>Columns.</code> The returned cell will contain an empty string.
    /// </summary>
    ICell this[int column] { get; }
  }

  /// <summary>
  /// This is the interface to a cell of a Excel- or OpenOffice Calc-document.
  /// </summary>
  public interface ICell : IReference
  {
    /// <summary>
    /// The value of the cell.
    /// If the cell contains a formula, the value is the result of the formula.
    /// </summary>
    string String { get; }

    /// <summary>
    /// Try to parse the <code>String</code> in a variable of type <code>T</code>.
    /// If <code>String</code> may not be parsed. a <code>SpreadSheetException</code> will be thrown.
    /// </summary>
    T Parse<T>();

    /// <summary>
    /// Same as <code>Parse<T>()</code>, but requries less writing.
    /// </summary>
    void Parse<T>(out T value);

    /// <summary>
    /// Same as <code>Parse<T>()</code>, but the interface may be used for reflection purposes.
    /// </summary>
    object Parse(Type type);
  }
  #endregion

  #region TableException
  /// <summary>
  /// If an exception is thrown, the exception should provide all needed information to write a meaningful error message to the user.
  /// The user which is responsible for the contents of the Excel-Sheet should be informed, which data in the sheet is referred by the message.
  /// </summary>
  public class SpreadSheetException : Exception
  {
    #region public
    /// <summary>
    /// For example:
    /// No column 'Spalte77'!
    /// </summary>
    public readonly string Msg = null;

    private readonly IReference Reference = null;

    /// <summary>
    /// For example:
    /// No column 'Spalte77'! Existing columns are (Spalte4|Spalte5|Spalte6|Spalte7|Spalte8). '../../zuluspreadsheet_test.ods', worksheet 'SheetA', table 'TableA', row 3
    /// </summary>
    public override string Message { get { return $"{Msg} Reference: {Reference.Reference}"; } }
    #endregion

    public SpreadSheetException(string msg, IReference reference = null)
      : base(msg)
    {
      Msg = msg;
      Reference = reference;
    }
  }
  #endregion // SpreadSheetException

  #region Reader: Excel and OpenOffice

  /// <summary>
  /// Base Class for ReaderExcel and ReaderOpenOffice.
  /// Logic for both:
  ///  - Exctracting the XML-File
  ///  - Detection of the Tables in the Worksheet.
  /// </summary>
  public abstract class SpreadSheetReaderFactory : IReference
  {
    #region public
    public List<IWorksheet> Worksheets { get; private set; }

    public string Filename { get; private set; }
    public string Name { get { return Path.GetFileName(Filename); } }
    public string Description { get { return $"file '{Name}'"; } }
    public string Reference { get { return Description; } }
    public INamedCells NamedCells { get; protected set; }
    public bool IsExcel { get; private set; }
    #endregion

    #region Constants
    public const string EXTENSION_ODS = "ods";
    public const string EXTENSION_XLSX = "xlsx";
    public const string EXTENSION_XLSM = "xlsm";
    public static readonly string[] EXTENSIONS_OPENOFFICE = new string[] { EXTENSION_ODS };
    public static readonly string[] EXTENSIONS_EXCEL = new string[] { EXTENSION_XLSX, EXTENSION_XLSM };
    private readonly Dictionary<int, string> dictCellsNone = new Dictionary<int, string>();
    #endregion

    #region factory
    /// <summary>
    /// Load the document from the filesystem into memory.
    /// </summary>
    /// <param name="filename">Der Filename</param>
    public static ISpreadSheetReader factory(string filename)
    {
      if (filename.EndsWith(EXTENSION_ODS))
      {
        return new ReaderOpenOffice(filename);
      }
      return new ReaderExcel(filename);
    }
    #endregion

    #region helper methods
    /// <summary>
    /// Example: ColumnNumberOffset0=0 -> First Row -> 'A'
    /// Example: ColumnNumberOffset0=1 -> Second Row -> 'B'
    /// Example: ColumnNumberOffset0=25 -> 'Z'
    /// Example: ColumnNumberOffset0=26 -> 'AA'
    /// Example: ColumnNumberOffset0=27 -> 'AB'
    /// </summary>
    public static string IntToAZ(int x)
    {
      string az = "";
      while (true)
      {
        int mod = x % 26;
        char mod_ = (char)(((int)'A') + mod);
        az = mod_ + az;
        x /= 26;
        if (x == 0)
        {
          break;
        }
        x--;
      }
      return az;
    }

    /// <summary>
    /// column  result
    /// B       1
    /// AB      27
    /// </summary>
    public static int AZtoInt(string sColumn)
    {
      int column0 = -1;
      for (int i = 0; i < sColumn.Length; i++)
      {
        column0++;
        column0 *= 26;
        column0 += sColumn[i] - 'A';
      }
      return column0;
    }

    private static Regex RegexCellAddress = new Regex(@"(?<column>[A-Z]+)(?<row>[0-9]+)");

    /// <summary>
    /// Examples
    ///   Text    Column0 Row0
    ///   A0      0       0
    ///   A1      0       1
    ///   B1      1       1
    ///   AB22    27      21
    /// </summary>
    public static CellAddress AZtoAddress(string text)
    {
      Match match = RegexCellAddress.Match(text);
      if (match == null)
      {
        throw new Exception($"'{text} is not a valid address for a cell!");
      }
      int row0 = int.Parse(match.Groups["row"].Value) - 1;
      string sColumn = match.Groups["column"].Value;
      int column0 = AZtoInt(sColumn); ;
      return new CellAddress(text, column0, row0);
    }

    private SpreadSheetReaderFactory(string filename, bool isExcel)
    {
      Filename = filename;
      IsExcel = isExcel;

      XmlNameTable xnt = new NameTable();
      nsMgr = new XmlNamespaceManager(xnt);
      populateNamespace();

      zipToOpen = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
      zipIn = new ZipArchive(zipToOpen, ZipArchiveMode.Read);
    }

    protected XmlDocument loadXml(string zipEntry)
    {
      XmlDocument xmlDoc = new XmlDocument();
      xmlDoc.PreserveWhitespace = true;
      ZipArchiveEntry entry = zipIn.GetEntry(zipEntry);
      XmlTextReader xtr = new XmlTextReader(entry.Open());
      xmlDoc.Load(xtr);
      return xmlDoc;
    }
    #endregion

    #region NamedCells
    private class NamedCellsClass : INamedCells
    {
      public string this[string name] { get { return d[name].String; } }
      private Dictionary<string, ICell> d = new Dictionary<string, ICell>();

      public string[] Names { get { return d.Keys.OrderBy(k => k).ToArray(); } }

      public void Add(string name, ICell cell)
      {
        d.Add(name, cell);
      }

      public ICell GetCell(string name)
      {
        return d[name];
      }
    }

    protected ICell getCell(Regex regexAddress, string address)
    {
      Match match = regexAddress.Match(address);
      Trace.Assert(match != null);
      string worksheet = match.Groups["worksheet"].Value;
      string column = match.Groups["column"].Value;
      string row = match.Groups["row"].Value;

      int column0 = AZtoInt(column);
      int row0 = int.Parse(row) - 1;

      return getCell(worksheet, column0, row0);
    }

    private ICell getCell(string worksheet, int column0, int row0)
    {
      IWorksheet worksheet_ = Worksheets.Find(ws => ws.Name == worksheet);
      Trace.Assert(worksheet_ != null);

      // Find row
      IRow row_ = worksheet_.Rows.Skip(row0).First();

      // Find column
      ICell cell = row_[column0];

      return cell;
    }
    #endregion

    #region ReaderExcel
    /// <summary>
    /// Reads a Excel-Document.
    /// A alternative implementation could have used: Open XML SDK.
    /// To understand this code, unzip a Excel-Document a look at the Xml-Documents.
    /// </summary>
    private class ReaderExcel : SpreadSheetReaderFactory, ISpreadSheetReader
    {
      private XmlNodeList sharedStrings;

      public ReaderExcel(string filename)
        : base(filename, isExcel: true)
      {
        XmlDocument xmlWorkbook = loadXml("xl/workbook.xml");
        Worksheets = loadWorksheets(xmlWorkbook);
        NamedCells = LoadNamedCells(xmlWorkbook);
      }

      protected override void populateNamespace()
      {
        nsMgr.AddNamespace("tns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        nsMgr.AddNamespace("tnsRels", "http://schemas.openxmlformats.org/package/2006/relationships");
      }

      private List<IWorksheet> loadWorksheets(XmlDocument xmlWorkbook)
      {
        List<IWorksheet> worksheets = new List<IWorksheet>();
        XmlDocument xmlSharedStrings = loadXml("xl/sharedStrings.xml");
        sharedStrings = xmlSharedStrings.SelectNodes("/tns:sst/tns:si/tns:t/text()", nsMgr);

        XmlDocument xmlWorkbookRels = loadXml("xl/_rels/workbook.xml.rels");
        XmlNodeList selectedNodes = xmlWorkbook.SelectNodes("/tns:workbook/tns:sheets/tns:sheet", nsMgr);
        foreach (XmlNode selectedNode in selectedNodes)
        {
          string worksheetName = selectedNode.Attributes["name"].Value;
          string worksheetId = selectedNode.Attributes["r:id"].Value;

          string xPath = @"/tns:Relationships/tns:Relationship[@Id=""" + worksheetId + @"""]";
          XmlNode targetNode = xmlWorkbookRels.SelectSingleNode(@"/tnsRels:Relationships/tnsRels:Relationship[@Id=""" + worksheetId + @"""]", nsMgr);
          string targetPath = targetNode.Attributes["Target"].Value;

          XmlDocument xmlWorksheet = loadXml("xl/" + targetPath);
          worksheets.Add(new Worksheet(this, xmlWorksheet, worksheetName));
        }
        return worksheets;
      }

      protected override IEnumerable<Row> rows(Worksheet worksheet, XmlNode nodeWorksheet)
      {
        int rowNumber0 = 0;
        foreach (XmlNode nodeRow in nodeWorksheet.SelectNodes("/tns:worksheet/tns:sheetData/tns:row", nsMgr))
        {
          int rowNumber0_ = int.Parse(nodeRow.Attributes["r"].Value) - 1;
          while (rowNumber0 < rowNumber0_)
          {
            yield return new Row(worksheet, dictCellsNone, rowNumber0);
            rowNumber0++;
          }

          int rowNumber = int.Parse(nodeRow.Attributes["r"].Value);
          Debug.Assert(rowNumber == rowNumber0 + 1);

          Dictionary<int, string> dictCells = readRow(nodeRow, rowNumber0);
          yield return new Row(worksheet, dictCells, rowNumber0); rowNumber0++;
        }
      }

      /// <summary>
      /// Loops over all cells of a row.
      /// If the cell has the 'number-columns-repeated'-attribute set, this method repeads the cell accordingly.
      /// The enumeration never ends! All cells to the right have the value "".
      /// </summary>
      private Dictionary<int, string> readRow(XmlNode nodeRow, int rowNumber0)
      {
        Trace.Assert(nodeRow.Name == "row");

        Dictionary<int, string> dictCells = new Dictionary<int, string>();

        foreach (XmlNode nodeCell in nodeRow.SelectNodes("./tns:c", nsMgr))
        {
          string address = nodeCell.Attributes["r"].Value;
          CellAddress cellAddress = SpreadSheetReaderFactory.AZtoAddress(address);
          Debug.Assert(cellAddress.Row0 == rowNumber0);
          int columnOffset0Actual = cellAddress.Column0;

          XmlNode nodeV = nodeCell.SelectSingleNode("./tns:v/text()", nsMgr);
          if (nodeV == null)
          {
            continue;
          }
          XmlNode nodeType = nodeCell.Attributes["t"];
          if (nodeType == null)
          {
            dictCells.Add(columnOffset0Actual, nodeV.Value);
            continue;
          }
          if (nodeType.Value == "s")
          {
            // "s": A shared String
            int sharedStringsIndex = int.Parse(nodeV.Value);
            dictCells.Add(columnOffset0Actual, sharedStrings[sharedStringsIndex].Value);
            continue;
          }
          throw new Exception("Programming error: Unknown type '" + nodeType.Value + "'!");
        }

        return dictCells;
      }

      #region NamedCells
      /// <summary>
      ///   <definedNames>
      ///       <definedName name = "NamedCellA" >'Named Cells'!$B$2</definedName>
      ///       <definedName name = "NamedCellB" >'Named Cells'!$B$4</definedName>
      ///   </definedNames>
      /// </summary>
      private INamedCells LoadNamedCells(XmlDocument xmlWorkbook)
      {
        NamedCellsClass namedCells = new NamedCellsClass();
        XmlNodeList xmlDefinedNames = xmlWorkbook.SelectNodes("/tns:workbook/tns:definedNames/tns:definedName", nsMgr);
        foreach (XmlNode xmlDefinedName in xmlDefinedNames)
        {
          string Name = xmlDefinedName.Attributes["name"].Value;
          string Address = xmlDefinedName.InnerText;
          ICell cell = getCell(regexAddress, Address);
          namedCells.Add(Name, cell);
        }
        return namedCells;
      }

      /// <summary>
      /// $NamedCells.$C$12  (OpenOffice)
      /// $'Named Cells'.$C$12  (OpenOffice)
      /// 'Named Cells'!$C$12  (Excel)
      /// </summary>
      private static Regex regexAddress = new Regex(@"^('?)(?<worksheet>.*?)('?\!\$)(?<column>.*?)\$(?<row>.*)$");
      #endregion
    }
    #endregion

    #region ReaderOpenOffice
    /// <summary>
    /// Reads a OpenOffice-Document.
    /// To understand this code, unzip a OpenOffice-Document a look at the Xml-Documents.
    /// </summary>
    private class ReaderOpenOffice : SpreadSheetReaderFactory, ISpreadSheetReader
    {
      public ReaderOpenOffice(string filename) : base(filename, isExcel: false)
      {
        XmlDocument xmlDocContent = loadXml("content.xml");
        Worksheets = loadWorksheets(xmlDocContent);
        NamedCells = LoadNamedCells(xmlDocContent);
      }

      protected override void populateNamespace()
      {
        nsMgr.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
        nsMgr.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
      }

      private List<IWorksheet> loadWorksheets(XmlDocument xmlDocContent)
      {
        List<IWorksheet> worksheets = new List<IWorksheet>();
        // Loop über alle Worksheets
        XmlNodeList nodesWorksheets = xmlDocContent.SelectNodes("/office:document-content/office:body/office:spreadsheet/table:table", nsMgr);
        foreach (XmlNode nodeWorksheet in nodesWorksheets)
        {
          string worksheetName = nodeWorksheet.Attributes["table:name"].Value;
          worksheets.Add(new Worksheet(this, nodeWorksheet, worksheetName));
        }
        return worksheets;
      }

      protected override IEnumerable<Row> rows(Worksheet worksheet, XmlNode nodeWorksheet)
      {
        int rowNumber0 = 0;
        foreach (XmlNode nodeRow in nodeWorksheet.SelectNodes("./table:table-row", nsMgr))
        {
          Dictionary<int, string> dictCells = readRow(nodeRow);
          int numberRowsRepeated = GetXmlAttribute(nodeRow, "table:number-rows-repeated", 1);
          if ((numberRowsRepeated > 1000) && (dictCells.Count == 0))
          {
            // So many empty rows might be the end of the table!
            yield break;
          }
          while (numberRowsRepeated-- > 0)
          {
            yield return new Row(worksheet, dictCells, rowNumber0);
            rowNumber0++;
          }
        }
      }

      /// <summary>
      /// Loops over all cells of a row.
      /// If the cell has the 'number-columns-repeated'-attribute set, this method repeads the cell accordingly.
      /// The enumeration never ends! All cells to the right have the value "".
      /// </summary>
      private Dictionary<int, string> readRow(XmlNode nodeRow)
      {
        Dictionary<int, string> dictCells = new Dictionary<int, string>();

        int cellnumber = 0;
        Trace.Assert(nodeRow.Name == "table:table-row");

        foreach (XmlNode nodeCell in nodeRow.SelectNodes("./table:table-cell", nsMgr))
        {
          int numberColumnsRepeated = GetXmlAttribute(nodeCell, "table:number-columns-repeated", 1);
          while (numberColumnsRepeated-- > 0)
          {
            if (nodeCell.InnerText.Length > 0)
            {
              dictCells.Add(cellnumber, nodeCell.InnerText);
            }
            cellnumber++;
          }
        }

        return dictCells;
      }

      /// <summary>
      /// Retrieves an attribute. For example "table:number-rows-repeated" in the following XML-None
      /// <table:table-row table:style-name="ro1" table:number-rows-repeated="3">
      /// </summary>
      private int GetXmlAttribute(XmlNode node, string attribute, int ifNotDefined)
      {
        XmlAttribute xmlAttribute = node.Attributes[attribute];
        if (xmlAttribute != null)
        {
          return int.Parse(xmlAttribute.Value);
        }
        return ifNotDefined;
      }

      #region NamedCells
      /// <summary>
      ///   <table:named-expressions>
      ///      <table:named-range table:name="QRcodeIban" table:base-cell-address="$Tabelle1.$C$4" table:cell-range-address="$Tabelle1.$C$12"/>
      ///      <table:named-range table:name="QRcodeContactName" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$13"/>
      ///      <table:named-range table:name="QRcodeContactZip" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$14"/>
      ///      <table:named-range table:name="QRcodeContactCity" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$15"/>
      ///      <table:named-range table:name="QRcodeContactCountry" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$16"/>
      ///      <table:named-range table:name="QRcodeContactStreet" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$17"/>
      ///      <table:named-range table:name="QRcodeContactNumber" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$18"/>
      ///      <table:named-range table:name="QRcodeReference" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$19"/>
      ///      <table:named-range table:name="QRcodeAmount" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$20"/>
      ///      <table:named-range table:name="QRcodeCurrency" table:base-cell-address="$Tabelle1.$C$5" table:cell-range-address="$Tabelle1.$C$21"/>
      ///    </table:named-expressions>
      ///  </summary>
      private INamedCells LoadNamedCells(XmlDocument xmlDocContent)
      {
        NamedCellsClass namedCells = new NamedCellsClass();
        XmlNodeList xmlNamedExpressons = xmlDocContent.SelectNodes("/office:document-content/office:body/office:spreadsheet/table:named-expressions/table:named-range", nsMgr);
        foreach (XmlNode xmlNamedExpression in xmlNamedExpressons)
        {
          string Name = xmlNamedExpression.Attributes["table:name"].Value;
          string Address = xmlNamedExpression.Attributes["table:cell-range-address"].Value;
          ICell cell = getCell(regexAddress, Address);
          namedCells.Add(Name, cell);
        }
        return namedCells;
      }

      /// <summary>
      /// $NamedCells.$C$12  (OpenOffice)
      /// $'Named Cells'.$C$12  (OpenOffice)
      /// 'Named Cells'!$C$12  (Excel)
      /// </summary>
      private static Regex regexAddress = new Regex(@"^(\$'?)(?<worksheet>.*?)('?\.\$)(?<column>.*?)(\$)(?<row>.*)$");
      #endregion
    }
    #endregion

    #region inner classes
    protected class Worksheet : IWorksheet
    {
      #region public
      public string Name { get; private set; }
      public string Description { get { return $"worksheet '{Name}'"; } }
      public string Reference { get { return $"{Description} in {readerFactory.Reference}"; } }
      public IEnumerable<IRow> Rows { get { return readerFactory.rows(this, nodeWorksheet); } }
      public bool IsExcel { get { return readerFactory.IsExcel; } }
      #endregion

      public Worksheet(SpreadSheetReaderFactory readerBase_, XmlNode nodeWorksheet_, string worksheetName)
      {
        nodeWorksheet = nodeWorksheet_;
        Name = worksheetName;
        readerFactory = readerBase_;
      }

      #region private
      private SpreadSheetReaderFactory readerFactory;
      private XmlNode nodeWorksheet;
      #endregion
    }

    protected class Row : IRow
    {
      #region public
      public string Name { get; private set; }
      public string Description { get { return $"row {RowNumber0 + 1}"; } }
      public string Reference { get { return $"{Description} in {worksheet.Reference}"; } }
      public int Columns { get; private set; }
      public int RowNumber0 { get; private set; }

      public Row(Worksheet worksheet_, Dictionary<int, string> dictCells_, int rowNumber0)
      {
        worksheet = worksheet_;
        dictCells = dictCells_;
        RowNumber0 = rowNumber0;
        Columns = 0;
        if (dictCells_.Count > 0)
        {
          Columns = dictCells_.Keys.Max() + 1;
        }
      }

      /// <summary>
      /// Get Value of a Table-Cell
      /// </summary>
      /// <param name="column"></param>
      /// <returns></returns>
      public ICell this[int column]
      {
        get
        {
          return getCell(column);
        }
      }
      #endregion

      #region private
      protected Cell getCell(int column)
      {
        string name = $"{IntToAZ(column)}{RowNumber0 + 1}";
        string value;
        if (!dictCells.TryGetValue(column, out value))
        {
          value = "";
        }
        return new Cell(value, name, worksheet);
      }

      private Dictionary<int, string> dictCells;
      private Worksheet worksheet;
      #endregion
    }

    public static ICell[] NewCellArray(int count)
    {
      return new SpreadSheetReaderFactory.Cell[count];
    }

    protected class Cell : ICell
    {
      #region public
      public string Name { get; private set; }
      public string Description { get { return $"cell '{Name}'"; } }
      public string Reference { get { return $"{Description} in {worksheet.Reference}"; } }
      public string String { get; private set; }
      #endregion

      public Cell(string value, string name, Worksheet worksheet_)
      {
        String = value;
        Name = name;
        worksheet = worksheet_;
      }

      #region Public Methods
      public T Parse<T>()
      {
        return (T)Parse(typeof(T));
      }

      /// <summary>
      /// Same as <code>Parse<T>()</code>, but requries less writing.
      /// </summary>
      public void Parse<T>(out T value)
      {
        value = (T)Parse(typeof(T));
      }

      /// <summary>
      /// The tryParse-Methods are cached under the assumption, that "type.GetMethod" is slower than a dictionary.
      /// </summary>
      private static Dictionary<Type, MethodInfo> cacheTryParse = new Dictionary<Type, MethodInfo>();

      public object Parse(Type type)
      {
        string s = String.Trim();
        if ("-" == s)
        {
          return Activator.CreateInstance(type); // correspondes to default(T)
        }
        try
        {
          if (type.IsEnum)
          {
            object enumValue = ParseEnum(type, s);
            return enumValue;
          }

          if ((type == typeof(DateTime)) && (worksheet.IsExcel))
          {
            // OpenOffice stores dates as strings (1992-02-12)
            // Excel stores dates as doubles (33646.043090277803).
            // This is the Excel-Version:
            double d = double.Parse(s);
            DateTime dateTime = DateTime.FromOADate(d);
            return dateTime;
          }

          /*
           * A UserType may define 'TryParse'.
           * 
           *  public struct SpreadSheetColumnReference
           *  {
           *    public static bool TryParse(string s, out SpreadSheetColumnReference result)
           *  }
           *
           * Many C# Classes do the same. For example string, int, DateTime
           */
          if (!cacheTryParse.TryGetValue(type, out MethodInfo tryParse))
          {
            tryParse = type.GetMethod("TryParse",
                 BindingFlags.Public | BindingFlags.Static,
                 null,
                 new[] { typeof(string), type.MakeByRefType() },
                 null);
            cacheTryParse[type] = tryParse;
          }

          if (tryParse != null)
          {
            object[] args = { s, null };
            try
            {
              if ((bool)tryParse.Invoke(null, args))
              {
                return args[1];
              }
            } catch (TargetInvocationException ex)
            {
                throw new SpreadSheetException($"'{s}' is not a valid {type.Name} ({ex.InnerException.Message})!", this);
            }
          }

          // CultureInfo.InvariantCulture: Avoid problem with the conversion of doubles or dates because of the culture settings
          return Convert.ChangeType(s, type, CultureInfo.InvariantCulture);
        }
        catch (System.FormatException)
        {
          string name = type.Name;
          if (Regex.IsMatch(name, @"^Int\d+"))
          {
            name = "integer";
          }
          if ((name == "Double") || (name == "Single"))
          {
            name = "float";
          }

          throw new SpreadSheetException($"'{s}' is not a valid {name}!", this);
        }
      }

      public T ParseEnum<T>(string value)
      {
        return (T)ParseEnum<T>(value);
      }
      #endregion

      #region Private Methods
      private object ParseEnum(Type type, string value)
      {
        System.Array possibleEnumValues = Enum.GetValues(type);
        object enumValue = extractEnumValue(possibleEnumValues, value);
        if (enumValue != null)
        {
          return enumValue;
        }
        object[] possibleEnumValues_ = possibleEnumValues.Cast<object>().ToArray();
        string values = string.Join(", ", possibleEnumValues_);
        // throw new TableException("'" + value + "' is not valid. Use one of " + values + "!", tableRow: this, tableColumn: excelTable.Columns[fieldInfo.Name]);
        throw new SpreadSheetException("'" + value + "' is not valid. Use one of " + values + "!", this);
      }

      private object extractEnumValue(System.Array possibleEnumV, string cellValue)
      {
        foreach (object v in possibleEnumV)
        {
          if (cellValue.Equals(v.ToString()))
          {
            return v;
          }
        }
        return null;
      }
      #endregion

      #region private members
      private Worksheet worksheet;
      #endregion
      #endregion
    }

    #region protected
    protected abstract IEnumerable<Row> rows(Worksheet worksheet, XmlNode nodeWorksheet);
    protected abstract void populateNamespace();
    #endregion

    #region private
    // TODO(HM): Don't forget to dispose
    private readonly ZipArchive zipIn;
    // TODO(HM): Don't forget to dispose
    private FileStream zipToOpen;
    private XmlNamespaceManager nsMgr;
    #endregion
  }

  #endregion
}
