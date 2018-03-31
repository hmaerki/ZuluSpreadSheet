//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using Zulu.Table.SpreadSheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Zulu.Table.Table
{
  #region Interfaces
  /// <summary>
  /// A ITableCollection stores all Tables from a SpreadSheet.
  /// </summary>
  public interface ITableCollection : IEnumerable<ITable>, IReference
  {
    string FileName { get; }
    string[] TableNames { get; }
    ITable this[string tableName] { get; }
    TableCollection.TypedTable<T> TypedRows<T>(string tableName = null) where T : ITableRowTyped, new();
  }

  public interface ITable : IEnumerable<ITableRow>, IReference
  {
    string ReferenceWithStartingRow { get; }
    IEnumerable<T> TypedRows<T>() where T : ITableRowTyped, new();
    ITableRow this[int row] { get; }
    string[] ColumnNames { get; }
    ITableCollection TableCollection { get; }
    /// <summary>The columns by 'Name'</summary>
    Dictionary<string, ITableColumn> Column { get; }
    IEnumerable<ITableColumn> Columns { get; }
  }

  public interface ITableColumn : IReference
  {
    ITable Table { get; }
    string ColumnAZ { get; }
  }

  public interface ITableRowTyped
  {
    ITableRow TableRow { get; set; }
  }

  public interface ITableRow : IReference
  {
    ITable Table { get; }
    /// <summary>Access for a value in a cell</summary>
    ICell this[string columnName] { get; }
    ICell this[ITableColumn column] { get; }
  }
  #endregion

  #region Implementation
  public class TableCollection : ITableCollection, IReference
  {
    #region static factories
    public static ITableCollection factory(string filename)
    {
      ISpreadSheetReader reader = SpreadSheetReaderFactory.factory(filename);
      return new TableCollection(reader);
    }
    public static ITableCollection factory(ISpreadSheetReader reader)
    {
      return new TableCollection(reader);
    }
    #endregion

    #region public
    public string FileName { get { return reader.Filename; } }
    public string[] TableNames { get; private set; }
    public IEnumerator<ITable> GetEnumerator() { return tables.Values.OrderBy(t=>t.Name).GetEnumerator(); }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return this.GetEnumerator(); }

    public string Name { get { return reader.Name; } }
    public string Description { get { return reader.Description; } }
    public string Reference { get { return reader.Reference; } }
    #endregion

    #region Generic Helpers
    private TableCollection(ISpreadSheetReader reader_)
    {
      reader = reader_;

      foreach (IWorksheet worksheet in reader.Worksheets)
      {
        LoopRows(worksheet);
      }

      TableNames = tables.Keys.Select(t => t).OrderBy(n => n).ToArray();
    }

    /// <summary>
    /// Access a table by name.
    /// </summary>
    public ITable this[string tableName]
    {
      get
      {
        Table table;
        if (tables.TryGetValue(tableName, out table))
        {
          return table;
        }
        string msg = "Table '" + tableName + "' does not exist!";
        msg += " Existing tables are (" + string.Join("|", TableNames) + ").";
        throw new SpreadSheetException(msg, this);
      }
    }

    /// <summary>
    /// No tablename is given as parameter.
    /// So we expect to get the tablename as a static const of the class.
    /// Example:
    ///    class Row {
    ///      public static string TableName = "Articles";
    ///      public string Name;
    ///      public int Price;
    ///    }
    /// We now get the tablename using reflection.
    /// </summary>
    public TypedTable<T> TypedRows<T>(string tableName = null) where T : ITableRowTyped, new()
    {
      if (tableName == null)
      {
        tableName = ExtractStaticMember<T>(MEMBER_TABLE_NAME, "Articles");
      }
      return new TypedTable<T>(this, tableName);
    }

    /// <summary>
    /// Extract a static member from a type using reflections.
    /// For example:
    /// <example>class T { public const string FileName = @"..\..\testcases.ods"; }</example>
    /// Usage:
    /// <example>exctractStaticMember<T>("FileName", "file_xy.ods")</example> will return <example>@"..\..\testcases.ods";</example>
    /// </summary>
    public static string ExtractStaticMember<T>(string memberName, string example = "sample") where T : ITableRowTyped, new()
    {
      FieldInfo fi = typeof(T).GetField(memberName, BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
      if (fi == null)
      {
        throw new Exception("Programmierfehler: Die Klasse '" + typeof(T).FullName + "' muss dies definieren: public const " + memberName + " = \"" + example + "\";");
      }
      string value = fi.GetValue(null) as string;
      return value;
    }

    protected void LoopRows(IWorksheet worksheet)
    {
      Table table = null;
      foreach (IRow row in worksheet.Rows)
      {
        if (row[CELLINDEX_TABLE].String.Equals(CELL_TABLE))
        {
          if (table == null)
          {
            table = new Table(this, row, worksheet);
            if (tables.ContainsKey(table.Name))
            {
              throw new SpreadSheetException("Several tables with name '" + table.Name + "'!", table);
            }
            tables.Add(table.Name, table);
            continue;
          }
          else
          {
            TableRow tableRow = new TableRow(table, row);
            table.AddRow(tableRow);
            continue;
          }
        }

        string cell_TABLE = row[CELLINDEX_TABLE].String;
        if (CELL_SKIP.Equals(cell_TABLE))
        {
          continue;
        }

        table = null;
      }
    }
    #endregion

    #region inner classes
    public class TypedTable<T> : IEnumerable<T> where T : ITableRowTyped, new()
    {
      #region public
      public IEnumerator<T> GetEnumerator()
      {
        return table.TypedRows<T>().GetEnumerator();
      }

      System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
      {
        return this.GetEnumerator();
      }
      #endregion

      public TypedTable(string fileName, string tableName)
      {
        ITableCollection tableCollection = TableCollection.factory(fileName);
        table = tableCollection[tableName];
      }
      public TypedTable(ITableCollection tableCollection = null, string tableName = null)
      {
        if (tableCollection == null)
        {
          string fileName = TableCollection.ExtractStaticMember<T>(TableCollection.MEMBER_FILE_NAME, "file_xy.ods");
          tableCollection = TableCollection.factory(fileName);
        }
        if (tableName == null)
        {
          tableName = TableCollection.ExtractStaticMember<T>(TableCollection.MEMBER_TABLE_NAME, "Articles");
        }
        try
        {
          table = tableCollection[tableName];
        }
        catch (KeyNotFoundException)
        {
          throw new SpreadSheetException("Table '" + tableName + "' does not exist!", tableCollection);
        }
      }

      #region private
      protected ITableCollection tableCollection { get; private set; }
      protected ITable table { get; private set; }
      #endregion
    }

    /// <summary>
    /// A Table within a Spreadsheet
    /// </summary>
    internal class Table : ITable, IReference
    {
      #region public
      public string Name { get; private set; }
      public string Description { get { return $"table '{Name}'"; } }
      public string Reference { get { return $"{Description} in {worksheet.Reference}"; } }

      /// <summary>The columns of the Table</summary>
      public string[] ColumnNames { get { return Column.Keys.OrderBy(c => c).ToArray(); } }
      public IEnumerable<ITableColumn> Columns { get { return Column.Values; } }
      public string ReferenceWithStartingRow { get { return Reference + " starting on row " + (startingRow.RowNumber0 + 1); } }

      /// <summary>The document we belong to</summary>
      public ITableCollection TableCollection { get; private set; }
      /// <summary>Table Name</summary>
      /// <summary>The columns by 'Name'</summary>
      public Dictionary<string, ITableColumn> Column { get; private set; }
      /// <summary>The columns by 'ColumnNumberOffset0'</summary>
      public List<TableColumn> ColumnsByColumnNumberOffset0 = new List<TableColumn>(); // TODO: Better Name, conflict with 'Columns'
                                                                                       /// <summary>The last column to the right</summary>
                                                                                       /// <summary>The rows of the Table</summary>
      public IEnumerator<ITableRow> GetEnumerator() { return rows.GetEnumerator(); }
      System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return this.GetEnumerator(); }
      public ITableRow this[int row] { get { return rows[row]; } }

      public IEnumerable<T> TypedRows<T>() where T : ITableRowTyped, new()
      {
        foreach (TableRow tableRow in rows)
        {
          T row = tableRow.fill<T>();
          row.TableRow = tableRow;
          yield return row;
        }
      }
      #endregion

      #region Helper: ExcelEnumerable

      /// <summary>
      /// No filename is given as parameter.
      /// So we expect to get the filename as a static const of the class.
      /// Example:
      ///    class Row {
      ///      public static string FileName = "xy.ods";
      ///      public string Name;
      ///      public int Price;
      ///    }
      /// We now get the filename using reflection.
      /// 
      /// We may use this method as a NUnit-<code>TestCaseSource</code>.
      /// </summary>
      internal class ExcelEnumerable<T> : IEnumerable<T> where T : ITableRowTyped, new()
      {
        private TypedTable<T> table;

        // NUnit needs a constructor without arguments...
        public ExcelEnumerable()
          : this(null, null)
        {
        }
        public ExcelEnumerable(string fileName = null, string tableName = null)
        {
          table = new TypedTable<T>(fileName, tableName);
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
          return GetEnumerator();
        }
        public IEnumerator<T> GetEnumerator()
        {
          return table.GetEnumerator();
        }
      }
      #endregion

      #region Constructor and Methods
      public Table(TableCollection tableCollection_, IRow startingRow_, IWorksheet worksheet_)
      {
        TableCollection = tableCollection_;
        worksheet = worksheet_;
        startingRow = startingRow_;
        rows = new List<ITableRow>();
        Column = new Dictionary<string, ITableColumn>();

        Name = startingRow_[CELLINDEX_TABLE_NAME].String.Trim();
        if (Name == "")
        {
          throw new SpreadSheetException("Column 'B' is empty but must contain a table-name.", this);
        }

        for (int i = CELLINDEX_DATA_START; i < startingRow_.Columns; i++)
        {
          ICell cell = startingRow_[i];

          if (cell.String.Length == 0)
          {
            // We arrived at the right-side of the table
            break;
          }
          if (cell.String == CELL_SKIP)
          {
            // '-' means: Skip this column
            continue;
          }
          TableColumn tableColumn = new TableColumn(this, cell, columnIndex: Column.Count, columnNumberOffset0: i);
          if (Column.ContainsKey(tableColumn.Name))
          {
            throw new SpreadSheetException("Duplicate column '" + tableColumn.Name + "'", tableColumn);
          }
          Column.Add(tableColumn.Name, tableColumn);
          ColumnsByColumnNumberOffset0.Add(tableColumn);
        }
      }

      public void AddRow(ITableRow tableRow)
      {
        rows.Add(tableRow);
      }
      #endregion

      #region Private staff
      private IWorksheet worksheet;
      private IRow startingRow;
      private List<ITableRow> rows;
      #endregion
    }

    #region TableColumn
    internal class TableColumn : ITableColumn, IReference
    {
      #region public
      /// <summary>The table we belong to</summary>
      public ITable Table { get; private set; }

      /// <summary>The columns name</summary>
      public string Name { get { return columnCell.String; } }
      public string Description { get { return $"column '{Name}'"; } }
      public string Reference { get { return $"{Description} in {Table.Reference}"; } }

      public readonly int ColumnIndex;
      public readonly int ColumnNumberOffset0;

      // TODO(HM): Loeschen?
      /// <summary>
      /// returns the column name in the 'A-Z'-Syntax
      /// </summary>
      public string ColumnAZ { get { return Zulu.Table.SpreadSheet.SpreadSheetReaderFactory.intToAZ(ColumnNumberOffset0); } }

      /// <summary>
      /// returns "column A7"
      /// </summary>
      // public string Reference { get { return "column " + ColumnAZ; } }
      #endregion

      public TableColumn(Table table, ICell columnCell_, int columnIndex, int columnNumberOffset0)
      {
        Table = table;
        columnCell = columnCell_;
        ColumnIndex = columnIndex;
        ColumnNumberOffset0 = columnNumberOffset0;
      }

      #region private
      private ICell columnCell;
      #endregion
    }
    #endregion

    #region TableRow
    /// <summary>
    /// The row of a table
    /// </summary>
    internal class TableRow : ITableRow
    {
      #region Properties
      /// <summary>The table we belong to</summary>
      public ITable Table { get { return table; } }

      public string Reference { get { return row.Reference; } }
      public string Name { get { return row.Name; } }
      public string Description { get { return row.Description; } }

      /// <summary>Access for a value in a cell</summary>
      public ICell this[string columnName] { get { return getCell(columnName); } }
      public ICell this[ITableColumn column] { get { return this[column.Name]; } }
      #endregion

      #region Constructor and Methods
      public TableRow(Table table_, IRow row_)
      {
        table = table_;
        row = row_;

        cells = new SpreadSheetReaderFactory.Cell[table_.Column.Count];

        {
          ICell cell = row_[TableCollection.CELLINDEX_TABLE_NAME];
          if (cell.String != "")
          {
            throw new SpreadSheetException("Expected a empty cell, but found '" + cell.String + "'!", cell);
          }
        }

        for (int i = 0; i < table_.Column.Count; i++)
        {
          TableColumn column = table_.ColumnsByColumnNumberOffset0[i];
          cells[column.ColumnIndex] = row_[column.ColumnNumberOffset0];
        }
      }
      #endregion

      /// <summary>
      /// Fill the members of a class with the values of the row
      /// </summary>
      /// <typeparam name="T"></typeparam>
      /// <returns></returns>
      public T fill<T>() where T : new()
      {
        T poco = new T();
        foreach (FieldInfo fieldInfo in typeof(T).GetFields())
        {
          if (fieldInfo.IsStatic)
          {
            // We are not interrested in static fields like "TableName".
            continue;
          }
          ICell cell = getCell(fieldInfo.Name);
          fieldInfo.SetValue(poco, cell.Parse(fieldInfo.FieldType));
          if (fieldInfo.FieldType == typeof(string))
          {
            fieldInfo.SetValue(poco, cell.String);
            continue;
          }
          // throw new Exception(typeof(T).Name + ": Feld '" + fieldInfo.Name + "': Unbekannter Datentyp '" + fieldInfo.FieldType.Name + "'!");
        }
        return poco;
      }

      public TableColumn getColumn(string columnName)
      {
        ITableColumn tableColumn;
        if (table.Column.TryGetValue(columnName, out tableColumn))
        {
          return (TableColumn)tableColumn;
        }

        string msg = $"No column '{columnName}'!";
        msg += $" Existing columns are ({string.Join("|", table.ColumnNames)}).";
        throw new SpreadSheetException(msg, this);
      }

      #region Private Staff
      private Table table;
      private IRow row;
      private ICell[] cells;
      protected ICell getCell(string columnName)
      {
        TableColumn tableColumn = getColumn(columnName);
        ICell cell = cells[tableColumn.ColumnIndex];
        return cell;
      }
      #endregion
    }
    #endregion

    #region TableCell
    /// <summary>
    /// The row of a table
    /// </summary>
    //internal class TableCell : SpreadSheet.Reader.SpreadSheetReaderFactory.Cell, ICell, IReference
    //{
    //  #region Properties
    //  /// <summary>The table we belong to</summary>
    //  public string Reference { get { return "TODO"; } }
    //  public String String { get; private set; }
    //  #endregion

    //  #region private
    //  private TableRow tableRow;
    //  private TableColumn tableColumn;
    //  #endregion

    //  public TableCell(TableRow tableRow_, TableColumn tableColumn_, String value) : base(value)
    //  {
    //    tableRow = tableRow_;
    //    tableColumn = tableColumn_;
    //    String = value;
    //  }

    //  protected override void throwException(string msg)
    //  {
    //    // throw new TableException("'" + s + "' is not a valid " + typeName + "!", tableRow: tableRow, tableColumn: tableColumn);
    //    throw new TableException(msg, tableRow: tableRow, tableColumn: tableColumn);
    //  }

    //  #region Access to cell values for string,float,bool,...
    //  public double Float { get { return getT<float>("float", 0.0f, s => float.Parse(s)); } }
    //  public double Double { get { return getT<double>("float", 0.0, s => double.Parse(s)); } }
    //  public int Int { get { return getT<int>("integer", 0, s => int.Parse(s)); } }
    //  public bool Bool { get { return getT<bool>("boolean", false, s => bool.Parse(s)); } }
    //  #endregion

    //  public T getT<T>(string typeName, T valueIfEmpty, Func<string, T> parse)
    //  {
    //    string s = String.Trim();
    //    if ("-" == s)
    //    {
    //      return valueIfEmpty;
    //    }
    //    try
    //    {
    //      return parse(s);
    //    }
    //    catch (Exception)
    //    {
    //      throw new TableException("'" + s + "' is not a valid " + typeName + "!", tableRow: m_tableRow, tableColumn: m_tableColumn);
    //    }
    //  }
    //}
    #endregion

    #endregion

    #region private constants
    private const string CELL_TABLE = "TABLE";
    private const string CELL_SKIP = "-";
    private const int CELLINDEX_TABLE = 0;
    private const int CELLINDEX_TABLE_NAME = 1;
    private const int CELLINDEX_DATA_START = 2;

    private const string MEMBER_FILE_NAME = "FileName";
    private const string MEMBER_TABLE_NAME = "TableName";
    #endregion

    #region private
    private ISpreadSheetReader reader;
    private Dictionary<string, Table> tables = new Dictionary<string, Table>();
    #endregion
  }
  #endregion // Implementation
}
