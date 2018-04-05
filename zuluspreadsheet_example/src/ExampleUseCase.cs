//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using System;
using System.Diagnostics;
using System.Linq;
using Zulu.Table.SpreadSheet;
using Zulu.Table.Table;
using Zulu.Table.TableDumpNamespace;

/// <summary>
/// This code repeats some code in other examples.
/// This code is meant to be copied into 'zuluspreadsheet_usecase.html'.
/// </summary>
namespace Zulu.Table.ExampleUseCase
{
  class Program
  {
    [TableName("Equipment")]
    class TableEquipment : ITableRowTyped
    {
      public enum EnumType { Voltmeter, Multimeter, Oscilloscope };
      public ITableRow TableRow { get; set; }

      public readonly int ID = 0;
      public readonly EnumType Type = EnumType.Multimeter;
      public readonly string Model = null;
      public readonly string Serial = null;
    }

    [TableName("Measurement")]
    class TableMeasurement : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }

      public readonly DateTime Date = DateTime.MinValue;
      public readonly string Operator = null;
      public readonly int Equipment = 0;
    }

    static void demoLinq(ITableCollection tables)
    {
      // The two tables in the spreadsheet
      var equipment = tables.TypedRows<TableEquipment>();
      var measurement = tables.TypedRows<TableMeasurement>();

      // Demonstrate table-access using 'class TableMeasurement' for iteration
      foreach (TableMeasurement row in measurement)
      {
        Debug.Print(row.Date + " " + row.Operator + " " + row.Equipment);
      }

      // Demonstrate table-access using 'class TableMeasurement' and linq
      // Return the date when Karl was measuring first
      var row_ = measurement.Where(row => row.Operator == "Karl").OrderBy(row => row.Date).First();
      // dateTime: 2017-06-05
      DateTime dateTime = row_.Date;
      // It may be handy to inform the user where this Date is coming from.
      // reference: "row 12 in worksheet 'SheetQuery' in file 'zuluspreadsheet_test.xlsx'"
      string reference = row_.TableRow.Reference;

      // Demonstrate table-access using 'class TableMeasurement' and linq-join
      // When was which equipment used by Karl?
      var list = from e in equipment
                 join m in measurement on e.ID equals m.Equipment
                 where (m.Operator == "Karl")
                 select new { Measurement = m, Equipment = e };
      foreach (var row in list)
      {
        string msg = $"Date '{row.Measurement.Date:yyyy-MM-dd}': {row.Equipment.Model} (See {row.Measurement.TableRow.Reference})";
        // msg: "Date '2017-06-20': Keysight 34460A (See row 9 in worksheet 'SheetQuery' in file 'zuluspreadsheet_test.xlsx')"
        // msg: "Date '2017-06-05': Fluke 787 (See row 12 in worksheet 'SheetQuery' in file 'zuluspreadsheet_test.xlsx')"
        Debug.Print(msg);
      }
    }

    static void demoTable(ITableCollection tables)
    {
      {
        // Demonstrate table-access using indexing
        ITable table = tables["Measurement"];
        ICell cell = table[2]["Operator"];
        table[2]["Date"].Parse(out DateTime dateTime);
      }

      {
        // Demonstrate table-access using Iterators
        foreach (ITable table in tables)
        {
          foreach (ITableRow row in table)
          {
            foreach (ITableColumn col in table.Columns)
            {
              string reference = col.Reference;
              ICell cell = row[col];
            }
          }
        }
      }
    }

    static void demoDump(ITableCollection tables)
    {
      TableDump dumper = new TableDump(tables);
      // Will create "zuluspreadsheet_test_dump.txt"
      dumper.dump();
    }

    public static void Main(string[] args)
    {
      ITableCollection tables = TableCollection.factory("zuluspreadsheet_test.xlsx");
      demoDump(tables);
      demoTable(tables);
      demoLinq(tables);
    }
  }
}

