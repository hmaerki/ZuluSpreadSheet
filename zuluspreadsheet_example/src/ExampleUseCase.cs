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
      // Demonstrate table-access using 'class TableMeasurement' and linq-join
      // When was which equipment used by Karl?
      var equipment = tables.TypedRows<TableEquipment>();
      var measurement = tables.TypedRows<TableMeasurement>();
      var list = from e in equipment
                 join m in measurement on e.ID equals m.Equipment
                 where (m.Operator == "Karl")
                 select new { Measurement = m, Equipment = e };
      foreach (var row in list)
      {
        Debug.Print($"Date '{row.Measurement.Date:yyyy-MM-dd}': {row.Equipment.Model} (See {row.Measurement.TableRow.Reference})");
      }

      // Demonstrate table-access using 'class TableMeasurement' for iteration
      foreach (TableMeasurement row in tables.TypedRows<TableMeasurement>())
      {
        Debug.Print(row.Date + " " + row.Operator + " " + row.Equipment);
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

