﻿using System;
using System.Diagnostics;
using System.Linq;
using Zulu.Table.SpreadSheet;
using Zulu.Table.Table;
using Zulu.Table.TableDumpNamespace;

namespace Zulu.Table.Example
{
  class DemoTable
  {
    public const string Filename = "./zuluspreadsheet_test.xlsx";

    /// <summary>
    /// This code expected a table as follows in OpenOffice-Calc or Excel:
    /// TABLE      TableC    name      age
    /// TABLE                Max       12
    /// TABLE                Moritz    13
    /// </summary>
    [TableName("TableC")]
    [FileName(Filename)]
    private class TableC : ITableRowTyped
    {
      public enum EnumGender { male, female };
      public ITableRow TableRow { get; set; }

      public readonly string Name = null;
      public readonly int Age = 0;
      public readonly EnumGender Gender = EnumGender.male;
      public readonly double Size = 0.0;
    }

    public void run()
    {
       ITableCollection tables = TableCollection.factory(Filename);

      {
        // Use indexing
        // Use Iterators
        {
          ITable table = tables["TableC"];
          ICell cell = table[2]["Age"];
          table[2]["Age"].Parse(out int age);
        }
      }

      {
        // Use Iterators
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

      {
        // Use a structure
        foreach (TableC row in tables.TypedRows<TableC>())
        {
          Debug.Print(row.Name + " " + row.Age + " " + row.Gender);
        }

        // This constructor uses the filename attribute: '[FileName(Filename)]'
        foreach (TableC row in new TableCollection.TypedTable<TableC>())
        {
          Debug.Print(row.Name + " " + row.Age + " " + row.Gender);
        }

        {
          int ottosAge = tables.TypedRows<TableC>().Where(row => row.Name == "Otto").First().Age;
          Debug.Print("ottosAge: " + ottosAge);
        }

        {
          var otto = from row in tables.TypedRows<TableC>() where row.Name == "Otto" select row.Age;
          Debug.Print("ottosAge: " + otto.First());
        }
      }
    }
  }

  class DemoTableLinq
  {
    [TableName("Equipment")]
    private class TableEquipment : ITableRowTyped
    {
      public enum EnumType { Voltmeter, Multimeter, Oscilloscope };
      public ITableRow TableRow { get; set; }

      public readonly int ID = 0;
      public readonly EnumType Type = EnumType.Multimeter;
      public readonly string Model = null;
      public readonly string Serial = null;
    }

    [TableName("Measurement")]
    private class TableMeasurement : ITableRowTyped
    {
      public ITableRow TableRow { get; set; }

      public readonly DateTime Date = DateTime.MinValue;
      public readonly string Operator = null;
      public readonly int Equipment = 0;
    }

    public void run()
    {
      ITableCollection tables = TableCollection.factory(DemoTable.Filename);

      {
        // linq-join:
        //  When was which equipment used by Karl?
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
      }
    }
  }

  class DemoTableDump
  {
    public void run()
    {
      ITableCollection tables = TableCollection.factory(DemoTable.Filename);

      TableDump dumper = new TableDump(tables);
      // dumper.dump(DemoTable.Filename.rename(".ods", "_dump.txt"));
      dumper.dump();
    }
  }
}
