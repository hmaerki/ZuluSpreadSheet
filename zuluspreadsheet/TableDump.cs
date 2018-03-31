//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using Zulu.Table.SpreadSheet;
using Zulu.Table.Table;
using System.Collections.Generic;
using System.IO;

namespace Zulu.Table.TableDumpNamespace
{
  public class TableDump
  {
    public const string DELIMIETER = "|";
    public const string ESCAPED_DELIMETER = @"\|";

    public TableDump(string filename) :
      this(TableCollection.factory(filename))
    {
    }

    public TableDump(ITableCollection tableCollection_)
    {
      tableCollection = tableCollection_;
    }

    public string dump(string filename = null)
    {
      if (filename == null)
      {
        // ..\..\..\zuluspreadsheet_test.ods => ..\..\..\zuluspreadsheet_test_dump.txt
        filename = Path.Combine(Path.GetDirectoryName(tableCollection.FileName), Path.GetFileNameWithoutExtension(tableCollection.FileName) + "_dump.txt");
      }

      using (StreamWriter sw = new StreamWriter(filename))
      {
        dump(sw, filename);
      }

      return filename;
    }

    public void dump(TextWriter tw, string filename = null)
    {
      foreach (ITable table in tableCollection)
      {
        tw.WriteLine();
        tw.WriteLine($"Table: {table.Name}");
        tw.WriteLine($"  {string.Join(DELIMIETER, table.ColumnNames)}");
        tw.WriteLine();

        foreach (ITableRow row in table)
        {
          List<string> cells = new List<string>();
          foreach (ITableColumn col in table.Columns)
          {
            string reference = col.Reference;
            ICell cell = row[col];
            cells.Add(cell.String.Replace(DELIMIETER, ESCAPED_DELIMETER));
          }
          tw.WriteLine($"  {string.Join(DELIMIETER, cells)}");
        }
      }
    }

    private ITableCollection tableCollection;
  }

}
