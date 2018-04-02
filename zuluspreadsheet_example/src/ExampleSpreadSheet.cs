//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using System.Diagnostics;
using Zulu.Table.CachedWorkSheetNamespace;
using Zulu.Table.SpreadSheet;

namespace Zulu.Table.Example
{
  class DemoSpreadSheet
  {
    private const string Filename = "../../../zuluspreadsheet_test.ods";

    enum GenderEnum { male, female };
    public void run()
    {
      exampleSpreadSheet();
      exampleCachedSpreadSheet();
    }

    /// <summary>
    /// This example loops over all worksheets and cell
    /// </summary>
    private void exampleSpreadSheet()
    {
      ISpreadSheetReader reader = SpreadSheetReaderFactory.factory(Filename);
      foreach (IWorksheet we in reader.Worksheets)
      {
        int counter = 0;
        Debug.Print("Worksheet: " + we.Name);
        Debug.Print("Worksheet Reference: " + we.Reference);
        foreach (IRow row in we.Rows)
        {
          for (int i = 0; i < row.Columns; i++)
          {
            ICell cell = row[i];
            if (cell.String != "")
            {
              if (counter++ < 6)
              {
                Debug.Print($"{cell.Reference}: {cell.String}");
              }
            }
          }
        }
      }
    }

    private void exampleCachedSpreadSheet()
    {
      CachedSpreadSheet spreadSheet = new CachedSpreadSheet(Filename);
      CachedWorksheet workSheet = spreadSheet["SheetA"];

      {
        // Access a Cell by AB-Notation
        ICell cell = workSheet["C5"];
        Debug.Print("Expecting 'Spalte4Zeile7':" + cell.String);
      }

      {
        // Access a Cell by row/column
        // Note: Index is 0based. [row][column].
        ICell cell = workSheet[4][2];
        Debug.Print("Expecting 'Spalte4Zeile7':" + cell.String);
      }

      {
        // Read an integer from a cell
        ICell cell = workSheet["C5"];
        try
        {
          int i = cell.Parse<int>();
        }
        catch (SpreadSheetException ex)
        {
          Debug.Print("Reading integer (failure): " + ex.Message);
        }
      }

      {
        // Read an integer from a cell
        ICell cell = workSheet["H4"];
        Debug.Print("Reading integer (ok): " + cell.Parse<int>());
      }

      {
        // Read an enumeration from a cell
        ICell cell = workSheet["E15"];
        try
        {
          GenderEnum gender = cell.Parse<GenderEnum>();
        }
        catch (SpreadSheetException ex)
        {
          Debug.Print("Reading enum (failure): " + ex.Message);
        }
      }

      {
        // Read an enumeration from a cell
        ICell cell = workSheet["E13"];
        Debug.Print("Reading enum (ok): " + cell.Parse<GenderEnum>().ToString());
      }
    }
  }
}
