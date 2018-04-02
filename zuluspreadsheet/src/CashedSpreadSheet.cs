//
// Project   ZuluSpreadSheet
// 
// (c) Copyright 2017 Solcept AG
// (c) Copyright 2002-2018, Hans Maerki, Maerki Informatik
// Distributed under the Boost Software License, Version 1.0. http://www.boost.org/LICENSE_1_0.txt)
//
using System.Collections.Generic;
using System.Linq;
using Zulu.Table.SpreadSheet;

namespace Zulu.Table.CachedWorkSheetNamespace
{
  #region CachedSpreadSheed
  public class CachedSpreadSheet
  {
    #region public
    public ISpreadSheetReader SpreadSheetReader { get; private set; }
    public INamedCells NamedCells { get { return SpreadSheetReader.NamedCells; } }

    public CachedWorksheet this[string worksheet]
    {
      get
      {
        try
        {
          return worksheets[worksheet];
        }
        catch (KeyNotFoundException)
        {
          string validWorksheets = string.Join(", ", SpreadSheetReader.Worksheets.Select(ws => ws.Name));
          throw new SpreadSheetException($"'{worksheet}' not found. Valid worksheets are {validWorksheets}!");
        }
      }
    }
    #endregion

    public CachedSpreadSheet(string filename)
      : this(SpreadSheetReaderFactory.factory(filename))
    {
    }

    public CachedSpreadSheet(ISpreadSheetReader reader)
    {
      SpreadSheetReader = reader;
      foreach (IWorksheet we in reader.Worksheets)
      {
        worksheets.Add(we.Name, new CachedWorksheet(we));
      }
    }

    #region private
    private Dictionary<string, CachedWorksheet> worksheets = new Dictionary<string, CachedWorksheet>();
    #endregion
  }

  public class CachedWorksheet
  {
    #region public
    public IWorksheet Worksheet { get; private set; }
    public int MaxRows { get; private set; }
    public int MaxColumns { get; private set; }

    /// <summary>
    /// Get the row for a given row number.
    /// </summary>
    public IRow this[int rowNumber0]
    {
      get
      {
        try
        {
          return dictRows[rowNumber0];
        }
        catch (KeyNotFoundException)
        {
          throw new SpreadSheetException($"Row {rowNumber0} is empty!", Worksheet);
        }
      }
    }

    /// <summary>
    /// Get the cell by a address of the form 'A5'
    /// </summary>
    public ICell this[string address]
    {
      get
      {
        CellAddress cellAddress = SpreadSheetReaderFactory.AZtoAddress(address);
        return this[cellAddress.Row0][cellAddress.Column0];
      }
    }
    #endregion

    public CachedWorksheet(IWorksheet worksheet)
    {
      Worksheet = worksheet;
      MaxColumns = 0;
      MaxRows = 0;
      foreach (IRow ri in worksheet.Rows)
      {
        MaxRows = ri.RowNumber0 + 1;
        if (ri.Columns >= MaxColumns)
        {
          MaxColumns = ri.Columns + 1;
        }
        dictRows[ri.RowNumber0] = ri;
      }
    }

    #region private
    private Dictionary<int, IRow> dictRows = new Dictionary<int, IRow>();
    #endregion
  }
  #endregion
}
