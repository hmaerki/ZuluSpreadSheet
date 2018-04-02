namespace Zulu.Table.Example
{
  class Program
  {
    static void Main(string[] args)
    {
      DemoTable demoTable = new DemoTable();
      demoTable.run();

      DemoTableLinq demoTableLinq = new DemoTableLinq();
      demoTableLinq.run();

      DemoTableDump demoTableDump = new DemoTableDump();
      demoTableDump.run();

      DemoSpreadSheet demoSpreadSheet = new DemoSpreadSheet();
      demoSpreadSheet.run();
    }
  }
}
