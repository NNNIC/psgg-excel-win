using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using psggExcelWinLib;
using System.IO;

namespace PSGGExcelWin
{
    class Program
    {
        static void Main(string[] args)
        {
            TEST02(args);

        }
        static void TEST01(string[] args)
        {
            var file =args[0];
            var work = new Work();
            if (work.Load(Path.GetFullPath(file)))
            {
                work.SetSheet("state-chart"); 
                var mr = work.MaxRow();
                var mc = work.MaxCol();
                for(var r = 1; r<=mr; r++)
                {
                    var s = string.Empty;
                    for(var c = 1; c<=mc; c++)
                    {
                        if (!string.IsNullOrEmpty(s)) s += "|";
                        var v = work.GetStr(r,c);
                        s += v!=null ? v : "-";
                    }
                    Console.WriteLine(r.ToString("000") + ":" + s);
                }
            }
            Console.WriteLine(work.latest_error);
            work.Dispose();
        }

        public class Cell {
            public int row;
            public int col;
            public string text;
        }
        static void TEST02(string[] args)
        {
            var cellist = new List<Cell>();

            var file =args[0];
            var work = new Work();
            if (work.Load(Path.GetFullPath(file)))
            {
                work.SetSheet("state-chart");

                var mr = work.MaxRow();
                var mc = work.MaxCol();
                for(var r = 1; r<=mr; r++)
                {
                    for(var c = 1; c<=mc; c++)
                    {
                        var v = work.GetStr(r,c);
                        if (v!=null)
                        {
                            var cell = new Cell();
                            cell.row = r;
                            cell.col = c;
                            cell.text = v;
                            cellist.Add(cell);
                        }
                    }
                }
                
                work.NewSheetForce("test");

                cellist.ForEach(c => {
                    work.SetStr(c.row,c.col,c.text);
                });
                work.WriteSheet();
                work.Save();
            }
            Console.WriteLine(work.latest_error);
            work.Dispose();
        }
        static void TEST03(string[] args)
        {

        }

    }
}

