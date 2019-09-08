using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _800Best.ExcelHelpModel
{
   public class MyExcel
    {
       
        public MyExcel()
            {
                this.StartRow = 0;
                this.CurrentRow = 0;
                this.LastRow = 0;
                this.SouceStartRow = 0;
            }

            public int StartRow { get; set; }

            public int SouceStartRow { get; set; }

            public int CurrentRow { get; set; }

            public int LastRow { get; set; }

            public int LastCellNum { get; set; }

            public int LastRowOffset { get; set; }

            public string SaveFile { get; set; }

            public List<string> AddFileNames { get; set; }
        }
    }



