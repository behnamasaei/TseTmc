using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDidebanToExcel
{
    public class Dideban
    {
        public string Stock { get; set; }
        public string Name { get; set; }
        public long Count { get; set; }
        public long Volume { get; set; }
        public long Cost { get; set; }
        public int Yesterday { get; set; }
        public int First { get; set; }

        // آخرین معامله
        public int LastTradePrice { get; set; }
        public decimal LastTradeChange { get; set; }
        public float LastTradePercent { get; set; }

        // قیمت پایانی
        public int AvgPrice { get; set; }
        public decimal AvgChange { get; set; }
        public float AvgPercent { get; set; }

        //
        public int LowestPrice { get; set; }
        public int HighestPrice { get; set; }

        //
        public decimal? EPS { get; set; }
        public decimal? PE { get; set; }

        // خرید
        public int BuyCount { get; set; }
        public long BuyVolume { get; set; }
        public int BuyPrice { get; set; }

        // فروش 
        public int SellCount { get; set; }
        public long SellVolume { get; set; }
        public int SellPrice { get; set; }

    }
}
