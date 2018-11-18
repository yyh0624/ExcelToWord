using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToWord
{

    public class DataModel
    {
       
        public string _fileName { get; set; }
        public string _allPriceCN { get; set; }
        public decimal? _allPriceNum { get; set; }

        public List<DetilsModel> _detils { get; set; } = new List<DetilsModel>();
    }
    public class DetilsModel
    {
        public string _wzmc { get; set; }
        public string _ggxh { get; set; }
        public string _jldw { get; set; }
        public int? sl { get; set; }
        public decimal? zj { get; set; }
        public decimal? dj { get; set; }
        public decimal? zjclf { get; set; }
        public decimal? wgcjf { get; set; }
        public decimal? rljdlf { get; set; }
        public decimal? zjrgf { get; set; }
        public decimal? fpssf { get; set; }
        public decimal? glfy { get; set; }
        public decimal? lr { get; set; }

        public decimal? sj { get; set; }
        public decimal? bjgjf { get; set; }
        public decimal? aztsf { get; set; }
        public decimal? jsfwf { get; set; }
        public decimal? yzf { get; set; }
        public string _pinpai { get; set; }
        public string _zxbz { get; set; }
        public string _chandi { get; set; }
    }
}
