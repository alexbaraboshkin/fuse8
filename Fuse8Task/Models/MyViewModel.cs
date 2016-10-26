using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Fuse8Task.Models
{
    public class MyViewModel
    {
        public int? OderID { get; set; }
        public DateTime? OrderDate { get; set; }
        public int? ProductId { get; set; }
        public short? Quantity { get; set; }
        public decimal? UnitPrice { get; set; }
    }
}