using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Fuse8Task.Models
{
    public class EmailDateViewModel
    {
        [Required]
        public string emailForSendReport { get; set; }
        [Required]
        public string datepicker { get; set; }
        [Required]
        public string datepicker1 { get; set; }
    }
}