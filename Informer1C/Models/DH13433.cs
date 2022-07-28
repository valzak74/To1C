namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class DH13433
    {
        [Key]
        [StringLength(9)]
        public string IDDOC { get; set; }

        [Required]
        [StringLength(9)]
        public string SP13424 { get; set; }

        [Required]
        [StringLength(150)]
        public string SP660 { get; set; }
    }
}
