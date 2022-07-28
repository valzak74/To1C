namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class RG4667
    {
        [Key]
        [Column(Order = 0)]
        public DateTime PERIOD { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(9)]
        public string SP4663 { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(9)]
        public string SP4664 { get; set; }

        [Key]
        [Column(Order = 3)]
        [StringLength(9)]
        public string SP4665 { get; set; }

        [Key]
        [Column(Order = 4, TypeName = "numeric")]
        public decimal SP13435 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP4666 { get; set; }
    }
}
