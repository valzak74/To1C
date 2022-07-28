namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class DT2457
    {
        [Key]
        [Column(Order = 0)]
        [StringLength(9)]
        public string IDDOC { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public short LINENO_ { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2446 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2447 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2448 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2449 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2450 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2451 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2454 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2452 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2455 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2453 { get; set; }
    }
}
