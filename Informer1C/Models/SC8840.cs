namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC8840
    {
        [Key]
        public int ROW_ID { get; set; }

        [Required]
        [StringLength(9)]
        public string ID { get; set; }

        [Required]
        [StringLength(5)]
        public string CODE { get; set; }

        [Required]
        [StringLength(25)]
        public string DESCR { get; set; }

        public bool ISMARK { get; set; }

        public int VERSTAMP { get; set; }

        [Required]
        [StringLength(9)]
        public string SP11193 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP11254 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP11467 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP11510 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP11668 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP12162 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP12932 { get; set; }
    }
}
