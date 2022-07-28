namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC4014
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
        [StringLength(50)]
        public string DESCR { get; set; }

        public bool ISMARK { get; set; }

        public int VERSTAMP { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4011 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4012 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4133 { get; set; }

        [Required]
        [StringLength(40)]
        public string SP5011 { get; set; }

        [Required]
        [StringLength(40)]
        public string SP5010 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10879 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10880 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP12015 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP13106 { get; set; }

        [Column(TypeName = "text")]
        public string SP4073 { get; set; }
    }
}
