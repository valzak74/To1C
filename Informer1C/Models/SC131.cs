namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC131
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
        [StringLength(2)]
        public string SP145 { get; set; }

        [Required]
        [StringLength(20)]
        public string SP135 { get; set; }

        public DateTime SP134 { get; set; }

        [Required]
        [StringLength(13)]
        public string SP141 { get; set; }

        [Required]
        [StringLength(4)]
        public string SP665 { get; set; }

        [Required]
        [StringLength(5)]
        public string SP140 { get; set; }

        [Required]
        [StringLength(2)]
        public string SP4101 { get; set; }

        [Required]
        [StringLength(2)]
        public string SP4102 { get; set; }

        [Required]
        [StringLength(4)]
        public string SP5350 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP4828 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP4829 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP6568 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP8375 { get; set; }

        [Column(TypeName = "text")]
        public string SP148 { get; set; }

        [Column(TypeName = "text")]
        public string SP149 { get; set; }

        [Column(TypeName = "text")]
        public string SP144 { get; set; }

        [Column(TypeName = "text")]
        public string SP5346 { get; set; }

        [Column(TypeName = "text")]
        public string SP593 { get; set; }

        [Column(TypeName = "text")]
        public string SP143 { get; set; }

        [Column(TypeName = "text")]
        public string SP664 { get; set; }

        [Column(TypeName = "text")]
        public string SP5347 { get; set; }

        [Column(TypeName = "text")]
        public string SP5348 { get; set; }

        [Column(TypeName = "text")]
        public string SP2905 { get; set; }

        [Column(TypeName = "text")]
        public string SP5349 { get; set; }

        [Column(TypeName = "text")]
        public string SP137 { get; set; }
    }
}
