namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC84
    {
        [Key]
        public int ROW_ID { get; set; }

        [Required]
        [StringLength(9)]
        public string ID { get; set; }

        [Required]
        [StringLength(9)]
        public string PARENTID { get; set; }

        [Required]
        [StringLength(9)]
        public string CODE { get; set; }

        [Required]
        [StringLength(50)]
        public string DESCR { get; set; }

        public byte ISFOLDER { get; set; }

        public bool ISMARK { get; set; }

        public int VERSTAMP { get; set; }

        [Required]
        [StringLength(30)]
        public string SP85 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP86 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP208 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2417 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP97 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP5066 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP94 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4427 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP103 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP104 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP8842 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP8845 { get; set; }

        [Required]
        [StringLength(120)]
        public string SP8848 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP8849 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP8899 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP9304 { get; set; }

        public DateTime SP9305 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10091 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10366 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10397 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10406 { get; set; }

        [Required]
        [StringLength(25)]
        public string SP10479 { get; set; }

        public DateTime SP10480 { get; set; }

        [Required]
        [StringLength(25)]
        public string SP10481 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10535 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10784 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP11534 { get; set; }

        [Required]
        [StringLength(180)]
        public string SP12309 { get; set; }

        [Required]
        [StringLength(20)]
        public string SP12643 { get; set; }

        [Required]
        [StringLength(20)]
        public string SP12992 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13277 { get; set; }

        [Column(TypeName = "text")]
        public string SP95 { get; set; }

        [Column(TypeName = "text")]
        public string SP101 { get; set; }

        [Column(TypeName = "text")]
        public string SP12310 { get; set; }
    }
}
