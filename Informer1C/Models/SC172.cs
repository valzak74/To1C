namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC172
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
        [StringLength(8)]
        public string CODE { get; set; }

        [Required]
        [StringLength(30)]
        public string DESCR { get; set; }

        public byte ISFOLDER { get; set; }

        public bool ISMARK { get; set; }

        public int VERSTAMP { get; set; }

        [Required]
        [StringLength(13)]
        public string SP521 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP667 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4137 { get; set; }

        [Required]
        [StringLength(40)]
        public string SP573 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4426 { get; set; }

        [Required]
        [StringLength(40)]
        public string SP572 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP583 { get; set; }

        [Required]
        [StringLength(20)]
        public string SP8380 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP9631 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP10379 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP12916 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP13072 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13073 { get; set; }

        [Column(TypeName = "text")]
        public string SP186 { get; set; }
    }
}
