namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class DT13433
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
        public string SP13425 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP13426 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13427 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13428 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13429 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13430 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13431 { get; set; }
    }
}
