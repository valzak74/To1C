namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("_1SCRDOC")]
    public partial class C_1SCRDOC
    {
        [Key]
        public int ROW_ID { get; set; }

        public int MDID { get; set; }

        [Required]
        [StringLength(23)]
        public string PARENTVAL { get; set; }

        [Required]
        [StringLength(23)]
        public string CHILD_DATE_TIME_IDDOC { get; set; }

        [Required]
        [StringLength(9)]
        public string CHILDID { get; set; }

        public byte FLAGS { get; set; }
    }
}
