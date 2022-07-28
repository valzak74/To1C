namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("_1SJOURN")]
    public partial class C_1SJOURN
    {
        [Key]
        public int ROW_ID { get; set; }

        public int IDJOURNAL { get; set; }

        [Required]
        [StringLength(9)]
        public string IDDOC { get; set; }

        public int IDDOCDEF { get; set; }

        public short APPCODE { get; set; }

        [Required]
        [StringLength(23)]
        public string DATE_TIME_IDDOC { get; set; }

        [Required]
        [StringLength(18)]
        public string DNPREFIX { get; set; }

        [Required]
        [StringLength(10)]
        public string DOCNO { get; set; }

        public byte CLOSED { get; set; }

        public bool ISMARK { get; set; }

        public int ACTCNT { get; set; }

        public int VERSTAMP { get; set; }

        public bool RF639 { get; set; }

        public bool RF464 { get; set; }

        public bool RF4667 { get; set; }

        public bool RF4674 { get; set; }

        public bool RF635 { get; set; }

        public bool RF3549 { get; set; }

        public bool RF4343 { get; set; }

        public bool RF8677 { get; set; }

        public bool RF8696 { get; set; }

        public bool RF405 { get; set; }

        public bool RF328 { get; set; }

        public bool RF351 { get; set; }

        public bool RF2964 { get; set; }

        public bool RF4335 { get; set; }

        public bool RF4314 { get; set; }

        public bool RF2351 { get; set; }

        public bool RF438 { get; set; }

        public bool RF4480 { get; set; }

        public bool RF8888 { get; set; }

        public bool RF8894 { get; set; }

        public bool RF9143 { get; set; }

        public bool RF9469 { get; set; }

        public bool RF9531 { get; set; }

        public bool RF9596 { get; set; }

        public bool RF9972 { get; set; }

        public bool RF9981 { get; set; }

        public bool RF9989 { get; set; }

        public bool RF10305 { get; set; }

        public bool RF10313 { get; set; }

        public bool RF10318 { get; set; }

        public bool RF10324 { get; set; }

        public bool RF10471 { get; set; }

        public bool RF10476 { get; set; }

        public bool RF11049 { get; set; }

        public bool RF11055 { get; set; }

        public bool RF11495 { get; set; }

        public bool RF11973 { get; set; }

        public bool RF12351 { get; set; }

        public bool RF12406 { get; set; }

        public bool RF12413 { get; set; }

        public bool RF12503 { get; set; }

        public bool RF12566 { get; set; }

        public bool RF12618 { get; set; }

        public bool RF12791 { get; set; }

        public bool RF12815 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP74 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP798 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4056 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP5365 { get; set; }

        [Required]
        [StringLength(2)]
        public string SP8662 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8663 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8664 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8665 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8666 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8720 { get; set; }

        [Required]
        [StringLength(30)]
        public string SP8723 { get; set; }

        public byte DS1946 { get; set; }

        public byte DS4757 { get; set; }

        public byte DS5722 { get; set; }
    }
}
