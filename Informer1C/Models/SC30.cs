namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SC30
    {
        [Key]
        public int ROW_ID { get; set; }

        [Required]
        [StringLength(9)]
        public string ID { get; set; }

        [Required]
        [StringLength(24)]
        public string CODE { get; set; }

        [Required]
        [StringLength(50)]
        public string DESCR { get; set; }

        public bool ISMARK { get; set; }

        public int VERSTAMP { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP9533 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP6811 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP5843 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP8122 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1949 { get; set; }

        public DateTime SP5578 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2643 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP9382 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2272 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1950 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1951 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4010 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2274 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2275 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP5727 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP873 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP9855 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1953 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1954 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP5354 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2336 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP1952 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP8272 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP9992 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP10808 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP11600 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP11726 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP11781 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP12835 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP12843 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP13116 { get; set; }
    }
}
