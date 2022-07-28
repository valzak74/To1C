namespace Informer1C.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class DH2457
    {
        [Key]
        [StringLength(9)]
        public string IDDOC { get; set; }

        [Required]
        [StringLength(13)]
        public string SP4433 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2621 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2434 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2435 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2436 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2437 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2439 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2440 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2441 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2442 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2443 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2444 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP2445 { get; set; }

        public DateTime SP2438 { get; set; }

        public DateTime SP4434 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4437 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP4760 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP7943 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP8681 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP8835 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP8910 { get; set; }

        [Required]
        [StringLength(9)]
        public string SP10382 { get; set; }

        [Required]
        [StringLength(7)]
        public string SP10864 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP10865 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP11556 { get; set; }

        [Required]
        [StringLength(10)]
        public string SP11557 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2451 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2452 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal SP2453 { get; set; }

        [Required]
        [StringLength(150)]
        public string SP660 { get; set; }
    }
}
