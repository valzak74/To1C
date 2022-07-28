namespace Informer1C.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class DbModel : DbContext
    {
        public DbModel()
            : base("name=DbModel")
        {
        }

        public virtual DbSet<C_1SCRDOC> C_1SCRDOC { get; set; }
        public virtual DbSet<C_1SJOURN> C_1SJOURN { get; set; }
        public virtual DbSet<DH13433> DH13433 { get; set; }
        public virtual DbSet<DH2457> DH2457 { get; set; }
        public virtual DbSet<DT13433> DT13433 { get; set; }
        public virtual DbSet<DT2457> DT2457 { get; set; }
        public virtual DbSet<RG4667> RG4667 { get; set; }
        public virtual DbSet<SC131> SC131 { get; set; }
        public virtual DbSet<SC172> SC172 { get; set; }
        public virtual DbSet<SC30> SC30 { get; set; }
        public virtual DbSet<SC4014> SC4014 { get; set; }
        public virtual DbSet<SC84> SC84 { get; set; }
        public virtual DbSet<SC8840> SC8840 { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<C_1SCRDOC>()
                .Property(e => e.PARENTVAL)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SCRDOC>()
                .Property(e => e.CHILD_DATE_TIME_IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SCRDOC>()
                .Property(e => e.CHILDID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.DATE_TIME_IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.DNPREFIX)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.DOCNO)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP74)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP798)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP4056)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP5365)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8662)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8663)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8664)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8665)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8666)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8720)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<C_1SJOURN>()
                .Property(e => e.SP8723)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH13433>()
                .Property(e => e.IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH13433>()
                .Property(e => e.SP13424)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH13433>()
                .Property(e => e.SP660)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP4433)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2621)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2434)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2435)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2436)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2437)
                .HasPrecision(9, 4);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2439)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2440)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2441)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2442)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2443)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2444)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2445)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP4437)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP4760)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP7943)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP8681)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP8835)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP8910)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP10382)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP10864)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP10865)
                .HasPrecision(4, 0);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP11556)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP11557)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2451)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2452)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP2453)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DH2457>()
                .Property(e => e.SP660)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13425)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13426)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13427)
                .HasPrecision(13, 3);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13428)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13429)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13430)
                .HasPrecision(3, 0);

            modelBuilder.Entity<DT13433>()
                .Property(e => e.SP13431)
                .HasPrecision(1, 0);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.IDDOC)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2446)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2447)
                .HasPrecision(13, 3);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2448)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2449)
                .HasPrecision(9, 3);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2450)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2451)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2454)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2452)
                .HasPrecision(14, 2);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2455)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<DT2457>()
                .Property(e => e.SP2453)
                .HasPrecision(14, 2);

            modelBuilder.Entity<RG4667>()
                .Property(e => e.SP4663)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<RG4667>()
                .Property(e => e.SP4664)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<RG4667>()
                .Property(e => e.SP4665)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<RG4667>()
                .Property(e => e.SP13435)
                .HasPrecision(1, 0);

            modelBuilder.Entity<RG4667>()
                .Property(e => e.SP4666)
                .HasPrecision(14, 5);

            modelBuilder.Entity<SC131>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP145)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP135)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP141)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP665)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP140)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP4101)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP4102)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP5350)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP4828)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP4829)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP6568)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP8375)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP148)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP149)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP144)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP5346)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP593)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP143)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP664)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP5347)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP5348)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP2905)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP5349)
                .IsUnicode(false);

            modelBuilder.Entity<SC131>()
                .Property(e => e.SP137)
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.PARENTID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP521)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP667)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP4137)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP573)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP4426)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP572)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP583)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP8380)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP9631)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP10379)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP12916)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP13072)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP13073)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC172>()
                .Property(e => e.SP186)
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP9533)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP6811)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP5843)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP8122)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1949)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP2643)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP9382)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP2272)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1950)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1951)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP4010)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP2274)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP2275)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP5727)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP873)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP9855)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1953)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1954)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP5354)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP2336)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP1952)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP8272)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP9992)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP10808)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP11600)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP11726)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP11781)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP12835)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP12843)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC30>()
                .Property(e => e.SP13116)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP4011)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP4012)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP4133)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP5011)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP5010)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP10879)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP10880)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP12015)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP13106)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC4014>()
                .Property(e => e.SP4073)
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.PARENTID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP85)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP86)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP208)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP2417)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP97)
                .HasPrecision(13, 3);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP5066)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP94)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP4427)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP103)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP104)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP8842)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP8845)
                .HasPrecision(7, 2);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP8848)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP8849)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP8899)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP9304)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10091)
                .HasPrecision(13, 3);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10366)
                .HasPrecision(12, 3);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10397)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10406)
                .HasPrecision(13, 3);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10479)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10481)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10535)
                .HasPrecision(3, 2);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP10784)
                .HasPrecision(4, 2);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP11534)
                .HasPrecision(10, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP12309)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP12643)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP12992)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP13277)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP95)
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP101)
                .IsUnicode(false);

            modelBuilder.Entity<SC84>()
                .Property(e => e.SP12310)
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.ID)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.CODE)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.DESCR)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP11193)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP11254)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP11467)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP11510)
                .HasPrecision(1, 0);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP11668)
                .HasPrecision(5, 2);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP12162)
                .IsFixedLength()
                .IsUnicode(false);

            modelBuilder.Entity<SC8840>()
                .Property(e => e.SP12932)
                .IsFixedLength()
                .IsUnicode(false);
        }
    }
}
