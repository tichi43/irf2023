using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

#nullable disable

namespace negyedikhet_excel.models
{
    public partial class RealEstateContext : DbContext
    {
        public RealEstateContext()
        {
        }

        public RealEstateContext(DbContextOptions<RealEstateContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Flat> Flats { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
                optionsBuilder.UseSqlServer("Data Source=bit.uni-corvinus.hu;Initial Catalog=RealEstate;Persist Security Info=True;Encrypt=False;User ID=hallgato;Password=Password123");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Flat>(entity =>
            {
                entity.HasKey(e => e.FlatSk)
                    .HasName("PK__Flat__7CD69E721CD5505F");

                entity.ToTable("Flat");

                entity.Property(e => e.FlatSk).HasColumnName("FlatSK");

                entity.Property(e => e.Code)
                    .IsRequired()
                    .HasMaxLength(6)
                    .IsFixedLength(true);

                entity.Property(e => e.NumberOfRooms).HasColumnType("decimal(18, 2)");

                entity.Property(e => e.Price).HasColumnType("decimal(18, 2)");

                entity.Property(e => e.Side)
                    .IsRequired()
                    .HasMaxLength(4)
                    .IsFixedLength(true);

                entity.Property(e => e.Vendor)
                    .IsRequired()
                    .HasMaxLength(20);
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
