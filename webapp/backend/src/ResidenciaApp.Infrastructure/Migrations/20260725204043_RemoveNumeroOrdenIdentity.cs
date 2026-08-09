using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ResidenciaApp.Infrastructure.Migrations
{
    /// <inheritdoc />
    public partial class RemoveNumeroOrdenIdentity : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NumeroOrden",
                table: "Reservas");

            migrationBuilder.AddColumn<int>(
                name: "NumeroOrden",
                table: "Reservas",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NumeroOrden",
                table: "Reservas");

            migrationBuilder.AddColumn<int>(
                name: "NumeroOrden",
                table: "Reservas",
                type: "int",
                nullable: false)
                .Annotation("SqlServer:Identity", "1, 1");
        }
    }
}
