using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ResidenciaApp.Infrastructure.Migrations
{
    /// <inheritdoc />
    public partial class AddHuespedFinalidadAndCategory : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "EmpleoCategoria",
                table: "Huespedes",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Finalidad",
                table: "Huespedes",
                type: "nvarchar(100)",
                maxLength: 100,
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "EmpleoCategoria",
                table: "Huespedes");

            migrationBuilder.DropColumn(
                name: "Finalidad",
                table: "Huespedes");
        }
    }
}
