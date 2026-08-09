using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ResidenciaApp.Infrastructure.Migrations
{
    /// <inheritdoc />
    public partial class AddFechaSolicitud : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "FechaSolicitud",
                table: "Reservas",
                type: "datetime2",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "FechaSolicitud",
                table: "Reservas");
        }
    }
}
