using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ResidenciaApp.Infrastructure.Migrations
{
    /// <inheritdoc />
    public partial class AddResidenciaIdToTarifa : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<Guid>(
                name: "ResidenciaId",
                table: "Tarifas",
                type: "uniqueidentifier",
                nullable: false,
                defaultValue: new Guid("00000000-0000-0000-0000-000000000000"));

            migrationBuilder.Sql("UPDATE Tarifas SET ResidenciaId = (SELECT TOP 1 Id FROM Residencias)");

            migrationBuilder.CreateIndex(
                name: "IX_Tarifas_ResidenciaId",
                table: "Tarifas",
                column: "ResidenciaId");

            migrationBuilder.AddForeignKey(
                name: "FK_Tarifas_Residencias_ResidenciaId",
                table: "Tarifas",
                column: "ResidenciaId",
                principalTable: "Residencias",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Tarifas_Residencias_ResidenciaId",
                table: "Tarifas");

            migrationBuilder.DropIndex(
                name: "IX_Tarifas_ResidenciaId",
                table: "Tarifas");

            migrationBuilder.DropColumn(
                name: "ResidenciaId",
                table: "Tarifas");
        }
    }
}
