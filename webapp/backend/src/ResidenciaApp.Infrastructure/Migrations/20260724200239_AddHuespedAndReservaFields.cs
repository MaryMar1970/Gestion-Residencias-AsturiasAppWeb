using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ResidenciaApp.Infrastructure.Migrations
{
    /// <inheritdoc />
    public partial class AddHuespedAndReservaFields : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "CodigosPostales",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    CodigoPostal = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Municipio = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: false),
                    Provincia = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CodigosPostales", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Huespedes",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    Dni = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: false),
                    Nombre = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    Apellidos = table.Column<string>(type: "nvarchar(150)", maxLength: 150, nullable: false),
                    Telefono = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: true),
                    Email = table.Column<string>(type: "nvarchar(200)", maxLength: 200, nullable: true),
                    Direccion = table.Column<string>(type: "nvarchar(300)", maxLength: 300, nullable: true),
                    CodigoPostal = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: true),
                    Municipio = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Provincia = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CentroOrigen = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Departamento = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TipoHuesped = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: false),
                    Empleo = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Situacion = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    EnListaNegra = table.Column<bool>(type: "bit", nullable: false),
                    MotivoListaNegra = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Notas = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CreadoEn = table.Column<DateTime>(type: "datetime2", nullable: false),
                    ActualizadoEn = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Huespedes", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Reservas",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    NumeroOrden = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    HabitacionId = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    HuespedId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    FechaEntrada = table.Column<DateOnly>(type: "date", nullable: false),
                    FechaSalida = table.Column<DateOnly>(type: "date", nullable: false),
                    Estado = table.Column<int>(type: "int", nullable: false),
                    Finalidad = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Empleo = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Evaluacion = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    NumPersonas = table.Column<int>(type: "int", nullable: false),
                    CamasSupletorias = table.Column<int>(type: "int", nullable: false),
                    EsBloqueo = table.Column<bool>(type: "bit", nullable: false),
                    MotivoBloqueo = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TarifaId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    TarifaNombreSnapshot = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    PrecioNocheAplicado = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    PorcentajeIvaAplicado = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    TotalNoches = table.Column<int>(type: "int", nullable: false),
                    ImporteBase = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    ImporteIva = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    ImporteTotal = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    Facturado = table.Column<bool>(type: "bit", nullable: false),
                    FacturaId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    Pagado = table.Column<bool>(type: "bit", nullable: false),
                    FormaPago = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: true),
                    FechaPago = table.Column<DateTime>(type: "datetime2", nullable: true),
                    Observaciones = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CreadoPorId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    CreadoEn = table.Column<DateTime>(type: "datetime2", nullable: false),
                    ActualizadoEn = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Reservas", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Reservas_Habitaciones_HabitacionId",
                        column: x => x.HabitacionId,
                        principalTable: "Habitaciones",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_Reservas_Huespedes_HuespedId",
                        column: x => x.HuespedId,
                        principalTable: "Huespedes",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                    table.ForeignKey(
                        name: "FK_Reservas_Tarifas_TarifaId",
                        column: x => x.TarifaId,
                        principalTable: "Tarifas",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                });

            migrationBuilder.CreateTable(
                name: "Facturas",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    NumeroFactura = table.Column<string>(type: "nvarchar(30)", maxLength: 30, nullable: false),
                    Serie = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    Ejercicio = table.Column<int>(type: "int", nullable: false),
                    NumeroOrden = table.Column<int>(type: "int", nullable: false),
                    ResidenciaId = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    ReservaId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    HuespedId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    FechaEmision = table.Column<DateOnly>(type: "date", nullable: false),
                    FechaVencimiento = table.Column<DateOnly>(type: "date", nullable: true),
                    FechaPago = table.Column<DateOnly>(type: "date", nullable: true),
                    DestinatarioNombre = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DestinatarioDni = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DestinatarioDireccion = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DestinatarioCp = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DestinatarioMunicipio = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    EmisorNombre = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    EmisorCif = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    EmisorDireccion = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    EmisorTelefono = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BaseImponible = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    PorcentajeIva = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    CuotaIva = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    Total = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    FormaPago = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Estado = table.Column<int>(type: "int", nullable: false),
                    Observaciones = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CreadoPorId = table.Column<Guid>(type: "uniqueidentifier", nullable: true),
                    CreadoEn = table.Column<DateTime>(type: "datetime2", nullable: false),
                    ActualizadoEn = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Facturas", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Facturas_Huespedes_HuespedId",
                        column: x => x.HuespedId,
                        principalTable: "Huespedes",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                    table.ForeignKey(
                        name: "FK_Facturas_Reservas_ReservaId",
                        column: x => x.ReservaId,
                        principalTable: "Reservas",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.SetNull);
                    table.ForeignKey(
                        name: "FK_Facturas_Residencias_ResidenciaId",
                        column: x => x.ResidenciaId,
                        principalTable: "Residencias",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "LineasFactura",
                columns: table => new
                {
                    Id = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    FacturaId = table.Column<Guid>(type: "uniqueidentifier", nullable: false),
                    Orden = table.Column<int>(type: "int", nullable: false),
                    Concepto = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Cantidad = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    Unidad = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    PrecioUnidad = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    Descuento = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    BaseLinea = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    PorcentajeIva = table.Column<decimal>(type: "decimal(5,2)", nullable: false),
                    CuotaIvaLinea = table.Column<decimal>(type: "decimal(10,2)", nullable: false),
                    TotalLinea = table.Column<decimal>(type: "decimal(10,2)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_LineasFactura", x => x.Id);
                    table.ForeignKey(
                        name: "FK_LineasFactura_Facturas_FacturaId",
                        column: x => x.FacturaId,
                        principalTable: "Facturas",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_CodigosPostales_CodigoPostal",
                table: "CodigosPostales",
                column: "CodigoPostal");

            migrationBuilder.CreateIndex(
                name: "IX_CodigosPostales_Municipio",
                table: "CodigosPostales",
                column: "Municipio");

            migrationBuilder.CreateIndex(
                name: "IX_Facturas_HuespedId",
                table: "Facturas",
                column: "HuespedId");

            migrationBuilder.CreateIndex(
                name: "IX_Facturas_ReservaId",
                table: "Facturas",
                column: "ReservaId");

            migrationBuilder.CreateIndex(
                name: "IX_Facturas_ResidenciaId",
                table: "Facturas",
                column: "ResidenciaId");

            migrationBuilder.CreateIndex(
                name: "IX_Facturas_Serie_Ejercicio_NumeroOrden",
                table: "Facturas",
                columns: new[] { "Serie", "Ejercicio", "NumeroOrden" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_Huespedes_Dni",
                table: "Huespedes",
                column: "Dni",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_LineasFactura_FacturaId",
                table: "LineasFactura",
                column: "FacturaId");

            migrationBuilder.CreateIndex(
                name: "IX_Reservas_HabitacionId",
                table: "Reservas",
                column: "HabitacionId");

            migrationBuilder.CreateIndex(
                name: "IX_Reservas_HuespedId",
                table: "Reservas",
                column: "HuespedId");

            migrationBuilder.CreateIndex(
                name: "IX_Reservas_TarifaId",
                table: "Reservas",
                column: "TarifaId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CodigosPostales");

            migrationBuilder.DropTable(
                name: "LineasFactura");

            migrationBuilder.DropTable(
                name: "Facturas");

            migrationBuilder.DropTable(
                name: "Reservas");

            migrationBuilder.DropTable(
                name: "Huespedes");
        }
    }
}
