using System.Text.Json;
using Microsoft.Extensions.Logging;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public interface IEvaluacionService
{
    (string EmpleoCategoria, string Evaluacion) EvaluarReserva(Huesped? huesped, string? finalidad, string? empleoInput = null, string? situacionInput = null);
}

public class EvaluacionService(IResidenciaDbContext db, ILogger<EvaluacionService> logger) : IEvaluacionService
{
    private static Dictionary<string, string>? _matrizCached;

    public static readonly Dictionary<string, string> MatrizPorDefecto = new()
    {
        { "comisión no indem.|gc|viogen", "1, 1, 1" },
        { "comisión no indem.|gc|activo", "1, 1, 2" },
        { "comisión no indem.|gc|reserva activo", "1, 1, 2" },
        { "destino|gc|activo", "1, 1, 3" },
        { "comisión|gc|activo", "1, 1, 4" },
        { "enfermedad|gc|retirado", "2, 8, 1" },
        { "comisión|alumno|activo", "1, 2, 2" },
        { "comisión|militar en gc|activo", "1, 3, 2" },
        { "destino|militar en gc|activo", "1, 3, 2" },
        { "comisión|funcionario en gc|activo", "1, 4, 2" },
        { "destino|funcionario en gc|activo", "1, 4, 2" },
        { "enfermedad|gc|activo", "2, 1, 1" },
        { "otros|gc|viogen", "2, 1, 2" },
        { "urgencia|gc|activo", "2, 1, 3" },
        { "sepelio|gc|activo", "2, 1, 4" },
        { "máx. estancia|gc|activo", "2, 1, 5" },
        { "otros|gc|asociación", "2, 1, 6" },
        { "otros|gc|activo", "2, 1, 7" },
        { "otros|gc|reserva activo", "2, 1, 8" },
        { "otros|alumno|activo", "2, 2, 7" },
        { "otros|gc|reserva", "2, 3, 7" },
        { "otros|militar en gc|activo", "2, 4, 7" },
        { "otros|funcionario en gc|activo", "2, 5, 7" },
        { "otros|gc|excedencia", "2, 6, 7" },
        { "otros|gc|especiales", "2, 7, 7" },
        { "otros|gc|retirado", "2, 8, 7" },
        { "otros|gc|viuda", "2, 9, 7" },
        { "otros|gc|huerfano", "2, 9, 7" },
        { "otros|militar no  gc|activo", "2, 10, 7" },
        { "otros|militar no  gc|reserva", "2, 10, 7" },
        { "otros|militar no  gc|retirado", "2, 10, 7" }
    };

    public static void InvalidarCacheMatriz()
    {
        _matrizCached = null;
    }

    private Dictionary<string, string> GetMatrizEvaluacion()
    {
        if (_matrizCached != null) return _matrizCached;

        try
        {
            var setting = db.SystemSettings.FirstOrDefault(s => s.Key == "MatrizEvaluacion");
            if (setting != null && !string.IsNullOrWhiteSpace(setting.Value))
            {
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(setting.Value);
                if (dict != null && dict.Count > 0)
                {
                    _matrizCached = dict;
                    return _matrizCached;
                }
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Error al deserializar MatrizEvaluacion de SystemSettings. Usando matriz por defecto.");
        }

        _matrizCached = MatrizPorDefecto;
        return _matrizCached;
    }

    public (string EmpleoCategoria, string Evaluacion) EvaluarReserva(Huesped? huesped, string? finalidad, string? empleoInput = null, string? situacionInput = null)
    {
        var empleoRaw = !string.IsNullOrWhiteSpace(empleoInput) ? empleoInput : huesped?.Empleo;
        var situacionRaw = !string.IsNullOrWhiteSpace(situacionInput) ? situacionInput : huesped?.Situacion;
        if (huesped is null && string.IsNullOrWhiteSpace(empleoRaw)) return (string.Empty, string.Empty);

        // 1. Mapear Empleo
        var e = empleoRaw?.Trim().ToLower() ?? string.Empty;
        string empleoCat = "GC";
        if (e.Contains("alumno"))
        {
            empleoCat = "Alumno";
        }
        else if (e.Contains("funcionario"))
        {
            empleoCat = "Funcionario en GC";
        }
        else if (e.Contains("militar en"))
        {
            empleoCat = "Militar en GC";
        }
        else if (e.Contains("militar no") || e.Contains("militar"))
        {
            empleoCat = "Militar no  GC";
        }
        else
        {
            empleoCat = "GC";
        }

        // 2. Mapear Finalidad
        var f = finalidad?.Trim().ToLower() ?? "otros";
        string finalidadMapeada = "Otros";
        if (f.Contains("comision no") || f.Contains("comisión no") || f.Contains("indem"))
        {
            finalidadMapeada = "Comisión NO indem.";
        }
        else if (f.Contains("comision") || f.Contains("comisión"))
        {
            finalidadMapeada = "Comisión";
        }
        else if (f.Contains("destino"))
        {
            finalidadMapeada = "Destino";
        }
        else if (f.Contains("enfermedad"))
        {
            finalidadMapeada = "Enfermedad";
        }
        else if (f.Contains("urgencia"))
        {
            finalidadMapeada = "Urgencia";
        }
        else if (f.Contains("sepelio"))
        {
            finalidadMapeada = "Sepelio";
        }
        else if (f.Contains("estancia"))
        {
            finalidadMapeada = "Máx. Estancia";
        }
        else
        {
            finalidadMapeada = "Otros";
        }

        // 3. Mapear Situación
        var s = situacionRaw?.Trim().ToLower() ?? "activo";

        string situacionMapeada = "Activo";
        if (s.Contains("viogen"))
        {
            situacionMapeada = "Viogen";
        }
        else if (s.Contains("asoc"))
        {
            situacionMapeada = "Asociación";
        }
        else if (s.Contains("reserva activo") || s.Contains("reserva activa"))
        {
            situacionMapeada = "Reserva activo";
        }
        else if (s.Contains("reserva"))
        {
            situacionMapeada = "Reserva";
        }
        else if (s.Contains("excedencia"))
        {
            situacionMapeada = "Excedencia";
        }
        else if (s.Contains("especial"))
        {
            situacionMapeada = "Especiales";
        }
        else if (s.Contains("retirado") || s.Contains("jubilado"))
        {
            situacionMapeada = "Retirado";
        }
        else if (s.Contains("viuda"))
        {
            situacionMapeada = "Viuda";
        }
        else if (s.Contains("huerfano") || s.Contains("huérfano"))
        {
            situacionMapeada = "Huerfano";
        }
        else
        {
            situacionMapeada = "Activo";
        }

        var clave = $"{finalidadMapeada.ToLower()}|{empleoCat.ToLower()}|{situacionMapeada.ToLower()}";

        var matriz = GetMatrizEvaluacion();

        if (matriz.TryGetValue(clave, out var eval))
        {
            logger.LogDebug("Evaluación calculada para clave {Clave}: {Evaluacion}", clave, eval);
            return (empleoCat, eval);
        }

        logger.LogWarning("No se encontró coincidencia en la matriz para la clave: {Clave}", clave);
        return (empleoCat, "(NO VÁLIDO)");
    }
}
