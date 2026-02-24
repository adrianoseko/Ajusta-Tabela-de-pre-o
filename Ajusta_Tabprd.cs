using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

// Program.cs - Minimal ASP.NET Core (.NET 9) web app converted from Ajusta_Tabprd.py
// Single-file application exposing web UI to list price tables, view/edit items, and export Excel.

// Models
record PriceTable(int IdTabPreco, string Nome);
record PriceItem(int IdTabPreco, int IdPrd, string CodigoPrd, string NomeFantasia, decimal? Preco, decimal? Custo, decimal? Margem, decimal? AdicFinanc);

// Service to interact with DB and perform operations
class PriceTableService : IDisposable
{
    private readonly string _connectionString;
    private readonly SqlConnection _connection;

    // Keep default DB server and credentials to preserve original behavior
    private const string DbDriver = "{SQL Server}"; // informational
    private const string DbServer = " Seu Servidor"; // original placeholder
    private const string DbUid = "login";
    private const string DbPwd = "senha";

    public PriceTableService(string[] args)
    {
        // Expect the same 3 arguments as original: banco (/d:...), usuario (/u:...), coligada (/c:...)
        if (args is null || args.Length < 3)
            throw new ArgumentException("Expected three command-line arguments: banco, usuario, coligada");

        // Keep same parsing logic (/d:, /u:, /c:)
        var banco = args[0].Split(new[] { "/d:" }, StringSplitOptions.None);
        var usuario = args[1].Split(new[] { "/u:" }, StringSplitOptions.None);
        var coligada = args[2].Split(new[] { "/c:" }, StringSplitOptions.None);

        if (banco.Length < 2) throw new ArgumentException("Invalid banco argument");

        // Build classic SQL Server connection string
        _connectionString = new SqlConnectionStringBuilder
        {
            DataSource = DbServer.Trim(),
            InitialCatalog = banco[1].Trim(),
            UserID = DbUid,
            Password = DbPwd,
            // disable encrypt for older servers like original environment likely
            Encrypt = false,
            TrustServerCertificate = true,
            MultipleActiveResultSets = false
        }.ConnectionString;

        _connection = new SqlConnection(_connectionString);
        // Connect eagerly so first request does not fail mysteriously
        try
        {
            _connection.Open();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Could not open database connection. See inner exception for details.", ex);
        }
    }

    public IEnumerable<PriceTable> GetPriceTables()
    {
        get
        {
            const string sql =
                "SELECT ZTC.IDTABPRECO, NOME from ZA_TTABPRECO ZTC " +
                "LEFT JOIN TTABPRECO TTP (NOLOCK) ON TTP.CODCOLIGADA = ZTC.CODCOLIGADA AND TTP.IDTABPRECO = ZTC.IDTABPRECO " +
                "where USADEFAULTABELA = 'N' and " +
                "TTP.IDTABPRECO > 3 AND TTP.ATIVA = 1 AND " +
                "CONVERT(VARCHAR(10) , GETDATE() , 126) >= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAINI , 126) AND" +
                "CONVERT(VARCHAR(10) , GETDATE() , 126) <= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAFIM , 126)";

            using var cmd = new SqlCommand(sql, _connection);
            using var reader = cmd.ExecuteReader();
            var results = new List<PriceTable>();
            while (reader.Read())
            {
                results.Add(new PriceTable(reader.GetInt32(0), reader.IsDBNull(1) ? string.Empty : reader.GetString(1)));
            }

            return results;
        }
    }

    public IEnumerable<PriceItem> GetItems(int idTabPreco)
    {
        string sql =
            "SELECT IDTABPRECO, ZTC.IDPRD , CODIGOPRD, NOMEFANTASIA, PRECO, CUSTO, MARGEM, ADIC_FINANC " +
            "from ZA_TTABPRECOITM ZTC " +
            "LEFT JOIN TPRD (NOLOCK) ON TPRD.CODCOLIGADA = ZTC.CODCOLIGADA AND TPRD.IDPRD = ZTC.IDPRD " +
            "where IDTABPRECO = @idtab ORDER BY NOMEFANTASIA";

        using var cmd = new SqlCommand(sql, _connection);
        cmd.Parameters.AddWithValue("@idtab", idTabPreco);
        using var reader = cmd.ExecuteReader();
        var list = new List<PriceItem>();
        while (reader.Read())
        {
            var idtab = reader.GetInt32(0);
            var idprd = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
            var cod = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
            var nome = reader.IsDBNull(3) ? string.Empty : reader.GetString(3);
            decimal? preco = reader.IsDBNull(4) ? null : reader.GetDecimal(4);
            decimal? custo = reader.IsDBNull(5) ? null : reader.GetDecimal(5);
            decimal? margem = reader.IsDBNull(6) ? null : reader.GetDecimal(6);
            decimal? adic = reader.IsDBNull(7) ? null : reader.GetDecimal(7);
            list.Add(new PriceItem(idtab, idprd, cod, nome, preco, custo, margem, adic));
        }

        return list;
    }

    public void UpdateItem(int idTabPreco, int idPrd, decimal? preco, decimal? custo, decimal? margem, decimal? adicFinanc)
    {
        // Preserve original update logic but use parameters to avoid injection
        const string sql =
            "UPDATE ZA_TTABPRECOITM SET PRECO = @preco, CUSTO = @custo, MARGEM = @margem, ADIC_FINANC = @adc " +
            "WHERE CODCOLIGADA = 5 AND IDPRD = @idprd and IDTABPRECO = @idtab";

        using var cmd = new SqlCommand(sql, _connection);
        cmd.Parameters.AddWithValue("@preco", preco.HasValue ? (object)preco.Value : DBNull.Value);
        cmd.Parameters.AddWithValue("@custo", custo.HasValue ? (object)custo.Value : DBNull.Value);
        cmd.Parameters.AddWithValue("@margem", margem.HasValue ? (object)margem.Value : DBNull.Value);
        cmd.Parameters.AddWithValue("@adc", adicFinanc.HasValue ? (object)adicFinanc.Value : DBNull.Value);
        cmd.Parameters.AddWithValue("@idprd", idPrd);
        cmd.Parameters.AddWithValue("@idtab", idTabPreco);

        var affected = cmd.ExecuteNonQuery();
        if (affected == 0)
            throw new InvalidOperationException("No rows were updated. Verify identifiers.");
    }

    public MemoryStream ExportToExcel(int idTabPreco, string sheetName)
    {
        var items = GetItems(idTabPreco).ToList();
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add(!string.IsNullOrWhiteSpace(sheetName) ? sheetName : "Tabela");

        // Header
        ws.Cell(1, 1).Value = "IDTABPRECO";
        ws.Cell(1, 2).Value = "IDPRD";
        ws.Cell(1, 3).Value = "CODIGOPRD";
        ws.Cell(1, 4).Value = "NOMEFANTASIA";
        ws.Cell(1, 5).Value = "PRECO";
        ws.Cell(1, 6).Value = "CUSTO";
        ws.Cell(1, 7).Value = "MARGEM";
        ws.Cell(1, 8).Value = "ADIC_FINANC";

        for (var i = 0; i < items.Count; i++)
        {
            var r = i + 2;
            var it = items[i];
            ws.Cell(r, 1).Value = it.IdTabPreco;
            ws.Cell(r, 2).Value = it.IdPrd;
            ws.Cell(r, 3).Value = it.CodigoPrd;
            ws.Cell(r, 4).Value = it.NomeFantasia;
            ws.Cell(r, 5).Value = it.Preco;
            ws.Cell(r, 6).Value = it.Custo;
            ws.Cell(r, 7).Value = it.Margem;
            ws.Cell(r, 8).Value = it.AdicFinanc;
        }

        ws.Columns().AdjustToContents();
        var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    public void Dispose()
    {
        _connection?.Dispose();
    }
}

// Helper HTML utilities
static class HtmlHelpers
{
    public static string Encode(object? value) => WebUtility.HtmlEncode(value?.ToString() ?? string.Empty);

    public static string RenderLayout(string title, string bodyHtml)
    {
        var sb = new StringBuilder();
        sb.Append("<!doctype html><html lang=\"pt-BR\"><head><meta charset=\"utf-8\"/><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>"
                 + $"<title>{Encode(title)}</title>"
                 + "<style>body{font-family:Segoe UI,Arial;background:#f8f9fa;padding:20px}table{border-collapse:collapse;width:100%;background:white}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background:#343a40;color:white}a.button{display:inline-block;padding:6px 12px;margin:4px 2px;background:#007bff;color:white;text-decoration:none;border-radius:4px}form.inline{display:inline}</style></head><body>");
        sb.Append($"<h2>{Encode(title)}</h2>");
        sb.Append(bodyHtml);
        sb.Append("</body></html>");
        return sb.ToString();
    }
}

// Setup web app
var builder = WebApplication.CreateBuilder(args);

// Register PriceTableService as singleton using command-line args
builder.Services.AddSingleton(new PriceTableService(args));

var app = builder.Build();

app.MapGet("/", (PriceTableService svc) =>
{
    IEnumerable<PriceTable> tables;
    try
    {
        tables = svc.GetPriceTables();
    }
    catch (Exception ex)
    {
        var errHtml = $"<div style=\"color:red;\">Erro ao carregar tabelas: {HtmlHelpers.Encode(ex.Message)}</div>";
        return Results.Content(HtmlHelpers.RenderLayout("Tabela de Preços - Erro", errHtml), "text/html");
    }

    var sb = new StringBuilder();
    sb.Append("<div>");

    if (!tables.Any())
    {
        sb.Append("<p>Nenhuma tabela de preço encontrada.</p>");
    }
    else
    {
        sb.Append("<ul>");
        foreach (var t in tables)
        {
            sb.Append($"<li><a href=\"/table/{t.IdTabPreco}\">{HtmlHelpers.Encode(t.Nome)} (ID {t.IdTabPreco})</a></li>");
        }
        sb.Append("</ul>");
    }

    sb.Append("<p>Comandos: selecione uma tabela para visualizar e editar itens. Use Exportar para baixar Excel.</p>");
    sb.Append("</div>");

    return Results.Content(HtmlHelpers.RenderLayout("CGA.NET - Tabela Preços", sb.ToString()), "text/html");
});

app.MapGet("/table/{id:int}", (int id, HttpRequest req, PriceTableService svc) =>
{
    var q = req.Query["q"].ToString() ?? string.Empty; // filter
    IEnumerable<PriceItem> items;
    try
    {
        items = svc.GetItems(id);
    }
    catch (Exception ex)
    {
        var err = $"<div style=\"color:red\">Erro ao carregar itens: {HtmlHelpers.Encode(ex.Message)}</div>";
        return Results.Content(HtmlHelpers.RenderLayout("Itens da Tabela - Erro", err), "text/html");
    }

    if (!string.IsNullOrWhiteSpace(q))
    {
        var up = q.ToUpperInvariant();
        items = items.Where(i => (i.NomeFantasia ?? string.Empty).ToUpperInvariant().Contains(up));
    }

    var sb = new StringBuilder();
    sb.Append($"<p><a href=\"/\" class=\"button\">Voltar</a> <a class=\"button\" href=\"/export/{id}\">Exportar</a></p>");
    sb.Append("<form method=\"get\" action=\"\"><label>Pesquisar: <input type=\"text\" name=\"q\" value=\"" + HtmlHelpers.Encode(q) + "\"/></label> <button type=\"submit\">Filtrar</button></form>");

    sb.Append("<table><thead><tr>");
    sb.Append("<th>ID Produto</th><th>Código Produto</th><th>Nome Produto</th><th>Preço</th><th>Custo</th><th>Margem</th><th>Adicional Financeiro</th><th>Ações</th>");
    sb.Append("</tr></thead><tbody>");

    foreach (var it in items)
    {
        sb.Append("<tr>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.IdPrd)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.CodigoPrd)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.NomeFantasia)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.Preco?.ToString() ?? string.Empty)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.Custo?.ToString() ?? string.Empty)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.Margem?.ToString() ?? string.Empty)}</td>");
        sb.Append($"<td>{HtmlHelpers.Encode(it.AdicFinanc?.ToString() ?? string.Empty)}</td>");
        sb.Append($"<td><a href=\"/edit/{id}/{it.IdPrd}\">Editar</a></td>");
        sb.Append("</tr>");
    }

    sb.Append("</tbody></table>");
    return Results.Content(HtmlHelpers.RenderLayout($"Itens da Tabela {id}", sb.ToString()), "text/html");
});

app.MapGet("/edit/{tableId:int}/{idPrd:int}", (int tableId, int idPrd, PriceTableService svc) =>
{
    PriceItem? item = svc.GetItems(tableId).FirstOrDefault(i => i.IdPrd == idPrd);
    if (item is null)
    {
        return Results.Content(HtmlHelpers.RenderLayout("Editar Item", $"<div style=\"color:red\">Item {idPrd} não encontrado na tabela {tableId}.</div>"), "text/html");
    }

    var sb = new StringBuilder();
    sb.Append($"<p><a href=\"/table/{tableId}\" class=\"button\">Voltar</a></p>");
    sb.Append("<form method=\"post\" action=\"/edit/" + tableId + "/" + idPrd + "\">\n");
    sb.Append("<div><label>Código do Produto: <input name=\"codigo\" value=\"" + HtmlHelpers.Encode(item.CodigoPrd) + "\" readonly /></label></div>");
    sb.Append("<div><label>Nome do Produto: <input name=\"nome\" value=\"" + HtmlHelpers.Encode(item.NomeFantasia) + "\" readonly /></label></div>");
    sb.Append("<div><label>Preço: <input name=\"preco\" value=\"" + HtmlHelpers.Encode(item.Preco?.ToString() ?? string.Empty) + "\" /></label></div>");
    sb.Append("<div><label>Custo: <input name=\"custo\" value=\"" + HtmlHelpers.Encode(item.Custo?.ToString() ?? string.Empty) + "\" /></label></div>");
    sb.Append("<div><label>Margem: <input name=\"margem\" value=\"" + HtmlHelpers.Encode(item.Margem?.ToString() ?? string.Empty) + "\" /></label></div>");
    sb.Append("<div><label>Adicional Financeiro: <input name=\"adic\" value=\"" + HtmlHelpers.Encode(item.AdicFinanc?.ToString() ?? string.Empty) + "\" /></label></div>");
    sb.Append("<div><button type=\"submit\">Gravar</button> <a href=\"/table/" + tableId + "\" class=\"button\">Cancelar</a></div>");
    sb.Append("</form>");

    return Results.Content(HtmlHelpers.RenderLayout($"Editar Item {idPrd}", sb.ToString()), "text/html");
});

app.MapPost("/edit/{tableId:int}/{idPrd:int}", async (HttpRequest req, int tableId, int idPrd, PriceTableService svc) =>
{
    // Read form values
    try
    {
        if (!req.HasFormContentType)
            return Results.BadRequest("Invalid form submission");

        var form = await req.ReadFormAsync();
        var precoStr = form["preco"].ToString();
        var custoStr = form["custo"].ToString();
        var margemStr = form["margem"].ToString();
        var adicStr = form["adic"].ToString();

        decimal? preco = ParseNullableDecimal(precoStr);
        decimal? custo = ParseNullableDecimal(custoStr);
        decimal? margem = ParseNullableDecimal(margemStr);
        decimal? adic = ParseNullableDecimal(adicStr);

        svc.UpdateItem(tableId, idPrd, preco, custo, margem, adic);

        // Redirect back to table view
        return Results.Redirect($"/table/{tableId}");
    }
    catch (Exception ex)
    {
        var err = $"<div style=\"color:red\">Erro ao gravar alterações: {HtmlHelpers.Encode(ex.Message)}</div>";
        return Results.Content(HtmlHelpers.RenderLayout("Erro ao Gravar", err), "text/html");
    }
});

app.MapGet("/export/{tableId:int}", (int tableId, PriceTableService svc) =>
{
    try
    {
        // Get a friendly sheet / filename using table data if available
        var tables = svc.GetPriceTables();
        var tableName = tables.FirstOrDefault(t => t.IdTabPreco == tableId)?.Nome ?? $"Tabela_{tableId}";
        var ms = svc.ExportToExcel(tableId, tableName);
        var fileName = SanitizedFilename(tableName) + ".xlsx";
        return Results.File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    catch (Exception ex)
    {
        var err = $"<div style=\"color:red\">Erro ao exportar: {HtmlHelpers.Encode(ex.Message)}</div>";
        return Results.Content(HtmlHelpers.RenderLayout("Exportação - Erro", err), "text/html");
    }
});

app.Run();

// Utilities
static decimal? ParseNullableDecimal(string? s)
{
    if (string.IsNullOrWhiteSpace(s)) return null;
    if (decimal.TryParse(s, out var v)) return v;
    return null;
}

static string SanitizedFilename(string name)
{
    var invalid = Path.GetInvalidFileNameChars();
    var clean = new string(name.Where(c => !invalid.Contains(c)).ToArray());
    return string.IsNullOrWhiteSpace(clean) ? "export" : clean.Replace(' ', '_');
}
