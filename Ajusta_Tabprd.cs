using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AjustaTabprd
{
    /// <summary>
    /// GUI application to view and edit price table items.
    /// Refactored: separates data access (repository) and business logic (service).
    /// The Form now only coordinates UI interactions.
    /// Behavior preserved from original version.
    /// </summary>
    public partial class PriceTableForm : Form
    {
        // Database connection placeholders (update accordingly)
        private const string DbServer = " Seu Servidor";
        private const string DbUid = "login";
        private const string DbPwd = "senha";

        private readonly string[] _banco;
        private readonly string[] _user;
        private readonly string[] _coligada;

        // UI controls
        private readonly ComboBox _tableCombo = new ComboBox();
        private readonly DataGridView _grid = new DataGridView();
        private readonly VScrollBar _vScroll = new VScrollBar();
        private readonly Button _saveButton = new Button();
        private readonly Button _cancelButton = new Button();
        private readonly Button _exportButton = new Button();
        private readonly Button _cancelToMenuButton = new Button();
        private readonly TextBox _filterBox = new TextBox();
        private readonly Label _filterLabel = new Label();

        // Edit fields
        private readonly Label _codeLabel = new Label();
        private readonly TextBox _codeBox = new TextBox();
        private readonly Label _nameLabel = new Label();
        private readonly TextBox _nameBox = new TextBox();
        private readonly Label _priceLabel = new Label();
        private readonly TextBox _priceBox = new TextBox();
        private readonly Label _costLabel = new Label();
        private readonly TextBox _costBox = new TextBox();
        private readonly Label _marginLabel = new Label();
        private readonly TextBox _marginBox = new TextBox();
        private readonly Label _addfLabel = new Label();
        private readonly TextBox _addfBox = new TextBox();

        private readonly BindingSource _itemsBinding = new BindingSource();

        // Architectural components
        private readonly IPriceTableRepository _repository;
        private readonly PriceTableService _service;

        // State
        private DataTable _tablesTable = new DataTable();
        private DataTable _itemsTable = new DataTable();
        private int _selectedTableIndex = -1;
        private object _selectedTableId = null;

        public PriceTableForm(string[] args)
        {
            // Validate and parse arguments (keeps original splitting behavior)
            if (args == null || args.Length < 4)
            {
                MessageBox.Show("Expected three command-line arguments: banco, usuario, coligada", "Arguments", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
            }

            _banco = args[1].Split(new[] { "/d:" }, StringSplitOptions.None);
            _user = args[2].Split(new[] { "/u:" }, StringSplitOptions.None);
            _coligada = args[3].Split(new[] { "/c:" }, StringSplitOptions.None);

            // Build connection string exactly as original
            var connStringBuilder = new SqlConnectionStringBuilder
            {
                DataSource = DbServer,
                InitialCatalog = _banco.Length > 1 ? _banco[1] : string.Empty,
                UserID = DbUid,
                Password = DbPwd,
                IntegratedSecurity = false,
                ConnectTimeout = 30
            };

            // Initialize repository and service
            try
            {
                _repository = new SqlPriceTableRepository(connStringBuilder.ConnectionString);
                _service = new PriceTableService(_repository);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível conectar ao banco de dados.\nEntre em contato com o suporte!", "Erro DB", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(1);
                return;
            }

            InitializeForm();

            // Load initial tables
            LoadTableList();
        }

        private void InitializeForm()
        {
            this.Text = "CGA.NET - Tabela Preços";
            this.ClientSize = new Size(800, 600);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Table ComboBox
            _tableCombo.Location = new Point(20, 10);
            _tableCombo.Size = new Size(400, 25);
            _tableCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _tableCombo.SelectedIndexChanged += (s, e) => OnTableSelected();
            this.Controls.Add(_tableCombo);

            // DataGridView
            _grid.Location = new Point(20, 50);
            _grid.Size = new Size(760, 425);
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _grid.CellDoubleClick += (s, e) => BeginEditSelected();
            this.Controls.Add(_grid);

            // Vertical scrollbar (visual parity)
            _vScroll.Location = new Point(783, 50);
            _vScroll.Size = new Size(17, 425);
            _vScroll.Visible = false; // rely on DataGridView's internal scroll
            this.Controls.Add(_vScroll);

            // Filter label and box
            _filterLabel.Text = "Pesquisar:";
            _filterLabel.Location = new Point(390, 15);
            _filterLabel.AutoSize = true;
            _filterLabel.Visible = false;
            this.Controls.Add(_filterLabel);

            _filterBox.Location = new Point(450, 15);
            _filterBox.Size = new Size(200, 20);
            _filterBox.Visible = false;
            _filterBox.TextChanged += (s, e) => ApplyFilter();
            this.Controls.Add(_filterBox);

            // Buttons
            _exportButton.Text = "Exportar";
            _exportButton.Size = new Size(80, 25);
            _exportButton.Location = new Point(700, 500);
            _exportButton.Click += (s, e) => Export();
            _exportButton.Visible = false;
            this.Controls.Add(_exportButton);

            _cancelToMenuButton.Text = "Cancelar";
            _cancelToMenuButton.Size = new Size(80, 25);
            _cancelToMenuButton.Location = new Point(700, 550);
            _cancelToMenuButton.Click += (s, e) => CancelToMenu();
            _cancelToMenuButton.Visible = false;
            this.Controls.Add(_cancelToMenuButton);

            _saveButton.Text = "Gravar";
            _saveButton.Size = new Size(80, 25);
            _saveButton.Location = new Point(700, 530);
            _saveButton.Click += (s, e) => SaveChanges();
            _saveButton.Visible = false;
            this.Controls.Add(_saveButton);

            _cancelButton.Text = "Cancelar";
            _cancelButton.Size = new Size(80, 25);
            _cancelButton.Location = new Point(700, 560);
            _cancelButton.Click += (s, e) => CancelEdit();
            _cancelButton.Visible = false;
            this.Controls.Add(_cancelButton);

            // Edit fields: code/name/price/cost/margin/addf
            _codeLabel.Text = "Código do Produto";
            _codeLabel.Location = new Point(50, 480);
            _codeLabel.AutoSize = true;
            _codeLabel.Visible = false;
            this.Controls.Add(_codeLabel);

            _codeBox.Location = new Point(50, 500);
            _codeBox.Size = new Size(120, 20);
            _codeBox.ReadOnly = true;
            _codeBox.Visible = false;
            this.Controls.Add(_codeBox);

            _nameLabel.Text = "Nome do Produto";
            _nameLabel.Location = new Point(200, 480);
            _nameLabel.AutoSize = true;
            _nameLabel.Visible = false;
            this.Controls.Add(_nameLabel);

            _nameBox.Location = new Point(200, 500);
            _nameBox.Size = new Size(380, 20);
            _nameBox.ReadOnly = true;
            _nameBox.Visible = false;
            this.Controls.Add(_nameBox);

            _priceLabel.Text = "Preço";
            _priceLabel.Location = new Point(50, 530);
            _priceLabel.AutoSize = true;
            _priceLabel.Visible = false;
            this.Controls.Add(_priceLabel);

            _priceBox.Location = new Point(50, 550);
            _priceBox.Size = new Size(100, 20);
            _priceBox.Visible = false;
            this.Controls.Add(_priceBox);

            _costLabel.Text = "Custo";
            _costLabel.Location = new Point(200, 530);
            _costLabel.AutoSize = true;
            _costLabel.Visible = false;
            this.Controls.Add(_costLabel);

            _costBox.Location = new Point(200, 550);
            _costBox.Size = new Size(100, 20);
            _costBox.Visible = false;
            this.Controls.Add(_costBox);

            _marginLabel.Text = "Margem";
            _marginLabel.Location = new Point(350, 530);
            _marginLabel.AutoSize = true;
            _marginLabel.Visible = false;
            this.Controls.Add(_marginLabel);

            _marginBox.Location = new Point(350, 550);
            _marginBox.Size = new Size(100, 20);
            _marginBox.Visible = false;
            this.Controls.Add(_marginBox);

            _addfLabel.Text = "Adicional Financeiro";
            _addfLabel.Location = new Point(500, 530);
            _addfLabel.AutoSize = true;
            _addfLabel.Visible = false;
            this.Controls.Add(_addfLabel);

            _addfBox.Location = new Point(500, 550);
            _addfBox.Size = new Size(100, 20);
            _addfBox.Visible = false;
            this.Controls.Add(_addfBox);

            // Info labels bottom
            var bancoLabel = new Label { Text = $"Banco: {BancoSafe()}", Location = new Point(50, 580), AutoSize = true };
            var coligadaLabel = new Label { Text = $"Coligada: {ColigadaSafe()}", Location = new Point(200, 580), AutoSize = true };
            var userLabel = new Label { Text = $"Usuario: {UserSafe()}", Location = new Point(280, 580), AutoSize = true };
            var versionLabel = new Label { Text = "Versão: 1.0", Location = new Point(500, 580), AutoSize = true };

            this.Controls.Add(bancoLabel);
            this.Controls.Add(coligadaLabel);
            this.Controls.Add(userLabel);
            this.Controls.Add(versionLabel);
        }

        private string BancoSafe() => _banco != null && _banco.Length > 1 ? _banco[1] : string.Empty;
        private string UserSafe() => _user != null && _user.Length > 1 ? _user[1] : string.Empty;
        private string ColigadaSafe() => _coligada != null && _coligada.Length > 1 ? _coligada[1] : string.Empty;

        private void LoadTableList()
        {
            try
            {
                _tablesTable = _service.FetchPriceTables();

                if (_tablesTable == null || _tablesTable.Rows.Count == 0)
                {
                    MessageBox.Show("Nenhuma tabela de preço encontrada.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _tableCombo.Items.Clear();
                foreach (DataRow row in _tablesTable.Rows)
                {
                    _tableCombo.Items.Add(row["NOME"].ToString());
                }

                if (_tableCombo.Items.Count > 0)
                {
                    _tableCombo.SelectedIndex = 0; // default like original
                    OnTableSelected();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível obter lista de tabelas!\nEntre em contato com o suporte!", "ERRO!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnTableSelected()
        {
            var selectedName = _tableCombo.SelectedItem as string;
            if (string.IsNullOrEmpty(selectedName) || _tablesTable == null) return;

            try
            {
                var rows = _tablesTable.AsEnumerable().ToList();
                var idx = rows.FindIndex(r => string.Equals(r.Field<string>("NOME"), selectedName, StringComparison.Ordinal));
                if (idx < 0) return;

                _selectedTableIndex = idx;
                _selectedTableId = _tablesTable.Rows[idx]["IDTABPRECO"];

                _itemsTable = _service.FetchItemsForTable(_selectedTableId);

                _itemsBinding.DataSource = _itemsTable;
                _grid.DataSource = _itemsBinding;

                ConfigureGridColumns();

                _filterLabel.Visible = true;
                _filterBox.Visible = true;
                _exportButton.Visible = true;
                _cancelToMenuButton.Visible = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível carregar itens da tabela!\nEntre em contato com o suporte!", "ERRO!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigureGridColumns()
        {
            try
            {
                foreach (DataGridViewColumn col in _grid.Columns)
                {
                    col.Width = 100;
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                SetColumnIfExists("IDPRD", "ID Produto", 80);
                SetColumnIfExists("CODIGOPRD", "Código Produto", 100);
                SetColumnIfExists("NOMEFANTASIA", "Nome Produto", 300);
                SetColumnIfExists("PRECO", "Preço", 60);
                SetColumnIfExists("CUSTO", "Custo", 60);
                SetColumnIfExists("MARGEM", "Margem", 60);
                SetColumnIfExists("ADIC_FINANC", "Adcional Financeiro", 100);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void SetColumnIfExists(string columnName, string headerText, int width)
        {
            if (_grid.Columns.Contains(columnName))
            {
                var col = _grid.Columns[columnName];
                col.HeaderText = headerText;
                col.Width = width;
            }
        }

        private void ApplyFilter()
        {
            try
            {
                if (_itemsTable == null) return;
                var filterText = _filterBox.Text?.ToUpperInvariant() ?? string.Empty;
                if (string.IsNullOrEmpty(filterText))
                {
                    _itemsBinding.RemoveFilter();
                }
                else
                {
                    var dv = _service.FilterItemsByName(_itemsTable, filterText);
                    _itemsBinding.DataSource = dv;
                    _grid.DataSource = _itemsBinding;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível obter filtro!\nEntre em contato com o suporte!", "ERRO!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BeginEditSelected()
        {
            try
            {
                if (_grid.SelectedRows == null || _grid.SelectedRows.Count == 0) return;
                var row = _grid.SelectedRows[0];
                var values = row.DataBoundItem as DataRowView;
                if (values == null) return;

                _exportButton.Visible = false;
                _cancelToMenuButton.Visible = false;
                _tableCombo.Enabled = false;
                _filterBox.Visible = false;
                _filterLabel.Visible = false;

                _codeBox.Text = values.Row.Field<string>("CODIGOPRD") ?? string.Empty;
                _nameBox.Text = values.Row.Field<string>("NOMEFANTASIA") ?? string.Empty;
                _priceBox.Text = Convert.ToString(values.Row["PRECO"] ?? string.Empty);
                _costBox.Text = Convert.ToString(values.Row["CUSTO"] ?? string.Empty);
                _marginBox.Text = Convert.ToString(values.Row["MARGEM"] ?? string.Empty);
                _addfBox.Text = Convert.ToString(values.Row["ADIC_FINANC"] ?? string.Empty);

                ShowEditWidgets(true);

                _grid.Enabled = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível selecionar item para edição.\nEntre em contato com o suporte!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowEditWidgets(bool show)
        {
            _codeLabel.Visible = show;
            _codeBox.Visible = show;
            _nameLabel.Visible = show;
            _nameBox.Visible = show;
            _priceLabel.Visible = show;
            _priceBox.Visible = show;
            _costLabel.Visible = show;
            _costBox.Visible = show;
            _marginLabel.Visible = show;
            _marginBox.Visible = show;
            _addfLabel.Visible = show;
            _addfBox.Visible = show;
            _saveButton.Visible = show;
            _cancelButton.Visible = show;
        }

        private void SaveChanges()
        {
            try
            {
                if (_grid.SelectedRows == null || _grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Nenhum item selecionado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var rowView = _grid.SelectedRows[0].DataBoundItem as DataRowView;
                if (rowView == null) return;

                var idprd = rowView.Row["IDPRD"];

                var preco = _priceBox.Text.Trim();
                var custo = _costBox.Text.Trim();
                var margem = _marginBox.Text.Trim();
                var adc = _addfBox.Text.Trim();

                // Clear edit fields in UI
                _priceBox.Text = string.Empty;
                _costBox.Text = string.Empty;
                _marginBox.Text = string.Empty;
                _addfBox.Text = string.Empty;

                var affected = _service.UpdateItem(_selectedTableId, idprd, preco, custo, margem, adc);
                Debug.WriteLine($"Rows updated: {affected}");

                MessageBox.Show("Alterações salvas com sucesso!", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                ShowEditWidgets(false);
                _grid.Enabled = true;
                _tableCombo.Enabled = true;
                _exportButton.Visible = true;
                _cancelToMenuButton.Visible = true;

                // Reload items to reflect change
                OnTableSelected();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível efetuar as alterações.\nEntre em contato com o suporte!", "Erro Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Export()
        {
            try
            {
                if (_itemsTable == null || _itemsTable.Rows.Count == 0)
                {
                    MessageBox.Show("Nada para exportar", "Exportação Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var fileBase = _tablesTable.Rows[_selectedTableIndex]["NOME"].ToString();
                var filename = fileBase + ".csv"; // Write CSV for portability
                var fullPath = Path.Combine(Directory.GetCurrentDirectory(), filename);

                _service.ExportToCsv(_itemsTable, fullPath);

                MessageBox.Show("Tabela exportada com sucesso!", "Exportação Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);

                var hardcodedPath = Path.Combine("C:", "Users", "admin", "Desktop", "Adriano", "Projetos_SW", "TABELA DE PREÇO", filename);

                try
                {
                    if (File.Exists(hardcodedPath))
                    {
                        Process.Start(new ProcessStartInfo { FileName = hardcodedPath, UseShellExecute = true });
                    }
                    else
                    {
                        Process.Start(new ProcessStartInfo { FileName = fullPath, UseShellExecute = true });
                    }
                }
                catch (Exception openEx)
                {
                    Debug.WriteLine(openEx);
                    // Not critical; user already informed
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Não foi possível exportar a tabela.\nEntre em contato com o suporte!", "Exportação Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CancelEdit()
        {
            try
            {
                _exportButton.Visible = true;
                _cancelToMenuButton.Visible = true;
                _filterBox.Visible = true;
                _filterLabel.Visible = true;

                ShowEditWidgets(false);
                _codeBox.Text = string.Empty;
                _nameBox.Text = string.Empty;
                _priceBox.Text = string.Empty;
                _costBox.Text = string.Empty;
                _marginBox.Text = string.Empty;
                _addfBox.Text = string.Empty;

                _grid.Enabled = true;
                _tableCombo.Enabled = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show("Entre em contato com o suporte!", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CancelToMenu()
        {
            _tableCombo.Enabled = true;
            _grid.DataSource = null;
            _itemsTable = new DataTable();
            _itemsBinding.DataSource = null;
            _grid.Rows.Clear();
            _grid.Visible = false;
            _exportButton.Visible = false;
            _cancelToMenuButton.Visible = false;
            _vScroll.Visible = false;
            _filterBox.Visible = false;
            _filterLabel.Visible = false;

            _tableCombo.Visible = true;
            _grid.Visible = true;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            try
            {
                (_repository as IDisposable)?.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        [STAThread]
        public static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var form = new PriceTableForm(args);
            Application.Run(form);
        }
    }

    // Service layer: business rules and orchestration
    public class PriceTableService
    {
        private readonly IPriceTableRepository _repository;

        public PriceTableService(IPriceTableRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public DataTable FetchPriceTables()
        {
            return _repository.GetPriceTables();
        }

        public DataTable FetchItemsForTable(object tableId)
        {
            return _repository.GetItemsByTableId(tableId);
        }

        public int UpdateItem(object tableId, object idPrd, string preco, string custo, string margem, string adc)
        {
            return _repository.UpdateItem(idPrd, tableId, preco, custo, margem, adc);
        }

        public DataView FilterItemsByName(DataTable itemsTable, string nameFilterUpper)
        {
            if (itemsTable == null) throw new ArgumentNullException(nameof(itemsTable));

            var dv = itemsTable.DefaultView;
            var escaped = SqlPriceTableRepository.EscapeLikeValue(nameFilterUpper);
            dv.RowFilter = $"CONVERT(NOMEFANTASIA, 'System.String') LIKE '%{escaped}%'";
            return dv;
        }

        public void ExportToCsv(DataTable table, string path)
        {
            CsvExporter.WriteDataTableToCsv(table, path);
        }
    }

    // Repository: data access
    public interface IPriceTableRepository : IDisposable
    {
        DataTable GetPriceTables();
        DataTable GetItemsByTableId(object tableId);
        int UpdateItem(object idPrd, object idTab, string preco, string custo, string margem, string adc);
    }

    public class SqlPriceTableRepository : IPriceTableRepository
    {
        private readonly SqlConnection _connection;
        private bool _disposed;

        public SqlPriceTableRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString)) throw new ArgumentException("connectionString");
            _connection = new SqlConnection(connectionString);
            _connection.Open();
        }

        public DataTable GetPriceTables()
        {
            const string sql =
                "SELECT ZTC.IDTABPRECO, NOME from ZA_TTABPRECO ZTC " +
                "LEFT JOIN TTABPRECO TTP (NOLOCK) ON TTP.CODCOLIGADA = ZTC.CODCOLIGADA AND TTP.IDTABPRECO = ZTC.IDTABPRECO " +
                "where USADEFAULTABELA = 'N' and " +
                "TTP.IDTABPRECO > 3 AND TTP.ATIVA = 1 AND " +
                "CONVERT(VARCHAR(10) , GETDATE() , 126) >= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAINI , 126) AND" +
                "CONVERT(VARCHAR(10) , GETDATE() , 126) <= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAFIM , 126)";

            var table = new DataTable();
            using (var adapter = new SqlDataAdapter(sql, _connection))
            {
                adapter.Fill(table);
            }

            return table;
        }

        public DataTable GetItemsByTableId(object tableId)
        {
            if (tableId == null) throw new ArgumentNullException(nameof(tableId));

            var sql =
                "SELECT  IDTABPRECO, ZTC.IDPRD , CODIGOPRD, NOMEFANTASIA, PRECO, CUSTO, MARGEM, ADIC_FINANC  from ZA_TTABPRECOITM ZTC " +
                "LEFT JOIN TPRD (NOLOCK) ON TPRD.CODCOLIGADA = ZTC.CODCOLIGADA AND TPRD.IDPRD = ZTC.IDPRD " +
                "where IDTABPRECO = @idtab ORDER BY NOMEFANTASIA";

            var table = new DataTable();
            using (var cmd = new SqlCommand(sql, _connection))
            {
                cmd.Parameters.AddWithValue("@idtab", tableId);
                using (var adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(table);
                }
            }

            return table;
        }

        public int UpdateItem(object idPrd, object idTab, string preco, string custo, string margem, string adc)
        {
            if (idPrd == null) throw new ArgumentNullException(nameof(idPrd));
            if (idTab == null) throw new ArgumentNullException(nameof(idTab));

            const string updateSql =
                "UPDATE ZA_TTABPRECOITM SET PRECO = @preco, CUSTO = @custo, MARGEM = @margem, ADIC_FINANC = @adc " +
                "WHERE CODCOLIGADA = 5 AND IDPRD = @idprd and IDTABPRECO = @idtab";

            using (var cmd = new SqlCommand(updateSql, _connection))
            {
                AddSqlParameterWithNullableDecimal(cmd, "@preco", preco);
                AddSqlParameterWithNullableDecimal(cmd, "@custo", custo);
                AddSqlParameterWithNullableDecimal(cmd, "@margem", margem);
                AddSqlParameterWithNullableDecimal(cmd, "@adc", adc);

                cmd.Parameters.AddWithValue("@idprd", idPrd);
                cmd.Parameters.AddWithValue("@idtab", idTab);

                return cmd.ExecuteNonQuery();
            }
        }

        internal static string EscapeLikeValue(string value)
        {
            if (value == null) return string.Empty;
            return value.Replace("%", "[%]").Replace("[", "[[]");
        }

        private static void AddSqlParameterWithNullableDecimal(SqlCommand cmd, string paramName, string value)
        {
            if (decimal.TryParse(value?.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dec))
            {
                cmd.Parameters.AddWithValue(paramName, dec);
            }
            else if (string.IsNullOrWhiteSpace(value))
            {
                cmd.Parameters.AddWithValue(paramName, DBNull.Value);
            }
            else
            {
                cmd.Parameters.AddWithValue(paramName, value);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            try
            {
                _connection?.Close();
                _connection?.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            _disposed = true;
        }
    }

    // CSV Exporter (single responsibility)
    public static class CsvExporter
    {
        public static void WriteDataTableToCsv(DataTable table, string filePath)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("filePath");

            using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                var columnNames = table.Columns.Cast<DataColumn>().Select(c => EscapeCsv(c.ColumnName));
                writer.WriteLine(string.Join(";", columnNames));

                foreach (DataRow row in table.Rows)
                {
                    var fields = table.Columns.Cast<DataColumn>().Select(c => EscapeCsv(Convert.ToString(row[c]) ?? string.Empty));
                    writer.WriteLine(string.Join(";", fields));
                }
            }
        }

        private static string EscapeCsv(string s)
        {
            if (s.Contains("\"")) s = s.Replace("\"", "\"\"");
            if (s.Contains(";") || s.Contains("\n") || s.Contains("\r") || s.Contains("\""))
            {
                return "\"" + s + "\"";
            }

            return s;
        }
    }
}
