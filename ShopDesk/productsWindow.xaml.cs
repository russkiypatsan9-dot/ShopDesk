using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

using ClosedXML.Excel;

namespace ShopDesk
{
    public partial class productsWindow : Window
    {
        private static readonly string ConnectionString =
            $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\repos\ShopDesk\ShopDesk\database\ShopDesk.accdb";

        private OleDbDataAdapter _productsAdapter, _categoryAdapter, _suppliersAdapter;
        private DataTable _productsTable, _categoryTable, _suppliersTable;

        // ============================================================
        //  НАСТРОЙКА СПРАВОЧНИКОВ (ComboBox) ПО ВАШЕЙ СХЕМЕ
        // ============================================================
        private static readonly Dictionary<string, (string Table, string KeyField, string DisplayField)>
            Lookups = new()
            {
                // В products колонка называется "Категория".
                // В category ключ "Код", а текст, вероятно, "Название категории".
                ["Категория"] = ("category", "id", "Название категории"),

                // В products колонка называется "Поставщик".
                // В suppliers ключ "Код", а название фирмы, вероятно, "Поставщик" (или "Название").
                ["Поставщик"] = ("suppliers", "id", "Поставщик")
            };

        private readonly Dictionary<string, List<LookupItem>> _lookupCache = new();

        private class LookupItem
        {
            public object Key { get; init; }
            public string Display { get; init; }
            public override string ToString() => Display;
        }

        private enum EditMode { None, Add, Update }
        private EditMode _editMode = EditMode.None;
        private string _editTableName;
        private DataGrid _editGrid;
        private DataRowView _editingRow;
        private string[] _editPkColumns;
        private HashSet<string> _editReadOnly;
        private readonly Dictionary<string, Control> _editControls = new();

        public productsWindow()
        {
            InitializeComponent();
            ReloadAll();
        }

        // ============================================================
        //  ЗАГРУЗКА
        // ============================================================

        private void LoadTableEditable(string tableName, DataGrid grid,
                                       out OleDbDataAdapter adapter, out DataTable table)
        {
            try
            {
                var connection = new OleDbConnection(ConnectionString);
                adapter = new OleDbDataAdapter($"SELECT * FROM [{tableName}]", connection);
                _ = new OleDbCommandBuilder(adapter)
                {
                    QuotePrefix = "[",
                    QuoteSuffix = "]"
                };
                table = new DataTable();
                adapter.Fill(table);
                grid.ItemsSource = table.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки {tableName}: {ex.Message}");
                adapter = null;
                table = null;
            }
        }

        private void ReloadAll()
        {
            LoadTableEditable("products", ProductsGrid, out _productsAdapter, out _productsTable);
            LoadTableEditable("category", CategoryGrid, out _categoryAdapter, out _categoryTable);
            LoadTableEditable("suppliers", SuppliersGrid, out _suppliersAdapter, out _suppliersTable);
            _lookupCache.Clear();
        }

        // ============================================================
        //  СХЕМА (С ЗАЩИТОЙ ОТ ОШИБКИ "НЕ ЗАДАН ПЕРВИЧНЫЙ КЛЮЧ")
        // ============================================================

        private (string[] pkColumns, DataTable schema, HashSet<string> autoInc)
            GetSchemaInfo(string tableName)
        {
            using var connection = new OleDbConnection(ConnectionString);
            connection.Open();

            var pk = new List<string>();
            var autoInc = new HashSet<string>();
            var allColumns = new List<string>();

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = $"SELECT * FROM [{tableName}]";
                using var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                var schema = reader.GetSchemaTable();

                if (schema != null)
                {
                    foreach (DataRow row in schema.Rows)
                    {
                        string colName = row["ColumnName"].ToString();
                        allColumns.Add(colName);

                        if (row["IsKey"] is bool isKey && isKey)
                            pk.Add(colName);

                        if (row["IsAutoIncrement"] is bool a && a)
                            autoInc.Add(colName);
                    }
                }
            }

            // --- ЖЕЛЕЗОБЕТОННАЯ СТРАХОВКА ---
            if (pk.Count == 0)
            {
                // 1. Ищем колонку с именем "Код" или "id" в списке автоинкрементов
                var fallbackKey = autoInc.FirstOrDefault(c =>
                    c.Equals("Код", StringComparison.OrdinalIgnoreCase) ||
                    c.Equals("id", StringComparison.OrdinalIgnoreCase));

                // 2. Если не нашли, берем первую колонку из автоинкрементов
                if (fallbackKey == null && autoInc.Count > 0)
                {
                    fallbackKey = autoInc.First();
                }

                // 3. Если автоинкрементов вообще нет, берем просто первую колонку таблицы
                if (fallbackKey == null && allColumns.Count > 0)
                {
                    fallbackKey = allColumns.First();
                }

                if (fallbackKey != null)
                {
                    pk.Add(fallbackKey);
                }
            }

            var dtSchema = new DataTable();
            using (var adapter = new OleDbDataAdapter($"SELECT * FROM [{tableName}] WHERE 1=0", connection))
                adapter.FillSchema(dtSchema, SchemaType.Source);

            return (pk.ToArray(), dtSchema, autoInc);
        }

        // ============================================================
        //  СПРАВОЧНИКИ
        // ============================================================

        private List<LookupItem> GetLookupItems(string columnName)
        {
            if (_lookupCache.TryGetValue(columnName, out var cached))
                return cached;

            if (!Lookups.TryGetValue(columnName, out var cfg))
                return null;

            var items = new List<LookupItem>();
            try
            {
                using var connection = new OleDbConnection(ConnectionString);
                connection.Open();
                using var cmd = connection.CreateCommand();
                cmd.CommandText = $"SELECT [{cfg.KeyField}], [{cfg.DisplayField}] " +
                                  $"FROM [{cfg.Table}] ORDER BY [{cfg.DisplayField}]";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    items.Add(new LookupItem
                    {
                        Key = reader[0],
                        Display = reader[1]?.ToString() ?? ""
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки справочника {cfg.Table}: {ex.Message}");
            }

            _lookupCache[columnName] = items;
            return items;
        }

        // ============================================================
        //  РЕДАКТОР
        // ============================================================

        private void OpenEditor(EditMode mode, string tableName, DataGrid grid)
        {
            try
            {
                var (pkCols, schema, autoInc) = GetSchemaInfo(tableName);

                if (pkCols.Length == 0)
                {
                    MessageBox.Show($"В таблице '{tableName}' не удалось определить первичный ключ.\n" +
                                    "Убедитесь, что в Access колонка 'Код' назначена первичным ключом.");
                    return;
                }

                _editMode = mode;
                _editTableName = tableName;
                _editGrid = grid;
                _editPkColumns = pkCols;
                _editControls.Clear();
                FieldsPanel.Children.Clear();

                DataRowView selectedRow = null;
                var initialValues = new Dictionary<string, object>();

                if (mode == EditMode.Update)
                {
                    if (grid.SelectedItem is not DataRowView row)
                    {
                        MessageBox.Show("Выберите строку для изменения.");
                        return;
                    }
                    selectedRow = row;

                    foreach (DataColumn col in schema.Columns)
                        initialValues[col.ColumnName] = row[col.ColumnName];

                    _editReadOnly = new HashSet<string>(pkCols);
                    foreach (var a in autoInc) _editReadOnly.Add(a);
                    EditorTitle.Text = $"Изменение записи — {tableName}";
                }
                else
                {
                    _editReadOnly = new HashSet<string>(autoInc);
                    EditorTitle.Text = $"Добавление записи — {tableName}";
                }

                _editingRow = selectedRow;

                foreach (DataColumn col in schema.Columns)
                {
                    if (_editReadOnly.Contains(col.ColumnName)) continue;

                    var sp = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Margin = new Thickness(0, 3, 0, 3)
                    };
                    sp.Children.Add(new TextBlock
                    {
                        Text = col.ColumnName,
                        Width = 140,
                        VerticalAlignment = VerticalAlignment.Center
                    });

                    Control input;

                    // ГЛАВНОЕ ИСПРАВЛЕНИЕ: проверяем, что таблица именно products
                    if (tableName == "products" && Lookups.ContainsKey(col.ColumnName))
                    {
                        var cb = new ComboBox
                        {
                            Width = 300,
                            IsEditable = true,
                            IsTextSearchEnabled = true,
                            DisplayMemberPath = nameof(LookupItem.Display),
                            SelectedValuePath = nameof(LookupItem.Key)
                        };
                        var items = GetLookupItems(col.ColumnName) ?? new List<LookupItem>();
                        foreach (var it in items) cb.Items.Add(it);

                        if (initialValues.TryGetValue(col.ColumnName, out var cur) &&
                            cur != null && cur != DBNull.Value)
                        {
                            var match = items.FirstOrDefault(
                                it => Equals(Normalize(it.Key), Normalize(cur)));
                            if (match != null) cb.SelectedItem = match;
                        }

                        input = cb;
                    }
                    else
                    {
                        var tb = new TextBox { Width = 300 };
                        if (initialValues.TryGetValue(col.ColumnName, out var v) &&
                            v != null && v != DBNull.Value)
                        {
                            tb.Text = v.ToString();
                        }
                        input = tb;
                    }

                    _editControls[col.ColumnName] = input;
                    sp.Children.Add(input);
                    FieldsPanel.Children.Add(sp);
                }

                EditorPanel.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка открытия редактора: {ex.Message}");
            }
        }

        private static object Normalize(object v) =>
            v is null or DBNull ? null : v.ToString();

        private void CloseEditor()
        {
            EditorPanel.Visibility = Visibility.Collapsed;
            FieldsPanel.Children.Clear();
            _editControls.Clear();
            _editMode = EditMode.None;
            _editTableName = null;
            _editGrid = null;
            _editingRow = null;
            _editPkColumns = null;
            _editReadOnly = null;
        }

        private Dictionary<string, object> CollectValues()
        {
            var result = new Dictionary<string, object>();
            foreach (var kv in _editControls)
            {
                if (kv.Value is ComboBox cb)
                {
                    result[kv.Key] = (cb.SelectedItem as LookupItem)?.Key ?? DBNull.Value;
                }
                else if (kv.Value is TextBox tb)
                {
                    result[kv.Key] = string.IsNullOrEmpty(tb.Text)
                        ? DBNull.Value
                        : tb.Text;
                }
            }
            return result;
        }

        private void EditorOk_Click(object sender, RoutedEventArgs e)
        {
            // Проверка на заполнение выпадающих списков
            foreach (var kv in _editControls)
            {
                if (kv.Value is ComboBox cb && cb.SelectedItem == null && _editMode == EditMode.Add)
                {
                    MessageBox.Show($"Выберите значение для «{kv.Key}».");
                    return;
                }
            }

            try
            {
                var values = CollectValues();

                if (_editMode == EditMode.Add)
                    DoInsert(values);
                else if (_editMode == EditMode.Update)
                    DoUpdate(values);

                CloseEditor();
                ReloadAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        private void EditorCancel_Click(object sender, RoutedEventArgs e) => CloseEditor();

        private void DoInsert(Dictionary<string, object> values)
        {
            var cols = new List<string>();
            var parms = new List<string>();
            var args = new List<object>();

            foreach (var kv in values)
            {
                cols.Add($"[{kv.Key}]");
                parms.Add("?");
                args.Add(kv.Value);
            }
            if (cols.Count == 0) return;

            string sql = $"INSERT INTO [{_editTableName}] ({string.Join(",", cols)}) " +
                         $"VALUES ({string.Join(",", parms)})";

            using var connection = new OleDbConnection(ConnectionString);
            connection.Open();
            using var cmd = connection.CreateCommand();
            cmd.CommandText = sql;
            foreach (var v in args) cmd.Parameters.AddWithValue("?", v);
            cmd.ExecuteNonQuery();
        }

        private void DoUpdate(Dictionary<string, object> values)
        {
            var setParts = new List<string>();
            var setValues = new List<object>();

            foreach (var kv in values)
            {
                setParts.Add($"[{kv.Key}] = ?");
                setValues.Add(kv.Value);
            }
            if (setParts.Count == 0) return;

            var whereParts = new List<string>();
            var whereValues = new List<object>();
            foreach (var pk in _editPkColumns)
            {
                whereParts.Add($"[{pk}] = ?");
                whereValues.Add(_editingRow[pk]);
            }

            string sql = $"UPDATE [{_editTableName}] SET {string.Join(",", setParts)} " +
                         $"WHERE {string.Join(" AND ", whereParts)}";

            using var connection = new OleDbConnection(ConnectionString);
            connection.Open();
            using var cmd = connection.CreateCommand();
            cmd.CommandText = sql;
            foreach (var v in setValues) cmd.Parameters.AddWithValue("?", v);
            foreach (var v in whereValues) cmd.Parameters.AddWithValue("?", v);
            cmd.ExecuteNonQuery();
        }

        // ============================================================
        //  УДАЛЕНИЕ / СОХРАНЕНИЕ
        // ============================================================

        private void DeleteRecord(string tableName, DataGrid grid)
        {
            if (grid.SelectedItem is not DataRowView row)
            {
                MessageBox.Show("Выберите строку для удаления.");
                return;
            }

            if (MessageBox.Show("Удалить выбранную запись?", "Подтверждение",
                    MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                var (pkCols, _, _) = GetSchemaInfo(tableName);
                if (pkCols.Length == 0)
                {
                    MessageBox.Show("В таблице нет PK.");
                    return;
                }

                var whereParts = new List<string>();
                var whereValues = new List<object>();
                foreach (var pk in pkCols)
                {
                    whereParts.Add($"[{pk}] = ?");
                    whereValues.Add(row[pk]);
                }

                string sql = $"DELETE FROM [{tableName}] WHERE {string.Join(" AND ", whereParts)}";

                using var connection = new OleDbConnection(ConnectionString);
                connection.Open();
                using var cmd = connection.CreateCommand();
                cmd.CommandText = sql;
                foreach (var v in whereValues) cmd.Parameters.AddWithValue("?", v);
                cmd.ExecuteNonQuery();

                ReloadAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка удаления: {ex.Message}\n" +
                                "Возможно, на запись ссылаются другие таблицы.");
            }
        }

        private void SaveTable(OleDbDataAdapter adapter, DataTable table)
        {
            if (adapter == null || table == null)
            {
                MessageBox.Show("Нет данных для сохранения.");
                return;
            }

            try
            {
                using var connection = new OleDbConnection(ConnectionString);
                connection.Open();
                adapter.SelectCommand.Connection = connection;
                adapter.Update(table);
                MessageBox.Show("Изменения сохранены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        // ============================================================
        //  ЭКСПОРТ
        // ============================================================

        private DataGrid GetActiveGrid() => MainTabs.SelectedIndex switch
        {
            0 => ProductsGrid,
            1 => CategoryGrid,
            2 => SuppliersGrid,
            _ => null
        };

        private string GetActiveTabName() =>
            (MainTabs.SelectedItem as TabItem)?.Header?.ToString() ?? "export";

        private void ExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetActiveGrid();
            if (grid?.ItemsSource is not DataView view || view.Table == null)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "CSV файл (*.csv)|*.csv",
                FileName = GetActiveTabName()
            };
            if (sfd.ShowDialog() != true) return;

            var sb = new StringBuilder();
            var cols = view.Table.Columns;

            for (int i = 0; i < cols.Count; i++)
            {
                sb.Append(Escape(cols[i].ColumnName));
                if (i < cols.Count - 1) sb.Append(';');
            }
            sb.AppendLine();

            foreach (DataRowView row in view)
            {
                for (int i = 0; i < cols.Count; i++)
                {
                    sb.Append(Escape(row[i]?.ToString() ?? ""));
                    if (i < cols.Count - 1) sb.Append(';');
                }
                sb.AppendLine();
            }

            File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
            MessageBox.Show("Экспортировано: " + sfd.FileName);

            static string Escape(string s) =>
                s.Contains(';') || s.Contains('"') || s.Contains('\n')
                    ? "\"" + s.Replace("\"", "\"\"") + "\""
                    : s;
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var grid = GetActiveGrid();
            if (grid?.ItemsSource is not DataView view || view.Table == null)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel файл (*.xlsx)|*.xlsx",
                FileName = GetActiveTabName()
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Данные");

                for (int c = 0; c < view.Table.Columns.Count; c++)
                {
                    var cell = ws.Cell(1, c + 1);
                    cell.Value = view.Table.Columns[c].ColumnName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                for (int r = 0; r < view.Count; r++)
                {
                    var row = view[r];
                    for (int c = 0; c < view.Table.Columns.Count; c++)
                    {
                        var val = row[c];
                        if (val is null || val == DBNull.Value) continue;

                        var cell = ws.Cell(r + 2, c + 1);
                        if (val is int or long or short or byte)
                            cell.Value = Convert.ToInt64(val);
                        else if (val is double or float or decimal)
                            cell.Value = Convert.ToDouble(val);
                        else if (val is DateTime dt)
                            cell.Value = dt;
                        else
                            cell.Value = val.ToString();
                    }
                }

                ws.Columns().AdjustToContents();
                wb.SaveAs(sfd.FileName);
                MessageBox.Show("Экспортировано: " + sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}");
            }
        }

        // ============================================================
        //  КНОПКИ
        // ============================================================

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            switch (MainTabs.SelectedIndex)
            {
                case 0: OpenEditor(EditMode.Add, "products", ProductsGrid); break;
                case 1: OpenEditor(EditMode.Add, "category", CategoryGrid); break;
                case 2: OpenEditor(EditMode.Add, "suppliers", SuppliersGrid); break;
            }
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            switch (MainTabs.SelectedIndex)
            {
                case 0: OpenEditor(EditMode.Update, "products", ProductsGrid); break;
                case 1: OpenEditor(EditMode.Update, "category", CategoryGrid); break;
                case 2: OpenEditor(EditMode.Update, "suppliers", SuppliersGrid); break;
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            switch (MainTabs.SelectedIndex)
            {
                case 0: DeleteRecord("products", ProductsGrid); break;
                case 1: DeleteRecord("category", CategoryGrid); break;
                case 2: DeleteRecord("suppliers", SuppliersGrid); break;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            switch (MainTabs.SelectedIndex)
            {
                case 0: SaveTable(_productsAdapter, _productsTable); break;
                case 1: SaveTable(_categoryAdapter, _categoryTable); break;
                case 2: SaveTable(_suppliersAdapter, _suppliersTable); break;
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e) => ReloadAll();

        private void MainTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_editMode != EditMode.None) CloseEditor();
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Завершить сессию?", "Выход",
                    MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;

            var login = new MainWindow();
            login.Show();
            Close();
        }
    }
}