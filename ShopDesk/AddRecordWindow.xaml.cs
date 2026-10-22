using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace ShopDesk
{
    public partial class AddRecordWindow : Window
    {
        // Храним object, чтобы не терять типы (int, DateTime и т.д.)
        public Dictionary<string, object> Values { get; } = new();
        public Dictionary<string, TextBox> Boxes { get; } = new();

        public AddRecordWindow(DataTable schema,
                               HashSet<string> readOnlyColumns,
                               Dictionary<string, object> initialValues = null)
        {
            InitializeComponent();

            foreach (DataColumn col in schema.Columns)
            {
                if (readOnlyColumns.Contains(col.ColumnName)) continue;

                var sp = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 4, 0, 4)
                };

                sp.Children.Add(new TextBlock
                {
                    Text = col.ColumnName,
                    Width = 130,
                    VerticalAlignment = VerticalAlignment.Center
                });

                var tb = new TextBox { Width = 240 };
                if (initialValues != null &&
                    initialValues.TryGetValue(col.ColumnName, out var v) &&
                    v != null && v != System.DBNull.Value)
                {
                    tb.Text = v.ToString();
                }

                Boxes[col.ColumnName] = tb;
                sp.Children.Add(tb);
                FieldsPanel.Children.Add(sp);
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            foreach (var kv in Boxes)
                Values[kv.Key] = kv.Value.Text;
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}