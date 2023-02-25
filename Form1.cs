using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace documentReading
{
    public partial class Home : Form
    {

        public DataTable dt = new DataTable();
        public DataTable dataTableGroup = new DataTable();
        public Dictionary<string, double > totalByGroup = new Dictionary<string, double >();
        public List<string> stringsGroups = new List<string>();

        public DataTable dataTableAge = new DataTable();
        public Dictionary<string, double> totalByAge = new Dictionary<string, double>();
        public List<string> stringsAge = new List<string>();

        public double total = 0;
        public Home()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button_load(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos de Excel|*.xlsx;*.xls|Todos los archivos|*.*";
            openFileDialog.Title = "Selecciona un archivo de Excel";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;

                // Abre el archivo de Excel
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excel.Workbooks.Open(fileName);
                Worksheet worksheet = workbook.ActiveSheet;

                // Lee los datos de las celdas
                Range range = worksheet.UsedRange;
                dt.Rows.Clear();
                dt.Columns.Clear();
                dt.Clear();
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    // Agrega el nombre de cada columna a la tabla
                    string columnName = ((Range)range.Cells[1, j]).Value2.ToString();
                    dt.Columns.Add(columnName, typeof(string));

                }

                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    // Agrega cada fila a la tabla
                    DataRow row = dt.NewRow();
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        
                        if ( j == 9 )
                        {
                            // Lee el valor de la celda y lo agrega a la fila
                            string cellValue = ((Range)range.Cells[i, j]).Value2.ToString();
                            row[j - 1] = cellValue;

                            if (totalByGroup.ContainsKey(cellValue))
                            {
                                double venta = (((Range)range.Cells[i, j + 1]).Value2);
                                totalByGroup[cellValue] += venta;
                            }
                            else
                            {
                                double venta = (((Range)range.Cells[i, j + 1]).Value2);
                                totalByGroup.Add(cellValue, venta);
                                stringsGroups.Add(cellValue);
                            }
                        }
                        else if ( j == 8)
                        {
                            // Lee el valor de la celda y lo agrega a la fila
                            double cellValue = ((Range)range.Cells[i, j]).Value2;
                            DateTime dateTime = DateTime.FromOADate(cellValue);
                            string data = dateTime.ToString("dd/MM/yyyy HH:mm:ss");
                            row[j - 1] = data;

                            string age = dateTime.Year.ToString();  
                            if (totalByAge.ContainsKey(age))
                            {
                                double venta = (((Range)range.Cells[i, j + 2]).Value2);
                                totalByAge[age] += venta;
                            }
                            else
                            {
                                double venta = (((Range)range.Cells[i, j + 2]).Value2);
                                totalByAge.Add(age, venta);
                                stringsGroups.Add(age);
                            }
                        }
                        else if ( j == 3)
                        {
                            // Lee el valor de la celda y lo agrega a la fila
                            double cellValue = ((Range)range.Cells[i, j]).Value2;
                            DateTime dateTime = DateTime.FromOADate(cellValue);
                            string data = dateTime.ToString("dd/MM/yyyy HH:mm:ss");
                            row[j - 1] = data;
                        }
                        else
                        {

                            // Lee el valor de la celda y lo agrega a la fila
                            string cellValue = ((Range)range.Cells[i, j]).Value2.ToString();
                            row[j - 1] = cellValue;

                        }

                    }
                    dt.Rows.Add(row);

                    // Asigna la tabla al componente dataView
                    dataView.DataSource = dt;

                    label3.Text = "Registros " + dt.Rows.Count.ToString();

                    label2.Text = "Cargando....";
                }

                //Total
                this.total = dt.AsEnumerable()
                    .Sum(x => Convert.ToDouble(x["Vendido"]));

                label2.Text = "Total: Q " + this.total.ToString();

                // Cierra el archivo de Excel
                workbook.Close();
                excel.Quit();

                MessageBox.Show("Finalizado con exito :))))))");
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataTableGroup.Rows.Count > 0)
            {
                dataView.DataSource = dataTableGroup;
            }
            else
            {
                dataTableGroup.Columns.Add("Grupo", typeof(string));
                dataTableGroup.Columns.Add("Total ventas", typeof(double));

                foreach (KeyValuePair<string, double> entry in this.totalByGroup)
                {
                    DataRow row = dataTableGroup.NewRow();
                    row[0] = entry.Key;
                    row[1] = entry.Value;
                    dataTableGroup.Rows.Add(row);
                }

                dataView.DataSource = dataTableGroup;
            }

            label2.Text = "";
            label3.Text = "Registros " + dataTableGroup.Rows.Count.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataView.DataSource = dt;

            //Total
            this.total = dt.AsEnumerable()
                .Sum(x => Convert.ToDouble(x["Vendido"]));

            label2.Text = "Total: Q " + this.total.ToString();

            label3.Text = "Registros " + dt.Rows.Count.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataTableAge.Rows.Count > 0)
            {
                dataView.DataSource = dataTableAge;
            }
            else
            {
                dataTableAge.Columns.Add("Año", typeof(string));
                dataTableAge.Columns.Add("Total ventas", typeof(double));

                foreach (KeyValuePair<string, double> entry in this.totalByAge)
                {
                    DataRow row = dataTableAge.NewRow();
                    row[0] = entry.Key;
                    row[1] = entry.Value;
                    dataTableAge.Rows.Add(row);
                }

                dataView.DataSource = dataTableAge;
            }

            label2.Text = "";
            label3.Text = "Registros " + dataTableAge.Rows.Count.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string textSearch = textBox1.Text;
            textBox1.Clear();

            if (textSearch.Length > 0)
            {
                // Nombre de la columna en la que buscar
                string nameColumFilter = "Nombre";

                // Verificar si la columna existe en la tabla
                if (!dt.Columns.Contains(nameColumFilter))
                {
                    MessageBox.Show($"La columna {nameColumFilter} no existe en la tabla");
                    return;
                }

                // Buscamos en la columna indicada los datos que contengan el texto introducido en el textbox
                DataRow[] rows = dt.Select($"{nameColumFilter} LIKE '%{textSearch}%'");

                // Crear la tabla que se mostrará en el DataView
                DataTable dtSearch = new DataTable();

                // Crear las columnas de la tabla
                foreach (DataColumn col in dt.Columns)
                {
                    dtSearch.Columns.Add(col.ColumnName, col.DataType);
                }

                // Añadir las filas a la tabla de búsqueda
                foreach (DataRow row in rows)
                {
                    dtSearch.ImportRow(row);
                }

                // Mostrar los datos en el DataView
                dataView.DataSource = dtSearch;

                //Total
                this.total = dtSearch.AsEnumerable()
                    .Sum(x => Convert.ToDouble(x["Vendido"]));

                label2.Text = "Total: Q " + this.total.ToString();

                label3.Text = "Registros " + dtSearch.Rows.Count.ToString();
            }
        }
        private void dataView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            Console.WriteLine("Valor de la celda: " + dataView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()
                + " - Fila: " + e.RowIndex
                + " - Columna: " + e.ColumnIndex
            );
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
