using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;


namespace InformeMensualApp
{
    public partial class ResultadosForm : Form
    {
        private object resultados;

        public ResultadosForm(Dictionary<string, int> ticketsPorWebService)
        {
            InitializeComponent();

            // Configura el DataGridView
            dataGridView1.DataSource = new BindingSource(resultados, null);
        }
    }
}
