using InformeMensualApp.Modelo;
using InformeMensualApp.Negocio;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InformeMensualApp
{
    public partial class MainFormulario : Form
    {
        public MainFormulario()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InfoTKT infoTKT = new InfoTKT();

            //configurar forma para levantar ruta del excel
            string rutaExcel = "";

            DateTime fechaDesde = new DateTime(2024, 1, 1);
            DateTime fechaHasta = new DateTime(2024, 1, 31);

            List<Ticket> listaTickets = infoTKT.LeerExcelYCargarModelo(rutaExcel);
            List<Ticket> filtro = infoTKT.FiltrarTickets(listaTickets, fechaDesde, fechaHasta);
             
        }
    }
}
