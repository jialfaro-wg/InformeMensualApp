using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InformeMensualApp.Modelo
{
    class Ticket
    {
        public string NroTicket { get; set; }
        public string TipoTicket { get; set; }
        public DateTime FechaAbierto { get; set; }
        public DateTime FechaCierre { get; set; }
        public string Estado { get; set; }  
        public string WebService { get; set; }
        //public string Problema { get; set; }
        //public string WebService { get; set; }
        


        //	Tipo-TKT	N°TK	Fecha-inicio	Fecha-cierre				Paìs			vinculo a tarjeta de Teams	Name WS									

    }
}
