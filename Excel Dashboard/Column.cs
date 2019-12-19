using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Dashboard
{
    public class Column
    {
        private string folio;
        private string ticket;
        private string nombreCliente;
        private string zona;
        private string unidad;
        private string chofer;
        private string salida;
        private string estatus;
        private string estatusCargando;
        private string estatusTrayecto;
        private string estatusEntregado;
        private string excelOrigen;

        public string Folio
        {
            get { return folio; }
            set { folio = value; }
        }

        public string ExcelOrigen
        {
            get { return excelOrigen; }
            set { excelOrigen = value; }
        }
        public string Ticket
        {
            get { return ticket; }
            set { ticket = value; }
        }

        public string NombreCliente
        {
            get { return nombreCliente; }
            set { nombreCliente = value; }
        }

        public string Zona
        {
            get { return zona; }
            set { zona = value; }
        }

        public string Unidad
        {
            get { return unidad; }
            set { unidad = value; }
        }

        public string Chofer
        {
            get { return chofer; }
            set { chofer = value; }
        }

        public string Salida
        {
            get { return salida; }
            set { salida = value; }
        }

        public string EstatusCargando
        {
            get { return estatusCargando; }
            set { estatusCargando = value; }
        }

        public string EstatusTrayecto
        {
            get { return estatusTrayecto; }
            set { estatusTrayecto = value; }
        }

        public string EstatusEntregado
        {
            get { return estatusEntregado; }
            set { estatusEntregado = value; }
        }

        public string Estatus
        {
            get { return estatus; }
            set { estatus = value; }
        }
    }
}
