using System;

namespace PruebaNominas
{
    class Empleados
    {
        public int Id { set; get; }
        public DateTime FechaIngreso { set; get; }
        public DateTime FechaSalida { set; get; }
        public string Cedula { set; get; }
        public string Apellido { set; get; }
        public string Nombre { set; get; }
        public double Sueldo_Mensual { set; get; }
        public string Area { set; get; }
        public bool Vacaciones { set; get; } = false;
        public bool CalcularDiasIESS { set; get; } = false;
    }

    class EmpleadosQ
    {
        public int Id { set; get; }
        public string Area { set; get; }
        public string Apellido { set; get; }
        public string Nombre { set; get; }
        public int DiasTrabajados { set; get; }
        //Ingresos
        public int Alim { set; get; } //Alimentacion I
        public int Transp { set; get; }
        public int Bono { set; get; }
        public int TarjetaIngresos { set; get; }
        public int HorasExtra { set; get; }
        public int Vacaciones { set; get; }
        public bool FondosReserva { set; get; }
        public bool DecimoTercero { set; get; }
        public bool DecimoCuarto { set; get; }
        //Egresos
        public int PrestHipot { set; get; }
        public int PrestQuiro { set; get; }
        public int PrestCia { set; get; }
        public int Multas { set; get; }
        public int ExtSalud { set; get; }
        public int TarjetaEgresos { set; get; }
        public int ContribucionSolidaria { set; get; }
        public int AnticipoQuincena { set; get; }
    }
}
