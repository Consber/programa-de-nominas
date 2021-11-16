using System;

namespace ProgramaNominas
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
        public double Alim { set; get; } //Alimentacion I
        public double Transp { set; get; }
        public double Bono { set; get; }
        public double TarjetaIngresos { set; get; }
        public double HorasExtra { set; get; }
        public double Vacaciones { set; get; }
        public bool FondosReserva { set; get; }
        public bool DecimoTercero { set; get; }
        public bool DecimoCuarto { set; get; }
        //Egresos
        public double PrestHipot { set; get; }
        public double PrestQuiro { set; get; }
        public double PrestCia { set; get; }
        public double Multas { set; get; }
        public double ExtSalud { set; get; }
        public double TarjetaEgresos { set; get; }
        public double ContribucionSolidaria { set; get; }
        public double AnticipoQuincena { set; get; }
    }
}
