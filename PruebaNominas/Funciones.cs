using System;
using System.Windows;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using Microsoft.Office;
using System.Runtime.InteropServices;   //GuidAttribute
using System.Reflection;                //Assembly
using System.Threading;                 //Mutex
using System.Security.AccessControl;    //MutexAccessRule
using System.Security.Principal;        //SecurityIdentifier

namespace PruebaNominas
{
    public partial class MainWindow
    {
        static string strSqlE = @"Select * from Empleados";
        static string strSqlC = @"Select * from Cuentas";
        static string strDB = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source = Empleados.accdb";
        string aBin = @"ciapp.bin";
        bool activo = false;
        bool checkEstilo;
        DateTime? fecha;

        static OleDbConnection con = new OleDbConnection(strDB);


        List<Cuentas> colCuentas;
        List<Empleados> colEmpleados;
        List<EmpleadosQ> colEmpleadosQ1;
        List<EmpleadosQ> colEmpleadosQ2;

        string posError = "Inicio";

        public void Empezar()
        {
            try
            {
                fecha = dtp_Fecha.SelectedDate;
                dtp_Fecha.Text = DateTime.Now.ToString("dd/MM/yyyy");
                fecha = dtp_Fecha.SelectedDate;

                nombreArchivo = Convert.ToDateTime(dtp_Fecha.ToString()).ToString("MMMM") + " - " + Convert.ToDateTime(dtp_Fecha.ToString()).ToString("yyyy") + ".xlsx";
                ruta = ruta + nombreArchivo;
                colCuentas = new List<Cuentas>();
                colEmpleados = new List<Empleados>();
                colEmpleadosQ1 = new List<EmpleadosQ>();
                colEmpleadosQ2 = new List<EmpleadosQ>();

                Lenguage("es");
                AppCerrada(activo);
                activo = true;
                Txtbox_Ruta.Text = ruta;
                contenidoRuta = ruta;
                Archivo(true);
                baseEjm();
                iniciandoPrograma = false;

                but_Actualizacion.IsChecked = Settings1.Default.IgnorarActualizacion;
            }
            catch (Exception c)
            {
                System.Windows.MessageBox.Show("Error en la seccion: " + posError + ". \nMensaje del error: " + c.Message, "Error");
                con.Close();
            }
            finally
            {
                con.Close();
            }
        }

        public void Lenguage(string idioma)
        {
            posError = "Carga de idiomas";
        }

        public void AppCerrada(bool activo)
        {
            char tmp = ' ';
            try
            {
                foreach (char c in File.ReadAllText(aBin))
                {
                    if (c == '0' || c == '1')
                    {
                        tmp = c;
                        continue;
                    }
                }
            }
            catch (Exception e)
            {
                Archivo(true);
            }
            finally
            {
                if (tmp == '1' && !activo)
                {
                    System.Windows.MessageBox.Show("Se cerro inesperadamente", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        public void Archivo(bool estado)
        {
            posError = "Leer Estado";
            using (StreamWriter file =
            new StreamWriter(aBin))
            {
                file.WriteLine(estado.GetHashCode());
            }
        }

        public string comandoSQL(int num, string strng)
        {
            switch (num)
            {
                case 0:
                    return strng;
                case 1:
                    return @"select * from " + strng;
                case 2:
                    return @"select * from Empleados where " + strng;
                default:
                    return @"select * from Empleados";
            }
        }

        public string comandoSQL(int num)
        {
            switch (num)
            {
                case 0:
                    return @"select * from Empleados";
                case 1:
                    return @"select * from Empleados where Area = Administracion";
                case 2:
                    return @"select * from Empleados where Area = Operativo";
                default:
                    return @"select * from Empleados";
            }
        }

        public void baseDatos()
        {
            posError = "Carga de base de datos";
            OleDbDataAdapter dA = new OleDbDataAdapter(strSqlE, con);
            DataSet ds = new DataSet();
            dA.Fill(ds, "[Order]");

        }

        public void baseEjm()
        {
            con.ConnectionString = strDB;
            con.Open();

            OleDbCommand cmdKlassen = new OleDbCommand(strSqlC, con);

            if (con.State == ConnectionState.Open)
            {
                OleDbDataReader KlasReader = null;
                KlasReader = cmdKlassen.ExecuteReader();
                posError = "Sacando archivos base de datos: Cuentas";
                while (KlasReader.Read())
                {
                    colCuentas.Add(new Cuentas()
                    {
                        FormaPago = KlasReader["FormaPago"].ToString(),
                        Banco = KlasReader["Banco"].ToString(),
                        TipoCuentaCheque = KlasReader["TipoCuentaCheque"].ToString(),
                        NumCuenta = KlasReader["NumCuenta"].ToString(),
                        Valor = KlasReader["Valor"].ToString(),
                        Identificacion = KlasReader["Identificacion"].ToString(),
                        TipoDocumento = KlasReader["TipoDocumento"].ToString(),
                        NUC = KlasReader["NUC"].ToString(),
                        Beneficiario = KlasReader["Beneficiario"].ToString(),
                        Telefono = KlasReader["Telefono"].ToString(),
                        Referencia = KlasReader["Referencia"].ToString(),
                        TipoDeRol = KlasReader["TipoDeRol"].ToString(),
                    });
                }

                cmdKlassen = new OleDbCommand(strSqlE, con);

                KlasReader = null;
                KlasReader = cmdKlassen.ExecuteReader();
                int countID = 1;
                while (KlasReader.Read())
                {
                    posError = "Sacando archivos base de datos: Empleados";
                    if (KlasReader["FechaSalida"].ToString() != "")
                    {
                        colEmpleados.Add(new Empleados()
                        { 
                            Id = countID,
                            FechaIngreso = Convert.ToDateTime(KlasReader["FechaIngreso"]),
                            FechaSalida = Convert.ToDateTime(KlasReader["FechaSalida"]),
                            Cedula = KlasReader["Cedula"].ToString(),
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Sueldo_Mensual = Convert.ToDouble(KlasReader["Sueldo Mensual"].ToString()),
                            Area = KlasReader["Area"].ToString()

                        });

                        colEmpleadosQ1.Add(new EmpleadosQ()
                        {
                            Id = countID,
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Area = KlasReader["Area"].ToString()
                        });

                        colEmpleadosQ2.Add(new EmpleadosQ()
                        {
                            Id = countID,
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Area = KlasReader["Area"].ToString()
                        });
                    }
                    else
                    {
                        colEmpleados.Add(new Empleados()
                        {
                            Id = countID,
                            FechaIngreso = Convert.ToDateTime(KlasReader["FechaIngreso"]),
                            Cedula = KlasReader["Cedula"].ToString(),
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Sueldo_Mensual = Convert.ToDouble(KlasReader["Sueldo Mensual"].ToString()),
                            Area = KlasReader["Area"].ToString()
                        });

                        colEmpleadosQ1.Add(new EmpleadosQ()
                        {
                            Id = countID,
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Area = KlasReader["Area"].ToString()
                        });

                        colEmpleadosQ2.Add(new EmpleadosQ()
                        {
                            Id = countID,
                            Apellido = KlasReader["Apellido"].ToString(),
                            Nombre = KlasReader["Nombre"].ToString(),
                            Area = KlasReader["Area"].ToString()
                        });
                    }
                    countID++;
                }
                con.Close();
            }
            else
            {
                System.Windows.MessageBox.Show("La conexion a la base tuvo un problema. Asegurese de que la base este en la carpeta correcta");
            }

            con.Close();
        }

        public void cargarBase(int i)
        {
            try
            {
                List<Cuentas> tmp = new List<Cuentas>();

                switch (i)
                {
                    case 0:
                        break;
                    case 1:
                        foreach (Cuentas c in colCuentas)
                        {
                            if (c.TipoDeRol == "Rol")
                            {
                                tmp.Add(c);
                            }
                        }
                        break;
                    case 2:
                        foreach (Cuentas c in colCuentas)
                        {
                            if (c.TipoDeRol == "Servicio")
                            {
                                tmp.Add(c);
                            }
                        }
                        break;
                }
            }
            catch (Exception c)
            {
                System.Windows.MessageBox.Show("Error en la seccion " + posError + ". \nMensaje del error: " + c.Message, "Error");
                con.Close();
            }
        }

        public int DiasMeses(int mes)
        {
            switch (mes)
            {
                case 1:
                    //Enero
                    return 31;
                case 2:
                    //Febrero
                    if (DateTime.IsLeapYear(Convert.ToInt32(fecha.Value.ToString("yyyy"))))
                    {
                        return 29;
                    }
                    else
                    {
                        return 28;
                    }
                case 3:
                    //Marzo
                    return 31;
                case 4:
                    //Abril
                    return 30;
                case 5:
                    //Mayo
                    return 31;
                case 6:
                    //Junio
                    return 30;
                case 7:
                    //Julio
                    return 31;
                case 8:
                    //Agosto
                    return 31;
                case 9:
                    //Septiembre
                    return 30;
                case 10:
                    //Octubre
                    return 31;
                case 11:
                    //Noviembre
                    return 30;
                case 12:
                    //Diciembre
                    return 31;
                default:
                    return 30;
            }
        }

        private string LetraColumna(int numero)
        {
            int dividend = numero;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private bool FileInUse(string _filePath)
        {
            try
            {
                using (StreamReader stream = new StreamReader(_filePath))
                {
                    return false;
                }
            }
            catch
            {
                return true;
            }
        }

        public void AbrirDireccion(string path)
        {
            SaveFileDialog sFD = new SaveFileDialog();
            sFD.InitialDirectory = rutaMes;
            sFD.Filter = "Excel (*.xlxs) | *.xlsx |Todos los archivos (*.*)|*.*";
            sFD.FileName = nombreArchivo;

            sFD.ShowDialog();

            if (sFD.FileName != nombreArchivo)
            {
                Txtbox_Ruta.Text = sFD.FileName;
                ruta = sFD.FileName;
                rutaMes = Path.GetDirectoryName(sFD.FileName);
            }
        }

        public void BarraTransicion(int progreso)
        {
            prg_Hilo.Value = progreso;
        }

        public void Invoke(int tran, string estado)
        {
            Dispatcher.Invoke(() => {
                BarraTransicion(tran);
                txt_progreso.Content = estado;
            });
        }

        public void AbrirArchivo(string path)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.Filter = "Archivo (*.pdat) | *.pdat";

            oFD.ShowDialog();

            if (File.Exists(path))
            {
                System.Windows.MessageBox.Show("paso");
                DTG_Empleados.ItemsSource = null;
                DTG_Empleados_1_Q.ItemsSource = null;
                DTG_Empleados_2_Q.ItemsSource = null;
                colEmpleados = null;

                colEmpleados = new List<Empleados>();

                using (BinaryReader reader = new BinaryReader(File.Open(path, FileMode.Open)))
                {
                    colEmpleados.Add(new Empleados()
                    {
                        Id = reader.ReadInt32(),
                        FechaIngreso = Convert.ToDateTime(reader.ReadString()),
                        FechaSalida = Convert.ToDateTime(reader.ReadString()),
                        Cedula = reader.ReadString(),
                        Apellido = reader.ReadString(),
                        Nombre = reader.ReadString(),
                        Sueldo_Mensual = reader.ReadDouble(),
                        Area = reader.ReadString()
                    });
                }

                DTG_Empleados.ItemsSource = colEmpleados;
            }
        }
        

        public void guardarEnBase()
        {
            SaveFileDialog sFD = new SaveFileDialog();
            sFD.InitialDirectory = rutaMes;
            sFD.Filter = "Archivo de datos (*.pdat) | *.pdat";

            sFD.ShowDialog();

            try
            {
                using (BinaryWriter writer = new BinaryWriter(File.Open(sFD.FileName, FileMode.Create)))
                {
                    foreach (Empleados lEmp in DTG_Empleados.Items)
                    {
                        writer.Write(lEmp.Id);
                        writer.Write(lEmp.FechaIngreso.ToString());
                        writer.Write(lEmp.FechaSalida.ToString());
                        writer.Write(lEmp.Cedula);
                        writer.Write(lEmp.Apellido);
                        writer.Write(lEmp.Nombre);
                        writer.Write(lEmp.Sueldo_Mensual);
                        writer.Write(lEmp.Area);
                        writer.Write(lEmp.Vacaciones);
                        writer.Write(lEmp.CalcularDiasIESS);
                        writer.Write(";");
                    }
                }
            }
            catch(Exception e)
            {

            }

            //System.Windows.MessageBox.Show("No se pudo guardar los archivos, intente otra carpeta o intente con privilegios elevados","Error al guardar",MessageBoxButton.OK,MessageBoxImage.Error);
        }

        public int TransformarBooleanos(bool b1, bool b2, bool b3, bool b4)
        {
            int tmp1, tmp2, tmp3, tmp4;

            if (b1)
            {
                tmp1 = 8;
            }
            else
            {
                tmp1 = 0;
            }

            if (b2)
            {
                tmp2 = 4;
            }
            else
            {
                tmp2 = 0;
            }

            if (b3)
            {
                tmp3 = 2;
            }
            else
            {
                tmp3 = 0;
            }

            if (b4)
            {
                tmp4 = 1;
            }
            else
            {
                tmp4 = 0;
            }

            return tmp1 + tmp2 + tmp3 + tmp4;
        }

        public void GuardarAjuste(double x, double y)
        {
            Settings1.Default.YVentana = y;
            Settings1.Default.XVentana = x;
            Settings1.Default.Save();
        }

        public string VersionAplicacion()
        {
            return System.Windows.Application.ResourceAssembly.GetName().Version.ToString();
        }

        public void BuscarVersion()
        {
            try
            {
                using (WebClient client = new WebClient())
                {
                    string version = client.DownloadString("https://raw.githubusercontent.com/");

                    if (version != VersionAplicacion())
                    {
                        System.Windows.MessageBox.Show("Hay una nueva version disponible, actualice la aplicacion para poder continuar", "Nueva version", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch
            {
                System.Windows.MessageBox.Show("No se pudo conectar con el servidor, intente mas tarde", "Error de conexion", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void BuscarVersion(bool actualizacion)
        {
            if (!actualizacion)
            {
                BuscarVersion();
            }
        }

        public void Mensaje(string mensaje)
        {
            System.Windows.MessageBox.Show(mensaje);
        }

        public void EscribirEnArchivoVersion()
        {
            // Ruta de la aplicacion sin el archivo
            string ruta = AppDomain.CurrentDomain.BaseDirectory;

            // Escribe en un archivo la version
            using (StreamWriter writer = new StreamWriter(ruta + "\\ver.txt"))
            {
                writer.WriteLine(VersionAplicacion());
            }

            Settings1.Default.Version = VersionAplicacion();
        }
    }
}
