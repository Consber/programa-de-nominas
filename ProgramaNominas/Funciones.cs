using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Forms;

namespace ProgramaNominas
{
    public partial class MainWindow
    {
        static string strSqlE = @"Select * from Empleados";
        static string strSqlC = @"Select * from Cuentas";
        static string strDB = @"Provider =Microsoft.ACE.OLEDB.12.0;Data Source = Empleados.accdb";
        string aBin = @"ciapp.bin";
        string rutaApp = Directory.GetCurrentDirectory() + @"\actualizadorPrograma.exe";
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
            Dispatcher.Invoke(() =>
            {
                BarraTransicion(tran);
                txt_progreso.Content = estado;
            });
        }

        private int TransformarBooleanos(bool a, bool b, bool c, bool d)
        {
            return (a ? 8 : 0) | (b ? 4 : 0) | (c ? 2 : 0) | (d ? 1 : 0);
        }

        public void GuardarAjuste(double x, double y)
        {
            Settings1.Default.YVentana = y;
            Settings1.Default.XVentana = x;
            Settings1.Default.Save();
        }
        string version;

        public string VersionAplicacion()
        {
#if DEBUG
            {
                try
                {
                    string rutaVer = Directory.GetCurrentDirectory() + @"\ver.txt";

                    using (StreamReader sr = new StreamReader(rutaVer))
                    {
                        return version = sr.ReadToEnd();
                    }
                }
                catch
                {
                    return System.Windows.Application.ResourceAssembly.GetName().Version.ToString();
                }
            }
#else
            return System.Windows.Application.ResourceAssembly.GetName().Version.ToString();
#endif
        }

        bool auto;

        public void BuscarVersion()
        {
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                using (WebClient client = new WebClient())
                {
                    #if DEBUG
                    {
                        version = "1.0.0.0";
                    }
                    #else
                    {
                        version = client.DownloadString("https://raw.githubusercontent.com/Consber/programa-de-nominas/main/Publico/ver.txt");
                    }
                    #endif

                    if(VersionAplicacion().CompareTo(QuitarEspacios(version)) < 0)
                    {

                        switch (System.Windows.MessageBox.Show("Hay una nueva version disponible, desea actualizar la aplicacion?", "Nueva version", MessageBoxButton.OKCancel, MessageBoxImage.Information))
                        {
                            case MessageBoxResult.OK:
                                #if !DEBUG
                                {
                                    AbrirPrograma(rutaApp, "-auto");
                                    Close();
                                }
                                #endif
                                break;
                            case MessageBoxResult.Cancel:
                                break;
                        }
                    }
                    else if (!auto)
                    {
                        System.Windows.MessageBox.Show("La version actual es la mas reciente", "Version actual", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch
            {
                System.Windows.MessageBox.Show("No se pudo conectar con el servidor, intente mas tarde", "Error de conexion", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void BuscarVersionAuto(bool actualizacion)
        {
            auto = true;
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
            #if RELEASE
            {
                using (StreamWriter writer = new StreamWriter(ruta + "\\ver.txt"))
                {
                    writer.WriteLine(VersionAplicacion());
                }

                Settings1.Default.Version = VersionAplicacion();
            }
            #endif
        }

        // Abrir programa con parametros
        public void AbrirPrograma(string ruta, string parametros)
        {
            System.Diagnostics.Process.Start(ruta, parametros);
        }

        // Quitar saltos de linea
        public string QuitarEspacios(string cadena)
        {
            return cadena.Replace("\n", "").Replace("\r", "");
        }
    }
}
