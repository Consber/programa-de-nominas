using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace ProgramaNominas
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool trabajando = false;
        string contenidoRuta = "";
        bool iniciandoPrograma = true;
        bool editado = false;
        Thread t1;
        bool finalCancelar = false;

        public MainWindow()
        {
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1)
            {
                MessageBox.Show("El programa se encuentra corriendo. Cierre el programa si desea iniciar de nuevo el programa", "Programa duplicado", MessageBoxButton.OK, MessageBoxImage.Stop);
                Close();
            }
            else
            {
                InitializeComponent();
                Empezar();

                EscribirEnArchivoVersion();

                Application.Current.MainWindow.Height = Settings1.Default.YVentana;
                Application.Current.MainWindow.Width = Settings1.Default.XVentana;

                DTG_Empleados.ItemsSource = null;
                DTG_Empleados_1_Q.ItemsSource = null;
                DTG_Empleados_2_Q.ItemsSource = null;
                DTG_Empleados.ItemsSource = colEmpleados;
                DTG_Empleados_1_Q.ItemsSource = colEmpleadosQ1;
                DTG_Empleados_2_Q.ItemsSource = colEmpleadosQ2;

                BuscarVersionAuto(Settings1.Default.IgnorarActualizacion);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!trabajando)
            {
                if (!FileInUse(ruta) || !File.Exists(ruta))
                {
                    tmpCB = Clipboard.GetText();
                    trabajando = true;
                    t1 = new Thread(() => { HojaExcel(ruta); });

                    check_estilo.IsEnabled = false;
                    dtp_Fecha.IsEnabled = false;
                    but_Direccion.IsEnabled = false;
                    Txtbox_Ruta.IsEnabled = false;
                    DTG_Empleados.IsEnabled = false;
                    DTG_Empleados_1_Q.IsEnabled = false;
                    DTG_Empleados_2_Q.IsEnabled = false;
                    but_generar.Content = "Cancelar";

                    if (check_estilo.IsChecked ?? true)
                        checkEstilo = true;
                    else
                        checkEstilo = false;

                    t1.Start();
                }
                else
                {
                    MessageBox.Show("Otro programa o proceso esta utilizando el archivo. Cierre el archivo para poder continuar", "Error al crear la planilla", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                intentandoCancelar = true;
                switch (MessageBox.Show("¿Desea cancelar la creacion de la planilla? La planilla seguira en progreso hasta que se cancele", "Cancelar planilla", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    case MessageBoxResult.Yes:
                        trabajando = false;
                        t1.Abort();
                        intentandoCancelar = false;
                        break;
                    default:
                        intentandoCancelar = false;
                        if (finalCancelar)
                        {
                            _wait.Set();
                            finalCancelar = false;
                        }
                        break;
                }
            }
        }

        private void check_estilo_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void but_Direccion_Click(object sender, RoutedEventArgs e)
        {
            AbrirDireccion(ruta);

        }

        private void Txtbox_Ruta_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Txtbox_Ruta.Text == "")
            {
                Txtbox_Ruta.Text = contenidoRuta;
            }
        }

        private void programaCerrar(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!trabajando)
            {
                Archivo(false);
                if (editado)
                {
                    switch (MessageBox.Show("¿Desea guardar los datos?", "Guardar datos", MessageBoxButton.YesNoCancel, MessageBoxImage.Question))
                    {
                        case MessageBoxResult.Yes:
                            GuardarAjuste(Application.Current.MainWindow.Height, Application.Current.MainWindow.Width);
                            break;
                        case MessageBoxResult.No:

                            break;
                        case MessageBoxResult.Cancel:
                            e.Cancel = true;
                            break;
                        default:
                            break;
                    }
                }
                GuardarAjuste(Application.Current.MainWindow.Width, Application.Current.MainWindow.Height);
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void dtp_Fecha_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            Txtbox_Ruta.Text = rutaMes + "\\" + (nombreArchivo = Convert.ToDateTime(dtp_Fecha.ToString()).ToString("MMMM") + " - " + Convert.ToDateTime(dtp_Fecha.ToString()).ToString("yyyy") + ".xlsx");

            if (!iniciandoPrograma)
            {
                ruta = rutaMes + "\\" + nombreArchivo;
                fecha = dtp_Fecha.SelectedDate;
            }
        }

        private void but_Ayuda_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Si no muestra ningun empleado, asegurarse de tener la base de datos en la direccion: " + AppDomain.CurrentDomain.BaseDirectory + "\n\nEl nombre de la base debe llamarse \"Empleados.accdb\" en el formato que se le da.\n\nPara soporte tecnico y guia en el programa, escribir al correo soporte@consber.com.ec y con una imagen del programa con la descripcion de la pregunta o error", "Ayuda de solucion de problemas", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void DTG_E_Cambios(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (!editado)
            {
                editado = true;
            }
        }

        private void but_Abrir_Click(object sender, RoutedEventArgs e)
        {
        }

        private void but_Guardar_Click(object sender, RoutedEventArgs e)
        {
        }

        private void cb_Version(object sender, RoutedEventArgs e)
        {
            if (Settings1.Default.IgnorarActualizacion)
            {
                Settings1.Default.IgnorarActualizacion = false;
            }
            else
            {
                Settings1.Default.IgnorarActualizacion = true;
            }
            Settings1.Default.Save();
        }

        private void cb_Version_Buscar(object sender, RoutedEventArgs e)
        {
            auto = false;
            BuscarVersion();
        }
    }
}
