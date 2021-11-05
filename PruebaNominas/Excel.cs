using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace PruebaNominas
{
    public partial class MainWindow
    {
        static string nombreArchivo = " - ";
        static string ruta = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\";
        static string rutaMes = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        static bool intentandoCancelar = false;
        private static AutoResetEvent _wait = new AutoResetEvent(false);
        NumberFormatInfo nfi = new NumberFormatInfo();
        string tmpCB;

        public void HojaExcel(string path)
        {
            try
            {
                //Comentario
                Dispatcher.Invoke(() => {
                    BarraTransicion(1);
                    txt_progreso.Content = "Estado: Inicio";
                });

                Excel.Application xlApp = new
                Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel no esta instalado apropiadamente");
                    return;
                }

                CultureInfo ci = new CultureInfo("es-EC");

                object misValue = Missing.Value;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = xlWorkBook.ActiveSheet as Excel.Worksheet;
            
                xlWorkSheet.Name = "Numero de cuentas";

                xlApp.DisplayAlerts = false;

                const string formatoContabilidad = "_ [$$-es-EC]* #,##0.00_ ;_ [$$-es-EC]* -#,##0.00_ ;_ [$$-es-EC]* \" - \"??_ ;_ @_ ";
                const string formatoSueldosVac = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \" - \"??_ ;_ @_ ";
                const string formatoFechaCorta = "mmm-yy";

                int countR = 0;
                int countS = 0;
                int countOt = 0;
                int countAl = 0;
                int countTmp;
                int tmp;

                int[] DiasTrabajados;
                bool[] checkFR;
                bool[] checkDT;
                bool[] checkDC;


                Color celdasN = (Color)ColorConverter.ConvertFromString("#FABF8F");
                var celdasR = System.Drawing.Color.FromArgb(230, 184, 183);
                var celdasNa = System.Drawing.Color.FromArgb(250, 191, 143);
                var celdasBlanco = System.Drawing.Color.FromArgb(255, 255, 255);

                #region Contadores
                foreach (Cuentas c in colCuentas)
                {
                    if (c.TipoDeRol == "Rol")
                    {
                        countR++;
                    }

                    if (c.TipoDeRol == "Servicio" || c.TipoDeRol == "Servicios")
                    {
                        countS++;
                    }

                    if (c.TipoDeRol != "Rol" && c.TipoDeRol != "Servicio")
                    {
                        countOt++;
                    }
                }

                int countA = 0;
                int countO = 0;

                foreach (Empleados c in colEmpleados)
                {
                    if (c.Area == "Administracion" || c.Area == "Administración")
                    {
                        countA++;
                    }

                    if (c.Area == "Operativo")
                    {
                        countO++;
                    }

                    if (c.Area == "Albañiles" || c.Area == "Albaniles")
                    {
                        countAl++;
                    }
                }
                #endregion

                var numCuentas = xlWorkBook.ActiveSheet as Excel.Worksheet;
                var roles = xlWorkBook.Sheets.Add(numCuentas, misValue, 1, misValue) as Excel.Worksheet;
                var resumen = xlWorkBook.Sheets.Add(numCuentas, misValue, 1, misValue) as Excel.Worksheet;
                var pQuincena = xlWorkBook.Sheets.Add(numCuentas, misValue, 1, misValue) as Excel.Worksheet;
                var sQuincena = xlWorkBook.Sheets.Add(numCuentas, misValue, 1, misValue) as Excel.Worksheet;
                var optCeldas = xlWorkBook.Sheets.Add(numCuentas, misValue, 1, misValue) as Excel.Worksheet;

                #region optCeldas

                #region Celdas
                optCeldas.Cells[2,2].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,4].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                
                optCeldas.Cells[2,6].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,6].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                
                optCeldas.Cells[2,8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,10].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,10].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,12].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,12].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,14].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,14].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,14].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,16].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,18].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,18].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                
                optCeldas.Cells[2,20].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,20].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,22].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,22].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,22].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,24].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,24].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,26].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,26].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,26].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                                
                optCeldas.Cells[2,28].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,28].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,28].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                optCeldas.Cells[2,30].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,30].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,30].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                optCeldas.Cells[2,30].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                #endregion

                #endregion

                Excel.Worksheet vacaciones;
                Invoke(5, "Estado: Generando valores planilla: Cuentas");
                #region Cuentas
                #region Cuentas Valores
                posError = "Creando planilla de Excel";

                int x = 1; //Letra
                int y = 3; //Numero

                var fechaAhora = DateTime.Now;

                numCuentas.Range[numCuentas.Cells[2, 1], numCuentas.Cells[2, 10]].Merge();

                numCuentas.Cells[y - 1, x] = "CUENTAS PERSONAL EN ROL";


                foreach (Cuentas c in colCuentas)
                {
                    if (c.TipoDeRol == "Rol")
                    {
                        numCuentas.Cells[y, x] = c.FormaPago;
                        numCuentas.Cells[y, x + 1] = "'" + c.Banco;
                        numCuentas.Cells[y, x + 2] = "'" + c.TipoCuentaCheque;
                        numCuentas.Cells[y, x + 3] = "'" + c.NumCuenta;
                        numCuentas.Cells[y, x + 4] = c.Valor;
                        numCuentas.Cells[y, x + 5] = c.Identificacion;
                        numCuentas.Cells[y, x + 6] = "'" + c.NUC;
                        numCuentas.Cells[y, x + 7] = c.Beneficiario;
                        numCuentas.Cells[y, x + 8] = c.Telefono;
                        numCuentas.Cells[y, x + 9] = c.Referencia;
                        
                        y++;
                    }
                }

                y += 5;

                numCuentas.Range[numCuentas.Cells[y, x], numCuentas.Cells[y, x + 9]].Merge();

                y++;

                numCuentas.Cells[y - 1, x] = "CUENTAS PERSONAL EN ROL";

                foreach (Cuentas c in colCuentas)
                {
                    if (c.TipoDeRol == "Servicio")
                    {
                        numCuentas.Cells[y, x] = c.FormaPago;
                        numCuentas.Cells[y, x + 1] = "'" + c.Banco;
                        numCuentas.Cells[y, x + 2] = "'" + c.TipoCuentaCheque;
                        numCuentas.Cells[y, x + 3] = "'" + c.NumCuenta;
                        numCuentas.Cells[y, x + 4] = c.Valor;
                        numCuentas.Cells[y, x + 5] = c.Identificacion;
                        numCuentas.Cells[y, x + 6] = "'" + c.NUC;
                        numCuentas.Cells[y, x + 7] = c.Beneficiario;
                        numCuentas.Cells[y, x + 8] = c.Telefono;
                        numCuentas.Cells[y, x + 9] = c.Referencia;
                        
                        y++;
                    }
                }

                y += 10;


                foreach (Cuentas c in colCuentas)
                {
                    if (c.TipoDeRol != "Rol" && c.TipoDeRol != "Servicio")
                    {
                        numCuentas.Cells[y, x] = c.FormaPago;
                        numCuentas.Cells[y, x + 1] = "'" + c.Banco;
                        numCuentas.Cells[y, x + 2] = "'" + c.TipoCuentaCheque;
                        numCuentas.Cells[y, x + 3] = "'" + c.NumCuenta;
                        numCuentas.Cells[y, x + 4] = c.Valor;
                        numCuentas.Cells[y, x + 5] = c.Identificacion;
                        numCuentas.Cells[y, x + 6] = "'" + c.NUC;
                        numCuentas.Cells[y, x + 7] = c.Beneficiario;
                        numCuentas.Cells[y, x + 8] = c.Telefono;
                        numCuentas.Cells[y, x + 9] = c.Referencia;

                        y++;
                    }
                }
                #endregion

                #region Valores Estilo
                if (checkEstilo)
                {
                    LineasCuadrosRango(true,true,true,true,2,1,2,10, numCuentas, Excel.XlLineStyle.xlContinuous);
                    numCuentas.Cells[2,1].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasNa);
                    LineasCuadrosRango(true,true,true,true,2 + countR + 6,1,2 + countR + 6,10, numCuentas, Excel.XlLineStyle.xlContinuous);
                    numCuentas.Cells[2 + countR + 6, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasNa);
                }
                    #endregion
                    #endregion

                Invoke(10, "Estado: Generando valores planilla: Primera Quincena");
                #region Primera Quincena
                #region Primera Quincena Valores
                pQuincena.Name = "Primera Quincena";
                y = 1;
                x = 1;

                #region Valores estaticos
                pQuincena.Cells[1, 1] = "CONSBER C.A. CONSTRUCTORA BERREZUETA";
                pQuincena.Cells[2, 1] = "DETALLE DE HABERES";
                pQuincena.Cells[3, 1] = "1ERA QUINCENA DE " + fecha.Value.ToString("MMMM", ci).ToUpper() + fecha.Value.ToString("yyyy", ci);
                
                pQuincena.Cells[5, 9] = "INGRESOS";

                pQuincena.Cells[5, 20] = "EGRESOS";
                
                pQuincena.Cells[6, 1] = "No";
                pQuincena.Cells[6, 2] = "Fecha Ingreso";
                pQuincena.Cells[6, 3] = "Fecha Salida";
                pQuincena.Cells[6, 4] = "Cedula";
                pQuincena.Cells[6, 5] = "Nombres";
                pQuincena.Cells[6, 6] = "Sueldo Mensual";
                pQuincena.Cells[6, 7] = "Valor Diario";
                pQuincena.Cells[6, 8] = "Dias Trab.";
                pQuincena.Cells[6, 9] = "Total Sueldo";
                pQuincena.Cells[6, 10] = "Alim. Quinc.";
                pQuincena.Cells[6, 11] = "Transporte";
                pQuincena.Cells[6, 12] = "Bono";
                pQuincena.Cells[6, 13] = "Tarjeta";
                pQuincena.Cells[6, 14] = "Horas Extras";
                pQuincena.Cells[6, 15] = "Vacaciones";
                pQuincena.Cells[6, 16] = "Fondo reserva";
                pQuincena.Cells[6, 17] = "Decimo tercero";
                pQuincena.Cells[6, 18] = "Decimo cuarto";
                pQuincena.Cells[6, 19] = "Total Ingresos";
                pQuincena.Cells[6, 20] = "Aportes IESS";
                pQuincena.Cells[6, 21] = "Prest. Hipot.";
                pQuincena.Cells[6, 22] = "Prest. Quiro";
                pQuincena.Cells[6, 23] = "Prest. Cía";
                pQuincena.Cells[6, 24] = "Multas";
                pQuincena.Cells[6, 25] = "Ext Salud";
                pQuincena.Cells[6, 26] = "Tarjeta";
                pQuincena.Cells[6, 27] = "Contribucion Solidaria";
                pQuincena.Cells[6, 28] = "Anticipo de quincena";
                pQuincena.Cells[6, 29] = "Total Egresos";
                pQuincena.Cells[6, 30] = "Total Recibir";

                pQuincena.Cells[8, 5] = "Administración";
                #endregion

                x = 1;
                y = 9;

                countTmp = 1;
                DiasTrabajados = new int[countA + countO + countAl];
                checkFR = new bool[countA + countO + countAl];
                checkDT = new bool[countA + countO + countAl];
                checkDC = new bool[countA + countO + countAl];

                tmp = y;

                foreach (EmpleadosQ e in DTG_Empleados_1_Q.Items)
                {
                    DiasTrabajados[e.Id - 1] = e.DiasTrabajados;

                    if (e.FondosReserva)
                    {
                        checkFR[e.Id - 1] = true;
                    }
                    else
                    {
                        checkFR[e.Id - 1] = false;
                    }

                    if (e.DecimoTercero)
                    {
                        checkDT[e.Id - 1] = true;
                    }
                    else
                    {
                        checkDT[e.Id - 1] = false;
                    }

                    if (e.DecimoCuarto)
                    {
                        checkDC[e.Id - 1] = true;
                    }
                    else
                    {
                        checkDC[e.Id - 1] = false;
                    }

                    if (e.Area == "Administracion")
                    {
                        pQuincena.Cells[y, x + 9] = e.Alim;
                        pQuincena.Cells[y, x + 10] = e.Transp;
                        pQuincena.Cells[y, x + 11] = e.Bono;
                        pQuincena.Cells[y, x + 12] = e.TarjetaIngresos;
                        pQuincena.Cells[y, x + 13] = e.HorasExtra;
                        pQuincena.Cells[y, x + 14] = e.Vacaciones;

                        pQuincena.Cells[y, x + 20] = e.PrestHipot;
                        pQuincena.Cells[y, x + 21] = e.PrestQuiro;
                        pQuincena.Cells[y, x + 22] = e.PrestCia;
                        pQuincena.Cells[y, x + 23] = e.Multas;
                        pQuincena.Cells[y, x + 24] = e.ExtSalud;
                        pQuincena.Cells[y, x + 25] = e.TarjetaEgresos;
                        pQuincena.Cells[y, x + 26] = e.ContribucionSolidaria;
                        pQuincena.Cells[y, x + 27] = e.AnticipoQuincena;

                        y++;
                    }

                    if (e.Area == "Operativo")
                    {
                        pQuincena.Cells[y + 3 , x + 9] = e.Alim;
                        pQuincena.Cells[y + 3 , x + 10] = e.Transp;
                        pQuincena.Cells[y + 3 , x + 11] = e.Bono;
                        pQuincena.Cells[y + 3 , x + 12] = e.TarjetaIngresos;
                        pQuincena.Cells[y + 3 , x + 13] = e.HorasExtra;
                        pQuincena.Cells[y + 3 , x + 14] = e.Vacaciones;
                        pQuincena.Cells[y + 3 , x + 20] = e.PrestHipot;
                        pQuincena.Cells[y + 3 , x + 21] = e.PrestQuiro;
                        pQuincena.Cells[y + 3 , x + 22] = e.PrestCia;
                        pQuincena.Cells[y + 3 , x + 23] = e.Multas;
                        pQuincena.Cells[y + 3 , x + 24] = e.ExtSalud;
                        pQuincena.Cells[y + 3 , x + 25] = e.TarjetaEgresos;
                        pQuincena.Cells[y + 3 , x + 26] = e.ContribucionSolidaria;
                        pQuincena.Cells[y + 3 , x + 27] = e.AnticipoQuincena;

                        y++;
                    }
                }

                y = tmp;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Administracion
                    if (e.Area == "Administracion")
                    {
                        //Valor
                        pQuincena.Cells[y, x] = countTmp;
                        countTmp++;
                        pQuincena.Cells[y, x + 1] = e.FechaIngreso;
                        pQuincena.Cells[y, x + 2] = e.FechaSalida;
                        if (pQuincena.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            pQuincena.Cells[y, x + 2] = "";
                        }
                        pQuincena.Cells[y, x + 3] = "'" + e.Cedula;
                        pQuincena.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        pQuincena.Range[pQuincena.Cells[y, x + 5], pQuincena.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        pQuincena.Cells[y, x + 5] = e.Sueldo_Mensual;
                        pQuincena.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/30";
                        pQuincena.Cells[y, x + 7] = DiasTrabajados[e.Id - 1];

                        //Total Ingresos
                        pQuincena.Range[pQuincena.Cells[y, x + 8], pQuincena.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        pQuincena.Cells[y, x + 8] = "=" + LetraColumna(x + 6) + y + "*" + LetraColumna(x + 7) + y;
                        
                        if (checkFR[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 15].Value = "=+(" + LetraColumna(x + 8) + y + "+" + LetraColumna(x + 8) + y + ") * 8.33 % ";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 15].Value = 0;
                        }

                        if (checkDT[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 16].Value = "=" + LetraColumna(x + 5) + y + "/12";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 16].Value = 0;
                        }

                        if (checkDC[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 17].Value = "=+(400 / 360) * " + LetraColumna(x + 7) + y;
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 17].Value = 0;
                        }

                        pQuincena.Cells[y, x + 18].Value = "=SUM(" + LetraColumna(x + 8) + y + ":" + LetraColumna(x + 17) + (y) + ")";

                        //Total Egresos
                        pQuincena.Cells[y, x + 19] = 0;
                        pQuincena.Range[pQuincena.Cells[y, x + 19], pQuincena.Cells[y, x + 29]].NumberFormat = formatoContabilidad;

                        pQuincena.Cells[y, x + 28] = "=SUM(" + LetraColumna(x + 19) + y + ":" + LetraColumna(x + 27) + (y) + ")";

                        pQuincena.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }

                pQuincena.Cells[8 + countA + 1, 5] = "Total Administración";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    pQuincena.Cells[8 + countA + 1, i] = "=SUM(" + LetraColumna(i) + 9 + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                pQuincena.Cells[8 + countA + 3, 5] = "Operativo";

                y += 3;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Operativo
                    if (e.Area == "Operativo")
                    {
                        //Valor
                        pQuincena.Cells[y, x] = countTmp;
                        countTmp++;
                        pQuincena.Cells[y, x + 1] = e.FechaIngreso;
                        pQuincena.Cells[y, x + 2] = e.FechaSalida;
                        if (pQuincena.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            pQuincena.Cells[y, x + 2] = "";
                        }
                        pQuincena.Cells[y, x + 3] = "'" + e.Cedula;
                        pQuincena.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        pQuincena.Range[pQuincena.Cells[y, x + 5], pQuincena.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        pQuincena.Cells[y, x + 5] = e.Sueldo_Mensual;
                        pQuincena.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/30";
                        pQuincena.Cells[y, x + 7] = DiasTrabajados[e.Id - 1];

                        //Total Ingresos
                        pQuincena.Range[pQuincena.Cells[y, x + 8], pQuincena.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        pQuincena.Cells[y, x + 8] = "=" + LetraColumna(x + 6) + y + "*" + LetraColumna(x + 7) + y;

                        if (checkFR[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 15].Value = "=+(" + LetraColumna(x + 8) + y + "+" + LetraColumna(x + 8) + y + ") * 8.33 % ";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 15].Value = 0;
                        }

                        if (checkDT[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 16].Value = "=" + LetraColumna(x + 5) + y + "/12";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 16].Value = 0;
                        }

                        if (checkDC[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 17].Value = "=+(400 / 360) * " + LetraColumna(x + 7) + y;
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 17].Value = 0;
                        }

                        pQuincena.Cells[y, x + 18].Value = "=SUM(" + LetraColumna(x + 8) + y + ":" + LetraColumna(x + 17) + (y) + ")";

                        //Total Egresos
                        pQuincena.Cells[y, x + 19] = 0;
                        pQuincena.Range[pQuincena.Cells[y, x + 19], pQuincena.Cells[y, x + 29]].NumberFormat = formatoContabilidad;

                        pQuincena.Cells[y, x + 28] = "=SUM(" + LetraColumna(x + 19) + y + ":" + LetraColumna(x + 27) + (y) + ")";

                        pQuincena.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }

                pQuincena.Cells[y, x + 4] = "Total Operativo";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    pQuincena.Cells[y, i] = "=SUM(" + LetraColumna(i) + (y - countO) + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                pQuincena.Cells[y + 2, 5] = "Total";
                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    pQuincena.Cells[y + 2, i] = "=" + LetraColumna(i) + (countA + 8 + 1) + " + " + LetraColumna(i) + (countA + countO + 12);
                }

                y += 9;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Albaniles
                    if (e.Area == "Albañiles")
                    {
                        //Valor
                        pQuincena.Cells[y, x] = countTmp;
                        countTmp++;
                        pQuincena.Cells[y, x + 1] = e.FechaIngreso;
                        pQuincena.Cells[y, x + 2] = e.FechaSalida;
                        if (pQuincena.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            pQuincena.Cells[y, x + 2] = "";
                        }
                        pQuincena.Cells[y, x + 3] = "'" + e.Cedula;
                        pQuincena.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        pQuincena.Cells[y, x + 5] = e.Sueldo_Mensual;
                        pQuincena.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/30";
                        pQuincena.Cells[y, x + 7] = DiasTrabajados[e.Id - 1];

                        //Total Ingresos
                        pQuincena.Cells[y, x + 8] = "=" + LetraColumna(x + 6) + y + "*" + LetraColumna(x + 7) + y;
                        pQuincena.Cells[y, x + 9] = 0;
                        pQuincena.Cells[y, x + 10].Value = 0;
                        pQuincena.Cells[y, x + 11].Value = 0;
                        pQuincena.Cells[y, x + 12].Value = 0;
                        pQuincena.Cells[y, x + 13].Value = 0;
                        pQuincena.Cells[y, x + 14].Value = 0;
                        if (checkFR[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 15].Value = "=+(" + LetraColumna(x + 8) + y + "+" + LetraColumna(x + 8) + y + ") * 8.33 % ";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 15].Value = 0;
                        }

                        if (checkDT[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 16].Value = "=" + LetraColumna(x + 5) + y + "/12";
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 16].Value = 0;
                        }

                        if (checkDC[e.Id - 1])
                        {
                            pQuincena.Cells[y, x + 17].Value = "=+(400 / 360) * " + LetraColumna(x + 7) + y;
                        }
                        else
                        {
                            pQuincena.Cells[y, x + 17].Value = 0;
                        }

                        pQuincena.Cells[y, x + 18].Value = "=SUM(" + LetraColumna(x + 8) + y + ":" + LetraColumna(x + 17) + (y) + ")";

                        //Total Egresos
                        pQuincena.Range[pQuincena.Cells[y, x + 19], pQuincena.Cells[y, x + 29]].NumberFormat = formatoContabilidad;
                        pQuincena.Cells[y, x + 19].Value = 0;
                        pQuincena.Cells[y, x + 20].Value = 0;
                        pQuincena.Cells[y, x + 21].Value = 0;
                        pQuincena.Cells[y, x + 22].Value = 0;
                        pQuincena.Cells[y, x + 23].Value = 0;
                        pQuincena.Cells[y, x + 24].Value = 0;
                        pQuincena.Cells[y, x + 25].Value = 0;
                        pQuincena.Cells[y, x + 26].Value = 0;
                        pQuincena.Cells[y, x + 27].Value = 0;

                        pQuincena.Cells[y, x + 28] = "=SUM(" + LetraColumna(x + 19) + y + ":" + LetraColumna(x + 27) + (y) + ")";

                        pQuincena.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }
                #endregion

                #region Valores Estilo
                if (checkEstilo)
                {
                    #region Lineas
                    LineasCuadrosRango(false, true, false, false, 3, 1, 3, 30, pQuincena, 4d, Excel.XlLineStyle.xlContinuous);

                    LineasCuadros(true, true, true, true, 8, 5, optCeldas, pQuincena);

                    for (int i = 9; i < 30; i++)
                    {
                        LineasCuadros(true, true, true, true, 5, i, optCeldas, pQuincena);
                    }

                    for (int i = 1; i < 31; i++)
                    {
                        LineasCuadros(true, true, true, true, 6, i, optCeldas, pQuincena);
                    }


                    for (int i = 9; i < 9 + countA; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, pQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + 1, i, optCeldas, pQuincena);
                        }
                    }

                    LineasCuadros(true, true, true, true, 8 + countA + 3, 5, optCeldas, pQuincena);

                    for (int i = 12 + countA; i < 12 + countA + countO; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, pQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 4, i, optCeldas, pQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 6, i, optCeldas, pQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + countAl + 13, i, optCeldas, pQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + countAl + 15, i, optCeldas, pQuincena);
                        }
                    }
                    #endregion

                    pQuincena.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    pQuincena.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    pQuincena.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    pQuincena.Range[pQuincena.Cells[1, 1], pQuincena.Cells[1, 30]].Merge();
                    pQuincena.Range[pQuincena.Cells[2, 1], pQuincena.Cells[2, 30]].Merge();
                    pQuincena.Range[pQuincena.Cells[3, 1], pQuincena.Cells[3, 30]].Merge();
                    pQuincena.Range[pQuincena.Cells[5, 9], pQuincena.Cells[5, 19]].Merge();
                    pQuincena.Range[pQuincena.Cells[5, 20], pQuincena.Cells[5, 29]].Merge();

                    pQuincena.Range[pQuincena.Cells[6, 1], pQuincena.Cells[6, 30]].WrapText = true;

                    pQuincena.Range[pQuincena.Cells[y, x + 5], pQuincena.Cells[y, x + 6]].NumberFormat = formatoContabilidad;

                    pQuincena.Range[pQuincena.Cells[y, x + 8], pQuincena.Cells[y, x + 18]].NumberFormat = formatoContabilidad;

                    pQuincena.Range[pQuincena.Cells[1, 1], pQuincena.Cells[3, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    pQuincena.Range[pQuincena.Cells[5, 9], pQuincena.Cells[5, 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[6, 1], pQuincena.Cells[6, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Cells[8, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[8 + countA + 1, 5], pQuincena.Cells[8 + countA + 1, 7]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[8 + countA + 1, 9], pQuincena.Cells[8 + countA + 1, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Cells[8 + countA + 3, x + 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y - 9 - countAl, x + 4], pQuincena.Cells[y - 9 - countAl, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y - 9 - countAl, x + 8], pQuincena.Cells[y - 9 - countAl, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y - 9 - countAl + 2, x + 4], pQuincena.Cells[y - 9 - countAl + 2, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y - 9 - countAl + 2, x + 8], pQuincena.Cells[y - 9 - countAl + 2, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y, x + 4], pQuincena.Cells[y, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y, x + 8], pQuincena.Cells[y, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y + 2, x + 4], pQuincena.Cells[y + 2, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    pQuincena.Range[pQuincena.Cells[y + 2, x + 8], pQuincena.Cells[y + 2, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    pQuincena.Cells.Font.Size = 16;

                    pQuincena.Columns.AutoFit();
                }
                #endregion
                #endregion

                Invoke(30, "Estado: Generando valores planilla: Segunda Quincena");
                #region Segunda Quincena
                sQuincena.Name = "Segunda Quincena";
                #region Segunda Quincena Valores
                y = 1;
                x = 1;
                sQuincena.Name = "Segunda Quincena";

                #region Valores estaticos
                sQuincena.Cells[1, 1] = "CONSBER C.A. CONSTRUCTORA BERREZUETA";
                sQuincena.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sQuincena.Cells[2, 1] = "DETALLE DE HABERES";
                sQuincena.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sQuincena.Cells[3, 1] = "2DA QUINCENA DE " + fecha.Value.ToString("MMMM", ci).ToUpper() + fecha.Value.ToString("yyyy", ci);
                sQuincena.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                sQuincena.Range[sQuincena.Cells[1, 1], sQuincena.Cells[1, 30]].Merge();
                sQuincena.Range[sQuincena.Cells[2, 1], sQuincena.Cells[2, 30]].Merge();
                sQuincena.Range[sQuincena.Cells[3, 1], sQuincena.Cells[3, 30]].Merge();

                sQuincena.Cells[5, 9] = "INGRESOS";
                sQuincena.Range[sQuincena.Cells[5, 9], sQuincena.Cells[5, 19]].Merge();

                sQuincena.Cells[5, 20] = "EGRESOS";
                sQuincena.Range[sQuincena.Cells[5, 20], sQuincena.Cells[5, 29]].Merge();

                sQuincena.Range[sQuincena.Cells[6, 1], sQuincena.Cells[6, 30]].WrapText = true;

                sQuincena.Cells[6, 1] = "No";
                sQuincena.Cells[6, 2] = "Fecha Ingreso";
                sQuincena.Cells[6, 3] = "Fecha Salida";
                sQuincena.Cells[6, 4] = "Cedula";
                sQuincena.Cells[6, 5] = "Nombres";
                sQuincena.Cells[6, 6] = "Sueldo Mensual";
                sQuincena.Cells[6, 7] = "Valor Diario";
                sQuincena.Cells[6, 8] = "Dias Trab.";
                sQuincena.Cells[6, 9] = "Total Sueldo";
                sQuincena.Cells[6, 10] = "Alim. Quinc.";
                sQuincena.Cells[6, 11] = "Transporte";
                sQuincena.Cells[6, 12] = "Bono";
                sQuincena.Cells[6, 13] = "Tarjeta";
                sQuincena.Cells[6, 14] = "Horas Extras";
                sQuincena.Cells[6, 15] = "Vacaciones";
                sQuincena.Cells[6, 16] = "Fondo reserva";
                sQuincena.Cells[6, 17] = "Decimo tercero";
                sQuincena.Cells[6, 18] = "Decimo cuarto";
                sQuincena.Cells[6, 19] = "Total Ingresos";
                sQuincena.Cells[6, 20] = "Aportes IESS";
                sQuincena.Cells[6, 21] = "Prest. Hipot.";
                sQuincena.Cells[6, 22] = "Prest. Quiro";
                sQuincena.Cells[6, 23] = "Prest. Cía";
                sQuincena.Cells[6, 24] = "Multas";
                sQuincena.Cells[6, 25] = "Ext Salud";
                sQuincena.Cells[6, 26] = "Tarjeta";
                sQuincena.Cells[6, 27] = "Contribucion Solidaria";
                sQuincena.Cells[6, 28] = "Anticipo de quincena";
                sQuincena.Cells[6, 29] = "Total Egresos";
                sQuincena.Cells[6, 30] = "Total Recibir";

                sQuincena.Cells[8, 5] = "Administración";
                #endregion

                x = 1;
                y = 9;

                countTmp = 1;
                DiasTrabajados = new int[countA + countO + countAl];
                checkFR = new bool[countA + countO + countAl];
                checkDT = new bool[countA + countO + countAl];
                checkDC = new bool[countA + countO + countAl];

                foreach (EmpleadosQ e in DTG_Empleados_2_Q.Items)
                {
                    DiasTrabajados[e.Id - 1] = e.DiasTrabajados;
                    if (e.FondosReserva)
                    {
                        checkFR[e.Id - 1] = true;
                    }
                    else
                    {
                        checkFR[e.Id - 1] = false;
                    }

                    if (e.DecimoTercero)
                    {
                        checkDT[e.Id - 1] = true;
                    }
                    else
                    {
                        checkDT[e.Id - 1] = false;
                    }

                    if (e.DecimoCuarto)
                    {
                        checkDC[e.Id - 1] = true;
                    }
                    else
                    {
                        checkDC[e.Id - 1] = false;
                    }
                }

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Administracion
                    if (e.Area == "Administracion")
                    {
                        //Valor
                        sQuincena.Cells[y, x] = countTmp;
                        countTmp++;
                        sQuincena.Cells[y, x + 1] = e.FechaIngreso;
                        sQuincena.Cells[y, x + 2] = e.FechaSalida;
                        if (sQuincena.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            sQuincena.Cells[y, x + 2] = "";
                        }
                        sQuincena.Cells[y, x + 3] = "'" + e.Cedula;
                        sQuincena.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        sQuincena.Range[sQuincena.Cells[y, x + 5], sQuincena.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 5] = e.Sueldo_Mensual;
                        sQuincena.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/30";
                        sQuincena.Cells[y, x + 7] = DiasTrabajados[e.Id - 1];

                        //Total Ingresos
                        sQuincena.Range[sQuincena.Cells[y, x + 8], sQuincena.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 8] = "=" + LetraColumna(x + 6) + y + "*" + LetraColumna(x + 7) + y;
                        sQuincena.Cells[y, x + 9] = 0;
                        sQuincena.Cells[y, x + 10].Value = 0;
                        sQuincena.Cells[y, x + 11].Value = 0;
                        sQuincena.Cells[y, x + 12].Value = 0;
                        sQuincena.Cells[y, x + 13].Value = 0;
                        sQuincena.Cells[y, x + 14].Value = 0;
                        if (checkFR[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 15].Value = "=+(" + LetraColumna(x + 8) + y + "+'Primera Quincena'!" + LetraColumna(x + 8) + y + ") * 8.33 % ";
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 15].Value = 0;
                        }

                        if (checkDT[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 16].Value = "=" + LetraColumna(x + 5) + y + "/12";
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 16].Value = 0;
                        }

                        if (checkDC[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 17].Value = "=+(400 / 360) * " + LetraColumna(x + 7) + y;
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 17].Value = 0;
                        }

                        sQuincena.Cells[y, x + 18].Value = "=SUM(" + LetraColumna(x + 8) + y + ":" + LetraColumna(x + 17) + (y) + ")";

                        //Total Egresos
                        sQuincena.Range[sQuincena.Cells[y, x + 19], sQuincena.Cells[y, x + 29]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 19].Value = "=" + LetraColumna(x + 5) + y + " * 9.45%";
                        sQuincena.Cells[y, x + 20].Value = 0;
                        sQuincena.Cells[y, x + 21].Value = 0;
                        sQuincena.Cells[y, x + 22].Value = 0;
                        sQuincena.Cells[y, x + 23].Value = 0;
                        sQuincena.Cells[y, x + 24].Value = 0;
                        sQuincena.Cells[y, x + 25].Value = 0;
                        sQuincena.Cells[y, x + 26].Value = 0;
                        sQuincena.Cells[y, x + 27].Value = 0;

                        sQuincena.Cells[y, x + 28] = "=SUM(" + LetraColumna(x + 19) + y + ":" + LetraColumna(x + 27) + (y) + ")";

                        sQuincena.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }

                    foreach (EmpleadosQ q in DTG_Empleados_1_Q.Items)
                    {

                    }
                    #endregion
                }
                sQuincena.Cells[8 + countA + 1, 5] = "Total Administración";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    sQuincena.Cells[8 + countA + 1, i] = "=SUM(" + LetraColumna(i) + 9 + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                sQuincena.Cells[8 + countA + 3, 5] = "Operativo";

                y += 3;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Operativo
                    if (e.Area == "Operativo")
                    {
                        //Valor
                        sQuincena.Cells[y, x] = countTmp;
                        countTmp++;
                        sQuincena.Cells[y, x + 1] = e.FechaIngreso;
                        sQuincena.Cells[y, x + 2] = e.FechaSalida;
                        if (sQuincena.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            sQuincena.Cells[y, x + 2] = "";
                        }
                        sQuincena.Cells[y, x + 3] = "'" + e.Cedula;
                        sQuincena.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        sQuincena.Range[sQuincena.Cells[y, x + 5], sQuincena.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 5] = e.Sueldo_Mensual;
                        sQuincena.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/30";
                        sQuincena.Cells[y, x + 7] = DiasTrabajados[e.Id - 1];

                        //Total Ingresos
                        sQuincena.Range[sQuincena.Cells[y, x + 8], sQuincena.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 8] = "=" + LetraColumna(x + 6) + y + "*" + LetraColumna(x + 7) + y;
                        sQuincena.Cells[y, x + 9] = 0;
                        sQuincena.Cells[y, x + 10].Value = 0;
                        sQuincena.Cells[y, x + 11].Value = 0;
                        sQuincena.Cells[y, x + 12].Value = 0;
                        sQuincena.Cells[y, x + 13].Value = 0;
                        sQuincena.Cells[y, x + 14].Value = 0;
                        if (checkFR[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 15].Value = "=+(" + LetraColumna(x + 8) + y + "+'Primera Quincena'!" + LetraColumna(x + 8) + y + ") * 8.33 % ";
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 15].Value = 0;
                        }

                        if (checkDT[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 16].Value = "=" + LetraColumna(x + 5) + y + "/12";
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 16].Value = 0;
                        }

                        if (checkDC[e.Id - 1])
                        {
                            sQuincena.Cells[y, x + 17].Value = "=+(400 / 360) * " + LetraColumna(x + 7) + y;
                        }
                        else
                        {
                            sQuincena.Cells[y, x + 17].Value = 0;
                        }

                        sQuincena.Cells[y, x + 18].Value = "=SUM(" + LetraColumna(x + 8) + y + ":" + LetraColumna(x + 17) + (y) + ")";

                        //Total Egresos
                        sQuincena.Range[sQuincena.Cells[y, x + 19], sQuincena.Cells[y, x + 29]].NumberFormat = formatoContabilidad;
                        sQuincena.Cells[y, x + 19].Value = 0;
                        sQuincena.Cells[y, x + 20].Value = 0;
                        sQuincena.Cells[y, x + 21].Value = 0;
                        sQuincena.Cells[y, x + 22].Value = 0;
                        sQuincena.Cells[y, x + 23].Value = 0;
                        sQuincena.Cells[y, x + 24].Value = 0;
                        sQuincena.Cells[y, x + 25].Value = 0;
                        sQuincena.Cells[y, x + 26].Value = 0;
                        sQuincena.Cells[y, x + 27].Value = 0;

                        sQuincena.Cells[y, x + 28] = "=SUM(" + LetraColumna(x + 19) + y + ":" + LetraColumna(x + 27) + (y) + ")";

                        sQuincena.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }

                sQuincena.Cells[y, x + 4] = "Total Operativo";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    sQuincena.Cells[y, i] = "=SUM(" + LetraColumna(i) + (y - countO) + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                sQuincena.Cells[y + 2, 5] = "Total";
                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    sQuincena.Cells[y + 2, i] = "=" + LetraColumna(i) + (countA + 8 + 1) + " + " + LetraColumna(i) + (countA + countO + 12);
                }
                #endregion

                #region Valores Estilo
                if (checkEstilo)
                {
                    LineasCuadrosRango(false, true, false, false, 3, 1, 3, 30, sQuincena, 4d, Excel.XlLineStyle.xlContinuous);

                    LineasCuadros(true, true, true, true, 8, 5, optCeldas, sQuincena);

                    for (int i = 9; i < 30; i++)
                    {
                        LineasCuadros(true, true, true, true, 5, i, optCeldas, sQuincena);
                    }

                    for (int i = 1; i < 31; i++)
                    {
                        LineasCuadros(true, true, true, true, 6, i, optCeldas, sQuincena);
                    }


                    for (int i = 9; i < 9 + countA; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, sQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + 1, i, optCeldas, sQuincena);
                        }
                    }

                    LineasCuadros(true, true, true, true, 8 + countA + 3, 5, optCeldas, sQuincena);

                    for (int i = 12 + countA; i < 12 + countA + countO; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, sQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 4, i, optCeldas, sQuincena);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 6, i, optCeldas, sQuincena);
                        }
                    }

                    sQuincena.Range[sQuincena.Cells[1, 1], sQuincena.Cells[3, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    sQuincena.Range[sQuincena.Cells[5, 9], sQuincena.Cells[5, 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[6, 1], sQuincena.Cells[6, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Cells[8, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[8 + countA + 1, 5], sQuincena.Cells[8 + countA + 1, 7]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[8 + countA + 1, 9], sQuincena.Cells[8 + countA + 1, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Cells[8 + countA + 3, x + 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[y, x + 4], sQuincena.Cells[y, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[y, x + 8], sQuincena.Cells[y, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[y + 2, x + 4], sQuincena.Cells[y + 2, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    sQuincena.Range[sQuincena.Cells[y + 2, x + 8], sQuincena.Cells[y + 2, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    sQuincena.Cells.Font.Size = 16;

                    sQuincena.Columns.AutoFit();

                }
                #endregion
                #endregion

                Invoke(50, "Estado: Generando valores planilla: Resumen");
                #region Resumen
                #region Resumen Valores
                y = 1;
                x = 1;
                resumen.Name = "Resumen";
            
                #region Valores estaticos
                resumen.Cells[1, 1] = "CONSBER C.A. CONSTRUCTORA BERREZUETA";
                resumen.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                resumen.Cells[2, 1] = "DETALLE DE HABERES";
                resumen.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                resumen.Cells[3, 1] = "Periodo: 1 de " + fecha.Value.ToString("MMMM", ci) + " de " + fecha.Value.ToString("yyyy", ci) + " - " + DiasMeses(Convert.ToInt32(fecha.Value.ToString("MM"))) + " de " + fecha.Value.ToString("MMMM", ci) + " de " + fecha.Value.ToString("yyyy", ci);
                resumen.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                resumen.Range[resumen.Cells[1, 1], resumen.Cells[1, 30]].Merge();
                resumen.Range[resumen.Cells[2, 1], resumen.Cells[2, 30]].Merge();
                resumen.Range[resumen.Cells[3, 1], resumen.Cells[3, 30]].Merge();

                resumen.Cells[5, 9] = "INGRESOS";
                resumen.Range[resumen.Cells[5, 9], resumen.Cells[5, 19]].Merge();

                resumen.Cells[5, 20] = "EGRESOS";
                resumen.Range[resumen.Cells[5, 20], resumen.Cells[5, 29]].Merge();

                resumen.Range[resumen.Cells[6, 1], resumen.Cells[6, 30]].WrapText = true;

                resumen.Cells[6, 1] = "No";
                resumen.Cells[6, 2] = "Fecha Ingreso";
                resumen.Cells[6, 3] = "Fecha Salida";
                resumen.Cells[6, 4] = "Cedula";
                resumen.Cells[6, 5] = "Nombres";
                resumen.Cells[6, 6] = "Sueldo Mensual";
                resumen.Cells[6, 7] = "Valor Diario";
                resumen.Cells[6, 8] = "Dias Trab.";
                resumen.Cells[6, 9] = "Total Sueldo";
                resumen.Cells[6, 10] = "Alim. Quinc.";
                resumen.Cells[6, 11] = "Transporte";
                resumen.Cells[6, 12] = "Bono";
                resumen.Cells[6, 13] = "Tarjeta";
                resumen.Cells[6, 14] = "Horas Extras";
                resumen.Cells[6, 15] = "Vacaciones";
                resumen.Cells[6, 16] = "Fondo reserva";
                resumen.Cells[6, 17] = "Decimo tercero";
                resumen.Cells[6, 18] = "Decimo cuarto";
                resumen.Cells[6, 19] = "Total Ingresos";
                resumen.Cells[6, 20] = "Aportes IESS";
                resumen.Cells[6, 21] = "Prest. Hipot.";
                resumen.Cells[6, 22] = "Prest. Quiro";
                resumen.Cells[6, 23] = "Prest. Cía";
                resumen.Cells[6, 24] = "Multas";
                resumen.Cells[6, 25] = "Ext Salud";
                resumen.Cells[6, 26] = "Tarjeta";
                resumen.Cells[6, 27] = "Contribucion Solidaria";
                resumen.Cells[6, 28] = "Anticipo de quincena";
                resumen.Cells[6, 29] = "Total Egresos";
                resumen.Cells[6, 30] = "Total Recibir";

                resumen.Cells[8, 5] = "Administración";
                #endregion

                x = 1;
                y = 9;

                countTmp = 1;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Administracion
                    if (e.Area == "Administracion")
                    {
                        //Valor
                        resumen.Cells[y, x] = countTmp;
                        countTmp++;
                        resumen.Cells[y, x + 1] = e.FechaIngreso;
                        resumen.Cells[y, x + 2] = e.FechaSalida;
                        if(resumen.Cells[y,x + 2].Value.ToString() == "0")
                        {
                            resumen.Cells[y, x + 2] = "";
                        }
                        resumen.Cells[y, x + 3] = "'" + e.Cedula;
                        resumen.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        resumen.Range[resumen.Cells[y, x + 5], resumen.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 5] = e.Sueldo_Mensual;
                        resumen.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/" + LetraColumna(8) + y;
                        resumen.Cells[y, x + 7] = "='Primera Quincena'!" + LetraColumna(x + 7) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 7) + y;

                        //Total Ingresos
                        resumen.Range[resumen.Cells[y, x + 8], resumen.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 8] = "='Primera Quincena'!" + LetraColumna(x + 8) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 8) + y;
                        resumen.Cells[y, x + 9] = "='Primera Quincena'!" + LetraColumna(x + 9) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 9) + y;
                        resumen.Cells[y, x + 10] = "='Primera Quincena'!" + LetraColumna(x + 10) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 10) + y;
                        resumen.Cells[y, x + 11] = "='Primera Quincena'!" + LetraColumna(x + 11) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 11) + y;
                        resumen.Cells[y, x + 12] = "='Primera Quincena'!" + LetraColumna(x + 12) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 12) + y;
                        resumen.Cells[y, x + 13] = "='Primera Quincena'!" + LetraColumna(x + 13) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 13) + y;
                        resumen.Cells[y, x + 14] = "='Primera Quincena'!" + LetraColumna(x + 14) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 14) + y;
                        resumen.Cells[y, x + 15] = "='Primera Quincena'!" + LetraColumna(x + 15) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 15) + y;
                        resumen.Cells[y, x + 16] = "='Primera Quincena'!" + LetraColumna(x + 16) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 16) + y;
                        resumen.Cells[y, x + 17] = "='Primera Quincena'!" + LetraColumna(x + 17) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 17) + y;

                        resumen.Cells[y, x + 18] = "='Primera Quincena'!" + LetraColumna(x + 18) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 18) + y;

                        //Total Egresos
                        resumen.Range[resumen.Cells[y, x + 19], resumen.Cells[y, x + 29]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 19] = "='Primera Quincena'!" + LetraColumna(x + 19) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 19) + y;
                        resumen.Cells[y, x + 20] = "='Primera Quincena'!" + LetraColumna(x + 20) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 20) + y;
                        resumen.Cells[y, x + 21] = "='Primera Quincena'!" + LetraColumna(x + 21) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 21) + y;
                        resumen.Cells[y, x + 22] = "='Primera Quincena'!" + LetraColumna(x + 22) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 22) + y;
                        resumen.Cells[y, x + 23] = "='Primera Quincena'!" + LetraColumna(x + 23) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 23) + y;
                        resumen.Cells[y, x + 24] = "='Primera Quincena'!" + LetraColumna(x + 24) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 24) + y;
                        resumen.Cells[y, x + 25] = "='Primera Quincena'!" + LetraColumna(x + 25) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 25) + y;
                        resumen.Cells[y, x + 26] = "='Primera Quincena'!" + LetraColumna(x + 26) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 26) + y;
                        resumen.Cells[y, x + 27] = "='Primera Quincena'!" + LetraColumna(x + 27) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 27) + y;

                        resumen.Cells[y, x + 28] = "='Primera Quincena'!" + LetraColumna(x + 28) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 28) + y;

                        resumen.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }
                resumen.Cells[8 + countA + 1, 5] = "Total Administración";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    resumen.Cells[8 + countA + 1, i] = "=SUM(" + LetraColumna(i) + 9 + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                resumen.Cells[8 + countA + 3, 5] = "Operativo";

                y += 3;

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    #region Operativo
                    if (e.Area == "Operativo")
                    {
                        //Valor
                        resumen.Cells[y, x] = countTmp;
                        countTmp++;
                        resumen.Cells[y, x + 1] = e.FechaIngreso;
                        resumen.Cells[y, x + 2] = e.FechaSalida;
                        if (resumen.Cells[y, x + 2].Value.ToString() == "0")
                        {
                            resumen.Cells[y, x + 2] = "";
                        }
                        resumen.Cells[y, x + 3] = "'" + e.Cedula;
                        resumen.Cells[y, x + 4] = e.Apellido + " " + e.Nombre;
                        resumen.Range[resumen.Cells[y, x + 5], resumen.Cells[y, x + 6]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 5] = e.Sueldo_Mensual;
                        resumen.Cells[y, x + 6] = "=" + LetraColumna(6) + y + "/" + LetraColumna(8) + y;
                        resumen.Cells[y, x + 7] = "='Primera Quincena'!" + LetraColumna(x + 7) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 7) + y;

                        //Total Ingresos
                        resumen.Range[resumen.Cells[y, x + 8], resumen.Cells[y, x + 18]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 8] = "='Primera Quincena'!" + LetraColumna(x + 8) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 8) + y;
                        resumen.Cells[y, x + 9] = "='Primera Quincena'!" + LetraColumna(x + 9) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 9) + y;
                        resumen.Cells[y, x + 10] = "='Primera Quincena'!" + LetraColumna(x + 10) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 10) + y;
                        resumen.Cells[y, x + 11] = "='Primera Quincena'!" + LetraColumna(x + 11) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 11) + y;
                        resumen.Cells[y, x + 12] = "='Primera Quincena'!" + LetraColumna(x + 12) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 12) + y;
                        resumen.Cells[y, x + 13] = "='Primera Quincena'!" + LetraColumna(x + 13) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 13) + y;
                        resumen.Cells[y, x + 14] = "='Primera Quincena'!" + LetraColumna(x + 14) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 14) + y;
                        resumen.Cells[y, x + 15] = "='Primera Quincena'!" + LetraColumna(x + 15) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 15) + y;
                        resumen.Cells[y, x + 16] = "='Primera Quincena'!" + LetraColumna(x + 16) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 16) + y;
                        resumen.Cells[y, x + 17] = "='Primera Quincena'!" + LetraColumna(x + 17) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 17) + y;

                        resumen.Cells[y, x + 18] = "='Primera Quincena'!" + LetraColumna(x + 18) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 18) + y;

                        //Total Egresos
                        resumen.Range[resumen.Cells[y, x + 19], resumen.Cells[y, x + 29]].NumberFormat = formatoContabilidad;
                        resumen.Cells[y, x + 19] = "='Primera Quincena'!" + LetraColumna(x + 19) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 19) + y;
                        resumen.Cells[y, x + 20] = "='Primera Quincena'!" + LetraColumna(x + 20) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 20) + y;
                        resumen.Cells[y, x + 21] = "='Primera Quincena'!" + LetraColumna(x + 21) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 21) + y;
                        resumen.Cells[y, x + 22] = "='Primera Quincena'!" + LetraColumna(x + 22) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 22) + y;
                        resumen.Cells[y, x + 23] = "='Primera Quincena'!" + LetraColumna(x + 23) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 23) + y;
                        resumen.Cells[y, x + 24] = "='Primera Quincena'!" + LetraColumna(x + 24) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 24) + y;
                        resumen.Cells[y, x + 25] = "='Primera Quincena'!" + LetraColumna(x + 25) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 25) + y;
                        resumen.Cells[y, x + 26] = "='Primera Quincena'!" + LetraColumna(x + 26) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 26) + y;
                        resumen.Cells[y, x + 27] = "='Primera Quincena'!" + LetraColumna(x + 27) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 27) + y;

                        resumen.Cells[y, x + 28] = "='Primera Quincena'!" + LetraColumna(x + 28) + y + "+" + "'Segunda Quincena'!" + LetraColumna(x + 28) + y;

                        resumen.Cells[y, x + 29] = "=" + LetraColumna(19) + y + " - " + LetraColumna(29) + (y);

                        y++;
                    }
                    #endregion
                }

                resumen.Cells[y, x + 4] = "Total Operativo";

                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    resumen.Cells[y, i] = "=SUM(" + LetraColumna(i) + (y - countO) + ":" + LetraColumna(i) + (y - 1) + ")";
                }

                resumen.Cells[y + 2, 5] = "Total";
                for (int i = 6; i < 31; i++)
                {
                    if (i == 8)
                    {
                        continue;
                    }
                    resumen.Cells[y + 2, i] = "=" + LetraColumna(i) + (countA + 8 + 1) + " + " + LetraColumna(i) + (countA + countO + 12);
                }
                #endregion

                #region Valores Estilo
                if (checkEstilo)
                {
                    LineasCuadrosRango(false, true, false, false, 3, 1, 3, 30, resumen, 4d, Excel.XlLineStyle.xlContinuous);

                    LineasCuadros(true, true, true, true, 8, 5, optCeldas, resumen);

                    for (int i = 9; i < 30; i++)
                    {
                        LineasCuadros(true, true, true, true, 5, i, optCeldas, resumen);
                    }

                    for (int i=1; i<31 ;i++)
                    {
                        LineasCuadros(true, true, true, true, 6, i, optCeldas, resumen);
                    }


                    for (int i = 9; i < 9 + countA; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, resumen);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + 1, i, optCeldas, resumen);
                        }
                    }

                    LineasCuadros(true, true, true, true, 8 + countA + 3, 5, optCeldas, resumen);

                    for (int i = 12 + countA; i < 12 + countA + countO; i++)
                    {
                        for (int j = 1; j < 31; j++)
                        {
                            LineasCuadros(true, true, true, true, i, j, optCeldas, resumen);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 4, i, optCeldas, resumen);
                        }
                    }

                    for (int i = 5; i < 31; i++)
                    {
                        if (i != 8)
                        {
                            LineasCuadros(true, true, true, true, 8 + countA + countO + 6, i, optCeldas, resumen);
                        }
                    }

                    resumen.Range[resumen.Cells[1, 1], resumen.Cells[3, 1]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    resumen.Range[resumen.Cells[5, 9], resumen.Cells[5, 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[6, 1], resumen.Cells[6, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Cells[8, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[8 + countA + 1, 5], resumen.Cells[8 + countA + 1, 7]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[8 + countA + 1, 9], resumen.Cells[8 + countA + 1, 30]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Cells[8 + countA + 3, x + 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[y, x + 4], resumen.Cells[y, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[y, x + 8], resumen.Cells[y, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[y + 2, x + 4], resumen.Cells[y + 2, x + 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);
                    resumen.Range[resumen.Cells[y + 2, x + 8], resumen.Cells[y + 2, x + 29]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasR);

                    resumen.Cells.Font.Size = 16;

                    resumen.Columns.AutoFit();
                }
                #endregion
                #endregion

                Invoke(90, "Estado: Generando valores planilla: Roles");
                #region Roles
                int xR = 1;
                int yR = 1;
                int separador = 40;

                roles.Name = "Roles";

                for (int i = 1; i < countA + countO; i++)
                {
                    if (xR >= 1 + 8 * 3)
                    {
                        xR = 1;
                        yR += separador;
                    }

                    #region Cuadro Roles

                    #region Nombres
                    roles.Shapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + "logo.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, (xR - 1) * 60 + 10, yR * 15, 50, 30);
                    roles.Shapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + "firma.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, (xR + 5) * 60 + 10, (yR + 35) * 15, 90, 50);
                    roles.Cells[yR + 1, xR] = "ROL DE PAGOS";
                    roles.Cells[yR + 1, xR].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    roles.Cells[yR + 2, xR] = "CONSBER C.A. CONSTRUCTORA BERREZUETA";
                    roles.Cells[yR + 2, xR].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    roles.Cells[yR + 3, xR] = "Periodo: 1 de " + fecha.Value.ToString("MMMM", ci) + " de " + fecha.Value.ToString("yyyy", ci) + " - " + DiasMeses(Convert.ToInt32(fecha.Value.ToString("MM"))) + " de " + fecha.Value.ToString("MMMM", ci) + " de " + fecha.Value.ToString("yyyy", ci);
                    roles.Cells[yR + 3, xR].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    roles.Cells[yR + 5, xR] = "Codigo:";
                    roles.Cells[yR + 5, xR + 2] = "Nombre:";
                    roles.Cells[yR + 6, xR] = "Sueldo Mensual:";
                    roles.Cells[yR + 6, xR + 2] = "Dias Trabajados";
                    roles.Cells[yR + 6, xR + 5] = "Cedula:";

                    roles.Cells[yR + 7, xR] = "Ingresos";
                    roles.Cells[yR + 8, xR] = "Sueldo:";
                    roles.Cells[yR + 9, xR] = "Alimentacion:";
                    roles.Cells[yR + 10, xR] = "Transporte:";
                    roles.Cells[yR + 11, xR] = "Bonificacion:";
                    roles.Cells[yR + 12, xR] = "Tarjeta:";
                    roles.Cells[yR + 13, xR] = "Sobretiempo:";
                    roles.Cells[yR + 14, xR] = "Vacaciones:";
                    roles.Cells[yR + 15, xR] = "Fdo Reserva:";
                    roles.Cells[yR + 16, xR] = "Decimo Tercero:";
                    roles.Cells[yR + 17, xR] = "Decimo Cuarto:";
                    roles.Cells[yR + 18, xR] = "Total Ingresos:";

                    roles.Cells[yR + 19, xR] = "Egresos";
                    roles.Cells[yR + 20, xR] = "Pago 1era Quincena:";
                    roles.Cells[yR + 21, xR] = "Iess 9.45 %:";
                    roles.Cells[yR + 22, xR] = "Préstamo Hipotecario:";
                    roles.Cells[yR + 23, xR] = "Préstamo Quirografario:";
                    roles.Cells[yR + 24, xR] = "Préstamo Compañia:";
                    roles.Cells[yR + 26, xR] = "Extension de Salud:";
                    roles.Cells[yR + 25, xR] = "Multas:";
                    roles.Cells[yR + 27, xR] = "Tarjeta:";
                    roles.Cells[yR + 28, xR] = "Contribucion Solidaria:";
                    roles.Cells[yR + 29, xR] = "Anticipo Sueldo:";

                    roles.Cells[yR + 30, xR].Value = "Total Egresos:";
                    roles.Cells[yR + 6, xR + 1].NumberFormat = formatoContabilidad;
                    roles.Range[roles.Cells[yR + 8, xR], roles.Cells[yR + 30, xR]].IndentLevel = 5;
                    roles.Range[roles.Cells[yR + 8, xR + 4], roles.Cells[yR + 29, xR + 4]].NumberFormat = formatoContabilidad;
                    roles.Cells[yR + 8, xR].IndentLevel = 2;
                    roles.Cells[yR + 19, xR].IndentLevel = 2;
                    roles.Cells[yR + 31, xR].Value = "Valor a Recibir:";
                    #endregion

                    #region Valores
                    var lookup = "=VLOOKUP(" + LetraColumna(xR + 1) + (yR + 5) + ",Resumen!$" + LetraColumna(1) + "$" + 8 + ":$" + LetraColumna(30) + "$" + (countA + countO + 11) + ",";
                    roles.Cells[yR + 5, xR + 1].Value = i;
                    roles.Cells[yR + 5, xR + 3].Value = lookup + "5,0)";
                    roles.Cells[yR + 6, xR + 1].Value = lookup + "6,0)";
                    roles.Cells[yR + 6, xR + 4].Value = lookup + "8,0)";
                    roles.Cells[yR + 6, xR + 6].Value = lookup + "4,0)";
                    roles.Cells[yR + 8, xR + 4].Value = lookup + "9,0)";
                    roles.Cells[yR + 9, xR + 4].Value = lookup + "10,0)";
                    roles.Cells[yR + 10, xR + 4].Value = lookup + "11,0)";
                    roles.Cells[yR + 11, xR + 4].Value = lookup + "12,0)";
                    roles.Cells[yR + 12, xR + 4].Value = lookup + "13,0)";
                    roles.Cells[yR + 13, xR + 4].Value = lookup + "14,0)";
                    roles.Cells[yR + 14, xR + 4].Value = lookup + "15,0)";
                    roles.Cells[yR + 15, xR + 4].Value = lookup + "16,0)";
                    roles.Cells[yR + 16, xR + 4].Value = lookup + "17,0)";
                    roles.Cells[yR + 17, xR + 4].Value = lookup + "18,0)";
                    roles.Cells[yR + 18, xR + 6].Value = "=SUM(" + LetraColumna(xR + 4) + (yR + 8) + ":" + LetraColumna(xR + 4) + (yR + 17) + ")";
                    roles.Cells[yR + 20, xR + 4].Value = lookup + "20,0)";
                    roles.Cells[yR + 21, xR + 4].Value = lookup + "21,0)";
                    roles.Cells[yR + 22, xR + 4].Value = lookup + "22,0)";
                    roles.Cells[yR + 23, xR + 4].Value = lookup + "23,0)";
                    roles.Cells[yR + 24, xR + 4].Value = lookup + "24,0)";
                    roles.Cells[yR + 25, xR + 4].Value = lookup + "25,0)";
                    roles.Cells[yR + 26, xR + 4].Value = lookup + "26,0)";
                    roles.Cells[yR + 27, xR + 4].Value = lookup + "27,0)";
                    roles.Cells[yR + 28, xR + 4].Value = lookup + "28,0)";
                    roles.Cells[yR + 29, xR + 4].Value = lookup + "29,0)";
                    roles.Cells[yR + 35, xR + 1].Value = lookup + "5,0)";
                    roles.Cells[yR + 30, xR + 6].Value = "=SUM(" + LetraColumna(xR + 4) + (yR + 20) + ":" + LetraColumna(xR + 4) + (yR + 29) + ")";
                    roles.Cells[yR + 31, xR + 6].Value = "=" + LetraColumna(xR + 6) + (yR + 18) + " + " + LetraColumna(xR + 6) + (yR + 30);
                    #endregion

                    #region Merge
                    roles.Range[roles.Cells[yR + 1, xR], roles.Cells[yR + 1, xR + 6]].Merge();
                    roles.Range[roles.Cells[yR + 2, xR], roles.Cells[yR + 2, xR + 6]].Merge();
                    roles.Range[roles.Cells[yR + 3, xR], roles.Cells[yR + 3, xR + 6]].Merge();
                    roles.Range[roles.Cells[yR + 5, xR + 3], roles.Cells[yR + 5, xR + 6]].Merge();
                    roles.Range[roles.Cells[yR + 6, xR + 2], roles.Cells[yR + 6, xR + 3]].Merge();
                    roles.Range[roles.Cells[yR + 35, xR + 1], roles.Cells[yR + 35, xR + 5]].Merge();
                    #endregion

                    #region Formato
                    /*roles[LetraColumna(xR + 4) + (yR + 8)].FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 9)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 10)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 11)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 12)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 13)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 14)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 15)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 16)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 17)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 6) + (yR + 18)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 20)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 21)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 22)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 23)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 24)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 25)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 26)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 27)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 28)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 4) + (yR + 29)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 6) + (yR + 30)].First().FormatString = formatoContabilidad;
                    roles[LetraColumna(xR + 6) + (yR + 31)].First().FormatString = formatoContabilidad;*/
                    #endregion

                    xR += 8;
                }

                #endregion

                xR = 1;
                yR = 1;

                if (checkEstilo)
                {
                    roles.Columns.AutoFit();
                    for (int i = 1; i < countA + countO; i++)
                    {
                        if (xR >= (8 * 3) - 1)
                        {
                            xR = 1;
                            yR += separador;
                        }

                        LineasCuadrosRango(false, true, false, false, yR + 3, xR, yR + 3, xR + 6, roles, 4d, Excel.XlLineStyle.xlDouble);
                        LineasCuadrosRango(false, true, false, false, yR + 4, xR, yR + 4, xR + 6, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadros(true, true, false, true, yR + 5, xR, optCeldas, roles);
                        LineasCuadros(true, true, true, true, yR + 5, xR + 1, optCeldas, roles);
                        LineasCuadros(true, true, true, true, yR + 5, xR + 2, optCeldas, roles);
                        LineasCuadros(true, true, true, true, yR + 5, xR + 3, optCeldas, roles);
                        LineasCuadrosRango(false, true, false, false, yR + 5, xR + 3, yR + 5, xR + 6, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadros(true, true, false, true, yR + 6, xR, optCeldas, roles);
                        LineasCuadros(true, true, true, true, yR + 6, xR + 1, optCeldas, roles);
                        LineasCuadrosRango(true, true, false, false, yR + 6, xR + 2, yR + 6, xR + 3, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadros(true, true, true, true, yR + 6, xR + 4, optCeldas,roles);
                        LineasCuadros(true, true, true, true, yR + 6, xR + 5, optCeldas, roles);
                        LineasCuadros(true, true, true, false, yR + 6, xR + 6, optCeldas, roles);

                        LineasCuadrosRango(false, false, false, true, yR + 7, xR + 2, yR + 17, xR + 2, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadrosRango(false, false, false, true, yR + 7, xR + 5, yR + 31, xR + 5, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadrosRango(true, true, false, false, yR + 18, xR, yR + 18, xR + 6, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadrosRango(false, false, false, true, yR + 19, xR + 2, yR + 30, xR + 2, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadrosRango(true, true, false, false, yR + 30, xR, yR + 30, xR + 6, roles, Excel.XlLineStyle.xlContinuous);
                        LineasCuadrosRango(false, true, false, false, yR + 31, xR, yR + 31, xR + 6, roles, Excel.XlLineStyle.xlContinuous);

                        LineasCuadrosRango(true, false, false, false, yR + 35, xR + 1, yR + 35, xR + 5, roles, Excel.XlLineStyle.xlContinuous);

                        roles.Cells[yR + 35, xR + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        roles.Range[roles.Cells[yR, xR], roles.Cells[yR + 39, xR + 7]].Interior.Color = System.Drawing.ColorTranslator.ToOle(celdasBlanco);

                        xR += 8;
                    }
                }
                #endregion

                Dispatcher.Invoke(() => {
                    Clipboard.SetText(tmpCB);
                    BarraTransicion(95);
                });

                optCeldas.Delete();

                #region Vacaciones

                foreach (Empleados e in DTG_Empleados.Items)
                {
                    if (e.Vacaciones)
                    {
                        #region Vacaciones
                        vacaciones = xlWorkBook.Sheets.Add(misValue, numCuentas, 1, misValue) as Excel.Worksheet;
                        if(e.Apellido.Length + e.Nombre.Length < 26)
                        {
                            vacaciones.Name = "VAC " + e.Apellido + " " + e.Nombre;
                        }
                        else if(e.Apellido.Length + e.Nombre.Split(' ')[0].Length < 28)
                        {
                            vacaciones.Name = "VAC " + e.Apellido + " " + e.Nombre.Split(' ')[0];
                        }
                        else if(e.Apellido.Split(' ')[0].Length + e.Nombre.Split(' ')[0].Length < 28)
                        {
                            vacaciones.Name = "VAC " + e.Apellido.Split(' ')[0] + " " + e.Nombre.Split(' ')[0];
                        }
                        else
                        {
                            vacaciones.Name = "VAC " + e.Apellido.Split(' ') + " " + e.Nombre.Substring(0, 1);
                        }

                        vacaciones.Cells[2, 1] = "CONSBER C.A.";
                        vacaciones.Cells[3, 1] = "CALCULO DE VACACIONES";
                        vacaciones.Cells[5, 1] = "NOMBRE:";
                        vacaciones.Cells[5, 2] = e.Nombre + " " + e.Apellido;
                        vacaciones.Cells[6, 1] = "CARGO:";
                        vacaciones.Cells[7, 1] = "FECHA DE INGRESO:";
                        vacaciones.Cells[7, 3] = "PERIODO:";
                        vacaciones.Cells[8, 1] = "DESDE:";
                        vacaciones.Cells[8, 3] = "HASTA:";
                        vacaciones.Cells[10, 2] = "N°";
                        vacaciones.Cells[10, 3] = "MESES";
                        vacaciones.Cells[10, 4] = "SUELDO";
                        for (int i=1; i<=12;i++)
                        {
                            vacaciones.Cells[10 + i, 2] = i;
                            vacaciones.Cells[10 + i, 3].numberFormat = formatoFechaCorta;
                            vacaciones.Cells[10 + i, 4].numberFormat = formatoSueldosVac;
                        }
                        vacaciones.Cells[24, 2] = "TOTAL GANADO";
                        vacaciones.Cells[24, 4] = "=SUM(D11:D22)";
                        vacaciones.Cells[26, 1] = "Vacaciones";
                        vacaciones.Cells[26, 3] = "15";
                        vacaciones.Cells[26, 4] = "=ROUND(D24/24,2)";
                        vacaciones.Cells[27, 1] = "Dias Vacaciones";
                        vacaciones.Cells[27, 3] = "1";
                        vacaciones.Cells[27, 4] = "=ROUND(D26/15*C27,2)";
                        vacaciones.Cells[28, 1] = "(-) Dscto. Iess";
                        vacaciones.Cells[28, 3] = "9.45%";
                        vacaciones.Cells[28, 4] = "=D27*C28";
                        vacaciones.Cells[30, 1] = "VALOR A PAGAR";
                        vacaciones.Cells[30, 4] = "=D27-D28";
                        vacaciones.Cells[34, 1] = "APROBADO POR:";
                        vacaciones.Cells[35, 1] = "Miriam Berrezueta V.";
                        vacaciones.Cells[34, 3] = "ELABORADO POR:";
                        vacaciones.Cells[35, 3] = "Dennisse Zevallos C.";
                        vacaciones.Cells[38, 1] = "RECIBI CONFORME:";
                        vacaciones.Cells[39, 1] = e.Nombre.ToUpper() + " " + e.Apellido.ToUpper();
                        #endregion
                    }
                }
                #endregion

                Invoke(98, "Estado: Esperando respuesta de cancelacion");
                if (intentandoCancelar)
                {
                    Dispatcher.Invoke(() => {
                        finalCancelar = true;
                    });
                    _wait.WaitOne();
                    Thread.Sleep(5);
                }


                if (File.Exists(ruta))
                {
                    Dispatcher.Invoke(() => {
                        but_generar.IsEnabled = false;
                    });
                    while (FileInUse(ruta))
                    {
                        if (MessageBox.Show("No se puede guardar, cierre el archivo antes de poder guardar", "Error al guardar", MessageBoxButton.OKCancel, MessageBoxImage.Error) == MessageBoxResult.Cancel)
                        {
                            Dispatcher.Invoke(() => {
                                BarraTransicion(0);
                                txt_progreso.Content = "Estado: Cancelado";
                                check_estilo.IsEnabled = true;
                                dtp_Fecha.IsEnabled = true;
                                but_Direccion.IsEnabled = true;
                                Txtbox_Ruta.IsEnabled = true;
                                but_generar.Content = "Generar";
                                trabajando = false;
                                DTG_Empleados.IsEnabled = true;
                                DTG_Empleados_1_Q.IsEnabled = true;
                                DTG_Empleados_2_Q.IsEnabled = true;
                            });
                            return;
                        }
                    }
                }

                xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                Dispatcher.Invoke(() => {
                    BarraTransicion(100);
                });
            
                MessageBox.Show("Hoja de nominas terminado", "Listo", MessageBoxButton.OK,MessageBoxImage.Asterisk);


                Dispatcher.Invoke(() => {
                    BarraTransicion(0);
                    txt_progreso.Content = "Estado: Finalizado";
                    check_estilo.IsEnabled = true;
                    dtp_Fecha.IsEnabled = true;
                    but_Direccion.IsEnabled = true;
                    Txtbox_Ruta.IsEnabled = true;
                    but_generar.Content = "Generar";
                    but_generar.IsEnabled = true;
                    trabajando = false;
                    DTG_Empleados.IsEnabled = true;
                    DTG_Empleados_1_Q.IsEnabled = true;
                    DTG_Empleados_2_Q.IsEnabled = true;
                });
            }
            catch (Exception c)
            {
                if (c.HResult == -2146233040)
                {
                    Dispatcher.Invoke(() => {
                        BarraTransicion(0);
                        txt_progreso.Content = "Estado: Planilla cancelada";
                        check_estilo.IsEnabled = true;
                        dtp_Fecha.IsEnabled = true;
                        but_Direccion.IsEnabled = true;
                        Txtbox_Ruta.IsEnabled = true;
                        but_generar.Content = "Generar";
                        but_generar.IsEnabled = true;
                        trabajando = false;
                        DTG_Empleados.IsEnabled = true;
                        DTG_Empleados_1_Q.IsEnabled = true;
                        DTG_Empleados_2_Q.IsEnabled = true;
                    });
                }
                else
                {
                    MessageBox.Show("Error en la seccion " + posError + ". \nMensaje del error: " + c.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Dispatcher.Invoke(() => {
                        BarraTransicion(0);
                        txt_progreso.Content = "Estado: Error al generar el archivo";
                        check_estilo.IsEnabled = true;
                        dtp_Fecha.IsEnabled = true;
                        but_Direccion.IsEnabled = true;
                        Txtbox_Ruta.IsEnabled = true;
                        but_generar.Content = "Generar";
                        but_generar.IsEnabled = true;
                        trabajando = false;
                        DTG_Empleados.IsEnabled = true;
                        DTG_Empleados_1_Q.IsEnabled = true;
                        DTG_Empleados_2_Q.IsEnabled = true;
                    });
                }
            }
        }
        
        public void LineasCuadros(bool arriba, bool abajo, bool izquierda, bool derecha, int y, int x, Excel.Worksheet ws, double grosor, Excel.XlLineStyle xs)
        {
            if (arriba)
            {
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = xs;
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = grosor;
            }

            if (abajo)
            {
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = xs;
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = grosor;
            }

            if (izquierda)
            {
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = xs;
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = grosor;
            }

            if (derecha)
            {
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = xs;
                ws.Cells[y, x].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = grosor;
            }
        }

        Excel.Range R1;

        public void LineasCuadros(bool arriba, bool abajo, bool izquierda, bool derecha, int y, int x, Excel.Worksheet ws, Excel.Worksheet target)
        {
            switch (TransformarBooleanos(arriba,abajo,izquierda,derecha))
            {
                case 1:
                    //0001 Derecha
                    R1 = (Excel.Range)ws.Cells[2, 2];
                    break;
                case 2:
                    //0010 Izquierda
                    R1 = (Excel.Range)ws.Cells[2, 4];
                    break;
                case 3:
                    //0011 Izquierda y derecha
                    R1 = (Excel.Range)ws.Cells[2, 6];
                    break;
                case 4:
                    //0100 Abajo
                    R1 = (Excel.Range)ws.Cells[2, 8];
                    break;
                case 5:
                    //0101 Abajo y derecha
                    R1 = (Excel.Range)ws.Cells[2, 10];
                    break;
                case 6:
                    //0110 Abajo e izquierda
                    R1 = (Excel.Range)ws.Cells[2, 12];
                    break;
                case 7:
                    //0111 Abajo, izquierda y derecha
                    R1 = (Excel.Range)ws.Cells[2, 14];
                    break;
                case 8:
                    //1000 Arriba
                    R1 = (Excel.Range)ws.Cells[2, 16];
                    break;
                case 9:
                    //1001 Arriba y derecha
                    R1 = (Excel.Range)ws.Cells[2, 18];
                    break;
                case 10:
                    //1010 Arriba e izquierda
                    R1 = (Excel.Range)ws.Cells[2, 20];
                    break;
                case 11:
                    //1011 Arriba, izquierda y derecha
                    R1 = (Excel.Range)ws.Cells[2, 22];
                    break;
                case 12:
                    //1100 Arriba y abajo
                    R1 = (Excel.Range)ws.Cells[2, 24];
                    break;
                case 13:
                    //1101 Arriba, abajo y derecha
                    R1 = (Excel.Range)ws.Cells[2, 26];
                    break;
                case 14:
                    //1110 Arriba, abajo e izquierda
                    R1 = (Excel.Range)ws.Cells[2, 28];
                    break;
                case 15:
                    //1111 Todos
                    R1 = (Excel.Range)ws.Cells[2, 30];
                    break;
                default:
                    //0000
                    MessageBox.Show("Mal ejecutado " + TransformarBooleanos(arriba, abajo, izquierda, derecha));
                    break;
            }
            //R1.Copy();
            Excel.Range R2 = (Excel.Range)target.Cells[y, x];
            try
            {
                R2.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            }
            catch
            {
                R1.Copy();
                R2.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            }
        }

        
        public void LineasCuadrosRango(bool arriba, bool abajo, bool izquierda, bool derecha, int y1, int x1, int y2, int x2, Excel.Worksheet ws, double grosor, Excel.XlLineStyle xs)
        {
            if (arriba)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = xs;
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = grosor;
            }

            if (abajo)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = xs;
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = grosor;
            }

            if (izquierda)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = xs;
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = grosor;
            }

            if (derecha)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = xs;
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = grosor;
            }
        }
        

        public void LineasCuadrosRango(bool arriba, bool abajo, bool izquierda, bool derecha, int y1, int x1, int y2, int x2, Excel.Worksheet ws, Excel.XlLineStyle xs)
        {
            if (arriba)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = xs;
            }

            if (abajo)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = xs;
            }

            if (izquierda)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = xs;
            }

            if (derecha)
            {
                ws.Range[ws.Cells[y1, x1], ws.Cells[y2, x2]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = xs;
            }
        }
    }
}
