using System;
using System.IO;
using System.Linq;
using System.Net;

namespace actualizadorPrograma
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string version = "";
                string rutaVer;
                string VersionAplicacion = "";
                string rutaAplicacion;

                #if DEBUG
                {
                    rutaVer = @"C:\Users\Javier Berrezueta\source\repos\PruebaNominas\actualizadorPrograma\bin\Debug\ver.txt";
                }
                #else
                {
                    rutaVer = Directory.GetCurrentDirectory() + @"\ver.txt";
                }
                #endif

                using (WebClient client = new WebClient())
                {
                    #if DEBUG
                    {
                        rutaAplicacion = "https://github.com/Consber/programa-de-nominas/raw/main/PruebaNominas/bin/Debug/Programa%20de%20nominas.exe";
                        version = client.DownloadString("https://raw.githubusercontent.com/Consber/programa-de-nominas/main/PruebaNominas/bin/Debug/ver.txt");
                    }
                    #else
                    {
                        rutaAplicacion = "https://github.com/Consber/programa-de-nominas/raw/main/PruebaNominas/bin/Release/Programa%20de%20nominas.exe";
                        version = client.DownloadString("https://raw.githubusercontent.com/Consber/programa-de-nominas/main/PruebaNominas/bin/Release/ver.txt");
                    }
                    #endif
                }

                // Lee archivo de version
                try
                {
                    using (StreamReader sr = new StreamReader(rutaVer))
                    {
                        VersionAplicacion = sr.ReadToEnd();
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("No se pudo leer el archivo de version. Asuminendo version 0.0.0.0");
                }
                finally
                {
                    VersionAplicacion = "0.0.0.0";
                }

                try
                {
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
                    }   
                }
                catch
                {

                }

                if (args.Contains("-auto"))
                {
                    if (!args.Contains("-forzar"))
                    {
                        if (VersionAplicacion.CompareTo(version) < 0)
                        {
                            using (var client = new WebClient())
                            {
                                client.DownloadFile(rutaAplicacion, "Programa de nominas.exe");
                            }
                            string rutaApp = AppDomain.CurrentDomain.BaseDirectory;

                            // Escribe en un archivo la version
                            using (StreamWriter writer = new StreamWriter(rutaApp + "\\ver.txt"))
                            {
                                writer.WriteLine(version);
                            }
                            System.Diagnostics.Process.Start(rutaAplicacion);
                        }
                    }
                    else
                    {
                        using (var client = new WebClient())
                        {
                            client.DownloadFile(rutaAplicacion, "Programa de nominas.exe");
                        }
                        string ruta = AppDomain.CurrentDomain.BaseDirectory;

                        // Escribe en un archivo la version
                        using (StreamWriter writer = new StreamWriter(ruta + "\\ver.txt"))
                        {
                            writer.WriteLine(version);
                        }
                    }
                }
                else
                {
                    if (VersionAplicacion.CompareTo(version) < 0)
                    {
                        Console.WriteLine("Hay una nueva version disponible, desea actualizar la aplicacion? (Si = Y)");

                        if (Console.ReadKey().KeyChar == 'y' || Console.ReadKey().KeyChar == 'Y')
                        {
                            Console.Clear();
                            Console.WriteLine("Actualizando...");
                            ServicePointManager.Expect100Continue = true;
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                            try
                            {
                                using (var client = new WebClient())
                                {
                                    client.DownloadFile(rutaAplicacion, "Programa de nominas.exe");
                                }
                                string ruta = AppDomain.CurrentDomain.BaseDirectory;
                                // Escribe en un archivo la version
                                using (StreamWriter writer = new StreamWriter(ruta + "ver.txt"))
                                {
                                    writer.WriteLine(version);
                                }
                                Console.WriteLine("Version actualizada a {0}\nPresione cualquier tecla para terminar", version);
                                Console.ReadKey();
                            }
                            catch
                            {
                                Console.WriteLine("No se pudo actualizar la aplicacion, intentelo de nuevo mas tarde\nPresione cualquier tecla para terminar");
                                Console.ReadKey();
                            }
                        }
                        else
                        {
                            Console.Clear();
                            Console.WriteLine("No se actualizo la aplicacion\nPresione cualquier tecla para terminar");
                            Console.ReadKey();
                        }
                    }
                    else
                    {
                        Console.WriteLine("La aplicacion esta actualizada\nPresione cualquier tecla para terminar");
                        Console.ReadKey();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
