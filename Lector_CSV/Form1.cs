using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Timers;
using System.Threading;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Threading.Tasks;

namespace Lector_CSV
{
    public partial class Form1 : Form
    {
        #region hilos
        static System.Timers.Timer hilo1 = new System.Timers.Timer();
        static System.Timers.Timer hilo2 = new System.Timers.Timer();
        static System.Timers.Timer hilo3 = new System.Timers.Timer();
        static System.Timers.Timer hilo4 = new System.Timers.Timer();
        static System.Timers.Timer hilo5 = new System.Timers.Timer();
        static System.Timers.Timer hilo6 = new System.Timers.Timer();
        static System.Timers.Timer hilo7 = new System.Timers.Timer();
        static System.Timers.Timer hilo8 = new System.Timers.Timer();
        static System.Timers.Timer hilo9 = new System.Timers.Timer();
        static System.Timers.Timer hilo10 = new System.Timers.Timer();
        static System.Timers.Timer hilo11 = new System.Timers.Timer();
        static System.Timers.Timer hilo12 = new System.Timers.Timer();
        static System.Timers.Timer hilo13 = new System.Timers.Timer();
        static System.Timers.Timer hilo14 = new System.Timers.Timer();
        static System.Timers.Timer hilo15 = new System.Timers.Timer();
        static System.Timers.Timer hilo16 = new System.Timers.Timer();

        //static System.Timers.Timer hilo17 = new System.Timers.Timer();
        //static System.Timers.Timer hilo18 = new System.Timers.Timer();



        static System.Timers.Timer hilo19 = new System.Timers.Timer();
        static System.Timers.Timer hilo20 = new System.Timers.Timer();

        static System.Timers.Timer hilo21 = new System.Timers.Timer();
        static System.Timers.Timer hilo22 = new System.Timers.Timer();
        static System.Timers.Timer hilo23 = new System.Timers.Timer();
        static System.Timers.Timer hilo24 = new System.Timers.Timer();
        static System.Timers.Timer hilo25 = new System.Timers.Timer();
        static System.Timers.Timer hilo26 = new System.Timers.Timer();


        #endregion
        delegate void DisplayEstado(string msg);
        delegate void DisplayGrd(DataTable tabla);
        delegate void DisplayLimpia();
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            #region

            hilo1.Interval = 1000 * 4; //60 segundos
            hilo1.Start();
            hilo1.Elapsed += new ElapsedEventHandler(traspasar1);
            hilo1.Enabled = true;

            hilo11.Interval = 7 * 1000; //60 segundos
            hilo11.Start();
            hilo11.Elapsed += new ElapsedEventHandler(traspasar2);
            hilo11.Enabled = true;

            hilo20.Interval = 11 * 1000; //60 segundos
            hilo20.Start();
            hilo20.Elapsed += new ElapsedEventHandler(traspasar3);
            hilo20.Enabled = true;

            hilo19.Interval = 16 * 1000; //60 segundos
            hilo19.Start();
            hilo19.Elapsed += new ElapsedEventHandler(traspasar4);
            hilo19.Enabled = true;

            hilo23.Interval = 24 * 1000; //60 segundos
            hilo23.Start();
            hilo23.Elapsed += new ElapsedEventHandler(traspasar5);
            hilo23.Enabled = true;

            hilo24.Interval = 30 * 1000; //60 segundos
            hilo24.Start();
            hilo24.Elapsed += new ElapsedEventHandler(traspasar6);
            hilo24.Enabled = true;

            hilo2.Interval = 1000 * 3; //60 segundos
            hilo2.Start();
            hilo2.Elapsed += new ElapsedEventHandler(precarga1);

            hilo9.Interval = 5 * 1000; //60 segundos
            hilo9.Start();
            hilo9.Elapsed += new ElapsedEventHandler(precarga2);

            hilo21.Interval = 10 * 1000; //60 segundos
            hilo21.Start();
            hilo21.Elapsed += new ElapsedEventHandler(precarga3);

            hilo22.Interval = 15 * 1000; //60 segundos
            hilo22.Start();
            hilo22.Elapsed += new ElapsedEventHandler(precarga4);

            hilo3.Interval = 1000 * 7; //60 segundos
            hilo3.Start();
            hilo3.Elapsed += new ElapsedEventHandler(Procesa_Excel_1);

            hilo10.Interval = 15 * 1000; //3 minutos
            hilo10.Start();
            hilo10.Elapsed += new ElapsedEventHandler(Procesa_Excel_2);



            //////////////////////////

            //hilo17.Interval = 44 * 1000; //60 segundos
            //hilo17.Start();
            //hilo17.Elapsed += new ElapsedEventHandler(Procesa_Excel_3);

            //hilo18.Interval = 80 * 1000; //3 minutos
            //hilo18.Start();
            //hilo18.Elapsed += new ElapsedEventHandler(Procesa_Excel_4);



            ///////////////////////



            hilo4.Interval = 1000 * 60 * 3; //2 segundos minutos // estaba en 9 segundos lo cambie a 2 para la proxima carga 10-09-19
            hilo4.Start();
            hilo4.Elapsed += new ElapsedEventHandler(ejecutar_envioProgramado);
            hilo4.Enabled = true;

            hilo5.Interval = (1000 * 60 * 60 * 6); //6 horas
            hilo5.Start();
            hilo5.Elapsed += new ElapsedEventHandler(Elimina_Desinscritos);
            hilo5.Enabled = true;

            ////hilo6.Interval = 1000; //60 segundos
            ////hilo6.Start();
            ////hilo6.Elapsed += new ElapsedEventHandler(PrecargaSMS);

            ////hilo7.Interval = 1000; //60 segundos
            ////hilo7.Start();
            ////hilo7.Elapsed += new ElapsedEventHandler(ejecutar_envioProgramadoSMS);

            ////hilo8.Interval = 1000 * 3600; //1 hora
            ////hilo8.Start();
            ////hilo8.Elapsed += new ElapsedEventHandler(actualizaSMS);

            hilo12.Interval = 1000 * 60 * 15; //1 hora
            hilo12.Start();
            hilo12.Elapsed += new ElapsedEventHandler(ExportarTXT1); // 1 minuto

            //hilo13.Interval = 1000; //1 hora
            //hilo13.Start();
            //hilo13.Elapsed += new ElapsedEventHandler(ExportarTXT2);//3 minuto


            hilo14.Interval = 1000 * 60 * 4; //1 hora
            hilo14.Start();
            hilo14.Elapsed += new ElapsedEventHandler(Procesa_reporte_1); // 1 minuto


            hilo15.Interval = 1000 * 75 * 2; //1 hora
            hilo15.Start();
            hilo15.Elapsed += new ElapsedEventHandler(Procesa_reporte_total); // 1 minuto


            hilo16.Interval = 1000 * 105; //1 hora
            hilo16.Start();
            hilo16.Elapsed += new ElapsedEventHandler(Procesa_reporte_regunico); // 1 minuto

            hilo25.Interval = 1000 * 60 * 4; //60 segundos
            hilo25.Start();
            hilo25.Elapsed += new ElapsedEventHandler(Unir_bases);

            hilo26.Interval = 1000 * 60 * 5; //60 segundos
            hilo26.Start();
            hilo26.Elapsed += new ElapsedEventHandler(Envio_CargaFinalizada);


            #endregion
        }    
        private void Precarga()
        {

            #region
            int CadaxFilas = 0;
            int segundos = 0;
            try
            {
                CadaxFilas = Convert.ToInt32(ConfigurationSettings.AppSettings["CadaxFilas"]);
                segundos = Convert.ToInt32(ConfigurationSettings.AppSettings["segundos"]);
            }
            catch (Exception)
            {
                CadaxFilas = 2500;
                segundos = 10;
            }
            #endregion
         //   string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where id_cliente=100 and id_grupo=35433"; //not in (select id_grupo from Grupo where sms=1) order by fecha";
              string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where estado=4 and id_grupo not in (select id_grupo from Grupo where sms=1) order by fecha";

            string a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1;
            DataTable listadoCsv = ConexionCall.SqlDTable(sqlList);
            int regs = listadoCsv.Rows.Count;
            int subreg = 0;
            if (regs > 0)
            {
                for (int i = 0; i < regs; i++)
                {
                    int errores = 0;
                    int blancos = 0;
                    int procesados = 0;
                    string carpetas = listadoCsv.Rows[i]["carpetas"].ToString();
                    string archivos = listadoCsv.Rows[i]["archivos"].ToString();
                    string id_grupo = listadoCsv.Rows[i]["id_grupo"].ToString();
                    string id_cliente = listadoCsv.Rows[i]["id_cliente"].ToString();
                    string id_temp = listadoCsv.Rows[i]["id_temp"].ToString();

                    int valida = ConexionCall.devuelveValorINT("SELECT count(*) FROM temporales  where id_temp=" + id_temp + " and estado=4");

                    if (valida > 0)
                    {
                        #region

                        string direccion = carpetas + archivos;
                        this.Invoke(new DisplayEstado(cargarRuta), direccion);
                        this.Invoke(new DisplayEstado(Progreso), "Validando Correos de Documento " + archivos);
                        string txt = archivos.ToLower();
                        string fileName = carpetas + "LOG_" + txt;
                        fileName = fileName.Replace(".xlsx", ".csv");
                        fileName = fileName.Replace(".xls", ".csv");
                        //    StreamWriter writer = File.AppendText(fileName,true,Encoding.UTF8);
                        StreamWriter writer = new StreamWriter(fileName, true, Encoding.UTF8);
                        if (i == 0)
                        {
                            writer.WriteLine("E-mail;b1;c1;d1;e1;f1;g1;h1;i1;j1;k1;Error;Línea");
                        }

                        #region subProceso
                        if (!string.IsNullOrEmpty(direccion))
                        {
                            if (File.Exists(direccion))
                            {
                                string ruta = carpetas + archivos;
                                DataTable maillist = arreglaColumnas(LeerArchivoPlano(new FileInfo(ruta), true));
                                subreg = maillist.Rows.Count;

                                // this.Invoke(new DisplayGrd(cargaGrid), maillist);
                                if (subreg > 0)
                                {
                                    ConexionCall actTemp = new ConexionCall();
                                    actTemp.ejecutorBase("update Temporales set estado=6 where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);

                                    for (int j = 0; j < subreg; j++)
                                    {
                                        a1 = maillist.Rows[j][0].ToString().Trim();
                                        b1 = maillist.Rows[j][1].ToString();
                                        c1 = maillist.Rows[j][2].ToString();
                                        d1 = maillist.Rows[j][3].ToString();
                                        e1 = maillist.Rows[j][4].ToString();
                                        f1 = maillist.Rows[j][5].ToString();
                                        g1 = maillist.Rows[j][6].ToString();
                                        h1 = maillist.Rows[j][7].ToString();
                                        i1 = maillist.Rows[j][8].ToString();
                                        j1 = maillist.Rows[j][9].ToString();
                                        k1 = maillist.Rows[j][10].ToString();

                                        #region

                                        if (!string.IsNullOrEmpty(a1))
                                        {
                                            if (validarEmail(a1))
                                            {
                                                this.Invoke(new DisplayEstado(numProcesado), "E-mail (" + (j + 1) + " de " + subreg + ") " + a1 + " válido ");
                                                procesados++;
                                            }
                                            else
                                            {
                                                errores++;
                                                this.Invoke(new DisplayEstado(numProcesado), "E-mail (" + (j + 1) + " de " + subreg + ")" + a1 + " no es válido");
                                                writer.WriteLine(a1+";"+b1+";"+c1+";"+d1+";"+e1+";"+f1+";"+g1+";"+h1+";"+i1+";"+j1+";"+k1+";correo no válido;" + (j + 1));
                                                //   writer.WriteLine("Error correo no válido  " + (j + 1) + ": " + a1);
                                            }
                                        }
                                        else
                                        {
                                            blancos++;
                                            this.Invoke(new DisplayEstado(numProcesado), "E-mail vacío o nulo");
                                            writer.WriteLine(a1 +";"+b1+";"+c1+";"+d1+";"+e1+";"+f1+";"+g1+";"+h1+";"+i1+";"+j1+";"+k1+";correo vacío o nulo;" + (j + 1));

                                            //  writer.WriteLine("Error correo vacío o nulo en línea " + (j + 1));

                                        }
                                        #endregion

                                        #region
                                        if (j % CadaxFilas == 0 && j != 0)
                                        {
                                            this.Invoke(new DisplayEstado(Progreso), "Detenido  " + segundos + " segundos");
                                            Thread.Sleep(segundos * 1000);
                                        }
                                        if (j % 100 == 0 && j != 0)
                                        {
                                            try
                                            {
                                                this.Invoke(new DisplayLimpia(limpiaMsg));
                                            }
                                            catch (Exception)
                                            { }
                                        }
                                        #endregion







        //aca podríamos agregar cada fila a un nuevo csv, que sea de la mitad de registros que el original 
                                        //if (j<subreg/2){ agregar a temp1    }else{ agregar a temp2     }
        //luego guardar comu nuevos temporales con archivo distinto al ser guardado
                                        //




                                    }
                                }

                                #region envia correo una vez terminado
                                if (errores == 0)
                                {
                                    writer.WriteLine("Se procesaron " + procesados + " correos;no hubo errores en los " + subreg + " registros revisados ;0");
                                }

                                int estado = 5;

                                if (errores >= subreg)
                                {
                                    writer.WriteLine("No se encontraron registros válidos");

                                    estado = 7;

                                }



                                writer.Dispose();
                                writer.Close();

                                ConexionCall actualiza = new ConexionCall();
                                actualiza.ejecutorBase("update Temporales set estado="+estado+",fecha=getdate() where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);




                                this.Invoke(new DisplayEstado(Progreso), "Documento " + archivos + " fue cambiado a estado revisado");
                                this.Invoke(new DisplayEstado(Progreso), "      ");

                                string email = buscaCorreo(id_cliente);
                                string from = System.Configuration.ConfigurationSettings.AppSettings["From_Carga"].ToString();
                                string nombreFrom = System.Configuration.ConfigurationSettings.AppSettings["Nombre_From"].ToString();
                                string mensaje = "<p>Proceso de Precarga del documento " + archivos + " ha finalizado correctamente</p></br></br>";
                                mensaje += "<br /><br /><p style='font-family: Arial; font-size:16px' ><b>Sistema hugoo.com<br />";
                                mensaje += "<a href='mailto:soporte@hugoo.com'>soporte@hugoo.com</a><br />";
                                mensaje += "<a href='http://clientes.hugoo.com/' >clientes.hugoo.com</a></b></p>";

                                if (!string.IsNullOrEmpty(email))
                                {
                                    enviar_Correo(email, from, nombreFrom, "Aviso Proceso de Precarga", mensaje);
                                }
                                #endregion
                            }
                            else
                            {
                                this.Invoke(new DisplayEstado(Progreso), "El archivo fue borrado fuera del sistema");
                            }

                        }
                        else
                        {
                            this.Invoke(new DisplayEstado(Progreso), "Dirección está en Nulo");
                        }

                        #endregion
                        #endregion

                    }

                }
                this.Invoke(new DisplayEstado(Progreso), "Total de " + regs + " Documentos Revisados");
            }

            try
            {
                this.Invoke(new DisplayLimpia(limpiaMsg));
            }
            catch (Exception) { }

        }



        private void Precargamulti()
        {
            #region
            int CadaxFilas = 0;
            int segundos = 0;
            try
            {
                CadaxFilas = Convert.ToInt32(ConfigurationSettings.AppSettings["CadaxFilas"]);
                segundos = Convert.ToInt32(ConfigurationSettings.AppSettings["segundos"]);
            }
            catch (Exception)
            {
                CadaxFilas = 2500;
                segundos = 10;
            }
            #endregion
            string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where estado=4 and id_grupo not in (select id_grupo from Grupo where sms=1) order by fecha";

            string a1 = "";
            DataTable listadoCsv = ConexionCall.SqlDTable(sqlList);
            int regs = listadoCsv.Rows.Count;
            int subreg = 0;
            if (regs > 0)
            {
                for (int i = 0; i < regs; i++)
                {
                    int errores = 0;
                    int blancos = 0;
                    int procesados = 0;
                    string carpetas = listadoCsv.Rows[i]["carpetas"].ToString();
                    string archivos = listadoCsv.Rows[i]["archivos"].ToString();
                    string id_grupo = listadoCsv.Rows[i]["id_grupo"].ToString();
                    string id_cliente = listadoCsv.Rows[i]["id_cliente"].ToString();
                    string id_temp = listadoCsv.Rows[i]["id_temp"].ToString();

                    int valida = ConexionCall.devuelveValorINT("SELECT count(*) FROM temporales  where id_temp=" + id_temp + " and estado=4");

                    if (valida > 0)
                    {
                        #region

                        string direccion = carpetas + archivos;
                        this.Invoke(new DisplayEstado(cargarRuta), direccion);
                        this.Invoke(new DisplayEstado(Progreso), "Validando Correos de Documento " + archivos);
                        string txt = archivos.ToLower();
                        string fileName = carpetas + "LOG_" + txt;
                        //    StreamWriter writer = File.AppendText(fileName,true,Encoding.UTF8);
                        StreamWriter writer = new StreamWriter(fileName, true, Encoding.UTF8);
                        if (i == 0)
                        {
                            writer.WriteLine("E-mail;Error;Línea");
                        }

                        #region subProceso
                        if (!string.IsNullOrEmpty(direccion))
                        {
                            if (File.Exists(direccion))
                            {
                                string ruta = carpetas + archivos;
                                DataTable maillist = arreglaColumnas(LeerArchivoPlano(new FileInfo(ruta), true));
                                subreg = maillist.Rows.Count;

                                // this.Invoke(new DisplayGrd(cargaGrid), maillist);
                                if (subreg > 0)
                                {
                                    ConexionCall actTemp = new ConexionCall();
                                    actTemp.ejecutorBase("update Temporales set estado=6 where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);

                                    for (int j = 0; j < subreg; j++)
                                    {
                                        a1 = maillist.Rows[j][0].ToString();
                                        #region

                                        if (!string.IsNullOrEmpty(a1))
                                        {
                                            if (validarEmail(a1))
                                            {
                                                this.Invoke(new DisplayEstado(numProcesado), "E-mail (" + (j + 1) + " de " + subreg + ") " + a1 + " válido ");
                                                procesados++;
                                            }
                                            else
                                            {
                                                errores++;
                                                this.Invoke(new DisplayEstado(numProcesado), "E-mail (" + (j + 1) + " de " + subreg + ")" + a1 + " no es válido");
                                                writer.WriteLine(a1 + ";correo no válido;" + (j + 1));
                                                //   writer.WriteLine("Error correo no válido  " + (j + 1) + ": " + a1);
                                            }
                                        }
                                        else
                                        {
                                            blancos++;
                                            this.Invoke(new DisplayEstado(numProcesado), "E-mail vacío o nulo");
                                            writer.WriteLine(a1 + ";correo vacío o nulo;" + (j + 1));

                                            //  writer.WriteLine("Error correo vacío o nulo en línea " + (j + 1));

                                        }
                                        #endregion

                                        #region
                                        if (j % CadaxFilas == 0 && j != 0)
                                        {
                                            this.Invoke(new DisplayEstado(Progreso), "Detenido  " + segundos + " segundos");
                                            Thread.Sleep(segundos * 1000);
                                        }
                                        if (j % 100 == 0 && j != 0)
                                        {
                                            try
                                            {
                                                this.Invoke(new DisplayLimpia(limpiaMsg));
                                            }
                                            catch (Exception)
                                            { }
                                        }
                                        #endregion







                                        //aca podríamos agregar cada fila a un nuevo csv, que sea de la mitad de registros que el original 
                                        //if (j<subreg/2){ agregar a temp1    }else{ agregar a temp2     }
                                        //luego guardar comu nuevos temporales con archivo distinto al ser guardado
                                        //




                                    }
                                }

                                #region envia correo una vez terminado
                                if (errores == 0)
                                {
                                    writer.WriteLine("Se procesaron " + procesados + " correos;no hubo errores en los " + subreg + " registros revisados ;0");
                                }

                                int estado = 5;

                                if (errores >= subreg)
                                {
                                    writer.WriteLine("No se encontraron registros válidos");

                                    estado = 7;

                                }



                                writer.Dispose();
                                writer.Close();

                                ConexionCall actualiza = new ConexionCall();
                                actualiza.ejecutorBase("update Temporales set estado=" + estado + ",fecha=getdate() where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);




                                this.Invoke(new DisplayEstado(Progreso), "Documento " + archivos + " fue cambiado a estado revisado");
                                this.Invoke(new DisplayEstado(Progreso), "      ");

                                string email = buscaCorreo(id_cliente);
                                string from = System.Configuration.ConfigurationSettings.AppSettings["From_Carga"].ToString();
                                string nombreFrom = System.Configuration.ConfigurationSettings.AppSettings["Nombre_From"].ToString();
                                string mensaje = "<p>Proceso de Precarga del documento " + archivos + " ha finalizado correctamente</p></br></br>";
                                mensaje += "<br /><br /><p style='font-family: Arial; font-size:16px' ><b>Sistema hugoo.com<br />";
                                mensaje += "<a href='mailto:soporte@hugoo.com'>soporte@hugoo.com</a><br />";
                                mensaje += "<a href='http://clientes.hugoo.com/' >clientes.hugoo.com</a></b></p>";

                                if (!string.IsNullOrEmpty(email))
                                {
                                    enviar_Correo(email, from, nombreFrom, "Aviso Proceso de Precarga", mensaje);
                                }
                                #endregion
                            }
                            else
                            {
                                this.Invoke(new DisplayEstado(Progreso), "El archivo fue borrado fuera del sistema");
                            }

                        }
                        else
                        {
                            this.Invoke(new DisplayEstado(Progreso), "Dirección está en Nulo");
                        }

                        #endregion
                        #endregion

                    }

                }
                this.Invoke(new DisplayEstado(Progreso), "Total de " + regs + " Documentos Revisados");
            }

            try
            {
                this.Invoke(new DisplayLimpia(limpiaMsg));
            }
            catch (Exception) { }

        }


        private void PrecargaSMS(object myObject, EventArgs myEventArgs)
        {
            hilo6.Stop();
            #region
            int CadaxFilas = 0;
            int segundos = 0;
            try
            {
                CadaxFilas = Convert.ToInt32(ConfigurationSettings.AppSettings["CadaxFilas"]);
                segundos = Convert.ToInt32(ConfigurationSettings.AppSettings["segundos"]);
            }
            catch (Exception)
            {
                CadaxFilas = 2500;
                segundos = 10;
            }
            #endregion
            string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente FROM temporales where estado=4 and id_grupo in (select id_grupo from Grupo where sms=1) order by fecha";
            string a1 = "";
            DataTable listadoCsv = ConexionCall.SqlDTable(sqlList);
            int regs = listadoCsv.Rows.Count;
            int subreg = 0;
            if (regs > 0)
            {
                for (int i = 0; i < regs; i++)
                {
                    int errores = 0;
                    int blancos = 0;
                    int procesados = 0;
                    string carpetas = listadoCsv.Rows[i]["carpetas"].ToString();
                    string archivos = listadoCsv.Rows[i]["archivos"].ToString();
                    string id_grupo = listadoCsv.Rows[i]["id_grupo"].ToString();
                    string id_cliente = listadoCsv.Rows[i]["id_cliente"].ToString();

                    string direccion = carpetas + archivos;
                    this.Invoke(new DisplayEstado(cargarRuta), direccion);
                    this.Invoke(new DisplayEstado(Progreso), "Validando números de Documento " + archivos);
                    string txt = archivos.ToLower();
                    //  txt = txt.Replace(".csv", ".txt");
                    string fileName = carpetas + "LOG_" + txt;
                    // StreamWriter writer = File.AppendText(fileName);
                    StreamWriter writer = new StreamWriter(fileName, true, Encoding.UTF8);
                    if (i == 0)
                    {
                        writer.WriteLine("Teléfono;Error;Línea");
                    }
                   
                    #region subProceso
                    if (!string.IsNullOrEmpty(direccion))
                    {
                        if (File.Exists(direccion))
                        {
                            string ruta = carpetas + archivos;
                            DataTable maillist = arreglaColumnas(LeerArchivoPlano(new FileInfo(ruta), true));
                            subreg = maillist.Rows.Count;

                            // this.Invoke(new DisplayGrd(cargaGrid), maillist);

                            for (int j = 0; j < subreg; j++)
                            {
                                a1 = maillist.Rows[j][0].ToString();
                                #region

                                if (!string.IsNullOrEmpty(a1))
                                {
                                    if (validarNumSMS(a1))
                                    {
                                        this.Invoke(new DisplayEstado(numProcesado), "Número (" + (j + 1) + " de " + subreg + ") " + a1 + " válido ");
                                        procesados++;
                                    }
                                    else
                                    {
                                        errores++;
                                        this.Invoke(new DisplayEstado(numProcesado), "Número (" + (j + 1) + " de " + subreg + ")" + a1 + " no es válido");
                                        // writer.WriteLine("Error Número no válido línea " + (j + 1) + ": " + a1);
                                        writer.WriteLine(a1 + ";Número no válido;" + (j + 1));

                                    }
                                }
                                else
                                {
                                    blancos++;
                                    this.Invoke(new DisplayEstado(numProcesado), "Número vacío o nulo");
                                    // writer.WriteLine("Error Número vacío o nulo en línea " + (j + 1));
                                    writer.WriteLine(a1 + ";Error Número vacío o nulo;" + (j + 1));

                                }
                                #endregion

                                #region
                                if (j % CadaxFilas == 0 && j != 0)
                                {
                                    this.Invoke(new DisplayEstado(Progreso), "Detenido  " + segundos + " segundos");
                                    Thread.Sleep(segundos * 1000);
                                }
                                if (j % 100 == 0 && j != 0)
                                {
                                    try
                                    {
                                        this.Invoke(new DisplayLimpia(limpiaMsg));
                                    }
                                    catch (Exception)
                                    { }
                                }
                                #endregion
                            }

                            #region envia correo una vez terminado
                            if (errores == 0)
                            {
                                //   writer.WriteLine("Se procesaron " + procesados + " correos, no hubieron errores en los " + subreg + " registos revisados");
                                writer.WriteLine("Se procesaron " + procesados + " números;no hubo errores en los " + subreg + " registros revisados;0");

                            }
                            /*     else
                                 {
                                    // writer.WriteLine(" ");
                                   //  writer.WriteLine("Se procesaron " + procesados + " correos, con " + errores + " errores en los " + subreg + " registos revisados");

                                 }*/
                            writer.Dispose();
                            writer.Close();

                            ConexionCall actualiza = new ConexionCall();
                            actualiza.ejecutorBase("update Temporales set estado =5,fecha=getdate() where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);

                            this.Invoke(new DisplayEstado(Progreso), "Documento " + archivos + " fue cambiado a estado revisado");
                            this.Invoke(new DisplayEstado(Progreso), "      ");

                            string email = buscaCorreo(id_cliente);
                            string from = System.Configuration.ConfigurationSettings.AppSettings["From_Carga"].ToString();
                            string nombreFrom = System.Configuration.ConfigurationSettings.AppSettings["Nombre_From"].ToString();
                            string mensaje = "<p>Proceso de Precarga del documento " + archivos + " ha finalizado correctamente</p></br></br>";
                            mensaje += "<br /><br /><p style='font-family: Arial; font-size:16px' ><b>Sistema hugoo.com<br />";
                            mensaje += "<a href='mailto:soporte@hugoo.com'>soporte@hugoo.com</a><br />";
                            mensaje += "<a href='http://clientes.hugoo.com/' >clientes.hugoo.com</a></b></p>";

                            if (!string.IsNullOrEmpty(email))
                            {
                                enviar_Correo(email, from, nombreFrom, "Aviso Proceso de Precarga", mensaje);
                            }
                            #endregion
                        }
                        else
                        {
                            this.Invoke(new DisplayEstado(Progreso), "El archivo fue borrado fuera del sistema");
                        }

                    }
                    else
                    {
                        this.Invoke(new DisplayEstado(Progreso), "Dirección está en Nulo");
                    }

                    #endregion
                }
                this.Invoke(new DisplayEstado(Progreso), "Total de " + regs + " Documentos Revisados");
            }

            try
            {
                Thread.Sleep(60 * 1000);
                this.Invoke(new DisplayLimpia(limpiaMsg));
                hilo6.Enabled = true;
                //  Precarga();
            }
            catch (Exception)
            {// Precarga();
            }

        }
        private void Ejecuta_Traspaso()
        {
            #region
            int CadaxFilas = 0;
            int segundos = 0;
            try
            {
                CadaxFilas = Convert.ToInt32(ConfigurationSettings.AppSettings["CadaxFilas"]);
                segundos = Convert.ToInt32(ConfigurationSettings.AppSettings["segundos"]);
            }
            catch (Exception)
            {
                CadaxFilas = 2500;
                segundos = 10;
            }
            #endregion


            DataTable tabnew = null;
            DataTable tabnew2 = null;
            DataTable tabnew3 = null;
            //estado 0 es pendiente, estado 1 en proceso, estado 2 es procesado  

            //  string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where id_grupo =33825";
            //  string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where estado=0  and id_cliente=100";
            string sqlList = "SELECT archivos,carpetas,id_grupo,id_cliente,id_temp FROM temporales where estado=0 order by fecha";

            #region  variables
            string a1, b1, c1, d1, e1, f1, g1, h1, i1, j1, k1, l1, m1, n1, o1, p1, q1, r1, s1, t1, u1, v1, w1, x1, y1, z1, aa1, ab1, ac1, ad1, ae1, af1;
            string ag1, ah1, ai1, aj1, ak1, al1, am1;

            string an1, ao1, ap1, aq1, ar1, as1, at1, au1, av1, aw1, ax1, ay1, az1;

            string ba1, bb1, bc1, bd1, be1, bf1, bg1, bh1, bi1, bj1, bk1, bl1, bm1, bn1, bo1, bp1, bq1, br1, bs1, bt1, bu1, bv1, bw1, bx1, by1, bz1;

            string ca1, cb1, cc1, cd1, ce1, cf1, cg1, ch1, ci1, cj1, ck1, cl1, cm1, cn1, co1, cp1, cq1, cr1, cs1, ct1, cu1, cv1, cw1, cx1, cy1, cz1;

            string da1, db1, dc1, dd1, de1, df1, dg1, dh1, di1, dj1, dk1, dl1, dm1, dn1, do1, dp1, dq1, dr1, ds1, dt1, du1, dv1, dw1, dx1, dy1, dz1;

            string ea1, eb1, ec1, ed1, ee1, ef1, eg1, eh1, ei1, ej1, ek1, el1, em1, en1, eo1, ep1, eq1, er1, es1, et1, eu1, ev1, ew1, ex1, ey1, ez1;

            string fa1, fb1, fc1, fd1, fe1, ff1, fg1, fh1, fi1, fj1, fk1, fl1, fm1, fn1, fo1, fp1, fq1, fr1, fs1, ft1, fu1, fv1, fw1, fx1, fy1, fz1;

            string ga1, gb1, gc1, gd1, ge1, gf1, gg1, gh1, gi1, gj1, gk1, gl1, gm1, gn1, go1, gp1, gq1, gr1, gs1, gt1, gu1, gv1, gw1, gx1, gy1, gz1;

            string ha1, hb1, hc1, hd1, he1, hf1, hg1, hh1, hi1, hj1, hk1, hl1, hm1, hn1, ho1, hp1, hq1, hr1, hs1, ht1, hu1, hv1, hw1, hx1, hy1, hz1;

            string ia1, ib1, ic1, id1, ie1, if1, ig1, ih1, ii1, ij1, ik1, il1, im1, in1, io1, ip1, iq1, ir1, is1, it1, iu1, iv1, iw1, ix1, iy1, iz1;

            #endregion

            DataTable listadoCsv = ConexionCall.SqlDTable(sqlList);
            int regs = listadoCsv.Rows.Count;
            int subreg = 0;
            try
            {
                #region
                if (regs > 0)
                {
                    for (int i = 0; i < regs; i++)
                    {
                        string carpetas =  listadoCsv.Rows[i]["carpetas"].ToString().Trim();
                        string archivos = listadoCsv.Rows[i]["archivos"].ToString().Trim();
                        string id_grupo = listadoCsv.Rows[i]["id_grupo"].ToString().Trim();
                        string id_cliente = listadoCsv.Rows[i]["id_cliente"].ToString().Trim();
                        string id_temp = listadoCsv.Rows[i]["id_temp"].ToString().Trim();

                        string direccion = carpetas + archivos;
                        this.Invoke(new DisplayEstado(cargarRuta), direccion);
                        this.Invoke(new DisplayEstado(Progreso), "Cargando datos del Documento " + archivos);

                        int valida3 = ConexionCall.devuelveValorINT("SELECT count(*) FROM temporales where id_cliente=" + id_cliente + " and estado=1");
                        int valida = ConexionCall.devuelveValorINT("SELECT count(*) FROM temporales where id_temp=" + id_temp + " and estado=0");
                       
                        if (valida > 0 && valida3 <= 3)
                        {
                            #region subProceso
                            if (!string.IsNullOrEmpty(direccion))
                            {
                                if (File.Exists(direccion))
                                {
                                    int existe = 0;
                                    int ingresado = 0;
                                    int noconecta = 0;

                                    string ruta = carpetas + archivos;

                                    int valida2 = ConexionCall.devuelveValorINT("SELECT count(*) FROM temporales where id_temp=" + id_temp + " and estado=0");
                                    if (valida2 > 0)
                                    {

                                    ConexionCall inserta = new ConexionCall();
                                    //inserta.ejecutorBase("update Temporales set estado=1,fecha=getdate() where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);

                                    inserta.ejecutorBase("update Temporales set estado=1,fecha=getdate() where  id_temp=" + id_temp );
                                        
                                    DataTable maillist = arreglaColumnas(LeerArchivoPlano(new FileInfo(ruta), true));
                                    subreg = maillist.Rows.Count;
                                    int validaSMS = ConexionCall.devuelveValorINT("select count(*) from grupo where sms=1 and id_grupo=" + id_grupo);
                                    if (validaSMS > 0)
                                    {
                                        #region traspaso Virtual SMS
                                        for (int j = 0; j < subreg; j++)
                                        {
                                            #region
                                            a1 = maillist.Rows[j][0].ToString().Trim();
                                            b1 = maillist.Rows[j][1].ToString().Replace("'", "’").Trim();
                                            c1 = maillist.Rows[j][2].ToString().Replace("'", "’").Trim();
                                            d1 = maillist.Rows[j][3].ToString().Replace("'", "’").Trim();
                                            e1 = maillist.Rows[j][4].ToString().Replace("'", "’").Trim();
                                            f1 = maillist.Rows[j][5].ToString().Replace("'", "’").Trim();
                                            g1 = maillist.Rows[j][6].ToString().Replace("'", "’").Trim();
                                            h1 = maillist.Rows[j][7].ToString().Replace("'", "’").Trim();
                                            i1 = maillist.Rows[j][8].ToString().Replace("'", "’").Trim();
                                            j1 = maillist.Rows[j][9].ToString().Replace("'", "’").Trim();
                                            k1 = maillist.Rows[j][10].ToString().Replace("'", "’").Trim();
                                            l1 = maillist.Rows[j][11].ToString().Replace("'", "’").Trim();
                                            m1 = maillist.Rows[j][12].ToString().Replace("'", "’").Trim();
                                            n1 = maillist.Rows[j][13].ToString().Replace("'", "’").Trim();
                                            o1 = maillist.Rows[j][14].ToString().Replace("'", "’").Trim();
                                            p1 = maillist.Rows[j][15].ToString().Replace("'", "’").Trim();
                                            q1 = maillist.Rows[j][16].ToString().Replace("'", "’").Trim();
                                            r1 = maillist.Rows[j][17].ToString().Replace("'", "’").Trim();
                                            s1 = maillist.Rows[j][18].ToString().Replace("'", "’").Trim();
                                            t1 = maillist.Rows[j][19].ToString().Replace("'", "’").Trim();
                                            u1 = maillist.Rows[j][20].ToString().Replace("'", "’").Trim();
                                            v1 = maillist.Rows[j][21].ToString().Replace("'", "’").Trim();
                                            w1 = maillist.Rows[j][22].ToString().Replace("'", "’").Trim();
                                            x1 = maillist.Rows[j][23].ToString().Replace("'", "’").Trim();
                                            y1 = maillist.Rows[j][24].ToString().Replace("'", "’").Trim();
                                            z1 = maillist.Rows[j][25].ToString().Replace("'", "’").Trim();

                                            aa1 = maillist.Rows[j][26].ToString().Replace("'", "’").Trim();
                                            ab1 = maillist.Rows[j][27].ToString().Replace("'", "’").Trim();
                                            ac1 = maillist.Rows[j][28].ToString().Replace("'", "’").Trim();
                                            ad1 = maillist.Rows[j][29].ToString().Replace("'", "’").Trim();
                                            ae1 = maillist.Rows[j][30].ToString().Replace("'", "’").Trim();
                                            af1 = maillist.Rows[j][31].ToString().Replace("'", "’").Trim();
                                            ag1 = maillist.Rows[j][32].ToString().Replace("'", "’").Trim();
                                            ah1 = maillist.Rows[j][33].ToString().Replace("'", "’").Trim();
                                            ai1 = maillist.Rows[j][34].ToString().Replace("'", "’").Trim();
                                            aj1 = maillist.Rows[j][35].ToString().Replace("'", "’").Trim();
                                            ak1 = maillist.Rows[j][36].ToString().Replace("'", "’").Trim();
                                            al1 = maillist.Rows[j][37].ToString().Replace("'", "’").Trim();
                                            am1 = maillist.Rows[j][38].ToString().Replace("'", "’").Trim();
                                            an1 = maillist.Rows[j][39].ToString().Replace("'", "’").Trim();
                                            ao1 = maillist.Rows[j][40].ToString().Replace("'", "’").Trim();
                                            ap1 = maillist.Rows[j][41].ToString().Replace("'", "’").Trim();
                                            aq1 = maillist.Rows[j][42].ToString().Replace("'", "’").Trim();
                                            ar1 = maillist.Rows[j][43].ToString().Replace("'", "’").Trim();
                                            as1 = maillist.Rows[j][44].ToString().Replace("'", "’").Trim();
                                            at1 = maillist.Rows[j][45].ToString().Replace("'", "’").Trim();
                                            au1 = maillist.Rows[j][46].ToString().Replace("'", "’").Trim();
                                            av1 = maillist.Rows[j][47].ToString().Replace("'", "’").Trim();
                                            aw1 = maillist.Rows[j][48].ToString().Replace("'", "’").Trim();
                                            ax1 = maillist.Rows[j][49].ToString().Replace("'", "’").Trim();
                                            ay1 = maillist.Rows[j][50].ToString().Replace("'", "’").Trim();
                                            az1 = maillist.Rows[j][51].ToString().Replace("'", "’").Trim();

                                            ba1 = maillist.Rows[j][52].ToString().Replace("'", "’").Trim();
                                            bb1 = maillist.Rows[j][53].ToString().Replace("'", "’").Trim();
                                            bc1 = maillist.Rows[j][54].ToString().Replace("'", "’").Trim();
                                            bd1 = maillist.Rows[j][55].ToString().Replace("'", "’").Trim();
                                            be1 = maillist.Rows[j][56].ToString().Replace("'", "’").Trim();
                                            bf1 = maillist.Rows[j][57].ToString().Replace("'", "’").Trim();
                                            bg1 = maillist.Rows[j][58].ToString().Replace("'", "’").Trim();
                                            bh1 = maillist.Rows[j][59].ToString().Replace("'", "’").Trim();
                                            bi1 = maillist.Rows[j][60].ToString().Replace("'", "’").Trim();
                                            bj1 = maillist.Rows[j][61].ToString().Replace("'", "’").Trim();
                                            bk1 = maillist.Rows[j][62].ToString().Replace("'", "’").Trim();
                                            bl1 = maillist.Rows[j][63].ToString().Replace("'", "’").Trim();
                                            bm1 = maillist.Rows[j][64].ToString().Replace("'", "’").Trim();
                                            bn1 = maillist.Rows[j][65].ToString().Replace("'", "’").Trim();
                                            bo1 = maillist.Rows[j][66].ToString().Replace("'", "’").Trim();
                                            bp1 = maillist.Rows[j][67].ToString().Replace("'", "’").Trim();
                                            bq1 = maillist.Rows[j][68].ToString().Replace("'", "’").Trim();
                                            br1 = maillist.Rows[j][69].ToString().Replace("'", "’").Trim();
                                            bs1 = maillist.Rows[j][70].ToString().Replace("'", "’").Trim();
                                            bt1 = maillist.Rows[j][71].ToString().Replace("'", "’").Trim();
                                            bu1 = maillist.Rows[j][72].ToString().Replace("'", "’").Trim();
                                            bv1 = maillist.Rows[j][73].ToString().Replace("'", "’").Trim();
                                            bw1 = maillist.Rows[j][74].ToString().Replace("'", "’").Trim();
                                            bx1 = maillist.Rows[j][75].ToString().Replace("'", "’").Trim();
                                            by1 = maillist.Rows[j][76].ToString().Replace("'", "’").Trim();
                                            bz1 = maillist.Rows[j][77].ToString().Replace("'", "’").Trim();

                                            ca1 = maillist.Rows[j][78].ToString().Replace("'", "’").Trim();
                                            cb1 = maillist.Rows[j][79].ToString().Replace("'", "’").Trim();
                                            cc1 = maillist.Rows[j][80].ToString().Replace("'", "’").Trim();
                                            cd1 = maillist.Rows[j][81].ToString().Replace("'", "’").Trim();
                                            ce1 = maillist.Rows[j][82].ToString().Replace("'", "’").Trim();
                                            cf1 = maillist.Rows[j][83].ToString().Replace("'", "’").Trim();
                                            cg1 = maillist.Rows[j][84].ToString().Replace("'", "’").Trim();
                                            ch1 = maillist.Rows[j][85].ToString().Replace("'", "’").Trim();
                                            ci1 = maillist.Rows[j][86].ToString().Replace("'", "’").Trim();
                                            cj1 = maillist.Rows[j][87].ToString().Replace("'", "’").Trim();
                                            ck1 = maillist.Rows[j][88].ToString().Replace("'", "’").Trim();
                                            cl1 = maillist.Rows[j][89].ToString().Replace("'", "’").Trim();
                                            cm1 = maillist.Rows[j][90].ToString().Replace("'", "’").Trim();
                                            cn1 = maillist.Rows[j][91].ToString().Replace("'", "’").Trim();
                                            co1 = maillist.Rows[j][92].ToString().Replace("'", "’").Trim();
                                            cp1 = maillist.Rows[j][93].ToString().Replace("'", "’").Trim();
                                            cq1 = maillist.Rows[j][94].ToString().Replace("'", "’").Trim();
                                            cr1 = maillist.Rows[j][95].ToString().Replace("'", "’").Trim();
                                            cs1 = maillist.Rows[j][96].ToString().Replace("'", "’").Trim();
                                            ct1 = maillist.Rows[j][97].ToString().Replace("'", "’").Trim();
                                            cu1 = maillist.Rows[j][98].ToString().Replace("'", "’").Trim();
                                            cv1 = maillist.Rows[j][99].ToString().Replace("'", "’").Trim();
                                            cw1 = maillist.Rows[j][100].ToString().Replace("'", "’").Trim();
                                            cx1 = maillist.Rows[j][101].ToString().Replace("'", "’").Trim();
                                            cy1 = maillist.Rows[j][102].ToString().Replace("'", "’").Trim();
                                            cz1 = maillist.Rows[j][103].ToString().Replace("'", "’").Trim();

                                            da1 = maillist.Rows[j][104].ToString().Replace("'", "’").Trim();
                                            db1 = maillist.Rows[j][105].ToString().Replace("'", "’").Trim();
                                            dc1 = maillist.Rows[j][106].ToString().Replace("'", "’").Trim();
                                            dd1 = maillist.Rows[j][107].ToString().Replace("'", "’").Trim();
                                            de1 = maillist.Rows[j][108].ToString().Replace("'", "’").Trim();
                                            df1 = maillist.Rows[j][109].ToString().Replace("'", "’").Trim();
                                            dg1 = maillist.Rows[j][110].ToString().Replace("'", "’").Trim();
                                            dh1 = maillist.Rows[j][111].ToString().Replace("'", "’").Trim();
                                            di1 = maillist.Rows[j][112].ToString().Replace("'", "’").Trim();
                                            dj1 = maillist.Rows[j][113].ToString().Replace("'", "’").Trim();
                                            dk1 = maillist.Rows[j][114].ToString().Replace("'", "’").Trim();
                                            dl1 = maillist.Rows[j][115].ToString().Replace("'", "’").Trim();
                                            dm1 = maillist.Rows[j][116].ToString().Replace("'", "’").Trim();
                                            dn1 = maillist.Rows[j][117].ToString().Replace("'", "’").Trim();
                                            do1 = maillist.Rows[j][118].ToString().Replace("'", "’").Trim();
                                            dp1 = maillist.Rows[j][119].ToString().Replace("'", "’").Trim();
                                            dq1 = maillist.Rows[j][120].ToString().Replace("'", "’").Trim();
                                            dr1 = maillist.Rows[j][121].ToString().Replace("'", "’").Trim();
                                            ds1 = maillist.Rows[j][122].ToString().Replace("'", "’").Trim();
                                            dt1 = maillist.Rows[j][123].ToString().Replace("'", "’").Trim();
                                            du1 = maillist.Rows[j][124].ToString().Replace("'", "’").Trim();
                                            dv1 = maillist.Rows[j][125].ToString().Replace("'", "’").Trim();
                                            dw1 = maillist.Rows[j][126].ToString().Replace("'", "’").Trim();
                                            dx1 = maillist.Rows[j][127].ToString().Replace("'", "’").Trim();
                                            dy1 = maillist.Rows[j][128].ToString().Replace("'", "’").Trim();
                                            dz1 = maillist.Rows[j][129].ToString().Replace("'", "’").Trim();

                                            ea1 = maillist.Rows[j][130].ToString().Replace("'", "’").Trim();
                                            eb1 = maillist.Rows[j][131].ToString().Replace("'", "’").Trim();
                                            ec1 = maillist.Rows[j][132].ToString().Replace("'", "’").Trim();
                                            ed1 = maillist.Rows[j][133].ToString().Replace("'", "’").Trim();
                                            ee1 = maillist.Rows[j][134].ToString().Replace("'", "’").Trim();
                                            ef1 = maillist.Rows[j][135].ToString().Replace("'", "’").Trim();
                                            eg1 = maillist.Rows[j][136].ToString().Replace("'", "’").Trim();
                                            eh1 = maillist.Rows[j][137].ToString().Replace("'", "’").Trim();
                                            ei1 = maillist.Rows[j][138].ToString().Replace("'", "’").Trim();
                                            ej1 = maillist.Rows[j][139].ToString().Replace("'", "’").Trim();
                                            ek1 = maillist.Rows[j][140].ToString().Replace("'", "’").Trim();
                                            el1 = maillist.Rows[j][141].ToString().Replace("'", "’").Trim();
                                            em1 = maillist.Rows[j][142].ToString().Replace("'", "’").Trim();
                                            en1 = maillist.Rows[j][143].ToString().Replace("'", "’").Trim();
                                            eo1 = maillist.Rows[j][144].ToString().Replace("'", "’").Trim();
                                            ep1 = maillist.Rows[j][145].ToString().Replace("'", "’").Trim();
                                            eq1 = maillist.Rows[j][146].ToString().Replace("'", "’").Trim();
                                            er1 = maillist.Rows[j][147].ToString().Replace("'", "’").Trim();
                                            es1 = maillist.Rows[j][148].ToString().Replace("'", "’").Trim();
                                            et1 = maillist.Rows[j][149].ToString().Replace("'", "’").Trim();
                                            eu1 = maillist.Rows[j][150].ToString().Replace("'", "’").Trim();
                                            ev1 = maillist.Rows[j][151].ToString().Replace("'", "’").Trim();
                                            ew1 = maillist.Rows[j][152].ToString().Replace("'", "’").Trim();
                                            ex1 = maillist.Rows[j][153].ToString().Replace("'", "’").Trim();
                                            ey1 = maillist.Rows[j][154].ToString().Replace("'", "’").Trim();
                                            ez1 = maillist.Rows[j][155].ToString().Replace("'", "’").Trim();

                                            fa1 = maillist.Rows[j][156].ToString().Replace("'", "’").Trim();
                                            fb1 = maillist.Rows[j][157].ToString().Replace("'", "’").Trim();
                                            fc1 = maillist.Rows[j][158].ToString().Replace("'", "’").Trim();
                                            fd1 = maillist.Rows[j][159].ToString().Replace("'", "’").Trim();
                                            fe1 = maillist.Rows[j][160].ToString().Replace("'", "’").Trim();
                                            ff1 = maillist.Rows[j][161].ToString().Replace("'", "’").Trim();
                                            fg1 = maillist.Rows[j][162].ToString().Replace("'", "’").Trim();
                                            fh1 = maillist.Rows[j][163].ToString().Replace("'", "’").Trim();
                                            fi1 = maillist.Rows[j][164].ToString().Replace("'", "’").Trim();
                                            fj1 = maillist.Rows[j][165].ToString().Replace("'", "’").Trim();
                                            fk1 = maillist.Rows[j][166].ToString().Replace("'", "’").Trim();
                                            fl1 = maillist.Rows[j][167].ToString().Replace("'", "’").Trim();
                                            fm1 = maillist.Rows[j][168].ToString().Replace("'", "’").Trim();
                                            fn1 = maillist.Rows[j][169].ToString().Replace("'", "’").Trim();
                                            fo1 = maillist.Rows[j][170].ToString().Replace("'", "’").Trim();
                                            fp1 = maillist.Rows[j][171].ToString().Replace("'", "’").Trim();
                                            fq1 = maillist.Rows[j][172].ToString().Replace("'", "’").Trim();
                                            fr1 = maillist.Rows[j][173].ToString().Replace("'", "’").Trim();
                                            fs1 = maillist.Rows[j][174].ToString().Replace("'", "’").Trim();
                                            ft1 = maillist.Rows[j][175].ToString().Replace("'", "’").Trim();
                                            fu1 = maillist.Rows[j][176].ToString().Replace("'", "’").Trim();
                                            fv1 = maillist.Rows[j][177].ToString().Replace("'", "’").Trim();
                                            fw1 = maillist.Rows[j][178].ToString().Replace("'", "’").Trim();
                                            fx1 = maillist.Rows[j][179].ToString().Replace("'", "’").Trim();
                                            fy1 = maillist.Rows[j][180].ToString().Replace("'", "’").Trim();
                                            fz1 = maillist.Rows[j][181].ToString().Replace("'", "’").Trim();

                                            ga1 = maillist.Rows[j][182].ToString().Replace("'", "’").Trim();
                                            gb1 = maillist.Rows[j][183].ToString().Replace("'", "’").Trim();
                                            gc1 = maillist.Rows[j][184].ToString().Replace("'", "’").Trim();
                                            gd1 = maillist.Rows[j][185].ToString().Replace("'", "’").Trim();
                                            ge1 = maillist.Rows[j][186].ToString().Replace("'", "’").Trim();
                                            gf1 = maillist.Rows[j][187].ToString().Replace("'", "’").Trim();
                                            gg1 = maillist.Rows[j][188].ToString().Replace("'", "’").Trim();
                                            gh1 = maillist.Rows[j][189].ToString().Replace("'", "’").Trim();
                                            gi1 = maillist.Rows[j][190].ToString().Replace("'", "’").Trim();
                                            gj1 = maillist.Rows[j][191].ToString().Replace("'", "’").Trim();
                                            gk1 = maillist.Rows[j][192].ToString().Replace("'", "’").Trim();
                                            gl1 = maillist.Rows[j][193].ToString().Replace("'", "’").Trim();
                                            gm1 = maillist.Rows[j][194].ToString().Replace("'", "’").Trim();
                                            gn1 = maillist.Rows[j][195].ToString().Replace("'", "’").Trim();
                                            go1 = maillist.Rows[j][196].ToString().Replace("'", "’").Trim();
                                            gp1 = maillist.Rows[j][197].ToString().Replace("'", "’").Trim();
                                            gq1 = maillist.Rows[j][198].ToString().Replace("'", "’").Trim();
                                            gr1 = maillist.Rows[j][199].ToString().Replace("'", "’").Trim();
                                            gs1 = maillist.Rows[j][200].ToString().Replace("'", "’").Trim();
                                            gt1 = maillist.Rows[j][201].ToString().Replace("'", "’").Trim();
                                            gu1 = maillist.Rows[j][202].ToString().Replace("'", "’").Trim();
                                            gv1 = maillist.Rows[j][203].ToString().Replace("'", "’").Trim();
                                            gw1 = maillist.Rows[j][204].ToString().Replace("'", "’").Trim();
                                            gx1 = maillist.Rows[j][205].ToString().Replace("'", "’").Trim();
                                            gy1 = maillist.Rows[j][206].ToString().Replace("'", "’").Trim();
                                            gz1 = maillist.Rows[j][207].ToString().Replace("'", "’").Trim();

                                            ha1 = maillist.Rows[j][208].ToString().Replace("'", "’").Trim();
                                            hb1 = maillist.Rows[j][209].ToString().Replace("'", "’").Trim();
                                            hc1 = maillist.Rows[j][210].ToString().Replace("'", "’").Trim();
                                            hd1 = maillist.Rows[j][211].ToString().Replace("'", "’").Trim();
                                            he1 = maillist.Rows[j][212].ToString().Replace("'", "’").Trim();
                                            hf1 = maillist.Rows[j][213].ToString().Replace("'", "’").Trim();
                                            hg1 = maillist.Rows[j][214].ToString().Replace("'", "’").Trim();
                                            hh1 = maillist.Rows[j][215].ToString().Replace("'", "’").Trim();
                                            hi1 = maillist.Rows[j][216].ToString().Replace("'", "’").Trim();
                                            hj1 = maillist.Rows[j][217].ToString().Replace("'", "’").Trim();
                                            hk1 = maillist.Rows[j][218].ToString().Replace("'", "’").Trim();
                                            hl1 = maillist.Rows[j][219].ToString().Replace("'", "’").Trim();
                                            hm1 = maillist.Rows[j][220].ToString().Replace("'", "’").Trim();
                                            hn1 = maillist.Rows[j][221].ToString().Replace("'", "’").Trim();
                                            ho1 = maillist.Rows[j][222].ToString().Replace("'", "’").Trim();
                                            hp1 = maillist.Rows[j][223].ToString().Replace("'", "’").Trim();
                                            hq1 = maillist.Rows[j][224].ToString().Replace("'", "’").Trim();
                                            hr1 = maillist.Rows[j][225].ToString().Replace("'", "’").Trim();
                                            hs1 = maillist.Rows[j][226].ToString().Replace("'", "’").Trim();
                                            ht1 = maillist.Rows[j][227].ToString().Replace("'", "’").Trim();
                                            hu1 = maillist.Rows[j][228].ToString().Replace("'", "’").Trim();
                                            hv1 = maillist.Rows[j][229].ToString().Replace("'", "’").Trim();
                                            hw1 = maillist.Rows[j][230].ToString().Replace("'", "’").Trim();
                                            hx1 = maillist.Rows[j][231].ToString().Replace("'", "’").Trim();
                                            hy1 = maillist.Rows[j][232].ToString().Replace("'", "’").Trim();
                                            hz1 = maillist.Rows[j][233].ToString().Replace("'", "’").Trim();

                                            ia1 = maillist.Rows[j][234].ToString().Replace("'", "’").Trim();
                                            ib1 = maillist.Rows[j][235].ToString().Replace("'", "’").Trim();
                                            ic1 = maillist.Rows[j][236].ToString().Replace("'", "’").Trim();
                                            id1 = maillist.Rows[j][237].ToString().Replace("'", "’").Trim();
                                            ie1 = maillist.Rows[j][238].ToString().Replace("'", "’").Trim();
                                            if1 = maillist.Rows[j][239].ToString().Replace("'", "’").Trim();
                                            ig1 = maillist.Rows[j][240].ToString().Replace("'", "’").Trim();
                                            ih1 = maillist.Rows[j][241].ToString().Replace("'", "’").Trim();
                                            ii1 = maillist.Rows[j][242].ToString().Replace("'", "’").Trim();
                                            ij1 = maillist.Rows[j][243].ToString().Replace("'", "’").Trim();
                                            ik1 = maillist.Rows[j][244].ToString().Replace("'", "’").Trim();
                                            il1 = maillist.Rows[j][245].ToString().Replace("'", "’").Trim();
                                            im1 = maillist.Rows[j][246].ToString().Replace("'", "’").Trim();
                                            in1 = maillist.Rows[j][247].ToString().Replace("'", "’").Trim();
                                            io1 = maillist.Rows[j][248].ToString().Replace("'", "’").Trim();
                                            ip1 = maillist.Rows[j][249].ToString().Replace("'", "’").Trim();
                                            iq1 = maillist.Rows[j][250].ToString().Replace("'", "’").Trim();
                                            ir1 = maillist.Rows[j][251].ToString().Replace("'", "’").Trim();
                                            is1 = maillist.Rows[j][252].ToString().Replace("'", "’").Trim();
                                            it1 = maillist.Rows[j][253].ToString().Replace("'", "’").Trim();
                                            iu1 = maillist.Rows[j][254].ToString().Replace("'", "’").Trim();
                                            iv1 = maillist.Rows[j][255].ToString().Replace("'", "’").Trim();
                                            iw1 = maillist.Rows[j][256].ToString().Replace("'", "’").Trim();
                                            ix1 = maillist.Rows[j][257].ToString().Replace("'", "’").Trim();
                                            iy1 = maillist.Rows[j][258].ToString().Replace("'", "’").Trim();
                                            iz1 = maillist.Rows[j][259].ToString().Replace("'", "’").Trim();

                                            #endregion
                                            this.Invoke(new DisplayEstado(numProcesado), "Documento " + archivos + " insertando registro en tabla virtual " + (j + 1) + " de " + subreg);
                                            #region


                                            string insertaSql = "INSERT INTO Email(Id_Grupo, a1,b1,activo,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1,aa1,ab1,ac1,ad1,ae1,af1,";
                                            insertaSql += "  ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                            insertaSql += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                            insertaSql += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                            insertaSql += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                            insertaSql += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                            insertaSql += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                            insertaSql += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                            insertaSql += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                            insertaSql += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";
                                            insertaSql += ") VALUES(" + id_grupo + ", '" + a1 + "', '" + b1 + "',1";
                                            insertaSql += " ,'" + c1 + "','" + d1 + "', '" + e1 + "','" + f1 + "', '" + g1 + "','" + h1 + "', '" + i1 + "','" + j1 + "', '" + k1 + "','" + l1 + "'";
                                            insertaSql += " ,'" + m1 + "','" + n1 + "','" + o1 + "', '" + p1 + "','" + q1 + "', '" + r1 + "','" + s1 + "', '" + t1 + "','" + u1 + "'";
                                            insertaSql += " ,'" + v1 + "','" + w1 + "', '" + x1 + "','" + y1 + "', '" + z1 + "','" + aa1 + "', '" + ab1 + "','" + ac1 + "', '" + ad1 + "','" + ae1 + "'";
                                            insertaSql += " ,'" + af1 + "','" + ag1 + "', '" + ah1 + "','" + ai1 + "', '" + aj1 + "','" + ak1 + "', '" + al1 + "','" + am1 + "'";
                                            insertaSql += " ,'" + an1 + "','" + ao1 + "', '" + ap1 + "','" + aq1 + "', '" + ar1 + "','" + as1 + "', '" + at1 + "','" + au1 + "','" + av1 + "', '" + aw1 + "','" + ax1 + "', '" + ay1 + "','" + az1 + "'";
                                            insertaSql += ",'" + ba1 + "', '" + bb1 + "','" + bc1 + "', '" + bd1 + "','" + be1 + "' ,'" + bf1 + "','" + bg1 + "', '" + bh1 + "','" + bi1 + "', '" + bj1 + "','" + bk1 + "', '" + bl1 + "','" + bm1 + "'";
                                            insertaSql += ",'" + bn1 + "', '" + bo1 + "','" + bp1 + "', '" + bq1 + "','" + br1 + "' ,'" + bs1 + "','" + bt1 + "', '" + bu1 + "','" + bv1 + "', '" + bw1 + "','" + bx1 + "', '" + by1 + "','" + bz1 + "'";
                                            insertaSql += ",'" + ca1 + "', '" + cb1 + "','" + cc1 + "', '" + cd1 + "','" + ce1 + "' ,'" + cf1 + "','" + cg1 + "', '" + ch1 + "','" + ci1 + "', '" + cj1 + "','" + ck1 + "', '" + cl1 + "','" + cm1 + "'";
                                            insertaSql += ",'" + cn1 + "', '" + co1 + "','" + cp1 + "', '" + cq1 + "','" + cr1 + "' ,'" + cs1 + "','" + ct1 + "', '" + cu1 + "','" + cv1 + "', '" + cw1 + "','" + cx1 + "', '" + cy1 + "','" + cz1 + "'";
                                            insertaSql += ",'" + da1 + "', '" + db1 + "','" + dc1 + "', '" + dd1 + "','" + de1 + "' ,'" + df1 + "','" + dg1 + "', '" + dh1 + "','" + di1 + "', '" + dj1 + "','" + dk1 + "', '" + dl1 + "','" + dm1 + "'";
                                            insertaSql += ",'" + dn1 + "', '" + do1 + "','" + dp1 + "', '" + dq1 + "','" + dr1 + "' ,'" + ds1 + "','" + dt1 + "', '" + du1 + "','" + dv1 + "', '" + dw1 + "','" + dx1 + "', '" + dy1 + "','" + dz1 + "'";
                                            insertaSql += ",'" + ea1 + "', '" + eb1 + "','" + ec1 + "', '" + ed1 + "','" + ee1 + "' ,'" + ef1 + "','" + eg1 + "', '" + eh1 + "','" + ei1 + "', '" + ej1 + "','" + ek1 + "', '" + el1 + "','" + em1 + "'";
                                            insertaSql += ",'" + en1 + "', '" + eo1 + "','" + ep1 + "', '" + eq1 + "','" + er1 + "' ,'" + es1 + "','" + et1 + "', '" + eu1 + "','" + ev1 + "', '" + ew1 + "','" + ex1 + "', '" + ey1 + "','" + ez1 + "'";
                                            insertaSql += ",'" + fa1 + "', '" + fb1 + "','" + fc1 + "', '" + fd1 + "','" + fe1 + "' ,'" + ff1 + "','" + fg1 + "', '" + fh1 + "','" + fi1 + "', '" + fj1 + "','" + fk1 + "', '" + fl1 + "','" + fm1 + "'";
                                            insertaSql += ",'" + fn1 + "', '" + fo1 + "','" + fp1 + "', '" + fq1 + "','" + fr1 + "' ,'" + fs1 + "','" + ft1 + "', '" + fu1 + "','" + fv1 + "', '" + fw1 + "','" + fx1 + "', '" + fy1 + "','" + fz1 + "'";
                                            insertaSql += ",'" + ga1 + "', '" + gb1 + "','" + gc1 + "', '" + gd1 + "','" + ge1 + "' ,'" + gf1 + "','" + gg1 + "', '" + gh1 + "','" + gi1 + "', '" + gj1 + "','" + gk1 + "', '" + gl1 + "','" + gm1 + "'";
                                            insertaSql += ",'" + gn1 + "', '" + go1 + "','" + gp1 + "', '" + gq1 + "','" + gr1 + "' ,'" + gs1 + "','" + gt1 + "', '" + gu1 + "','" + gv1 + "', '" + gw1 + "','" + gx1 + "', '" + gy1 + "','" + gz1 + "'";
                                            insertaSql += ",'" + ha1 + "', '" + hb1 + "','" + hc1 + "', '" + hd1 + "','" + he1 + "' ,'" + hf1 + "','" + hg1 + "', '" + hh1 + "','" + hi1 + "', '" + hj1 + "','" + hk1 + "', '" + hl1 + "','" + hm1 + "'";
                                            insertaSql += ",'" + hn1 + "', '" + ho1 + "','" + hp1 + "', '" + hq1 + "','" + hr1 + "' ,'" + hs1 + "','" + ht1 + "', '" + hu1 + "','" + hv1 + "', '" + hw1 + "','" + hx1 + "', '" + hy1 + "','" + hz1 + "'";
                                            insertaSql += ",'" + ia1 + "', '" + ib1 + "','" + ic1 + "', '" + id1 + "','" + ie1 + "' ,'" + if1 + "','" + ig1 + "', '" + ih1 + "','" + ii1 + "', '" + ij1 + "','" + ik1 + "', '" + il1 + "','" + im1 + "'";
                                            insertaSql += ",'" + in1 + "', '" + io1 + "','" + ip1 + "', '" + iq1 + "','" + ir1 + "' ,'" + is1 + "','" + it1 + "', '" + iu1 + "','" + iv1 + "', '" + iw1 + "','" + ix1 + "', '" + iy1 + "','" + iz1 + "'";

                                            insertaSql += "  ) ";





                                            //string insertaSql = "INSERT INTO Email_virtual (Id_Grupo, a1,b1,activo,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1,aa1,ab1,ac1,ad1,ae1,af1,";
                                            //insertaSql += "  ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                            //insertaSql += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                            //insertaSql += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                            //insertaSql += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                            //insertaSql += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                            //insertaSql += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                            //insertaSql += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                            //insertaSql += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                            //insertaSql += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";

                                            //insertaSql += ") VALUES(" + id_grupo + ", '" + a1 + "', '" + b1 + "','1'";
                                            //insertaSql += " ,'" + c1 + "','" + d1 + "', '" + e1 + "','" + f1 + "', '" + g1 + "','" + h1 + "', '" + i1 + "','" + j1 + "', '" + k1 + "','" + l1 + "'";
                                            //insertaSql += " ,'" + m1 + "','" + n1 + "','" + o1 + "', '" + p1 + "','" + q1 + "', '" + r1 + "','" + s1 + "', '" + t1 + "','" + u1 + "'";
                                            //insertaSql += " ,'" + v1 + "','" + w1 + "', '" + x1 + "','" + y1 + "', '" + z1 + "','" + aa1 + "', '" + ab1 + "','" + ac1 + "', '" + ad1 + "','" + ae1 + "'";
                                            //insertaSql += " ,'" + af1 + "','" + ag1 + "', '" + ah1 + "','" + ai1 + "', '" + aj1 + "','" + ak1 + "', '" + al1 + "','" + am1 + "'";
                                            //insertaSql += " ,'" + an1 + "','" + ao1 + "', '" + ap1 + "','" + aq1 + "', '" + ar1 + "','" + as1 + "', '" + at1 + "','" + au1 + "','" + av1 + "', '" + aw1 + "','" + ax1 + "', '" + ay1 + "','" + az1 + "'";
                                            //insertaSql += "','" + ba1 + "', '" + bb1 + "','" + bc1 + "', '" + bd1 + "','" + be1 + "' ,'" + bf1 + "','" + bg1 + "', '" + bh1 + "','" + bi1 + "', '" + bj1 + "','" + bk1 + "', '" + bl1 + "','" + bm1 + "'";
                                            //insertaSql += "','" + bn1 + "', '" + bo1 + "','" + bp1 + "', '" + bq1 + "','" + br1 + "' ,'" + bs1 + "','" + bt1 + "', '" + bu1 + "','" + bv1 + "', '" + bw1 + "','" + bx1 + "', '" + by1 + "','" + bz1 + "'";
                                            //insertaSql += "','" + ca1 + "', '" + cb1 + "','" + cc1 + "', '" + cd1 + "','" + ce1 + "' ,'" + cf1 + "','" + cg1 + "', '" + ch1 + "','" + ci1 + "', '" + cj1 + "','" + ck1 + "', '" + cl1 + "','" + cm1 + "'";
                                            //insertaSql += "','" + cn1 + "', '" + co1 + "','" + cp1 + "', '" + cq1 + "','" + cr1 + "' ,'" + cs1 + "','" + ct1 + "', '" + cu1 + "','" + cv1 + "', '" + cw1 + "','" + cx1 + "', '" + cy1 + "','" + cz1 + "'";
                                            //insertaSql += "','" + da1 + "', '" + db1 + "','" + dc1 + "', '" + dd1 + "','" + de1 + "' ,'" + df1 + "','" + dg1 + "', '" + dh1 + "','" + di1 + "', '" + dj1 + "','" + dk1 + "', '" + dl1 + "','" + dm1 + "'";
                                            //insertaSql += "','" + dn1 + "', '" + do1 + "','" + dp1 + "', '" + dq1 + "','" + dr1 + "' ,'" + ds1 + "','" + dt1 + "', '" + du1 + "','" + dv1 + "', '" + dw1 + "','" + dx1 + "', '" + dy1 + "','" + dz1 + "'";
                                            //insertaSql += "','" + ea1 + "', '" + eb1 + "','" + ec1 + "', '" + ed1 + "','" + ee1 + "' ,'" + ef1 + "','" + eg1 + "', '" + eh1 + "','" + ei1 + "', '" + ej1 + "','" + ek1 + "', '" + el1 + "','" + em1 + "'";
                                            //insertaSql += "','" + en1 + "', '" + eo1 + "','" + ep1 + "', '" + eq1 + "','" + er1 + "' ,'" + es1 + "','" + et1 + "', '" + eu1 + "','" + ev1 + "', '" + ew1 + "','" + ex1 + "', '" + ey1 + "','" + ez1 + "'";
                                            //insertaSql += "','" + fa1 + "', '" + fb1 + "','" + fc1 + "', '" + fd1 + "','" + fe1 + "' ,'" + ff1 + "','" + fg1 + "', '" + fh1 + "','" + fi1 + "', '" + fj1 + "','" + fk1 + "', '" + fl1 + "','" + fm1 + "'";
                                            //insertaSql += "','" + fn1 + "', '" + fo1 + "','" + fp1 + "', '" + fq1 + "','" + fr1 + "' ,'" + fs1 + "','" + ft1 + "', '" + fu1 + "','" + fv1 + "', '" + fw1 + "','" + fx1 + "', '" + fy1 + "','" + fz1 + "'";
                                            //insertaSql += "','" + ga1 + "', '" + gb1 + "','" + gc1 + "', '" + gd1 + "','" + ge1 + "' ,'" + gf1 + "','" + gg1 + "', '" + gh1 + "','" + gi1 + "', '" + gj1 + "','" + gk1 + "', '" + gl1 + "','" + gm1 + "'";
                                            //insertaSql += "','" + gn1 + "', '" + go1 + "','" + gp1 + "', '" + gq1 + "','" + gr1 + "' ,'" + gs1 + "','" + gt1 + "', '" + gu1 + "','" + gv1 + "', '" + gw1 + "','" + gx1 + "', '" + gy1 + "','" + gz1 + "'";
                                            //insertaSql += "','" + ha1 + "', '" + hb1 + "','" + hc1 + "', '" + hd1 + "','" + he1 + "' ,'" + hf1 + "','" + hg1 + "', '" + hh1 + "','" + hi1 + "', '" + hj1 + "','" + hk1 + "', '" + hl1 + "','" + hm1 + "'";
                                            //insertaSql += "','" + hn1 + "', '" + ho1 + "','" + hp1 + "', '" + hq1 + "','" + hr1 + "' ,'" + hs1 + "','" + ht1 + "', '" + hu1 + "','" + hv1 + "', '" + hw1 + "','" + hx1 + "', '" + hy1 + "','" + hz1 + "'";
                                            //insertaSql += "','" + ia1 + "', '" + ib1 + "','" + ic1 + "', '" + id1 + "','" + ie1 + "' ,'" + if1 + "','" + ig1 + "', '" + ih1 + "','" + ii1 + "', '" + ij1 + "','" + ik1 + "', '" + il1 + "','" + im1 + "'";
                                            //insertaSql += "','" + in1 + "', '" + io1 + "','" + ip1 + "', '" + iq1 + "','" + ir1 + "' ,'" + is1 + "','" + it1 + "', '" + iu1 + "','" + iv1 + "', '" + iw1 + "','" + ix1 + "', '" + iy1 + "','" + iz1 + "'";

                                            //insertaSql += "  ) ";
                                            if (inserta.ejecutorBase(insertaSql))
                                            {
                                                this.Invoke(new DisplayEstado(Progreso), "Tabla virtual de Ingreso (" + (j + 1) + " de " + subreg + ") de   " + b1 + " - " + a1 + " fue correcto");
                                            }

                                            #endregion
                                        }

                                        #endregion
                                        #region sms
                                        string sqlTras = "SELECT a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1";
                                        sqlTras += ",aa1,ab1 ,ac1 ,ad1,ae1,af1,ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                        sqlTras += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                        sqlTras += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                        sqlTras += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                        sqlTras += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                        sqlTras += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                        sqlTras += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                        sqlTras += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                        sqlTras += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";
                                        

                                        sqlTras += " FROM Email_virtual ";
                                        sqlTras += " where Id_Grupo=" + id_grupo;
                                        sqlTras += " and a1 not in (SELECT a1 FROM Email where id_grupo=" + id_grupo + " group by a1) ";
                                        sqlTras += " group by  ";
                                        sqlTras += " a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1";
                                        sqlTras += ",aa1,ab1 ,ac1 ,ad1,ae1,af1,ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                        sqlTras += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                        sqlTras += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                        sqlTras += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                        sqlTras += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                        sqlTras += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                        sqlTras += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                        sqlTras += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                        sqlTras += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";

                                        maillist = ConexionCall.SqlDTable(sqlTras);
                                        subreg = maillist.Rows.Count;
                                        #endregion
                                    }
                                    //id_estado 1 es dinamico y 2 es estatico
                                    int validaEstatico = ConexionCall.devuelveValorINT("select count(*) from Grupo where id_estado=2 and id_grupo=" + id_grupo);

                                    #region Transpado estatico // no se utiliza 2019

                                    if (validaSMS == 0 && validaEstatico > 0)
                                    {
                                        #region traspaso Virtual Estatico
                                        for (int l = 0; l < subreg; l++)
                                        {
                                            #region
                                            a1 = maillist.Rows[l][0].ToString().Trim();
                                            b1 = maillist.Rows[l][1].ToString().Trim();
                                            c1 = maillist.Rows[l][2].ToString().Trim();
                                            d1 = maillist.Rows[l][3].ToString().Trim();
                                            e1 = maillist.Rows[l][4].ToString().Trim();
                                            f1 = maillist.Rows[l][5].ToString().Trim();
                                            g1 = maillist.Rows[l][6].ToString().Trim();
                                            h1 = maillist.Rows[l][7].ToString().Trim();
                                            i1 = maillist.Rows[l][8].ToString().Trim();
                                            j1 = maillist.Rows[l][9].ToString().Trim();
                                            k1 = maillist.Rows[l][10].ToString().Trim();
                                            l1 = maillist.Rows[l][11].ToString().Trim();
                                            m1 = maillist.Rows[l][12].ToString().Trim();
                                            n1 = maillist.Rows[l][13].ToString().Trim();
                                            o1 = maillist.Rows[l][14].ToString().Trim();
                                            p1 = maillist.Rows[l][15].ToString().Trim();
                                            q1 = maillist.Rows[l][16].ToString().Trim();
                                            r1 = maillist.Rows[l][17].ToString().Trim();
                                            s1 = maillist.Rows[l][18].ToString().Trim();
                                            t1 = maillist.Rows[l][19].ToString().Trim();
                                            u1 = maillist.Rows[l][20].ToString().Trim();
                                            v1 = maillist.Rows[l][21].ToString().Trim();
                                            w1 = maillist.Rows[l][22].ToString().Trim();
                                            x1 = maillist.Rows[l][23].ToString().Trim();
                                            y1 = maillist.Rows[l][24].ToString().Trim();
                                            z1 = maillist.Rows[l][25].ToString().Trim();
                                            aa1 = maillist.Rows[l][26].ToString().Trim();
                                            ab1 = maillist.Rows[l][27].ToString().Trim();
                                            ac1 = maillist.Rows[l][28].ToString().Trim();
                                            ad1 = maillist.Rows[l][29].ToString().Trim();
                                            ae1 = maillist.Rows[l][30].ToString().Trim();
                                            af1 = maillist.Rows[l][31].ToString().Trim();
                                            ag1 = maillist.Rows[l][32].ToString().Trim();
                                            ah1 = maillist.Rows[l][33].ToString().Trim();
                                            ai1 = maillist.Rows[l][34].ToString().Trim();
                                            aj1 = maillist.Rows[l][35].ToString().Trim();
                                            ak1 = maillist.Rows[l][36].ToString().Trim();
                                            al1 = maillist.Rows[l][37].ToString().Trim();
                                            am1 = maillist.Rows[l][38].ToString().Trim();

                                            an1 = maillist.Rows[l][39].ToString().Trim();
                                            ao1 = maillist.Rows[l][40].ToString().Trim();
                                            ap1 = maillist.Rows[l][41].ToString().Trim();
                                            aq1 = maillist.Rows[l][42].ToString().Trim();
                                            ar1 = maillist.Rows[l][43].ToString().Trim();
                                            as1 = maillist.Rows[l][44].ToString().Trim();
                                            at1 = maillist.Rows[l][45].ToString().Trim();
                                            au1 = maillist.Rows[l][46].ToString().Trim();
                                            av1 = maillist.Rows[l][47].ToString().Trim();
                                            aw1 = maillist.Rows[l][48].ToString().Trim();
                                            ax1 = maillist.Rows[l][49].ToString().Trim();
                                            ay1 = maillist.Rows[l][50].ToString().Trim();
                                            az1 = maillist.Rows[l][51].ToString().Trim();

                                            ba1 = maillist.Rows[l][52].ToString().Trim();
                                            bb1 = maillist.Rows[l][53].ToString().Trim();
                                            bc1 = maillist.Rows[l][54].ToString().Trim();
                                            bd1 = maillist.Rows[l][55].ToString().Trim();
                                            be1 = maillist.Rows[l][56].ToString().Trim();
                                            bf1 = maillist.Rows[l][57].ToString().Trim();
                                            bg1 = maillist.Rows[l][58].ToString().Trim();
                                            bh1 = maillist.Rows[l][59].ToString().Trim();
                                            bi1 = maillist.Rows[l][60].ToString().Trim();
                                            bj1 = maillist.Rows[l][61].ToString().Trim();
                                            bk1 = maillist.Rows[l][62].ToString().Trim();
                                            bl1 = maillist.Rows[l][63].ToString().Trim();
                                            bm1 = maillist.Rows[l][64].ToString().Trim();
                                            bn1 = maillist.Rows[l][65].ToString().Trim();
                                            bo1 = maillist.Rows[l][66].ToString().Trim();
                                            bp1 = maillist.Rows[l][67].ToString().Trim();
                                            bq1 = maillist.Rows[l][68].ToString().Trim();
                                            br1 = maillist.Rows[l][69].ToString().Trim();
                                            bs1 = maillist.Rows[l][70].ToString().Trim();
                                            bt1 = maillist.Rows[l][71].ToString().Trim();
                                            bu1 = maillist.Rows[l][72].ToString().Trim();
                                            bv1 = maillist.Rows[l][73].ToString().Trim();
                                            bw1 = maillist.Rows[l][74].ToString().Trim();
                                            bx1 = maillist.Rows[l][75].ToString().Trim();
                                            by1 = maillist.Rows[l][76].ToString().Trim();
                                            bz1 = maillist.Rows[l][77].ToString().Trim();

                                            ca1 = maillist.Rows[l][78].ToString().Trim();
                                            cb1 = maillist.Rows[l][79].ToString().Trim();
                                            cc1 = maillist.Rows[l][80].ToString().Trim();
                                            cd1 = maillist.Rows[l][81].ToString().Trim();
                                            ce1 = maillist.Rows[l][82].ToString().Trim();
                                            cf1 = maillist.Rows[l][83].ToString().Trim();
                                            cg1 = maillist.Rows[l][84].ToString().Trim();
                                            ch1 = maillist.Rows[l][85].ToString().Trim();
                                            ci1 = maillist.Rows[l][86].ToString().Trim();
                                            cj1 = maillist.Rows[l][87].ToString().Trim();
                                            ck1 = maillist.Rows[l][88].ToString().Trim();
                                            cl1 = maillist.Rows[l][89].ToString().Trim();
                                            cm1 = maillist.Rows[l][90].ToString().Trim();
                                            cn1 = maillist.Rows[l][91].ToString().Trim();
                                            co1 = maillist.Rows[l][92].ToString().Trim();
                                            cp1 = maillist.Rows[l][93].ToString().Trim();
                                            cq1 = maillist.Rows[l][94].ToString().Trim();
                                            cr1 = maillist.Rows[l][95].ToString().Trim();
                                            cs1 = maillist.Rows[l][96].ToString().Trim();
                                            ct1 = maillist.Rows[l][97].ToString().Trim();
                                            cu1 = maillist.Rows[l][98].ToString().Trim();
                                            cv1 = maillist.Rows[l][99].ToString().Trim();
                                            cw1 = maillist.Rows[l][100].ToString().Trim();
                                            cx1 = maillist.Rows[l][101].ToString().Trim();
                                            cy1 = maillist.Rows[l][102].ToString().Trim();
                                            cz1 = maillist.Rows[l][103].ToString().Trim();

                                            da1 = maillist.Rows[l][104].ToString().Trim();
                                            db1 = maillist.Rows[l][105].ToString().Trim();
                                            dc1 = maillist.Rows[l][106].ToString().Trim();
                                            dd1 = maillist.Rows[l][107].ToString().Trim();
                                            de1 = maillist.Rows[l][108].ToString().Trim();
                                            df1 = maillist.Rows[l][109].ToString().Trim();
                                            dg1 = maillist.Rows[l][110].ToString().Trim();
                                            dh1 = maillist.Rows[l][111].ToString().Trim();
                                            di1 = maillist.Rows[l][112].ToString().Trim();
                                            dj1 = maillist.Rows[l][113].ToString().Trim();
                                            dk1 = maillist.Rows[l][114].ToString().Trim();
                                            dl1 = maillist.Rows[l][115].ToString().Trim();
                                            dm1 = maillist.Rows[l][116].ToString().Trim();
                                            dn1 = maillist.Rows[l][117].ToString().Trim();
                                            do1 = maillist.Rows[l][118].ToString().Trim();
                                            dp1 = maillist.Rows[l][119].ToString().Trim();
                                            dq1 = maillist.Rows[l][120].ToString().Trim();
                                            dr1 = maillist.Rows[l][121].ToString().Trim();
                                            ds1 = maillist.Rows[l][122].ToString().Trim();
                                            dt1 = maillist.Rows[l][123].ToString().Trim();
                                            du1 = maillist.Rows[l][124].ToString().Trim();
                                            dv1 = maillist.Rows[l][125].ToString().Trim();
                                            dw1 = maillist.Rows[l][126].ToString().Trim();
                                            dx1 = maillist.Rows[l][127].ToString().Trim();
                                            dy1 = maillist.Rows[l][128].ToString().Trim();
                                            dz1 = maillist.Rows[l][129].ToString().Trim();

                                            ea1 = maillist.Rows[l][130].ToString().Trim();
                                            eb1 = maillist.Rows[l][131].ToString().Trim();
                                            ec1 = maillist.Rows[l][132].ToString().Trim();
                                            ed1 = maillist.Rows[l][133].ToString().Trim();
                                            ee1 = maillist.Rows[l][134].ToString().Trim();
                                            ef1 = maillist.Rows[l][135].ToString().Trim();
                                            eg1 = maillist.Rows[l][136].ToString().Trim();
                                            eh1 = maillist.Rows[l][137].ToString().Trim();
                                            ei1 = maillist.Rows[l][138].ToString().Trim();
                                            ej1 = maillist.Rows[l][139].ToString().Trim();
                                            ek1 = maillist.Rows[l][140].ToString().Trim();
                                            el1 = maillist.Rows[l][141].ToString().Trim();
                                            em1 = maillist.Rows[l][142].ToString().Trim();
                                            en1 = maillist.Rows[l][143].ToString().Trim();
                                            eo1 = maillist.Rows[l][144].ToString().Trim();
                                            ep1 = maillist.Rows[l][145].ToString().Trim();
                                            eq1 = maillist.Rows[l][146].ToString().Trim();
                                            er1 = maillist.Rows[l][147].ToString().Trim();
                                            es1 = maillist.Rows[l][148].ToString().Trim();
                                            et1 = maillist.Rows[l][149].ToString().Trim();
                                            eu1 = maillist.Rows[l][150].ToString().Trim();
                                            ev1 = maillist.Rows[l][151].ToString().Trim();
                                            ew1 = maillist.Rows[l][152].ToString().Trim();
                                            ex1 = maillist.Rows[l][153].ToString().Trim();
                                            ey1 = maillist.Rows[l][154].ToString().Trim();
                                            ez1 = maillist.Rows[l][155].ToString().Trim();

                                            fa1 = maillist.Rows[l][156].ToString().Trim();
                                            fb1 = maillist.Rows[l][157].ToString().Trim();
                                            fc1 = maillist.Rows[l][158].ToString().Trim();
                                            fd1 = maillist.Rows[l][159].ToString().Trim();
                                            fe1 = maillist.Rows[l][160].ToString().Trim();
                                            ff1 = maillist.Rows[l][161].ToString().Trim();
                                            fg1 = maillist.Rows[l][162].ToString().Trim();
                                            fh1 = maillist.Rows[l][163].ToString().Trim();
                                            fi1 = maillist.Rows[l][164].ToString().Trim();
                                            fj1 = maillist.Rows[l][165].ToString().Trim();
                                            fk1 = maillist.Rows[l][166].ToString().Trim();
                                            fl1 = maillist.Rows[l][167].ToString().Trim();
                                            fm1 = maillist.Rows[l][168].ToString().Trim();
                                            fn1 = maillist.Rows[l][169].ToString().Trim();
                                            fo1 = maillist.Rows[l][170].ToString().Trim();
                                            fp1 = maillist.Rows[l][171].ToString().Trim();
                                            fq1 = maillist.Rows[l][172].ToString().Trim();
                                            fr1 = maillist.Rows[l][173].ToString().Trim();
                                            fs1 = maillist.Rows[l][174].ToString().Trim();
                                            ft1 = maillist.Rows[l][175].ToString().Trim();
                                            fu1 = maillist.Rows[l][176].ToString().Trim();
                                            fv1 = maillist.Rows[l][177].ToString().Trim();
                                            fw1 = maillist.Rows[l][178].ToString().Trim();
                                            fx1 = maillist.Rows[l][179].ToString().Trim();
                                            fy1 = maillist.Rows[l][180].ToString().Trim();
                                            fz1 = maillist.Rows[l][181].ToString().Trim();

                                            ga1 = maillist.Rows[l][182].ToString().Trim();
                                            gb1 = maillist.Rows[l][183].ToString().Trim();
                                            gc1 = maillist.Rows[l][184].ToString().Trim();
                                            gd1 = maillist.Rows[l][185].ToString().Trim();
                                            ge1 = maillist.Rows[l][186].ToString().Trim();
                                            gf1 = maillist.Rows[l][187].ToString().Trim();
                                            gg1 = maillist.Rows[l][188].ToString().Trim();
                                            gh1 = maillist.Rows[l][189].ToString().Trim();
                                            gi1 = maillist.Rows[l][190].ToString().Trim();
                                            gj1 = maillist.Rows[l][191].ToString().Trim();
                                            gk1 = maillist.Rows[l][192].ToString().Trim();
                                            gl1 = maillist.Rows[l][193].ToString().Trim();
                                            gm1 = maillist.Rows[l][194].ToString().Trim();
                                            gn1 = maillist.Rows[l][195].ToString().Trim();
                                            go1 = maillist.Rows[l][196].ToString().Trim();
                                            gp1 = maillist.Rows[l][197].ToString().Trim();
                                            gq1 = maillist.Rows[l][198].ToString().Trim();
                                            gr1 = maillist.Rows[l][199].ToString().Trim();
                                            gs1 = maillist.Rows[l][200].ToString().Trim();
                                            gt1 = maillist.Rows[l][201].ToString().Trim();
                                            gu1 = maillist.Rows[l][202].ToString().Trim();
                                            gv1 = maillist.Rows[l][203].ToString().Trim();
                                            gw1 = maillist.Rows[l][204].ToString().Trim();
                                            gx1 = maillist.Rows[l][205].ToString().Trim();
                                            gy1 = maillist.Rows[l][206].ToString().Trim();
                                            gz1 = maillist.Rows[l][207].ToString().Trim();

                                            ha1 = maillist.Rows[l][208].ToString().Trim();
                                            hb1 = maillist.Rows[l][209].ToString().Trim();
                                            hc1 = maillist.Rows[l][210].ToString().Trim();
                                            hd1 = maillist.Rows[l][211].ToString().Trim();
                                            he1 = maillist.Rows[l][212].ToString().Trim();
                                            hf1 = maillist.Rows[l][213].ToString().Trim();
                                            hg1 = maillist.Rows[l][214].ToString().Trim();
                                            hh1 = maillist.Rows[l][215].ToString().Trim();
                                            hi1 = maillist.Rows[l][216].ToString().Trim();
                                            hj1 = maillist.Rows[l][217].ToString().Trim();
                                            hk1 = maillist.Rows[l][218].ToString().Trim();
                                            hl1 = maillist.Rows[l][219].ToString().Trim();
                                            hm1 = maillist.Rows[l][220].ToString().Trim();
                                            hn1 = maillist.Rows[l][221].ToString().Trim();
                                            ho1 = maillist.Rows[l][222].ToString().Trim();
                                            hp1 = maillist.Rows[l][223].ToString().Trim();
                                            hq1 = maillist.Rows[l][224].ToString().Trim();
                                            hr1 = maillist.Rows[l][225].ToString().Trim();
                                            hs1 = maillist.Rows[l][226].ToString().Trim();
                                            ht1 = maillist.Rows[l][227].ToString().Trim();
                                            hu1 = maillist.Rows[l][228].ToString().Trim();
                                            hv1 = maillist.Rows[l][229].ToString().Trim();
                                            hw1 = maillist.Rows[l][230].ToString().Trim();
                                            hx1 = maillist.Rows[l][231].ToString().Trim();
                                            hy1 = maillist.Rows[l][232].ToString().Trim();
                                            hz1 = maillist.Rows[l][233].ToString().Trim();

                                            ia1 = maillist.Rows[l][234].ToString().Trim();
                                            ib1 = maillist.Rows[l][235].ToString().Trim();
                                            ic1 = maillist.Rows[l][236].ToString().Trim();
                                            id1 = maillist.Rows[l][237].ToString().Trim();
                                            ie1 = maillist.Rows[l][238].ToString().Trim();
                                            if1 = maillist.Rows[l][239].ToString().Trim();
                                            ig1 = maillist.Rows[l][240].ToString().Trim();
                                            ih1 = maillist.Rows[l][241].ToString().Trim();
                                            ii1 = maillist.Rows[l][242].ToString().Trim();
                                            ij1 = maillist.Rows[l][243].ToString().Trim();
                                            ik1 = maillist.Rows[l][244].ToString().Trim();
                                            il1 = maillist.Rows[l][245].ToString().Trim();
                                            im1 = maillist.Rows[l][246].ToString().Trim();
                                            in1 = maillist.Rows[l][247].ToString().Trim();
                                            io1 = maillist.Rows[l][248].ToString().Trim();
                                            ip1 = maillist.Rows[l][249].ToString().Trim();
                                            iq1 = maillist.Rows[l][250].ToString().Trim();
                                            ir1 = maillist.Rows[l][251].ToString().Trim();
                                            is1 = maillist.Rows[l][252].ToString().Trim();
                                            it1 = maillist.Rows[l][253].ToString().Trim();
                                            iu1 = maillist.Rows[l][254].ToString().Trim();
                                            iv1 = maillist.Rows[l][255].ToString().Trim();
                                            iw1 = maillist.Rows[l][256].ToString().Trim();
                                            ix1 = maillist.Rows[l][257].ToString().Trim();
                                            iy1 = maillist.Rows[l][258].ToString().Trim();
                                            iz1 = maillist.Rows[l][259].ToString().Trim();



                                            #endregion
                                            this.Invoke(new DisplayEstado(numProcesado), "Documento " + archivos + " insertando registro en tabla Estático virtual " + (l + 1) + " de " + subreg);
                                            #region


                                            string insertaSql = "INSERT INTO Email_virtual (Id_Grupo, a1,b1,activo,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1,aa1,ab1,ac1,ad1,ae1,af1,";
                                            insertaSql += "  ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                            insertaSql += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                            insertaSql += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                            insertaSql += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                            insertaSql += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                            insertaSql += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                            insertaSql += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                            insertaSql += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                            insertaSql += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";

                                            insertaSql += ") VALUES(" + id_grupo + ", '" + a1.Trim() + "', '" + b1 + "','1'";
                                            insertaSql += " ,'" + c1 + "','" + d1 + "', '" + e1 + "','" + f1 + "', '" + g1 + "','" + h1 + "', '" + i1 + "','" + j1 + "', '" + k1 + "','" + l1 + "'";
                                            insertaSql += " ,'" + m1 + "','" + n1 + "','" + o1 + "', '" + p1 + "','" + q1 + "', '" + r1 + "','" + s1 + "', '" + t1 + "','" + u1 + "'";
                                            insertaSql += " ,'" + v1 + "','" + w1 + "', '" + x1 + "','" + y1 + "', '" + z1 + "','" + aa1 + "', '" + ab1 + "','" + ac1 + "', '" + ad1 + "','" + ae1 + "'";
                                            insertaSql += " ,'" + af1 + "','" + ag1 + "', '" + ah1 + "','" + ai1 + "', '" + aj1 + "','" + ak1 + "', '" + al1 + "','" + am1 + "'";
                                            insertaSql += " ,'" + an1 + "','" + ao1 + "', '" + ap1 + "','" + aq1 + "', '" + ar1 + "','" + as1 + "', '" + at1 + "','" + au1 + "','" + av1 + "', '" + aw1 + "','" + ax1 + "', '" + ay1 + "','" + az1 + "'";
                                            insertaSql += ",'" + ba1 + "', '" + bb1 + "','" + bc1 + "', '" + bd1 + "','" + be1 + "' ,'" + bf1 + "','" + bg1 + "', '" + bh1 + "','" + bi1 + "', '" + bj1 + "','" + bk1 + "', '" + bl1 + "','" + bm1 + "'";
                                            insertaSql += ",'" + bn1 + "', '" + bo1 + "','" + bp1 + "', '" + bq1 + "','" + br1 + "' ,'" + bs1 + "','" + bt1 + "', '" + bu1 + "','" + bv1 + "', '" + bw1 + "','" + bx1 + "', '" + by1 + "','" + bz1 + "'";
                                            insertaSql += ",'" + ca1 + "', '" + cb1 + "','" + cc1 + "', '" + cd1 + "','" + ce1 + "' ,'" + cf1 + "','" + cg1 + "', '" + ch1 + "','" + ci1 + "', '" + cj1 + "','" + ck1 + "', '" + cl1 + "','" + cm1 + "'";
                                            insertaSql += ",'" + cn1 + "', '" + co1 + "','" + cp1 + "', '" + cq1 + "','" + cr1 + "' ,'" + cs1 + "','" + ct1 + "', '" + cu1 + "','" + cv1 + "', '" + cw1 + "','" + cx1 + "', '" + cy1 + "','" + cz1 + "'";
                                            insertaSql += ",'" + da1 + "', '" + db1 + "','" + dc1 + "', '" + dd1 + "','" + de1 + "' ,'" + df1 + "','" + dg1 + "', '" + dh1 + "','" + di1 + "', '" + dj1 + "','" + dk1 + "', '" + dl1 + "','" + dm1 + "'";
                                            insertaSql += ",'" + dn1 + "', '" + do1 + "','" + dp1 + "', '" + dq1 + "','" + dr1 + "' ,'" + ds1 + "','" + dt1 + "', '" + du1 + "','" + dv1 + "', '" + dw1 + "','" + dx1 + "', '" + dy1 + "','" + dz1 + "'";
                                            insertaSql += ",'" + ea1 + "', '" + eb1 + "','" + ec1 + "', '" + ed1 + "','" + ee1 + "' ,'" + ef1 + "','" + eg1 + "', '" + eh1 + "','" + ei1 + "', '" + ej1 + "','" + ek1 + "', '" + el1 + "','" + em1 + "'";
                                            insertaSql += ",'" + en1 + "', '" + eo1 + "','" + ep1 + "', '" + eq1 + "','" + er1 + "' ,'" + es1 + "','" + et1 + "', '" + eu1 + "','" + ev1 + "', '" + ew1 + "','" + ex1 + "', '" + ey1 + "','" + ez1 + "'";
                                            insertaSql += ",'" + fa1 + "', '" + fb1 + "','" + fc1 + "', '" + fd1 + "','" + fe1 + "' ,'" + ff1 + "','" + fg1 + "', '" + fh1 + "','" + fi1 + "', '" + fj1 + "','" + fk1 + "', '" + fl1 + "','" + fm1 + "'";
                                            insertaSql += ",'" + fn1 + "', '" + fo1 + "','" + fp1 + "', '" + fq1 + "','" + fr1 + "' ,'" + fs1 + "','" + ft1 + "', '" + fu1 + "','" + fv1 + "', '" + fw1 + "','" + fx1 + "', '" + fy1 + "','" + fz1 + "'";
                                            insertaSql += ",'" + ga1 + "', '" + gb1 + "','" + gc1 + "', '" + gd1 + "','" + ge1 + "' ,'" + gf1 + "','" + gg1 + "', '" + gh1 + "','" + gi1 + "', '" + gj1 + "','" + gk1 + "', '" + gl1 + "','" + gm1 + "'";
                                            insertaSql += ",'" + gn1 + "', '" + go1 + "','" + gp1 + "', '" + gq1 + "','" + gr1 + "' ,'" + gs1 + "','" + gt1 + "', '" + gu1 + "','" + gv1 + "', '" + gw1 + "','" + gx1 + "', '" + gy1 + "','" + gz1 + "'";
                                            insertaSql += ",'" + ha1 + "', '" + hb1 + "','" + hc1 + "', '" + hd1 + "','" + he1 + "' ,'" + hf1 + "','" + hg1 + "', '" + hh1 + "','" + hi1 + "', '" + hj1 + "','" + hk1 + "', '" + hl1 + "','" + hm1 + "'";
                                            insertaSql += ",'" + hn1 + "', '" + ho1 + "','" + hp1 + "', '" + hq1 + "','" + hr1 + "' ,'" + hs1 + "','" + ht1 + "', '" + hu1 + "','" + hv1 + "', '" + hw1 + "','" + hx1 + "', '" + hy1 + "','" + hz1 + "'";
                                            insertaSql += ",'" + ia1 + "', '" + ib1 + "','" + ic1 + "', '" + id1 + "','" + ie1 + "' ,'" + if1 + "','" + ig1 + "', '" + ih1 + "','" + ii1 + "', '" + ij1 + "','" + ik1 + "', '" + il1 + "','" + im1 + "'";
                                            insertaSql += ",'" + in1 + "', '" + io1 + "','" + ip1 + "', '" + iq1 + "','" + ir1 + "' ,'" + is1 + "','" + it1 + "', '" + iu1 + "','" + iv1 + "', '" + iw1 + "','" + ix1 + "', '" + iy1 + "','" + iz1 + "'";




                                            insertaSql += "  ) ";
                                            if (inserta.ejecutorBase(insertaSql))
                                            {
                                                this.Invoke(new DisplayEstado(Progreso), "Tabla virtual de Ingreso (" + (l + 1) + " de " + subreg + ") de   " + b1 + " - " + a1 + " fue correcto");
                                            }

                                            #endregion
                                        }
                                        string sqlTras = "SELECT a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1";
                                        sqlTras += ",aa1,ab1 ,ac1 ,ad1,ae1,af1,ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                        sqlTras += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                        sqlTras += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                        sqlTras += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                        sqlTras += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                        sqlTras += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                        sqlTras += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                        sqlTras += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                        sqlTras += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";


                                        sqlTras += " FROM Email_virtual ";
                                        sqlTras += " where Id_Grupo=" + id_grupo;
                                        sqlTras += " and a1 not in (SELECT a1 FROM Email where id_grupo=" + id_grupo + " group by a1) ";
                                        sqlTras += " group by  ";
                                        sqlTras += " a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1";
                                        sqlTras += ",aa1,ab1 ,ac1 ,ad1,ae1,af1,ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                                        sqlTras += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                                        sqlTras += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                                        sqlTras += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                                        sqlTras += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                                        sqlTras += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                                        sqlTras += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                                        sqlTras += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                                        sqlTras += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";

                                        maillist = ConexionCall.SqlDTable(sqlTras);
                                        subreg = maillist.Rows.Count;
                                        #endregion
                                    }

                                        #endregion

                                        #region traspaso real


                                        int numeroserver = 10;

                                        int mod = subreg % numeroserver;
                                        int gru1 = subreg / numeroserver;
                                        int gru2 = gru1;
                                        int gru3 = gru1;

                                        tabnew = maillist.Clone();
                                        tabnew2 = maillist.Clone();
                                        tabnew3 = maillist.Clone();

                                        int cont1 = 0;
                                        while (cont1 < gru1)
                                        {
                                            try
                                            {
                                                tabnew.ImportRow(maillist.Rows[cont1]);
                                                cont1++;
                                            }
                                            catch (Exception e)
                                            { }
                                        }
                                        int g = 0;
                                        while (g < gru2)
                                        {
                                            try
                                            {
                                                tabnew2.ImportRow(maillist.Rows[cont1]);
                                                cont1++;
                                                g++;
                                            }
                                            catch (Exception e)
                                            { }
                                        }

                                        int p = 0;
                                        while (p < gru3)
                                        {
                                            try
                                            {
                                                tabnew3.ImportRow(maillist.Rows[cont1]);
                                                cont1++;
                                                p++;
                                            }
                                            catch (Exception e)
                                            { }
                                        }

                                        Task[] tasks2 = new Task[numeroserver];
                                        tasks2[0] = Task.Factory.StartNew(() => CargabaseHilo(tabnew, subreg, id_grupo, id_cliente, id_temp, archivos, validaEstatico, validaSMS));
                                        tasks2[1] = Task.Factory.StartNew(() => CargabaseHilo(tabnew2, subreg, id_grupo, id_cliente, id_temp, archivos, validaEstatico, validaSMS));
                                        tasks2[2] = Task.Factory.StartNew(() => CargabaseHilo(tabnew3, subreg, id_grupo, id_cliente, id_temp, archivos, validaEstatico, validaSMS));
                                      

                                        //Aquí se termina la ejecución Simultanea de los hilos de envío de correo
                                        Task.WaitAll(tasks2);

                                      //  CargabaseHilo(tabnew, subreg,id_grupo,id_cliente,id_temp,archivos,validaEstatico,validaSMS);
                                 
                                        #endregion
                                        // inserta.ejecutorBase("update grupo set porc=100 where  id_temp=" + id_temp);
                                        int estado = 2;
                                        int porcent = 0;
                                        //if (!pausa)
                                        //{ estado = 8; }
                                        // inserta.ejecutorBase("update Temporales set estado=2,fecha=getdate(), porc=100 where archivos='" + archivos + "' and carpetas='" + carpetas + "' and Id_Grupo=" + id_grupo + " and id_cliente=" + id_cliente);
                                        inserta.ejecutorBase("update Temporales set estado="+estado+",fecha=getdate(), porc="+porcent+" where  id_temp=" + id_temp);

                                    inserta.ejecutorBase("update Grupo set n_registros=" + ingresado + ", fecha_modificacion=getdate() where id_grupo=" + id_grupo);
                                   
                                    this.Invoke(new DisplayEstado(Progreso), "Documento " + archivos + " fue cambiado a estado procesado");
                                    this.Invoke(new DisplayEstado(Progreso), "Total registros " + subreg + " repetidos " + existe + " ingresados " + ingresado + " no ingresados por errores " + noconecta);
                                    this.Invoke(new DisplayEstado(Progreso), "      ");
                                    inserta.ejecutorBase("delete FROM Email_virtual where id_grupo=" + id_grupo);

                                    string email = buscaCorreo(id_cliente);
                                    string nombreFrom = System.Configuration.ConfigurationSettings.AppSettings["Nombre_From"].ToString().Trim();
                                    string from = System.Configuration.ConfigurationSettings.AppSettings["From_Carga"].ToString().Trim();
                                    string mensaje = "<p>Proceso de carga del documento " + archivos + " ha finalizado correctamente</p></br></br>";
                                    mensaje += "<br /><br /><p style='font-family: Arial; font-size:16px' ><b>Sistema hugoo.com<br />";
                                    mensaje += "<a href='mailto:soporte@hugoo.com'>soporte@hugoo.com</a><br />";
                                    mensaje += "<a href='http://clientes.hugoo.com/' >clientes.hugoo.com</a></b></p>";



                                    if (!string.IsNullOrEmpty(email))
                                    {
                                        enviar_Correo(email, from, nombreFrom, "Aviso Proceso de Carga", mensaje);
                                    }




                                }
                                }
                                else
                                {

                                    this.Invoke(new DisplayEstado(Progreso), "El archivo fue borrado fuera del sistema");
                                }

                            }
                            else
                            {
                                this.Invoke(new DisplayEstado(Progreso), "Dirección está en Nulo");
                            }


                            #endregion
                        }


                    }
                    this.Invoke(new DisplayEstado(Progreso), "Total de Documentos " + regs);
                }
                else
                {
                    this.Invoke(new DisplayEstado(numProcesado), "No hay registros que ingresar a las " + DateTime.Now);
                    try
                    {
                        this.Invoke(new DisplayEstado(Progreso), "No hay registros que ingresar a las " + DateTime.Now);
                    }
                    catch (Exception)
                    { }
                }


                try
                {

                    this.Invoke(new DisplayLimpia(limpiaMsg));
                }
                catch (Exception) { }
                #endregion

            }
            catch(Exception ex) { }

            

        }


        public void CargabaseHilo(DataTable maillist, int subreg, string id_grupo, string id_cliente, string id_temp, string archivos, int validaEstatico, int validaSMS)
        {
            #region
            int CadaxFilas = 0;
            int segundos = 0;
            try
            {
                CadaxFilas = Convert.ToInt32(ConfigurationSettings.AppSettings["CadaxFilas"]);
                segundos = Convert.ToInt32(ConfigurationSettings.AppSettings["segundos"]);
            }
            catch (Exception)
            {
                CadaxFilas = 2500;
                segundos = 10;
            }
            #endregion
            ConexionCall inserta = new ConexionCall();

            #region  variables
            int ingresado = 0;
            int noconecta = 0;
            int existe = 0;

            string a1, b1, c1, d1, e1, f1, g1, h1, i1, j1, k1, l1, m1, n1, o1, p1, q1, r1, s1, t1, u1, v1, w1, x1, y1, z1, aa1, ab1, ac1, ad1, ae1, af1;
            string ag1, ah1, ai1, aj1, ak1, al1, am1;

            string an1, ao1, ap1, aq1, ar1, as1, at1, au1, av1, aw1, ax1, ay1, az1;

            string ba1, bb1, bc1, bd1, be1, bf1, bg1, bh1, bi1, bj1, bk1, bl1, bm1, bn1, bo1, bp1, bq1, br1, bs1, bt1, bu1, bv1, bw1, bx1, by1, bz1;

            string ca1, cb1, cc1, cd1, ce1, cf1, cg1, ch1, ci1, cj1, ck1, cl1, cm1, cn1, co1, cp1, cq1, cr1, cs1, ct1, cu1, cv1, cw1, cx1, cy1, cz1;

            string da1, db1, dc1, dd1, de1, df1, dg1, dh1, di1, dj1, dk1, dl1, dm1, dn1, do1, dp1, dq1, dr1, ds1, dt1, du1, dv1, dw1, dx1, dy1, dz1;

            string ea1, eb1, ec1, ed1, ee1, ef1, eg1, eh1, ei1, ej1, ek1, el1, em1, en1, eo1, ep1, eq1, er1, es1, et1, eu1, ev1, ew1, ex1, ey1, ez1;

            string fa1, fb1, fc1, fd1, fe1, ff1, fg1, fh1, fi1, fj1, fk1, fl1, fm1, fn1, fo1, fp1, fq1, fr1, fs1, ft1, fu1, fv1, fw1, fx1, fy1, fz1;

            string ga1, gb1, gc1, gd1, ge1, gf1, gg1, gh1, gi1, gj1, gk1, gl1, gm1, gn1, go1, gp1, gq1, gr1, gs1, gt1, gu1, gv1, gw1, gx1, gy1, gz1;

            string ha1, hb1, hc1, hd1, he1, hf1, hg1, hh1, hi1, hj1, hk1, hl1, hm1, hn1, ho1, hp1, hq1, hr1, hs1, ht1, hu1, hv1, hw1, hx1, hy1, hz1;

            string ia1, ib1, ic1, id1, ie1, if1, ig1, ih1, ii1, ij1, ik1, il1, im1, in1, io1, ip1, iq1, ir1, is1, it1, iu1, iv1, iw1, ix1, iy1, iz1;

            #endregion
            bool valFormato = false;
            int mailAnterior = ConexionCall.devuelveValorINT("SELECT top 1 id_email FROM Email where Id_Grupo=" + id_grupo);


            int validaDesinscritos = ConexionCall.devuelveValorINT("select id_grupo_origen from Verificar_Desinscritos where id_grupo=" + id_grupo + " and Id_cliente=" + id_cliente);
            //si es -1 desinscribir de todo el cliente, sie es un valor específico solo sacar los desinscritos de dicho grupo       

            bool pausa = true;
            int porcent = 100;
            for (int k = 0; k < subreg; k++)
            {
                if (pausa)
                {

                    if ((k % 1000) == 0 && k != 0)
                    {
                        string sqlvalida = "select detener from temporales where id_temp = " + id_temp;
                        int detener = ConexionCall.devuelveValorINT(sqlvalida);

                        if (detener > 0)
                        {
                            pausa = false;
                            porcent = Convert.ToInt32((100 * k) / subreg);
                        }
                    }

                    #region carga variables
                    a1 = maillist.Rows[k][0].ToString().Trim();
                    b1 = maillist.Rows[k][1].ToString().Replace("'", "’").Trim();
                    c1 = maillist.Rows[k][2].ToString().Replace("'", "’").Trim();
                    d1 = maillist.Rows[k][3].ToString().Replace("'", "’").Trim();
                    e1 = maillist.Rows[k][4].ToString().Replace("'", "’").Trim();
                    f1 = maillist.Rows[k][5].ToString().Replace("'", "’").Trim();
                    g1 = maillist.Rows[k][6].ToString().Replace("'", "’").Trim();
                    h1 = maillist.Rows[k][7].ToString().Replace("'", "’").Trim();
                    i1 = maillist.Rows[k][8].ToString().Replace("'", "’").Trim();
                    j1 = maillist.Rows[k][9].ToString().Replace("'", "’").Trim();
                    k1 = maillist.Rows[k][10].ToString().Replace("'", "’").Trim();
                    l1 = maillist.Rows[k][11].ToString().Replace("'", "’").Trim();
                    m1 = maillist.Rows[k][12].ToString().Replace("'", "’").Trim();
                    n1 = maillist.Rows[k][13].ToString().Replace("'", "’").Trim();
                    o1 = maillist.Rows[k][14].ToString().Replace("'", "’").Trim();
                    p1 = maillist.Rows[k][15].ToString().Replace("'", "’").Trim();
                    q1 = maillist.Rows[k][16].ToString().Replace("'", "’").Trim();
                    r1 = maillist.Rows[k][17].ToString().Replace("'", "’").Trim();
                    s1 = maillist.Rows[k][18].ToString().Replace("'", "’").Trim();
                    t1 = maillist.Rows[k][19].ToString().Replace("'", "’").Trim();
                    u1 = maillist.Rows[k][20].ToString().Replace("'", "’").Trim();
                    v1 = maillist.Rows[k][21].ToString().Replace("'", "’").Trim();
                    w1 = maillist.Rows[k][22].ToString().Replace("'", "’").Trim();
                    x1 = maillist.Rows[k][23].ToString().Replace("'", "’").Trim();
                    y1 = maillist.Rows[k][24].ToString().Replace("'", "’").Trim();
                    z1 = maillist.Rows[k][25].ToString().Replace("'", "’").Trim();
                    aa1 = maillist.Rows[k][26].ToString().Replace("'", "’").Trim();
                    ab1 = maillist.Rows[k][27].ToString().Replace("'", "’").Trim();
                    ac1 = maillist.Rows[k][28].ToString().Replace("'", "’").Trim();
                    ad1 = maillist.Rows[k][29].ToString().Replace("'", "’").Trim();
                    ae1 = maillist.Rows[k][30].ToString().Replace("'", "’").Trim();
                    af1 = maillist.Rows[k][31].ToString().Replace("'", "’").Trim();
                    ag1 = maillist.Rows[k][32].ToString().Replace("'", "’").Trim();
                    ah1 = maillist.Rows[k][33].ToString().Replace("'", "’").Trim();
                    ai1 = maillist.Rows[k][34].ToString().Replace("'", "’").Trim();
                    aj1 = maillist.Rows[k][35].ToString().Replace("'", "’").Trim();
                    ak1 = maillist.Rows[k][36].ToString().Replace("'", "’").Trim();
                    al1 = maillist.Rows[k][37].ToString().Replace("'", "’").Trim();
                    am1 = maillist.Rows[k][38].ToString().Replace("'", "’").Trim();


                    an1 = maillist.Rows[k][39].ToString().Replace("'", "’").Trim();
                    ao1 = maillist.Rows[k][40].ToString().Replace("'", "’").Trim();
                    ap1 = maillist.Rows[k][41].ToString().Replace("'", "’").Trim();
                    aq1 = maillist.Rows[k][42].ToString().Replace("'", "’").Trim();
                    ar1 = maillist.Rows[k][43].ToString().Replace("'", "’").Trim();
                    as1 = maillist.Rows[k][44].ToString().Replace("'", "’").Trim();
                    at1 = maillist.Rows[k][45].ToString().Replace("'", "’").Trim();
                    au1 = maillist.Rows[k][46].ToString().Replace("'", "’").Trim();
                    av1 = maillist.Rows[k][47].ToString().Replace("'", "’").Trim();
                    aw1 = maillist.Rows[k][48].ToString().Replace("'", "’").Trim();
                    ax1 = maillist.Rows[k][49].ToString().Replace("'", "’").Trim();
                    ay1 = maillist.Rows[k][50].ToString().Replace("'", "’").Trim();
                    az1 = maillist.Rows[k][51].ToString().Replace("'", "’").Trim();

                    ba1 = maillist.Rows[k][52].ToString().Replace("'", "’").Trim();
                    bb1 = maillist.Rows[k][53].ToString().Replace("'", "’").Trim();
                    bc1 = maillist.Rows[k][54].ToString().Replace("'", "’").Trim();
                    bd1 = maillist.Rows[k][55].ToString().Replace("'", "’").Trim();
                    be1 = maillist.Rows[k][56].ToString().Replace("'", "’").Trim();
                    bf1 = maillist.Rows[k][57].ToString().Replace("'", "’").Trim();
                    bg1 = maillist.Rows[k][58].ToString().Replace("'", "’").Trim();
                    bh1 = maillist.Rows[k][59].ToString().Replace("'", "’").Trim();
                    bi1 = maillist.Rows[k][60].ToString().Replace("'", "’").Trim();
                    bj1 = maillist.Rows[k][61].ToString().Replace("'", "’").Trim();
                    bk1 = maillist.Rows[k][62].ToString().Replace("'", "’").Trim();
                    bl1 = maillist.Rows[k][63].ToString().Replace("'", "’").Trim();
                    bm1 = maillist.Rows[k][64].ToString().Replace("'", "’").Trim();
                    bn1 = maillist.Rows[k][65].ToString().Replace("'", "’").Trim();
                    bo1 = maillist.Rows[k][66].ToString().Replace("'", "’").Trim();
                    bp1 = maillist.Rows[k][67].ToString().Replace("'", "’").Trim();
                    bq1 = maillist.Rows[k][68].ToString().Replace("'", "’").Trim();
                    br1 = maillist.Rows[k][69].ToString().Replace("'", "’").Trim();
                    bs1 = maillist.Rows[k][70].ToString().Replace("'", "’").Trim();
                    bt1 = maillist.Rows[k][71].ToString().Replace("'", "’").Trim();
                    bu1 = maillist.Rows[k][72].ToString().Replace("'", "’").Trim();
                    bv1 = maillist.Rows[k][73].ToString().Replace("'", "’").Trim();
                    bw1 = maillist.Rows[k][74].ToString().Replace("'", "’").Trim();
                    bx1 = maillist.Rows[k][75].ToString().Replace("'", "’").Trim();
                    by1 = maillist.Rows[k][76].ToString().Replace("'", "’").Trim();
                    bz1 = maillist.Rows[k][77].ToString().Replace("'", "’").Trim();

                    ca1 = maillist.Rows[k][78].ToString().Replace("'", "’").Trim();
                    cb1 = maillist.Rows[k][79].ToString().Replace("'", "’").Trim();
                    cc1 = maillist.Rows[k][80].ToString().Replace("'", "’").Trim();
                    cd1 = maillist.Rows[k][81].ToString().Replace("'", "’").Trim();
                    ce1 = maillist.Rows[k][82].ToString().Replace("'", "’").Trim();
                    cf1 = maillist.Rows[k][83].ToString().Replace("'", "’").Trim();
                    cg1 = maillist.Rows[k][84].ToString().Replace("'", "’").Trim();
                    ch1 = maillist.Rows[k][85].ToString().Replace("'", "’").Trim();
                    ci1 = maillist.Rows[k][86].ToString().Replace("'", "’").Trim();
                    cj1 = maillist.Rows[k][87].ToString().Replace("'", "’").Trim();
                    ck1 = maillist.Rows[k][88].ToString().Replace("'", "’").Trim();
                    cl1 = maillist.Rows[k][89].ToString().Replace("'", "’").Trim();
                    cm1 = maillist.Rows[k][90].ToString().Replace("'", "’").Trim();
                    cn1 = maillist.Rows[k][91].ToString().Replace("'", "’").Trim();
                    co1 = maillist.Rows[k][92].ToString().Replace("'", "’").Trim();
                    cp1 = maillist.Rows[k][93].ToString().Replace("'", "’").Trim();
                    cq1 = maillist.Rows[k][94].ToString().Replace("'", "’").Trim();
                    cr1 = maillist.Rows[k][95].ToString().Replace("'", "’").Trim();
                    cs1 = maillist.Rows[k][96].ToString().Replace("'", "’").Trim();
                    ct1 = maillist.Rows[k][97].ToString().Replace("'", "’").Trim();
                    cu1 = maillist.Rows[k][98].ToString().Replace("'", "’").Trim();
                    cv1 = maillist.Rows[k][99].ToString().Replace("'", "’").Trim();
                    cw1 = maillist.Rows[k][100].ToString().Replace("'", "’").Trim();
                    cx1 = maillist.Rows[k][101].ToString().Replace("'", "’").Trim();
                    cy1 = maillist.Rows[k][102].ToString().Replace("'", "’").Trim();
                    cz1 = maillist.Rows[k][103].ToString().Replace("'", "’").Trim();

                    da1 = maillist.Rows[k][104].ToString().Replace("'", "’").Trim();
                    db1 = maillist.Rows[k][105].ToString().Replace("'", "’").Trim();
                    dc1 = maillist.Rows[k][106].ToString().Replace("'", "’").Trim();
                    dd1 = maillist.Rows[k][107].ToString().Replace("'", "’").Trim();
                    de1 = maillist.Rows[k][108].ToString().Replace("'", "’").Trim();
                    df1 = maillist.Rows[k][109].ToString().Replace("'", "’").Trim();
                    dg1 = maillist.Rows[k][110].ToString().Replace("'", "’").Trim();
                    dh1 = maillist.Rows[k][111].ToString().Replace("'", "’").Trim();
                    di1 = maillist.Rows[k][112].ToString().Replace("'", "’").Trim();
                    dj1 = maillist.Rows[k][113].ToString().Replace("'", "’").Trim();
                    dk1 = maillist.Rows[k][114].ToString().Replace("'", "’").Trim();
                    dl1 = maillist.Rows[k][115].ToString().Replace("'", "’").Trim();
                    dm1 = maillist.Rows[k][116].ToString().Replace("'", "’").Trim();
                    dn1 = maillist.Rows[k][117].ToString().Replace("'", "’").Trim();
                    do1 = maillist.Rows[k][118].ToString().Replace("'", "’").Trim();
                    dp1 = maillist.Rows[k][119].ToString().Replace("'", "’").Trim();
                    dq1 = maillist.Rows[k][120].ToString().Replace("'", "’").Trim();
                    dr1 = maillist.Rows[k][121].ToString().Replace("'", "’").Trim();
                    ds1 = maillist.Rows[k][122].ToString().Replace("'", "’").Trim();
                    dt1 = maillist.Rows[k][123].ToString().Replace("'", "’").Trim();
                    du1 = maillist.Rows[k][124].ToString().Replace("'", "’").Trim();
                    dv1 = maillist.Rows[k][125].ToString().Replace("'", "’").Trim();
                    dw1 = maillist.Rows[k][126].ToString().Replace("'", "’").Trim();
                    dx1 = maillist.Rows[k][127].ToString().Replace("'", "’").Trim();
                    dy1 = maillist.Rows[k][128].ToString().Replace("'", "’").Trim();
                    dz1 = maillist.Rows[k][129].ToString().Replace("'", "’").Trim();

                    ea1 = maillist.Rows[k][130].ToString().Replace("'", "’").Trim();
                    eb1 = maillist.Rows[k][131].ToString().Replace("'", "’").Trim();
                    ec1 = maillist.Rows[k][132].ToString().Replace("'", "’").Trim();
                    ed1 = maillist.Rows[k][133].ToString().Replace("'", "’").Trim();
                    ee1 = maillist.Rows[k][134].ToString().Replace("'", "’").Trim();
                    ef1 = maillist.Rows[k][135].ToString().Replace("'", "’").Trim();
                    eg1 = maillist.Rows[k][136].ToString().Replace("'", "’").Trim();
                    eh1 = maillist.Rows[k][137].ToString().Replace("'", "’").Trim();
                    ei1 = maillist.Rows[k][138].ToString().Replace("'", "’").Trim();
                    ej1 = maillist.Rows[k][139].ToString().Replace("'", "’").Trim();
                    ek1 = maillist.Rows[k][140].ToString().Replace("'", "’").Trim();
                    el1 = maillist.Rows[k][141].ToString().Replace("'", "’").Trim();
                    em1 = maillist.Rows[k][142].ToString().Replace("'", "’").Trim();
                    en1 = maillist.Rows[k][143].ToString().Replace("'", "’").Trim();
                    eo1 = maillist.Rows[k][144].ToString().Replace("'", "’").Trim();
                    ep1 = maillist.Rows[k][145].ToString().Replace("'", "’").Trim();
                    eq1 = maillist.Rows[k][146].ToString().Replace("'", "’").Trim();
                    er1 = maillist.Rows[k][147].ToString().Replace("'", "’").Trim();
                    es1 = maillist.Rows[k][148].ToString().Replace("'", "’").Trim();
                    et1 = maillist.Rows[k][149].ToString().Replace("'", "’").Trim();
                    eu1 = maillist.Rows[k][150].ToString().Replace("'", "’").Trim();
                    ev1 = maillist.Rows[k][151].ToString().Replace("'", "’").Trim();
                    ew1 = maillist.Rows[k][152].ToString().Replace("'", "’").Trim();
                    ex1 = maillist.Rows[k][153].ToString().Replace("'", "’").Trim();
                    ey1 = maillist.Rows[k][154].ToString().Replace("'", "’").Trim();
                    ez1 = maillist.Rows[k][155].ToString().Replace("'", "’").Trim();

                    fa1 = maillist.Rows[k][156].ToString().Replace("'", "’").Trim();
                    fb1 = maillist.Rows[k][157].ToString().Replace("'", "’").Trim();
                    fc1 = maillist.Rows[k][158].ToString().Replace("'", "’").Trim();
                    fd1 = maillist.Rows[k][159].ToString().Replace("'", "’").Trim();
                    fe1 = maillist.Rows[k][160].ToString().Replace("'", "’").Trim();
                    ff1 = maillist.Rows[k][161].ToString().Replace("'", "’").Trim();
                    fg1 = maillist.Rows[k][162].ToString().Replace("'", "’").Trim();
                    fh1 = maillist.Rows[k][163].ToString().Replace("'", "’").Trim();
                    fi1 = maillist.Rows[k][164].ToString().Replace("'", "’").Trim();
                    fj1 = maillist.Rows[k][165].ToString().Replace("'", "’").Trim();
                    fk1 = maillist.Rows[k][166].ToString().Replace("'", "’").Trim();
                    fl1 = maillist.Rows[k][167].ToString().Replace("'", "’").Trim();
                    fm1 = maillist.Rows[k][168].ToString().Replace("'", "’").Trim();
                    fn1 = maillist.Rows[k][169].ToString().Replace("'", "’").Trim();
                    fo1 = maillist.Rows[k][170].ToString().Replace("'", "’").Trim();
                    fp1 = maillist.Rows[k][171].ToString().Replace("'", "’").Trim();
                    fq1 = maillist.Rows[k][172].ToString().Replace("'", "’").Trim();
                    fr1 = maillist.Rows[k][173].ToString().Replace("'", "’").Trim();
                    fs1 = maillist.Rows[k][174].ToString().Replace("'", "’").Trim();
                    ft1 = maillist.Rows[k][175].ToString().Replace("'", "’").Trim();
                    fu1 = maillist.Rows[k][176].ToString().Replace("'", "’").Trim();
                    fv1 = maillist.Rows[k][177].ToString().Replace("'", "’").Trim();
                    fw1 = maillist.Rows[k][178].ToString().Replace("'", "’").Trim();
                    fx1 = maillist.Rows[k][179].ToString().Replace("'", "’").Trim();
                    fy1 = maillist.Rows[k][180].ToString().Replace("'", "’").Trim();
                    fz1 = maillist.Rows[k][181].ToString().Replace("'", "’").Trim();

                    ga1 = maillist.Rows[k][182].ToString().Replace("'", "’").Trim();
                    gb1 = maillist.Rows[k][183].ToString().Replace("'", "’").Trim();
                    gc1 = maillist.Rows[k][184].ToString().Replace("'", "’").Trim();
                    gd1 = maillist.Rows[k][185].ToString().Replace("'", "’").Trim();
                    ge1 = maillist.Rows[k][186].ToString().Replace("'", "’").Trim();
                    gf1 = maillist.Rows[k][187].ToString().Replace("'", "’").Trim();
                    gg1 = maillist.Rows[k][188].ToString().Replace("'", "’").Trim();
                    gh1 = maillist.Rows[k][189].ToString().Replace("'", "’").Trim();
                    gi1 = maillist.Rows[k][190].ToString().Replace("'", "’").Trim();
                    gj1 = maillist.Rows[k][191].ToString().Replace("'", "’").Trim();
                    gk1 = maillist.Rows[k][192].ToString().Replace("'", "’").Trim();
                    gl1 = maillist.Rows[k][193].ToString().Replace("'", "’").Trim();
                    gm1 = maillist.Rows[k][194].ToString().Replace("'", "’").Trim();
                    gn1 = maillist.Rows[k][195].ToString().Replace("'", "’").Trim();
                    go1 = maillist.Rows[k][196].ToString().Replace("'", "’").Trim();
                    gp1 = maillist.Rows[k][197].ToString().Replace("'", "’").Trim();
                    gq1 = maillist.Rows[k][198].ToString().Replace("'", "’").Trim();
                    gr1 = maillist.Rows[k][199].ToString().Replace("'", "’").Trim();
                    gs1 = maillist.Rows[k][200].ToString().Replace("'", "’").Trim();
                    gt1 = maillist.Rows[k][201].ToString().Replace("'", "’").Trim();
                    gu1 = maillist.Rows[k][202].ToString().Replace("'", "’").Trim();
                    gv1 = maillist.Rows[k][203].ToString().Replace("'", "’").Trim();
                    gw1 = maillist.Rows[k][204].ToString().Replace("'", "’").Trim();
                    gx1 = maillist.Rows[k][205].ToString().Replace("'", "’").Trim();
                    gy1 = maillist.Rows[k][206].ToString().Replace("'", "’").Trim();
                    gz1 = maillist.Rows[k][207].ToString().Replace("'", "’").Trim();

                    ha1 = maillist.Rows[k][208].ToString().Replace("'", "’").Trim();
                    hb1 = maillist.Rows[k][209].ToString().Replace("'", "’").Trim();
                    hc1 = maillist.Rows[k][210].ToString().Replace("'", "’").Trim();
                    hd1 = maillist.Rows[k][211].ToString().Replace("'", "’").Trim();
                    he1 = maillist.Rows[k][212].ToString().Replace("'", "’").Trim();
                    hf1 = maillist.Rows[k][213].ToString().Replace("'", "’").Trim();
                    hg1 = maillist.Rows[k][214].ToString().Replace("'", "’").Trim();
                    hh1 = maillist.Rows[k][215].ToString().Replace("'", "’").Trim();
                    hi1 = maillist.Rows[k][216].ToString().Replace("'", "’").Trim();
                    hj1 = maillist.Rows[k][217].ToString().Replace("'", "’").Trim();
                    hk1 = maillist.Rows[k][218].ToString().Replace("'", "’").Trim();
                    hl1 = maillist.Rows[k][219].ToString().Replace("'", "’").Trim();
                    hm1 = maillist.Rows[k][220].ToString().Replace("'", "’").Trim();
                    hn1 = maillist.Rows[k][221].ToString().Replace("'", "’").Trim();
                    ho1 = maillist.Rows[k][222].ToString().Replace("'", "’").Trim();
                    hp1 = maillist.Rows[k][223].ToString().Replace("'", "’").Trim();
                    hq1 = maillist.Rows[k][224].ToString().Replace("'", "’").Trim();
                    hr1 = maillist.Rows[k][225].ToString().Replace("'", "’").Trim();
                    hs1 = maillist.Rows[k][226].ToString().Replace("'", "’").Trim();
                    ht1 = maillist.Rows[k][227].ToString().Replace("'", "’").Trim();
                    hu1 = maillist.Rows[k][228].ToString().Replace("'", "’").Trim();
                    hv1 = maillist.Rows[k][229].ToString().Replace("'", "’").Trim();
                    hw1 = maillist.Rows[k][230].ToString().Replace("'", "’").Trim();
                    hx1 = maillist.Rows[k][231].ToString().Replace("'", "’").Trim();
                    hy1 = maillist.Rows[k][232].ToString().Replace("'", "’").Trim();
                    hz1 = maillist.Rows[k][233].ToString().Replace("'", "’").Trim();

                    ia1 = maillist.Rows[k][234].ToString().Replace("'", "’").Trim();
                    ib1 = maillist.Rows[k][235].ToString().Replace("'", "’").Trim();
                    ic1 = maillist.Rows[k][236].ToString().Replace("'", "’").Trim();
                    id1 = maillist.Rows[k][237].ToString().Replace("'", "’").Trim();
                    ie1 = maillist.Rows[k][238].ToString().Replace("'", "’").Trim();
                    if1 = maillist.Rows[k][239].ToString().Replace("'", "’").Trim();
                    ig1 = maillist.Rows[k][240].ToString().Replace("'", "’").Trim();
                    ih1 = maillist.Rows[k][241].ToString().Replace("'", "’").Trim();
                    ii1 = maillist.Rows[k][242].ToString().Replace("'", "’").Trim();
                    ij1 = maillist.Rows[k][243].ToString().Replace("'", "’").Trim();
                    ik1 = maillist.Rows[k][244].ToString().Replace("'", "’").Trim();
                    il1 = maillist.Rows[k][245].ToString().Replace("'", "’").Trim();
                    im1 = maillist.Rows[k][246].ToString().Replace("'", "’").Trim();
                    in1 = maillist.Rows[k][247].ToString().Replace("'", "’").Trim();
                    io1 = maillist.Rows[k][248].ToString().Replace("'", "’").Trim();
                    ip1 = maillist.Rows[k][249].ToString().Replace("'", "’").Trim();
                    iq1 = maillist.Rows[k][250].ToString().Replace("'", "’").Trim();
                    ir1 = maillist.Rows[k][251].ToString().Replace("'", "’").Trim();
                    is1 = maillist.Rows[k][252].ToString().Replace("'", "’").Trim();
                    it1 = maillist.Rows[k][253].ToString().Replace("'", "’").Trim();
                    iu1 = maillist.Rows[k][254].ToString().Replace("'", "’").Trim();
                    iv1 = maillist.Rows[k][255].ToString().Replace("'", "’").Trim();
                    iw1 = maillist.Rows[k][256].ToString().Replace("'", "’").Trim();
                    ix1 = maillist.Rows[k][257].ToString().Replace("'", "’").Trim();
                    iy1 = maillist.Rows[k][258].ToString().Replace("'", "’").Trim();
                    iz1 = maillist.Rows[k][259].ToString().Replace("'", "’").Trim();


                    #endregion
                    this.Invoke(new DisplayEstado(numProcesado), "Documento " + archivos + " insertando registro " + (k + 1) + " de " + subreg);
                    #region valida formato de mail o numero
                    if (!string.IsNullOrEmpty(a1))
                    {
                        valFormato = true;
                    }
                    else
                    {
                        valFormato = false;
                    }
                    #endregion
                    if (valFormato)
                    {
                        #region
                        #region valida existencia de a1
                        int validaCampos = 0;
                        if (validaEstatico > 0 && mailAnterior > 0)
                        {
                            validaCampos = ConexionCall.devuelveValorINT("SELECT count(*) FROM Email where a1='" + a1 + "' and Id_Grupo=" + id_grupo);
                        }

                        //ver si el grupo esta en tabla verificar_desinscritos


                        if (validaCampos > 0)
                        {
                            existe++;
                            this.Invoke(new DisplayEstado(Progreso), "Registro " + b1 + " - " + a1 + " ya existe");
                        }

                        #endregion

                        else
                        {

                            // validarEmail q el1 mail noconecta este en1 lista de1 erroneos
                            int error = ConexionCall.devuelveValorINT("select count(*)  from email_errores where email='" + a1 + "'");

                            int activo = 1;
                            if (error > 0) { activo = 0; }

                            int desinsc = 0;
                            if (validaDesinscritos != 0)
                            {


                                if (validaDesinscritos == -1)//si es -1 sacamos desinscritos de todos los grupos del cliente
                                {

                                    desinsc = ConexionCall.devuelveValorINT("select count(*) from Email_desincritos where id_cliente=" + id_cliente + " and mail ='" + a1 + "'");
                                }
                                else
                                {//si es un grupo específico sacamos solo desinscritos de ese grupo

                                    desinsc = ConexionCall.devuelveValorINT("select count(*)  from Email_desincritos where id_grupo=" + validaDesinscritos + " and mail ='" + a1 + "'");

                                }


                                validaCampos += desinsc;
                                if (desinsc > 0)
                                {
                                    this.Invoke(new DisplayEstado(Progreso), "Registro " + b1 + " - " + a1 + " está desinscrito de listas del cliente");
                                }

                            }



                            if (validaSMS > 0)
                            {
                                valFormato = validarNumSMS(a1);
                            }
                            else
                            {
                                valFormato = validarEmail(a1);
                            }

                            if (!valFormato) { activo = 0; }
                            if (desinsc > 0) { activo = 0; }


                            string insertaSql = "INSERT INTO Email(Id_Grupo, a1,b1,activo,c1,d1,e1,f1,g1,h1,i1,j1,k1,l1,m1,n1,o1,p1,q1,r1,s1,t1,u1,v1,w1,x1,y1,z1,aa1,ab1,ac1,ad1,ae1,af1,";
                            insertaSql += "  ag1,ah1,ai1,aj1,ak1,al1,am1,an1,ao1 ,ap1 ,aq1,ar1,as1,at1,au1,av1,aw1,ax1,ay1,az1";
                            insertaSql += ",ba1,bb1 ,bc1 ,bd1,be1,bf1,bg1,bh1,bi1,bj1,bk1,bl1,bm1,bn1,bo1 ,bp1 ,bq1,br1,bs1,bt1,bu1,bv1,bw1,bx1,by1,bz1";
                            insertaSql += ",ca1,cb1 ,cc1 ,cd1,ce1,cf1,cg1,ch1,ci1,cj1,ck1,cl1,cm1,cn1,co1 ,cp1 ,cq1,cr1,cs1,ct1,cu1,cv1,cw1,cx1,cy1,cz1";
                            insertaSql += ",da1,db1 ,dc1 ,dd1,de1,df1,dg1,dh1,di1,dj1,dk1,dl1,dm1,dn1,do1 ,dp1 ,dq1,dr1,ds1,dt1,du1,dv1,dw1,dx1,dy1,dz1";
                            insertaSql += ",ea1,eb1 ,ec1 ,ed1,ee1,ef1,eg1,eh1,ei1,ej1,ek1,el1,em1,en1,eo1 ,ep1 ,eq1,er1,es1,et1,eu1,ev1,ew1,ex1,ey1,ez1";
                            insertaSql += ",fa1,fb1 ,fc1 ,fd1,fe1,ff1,fg1,fh1,fi1,fj1,fk1,fl1,fm1,fn1,fo1 ,fp1 ,fq1,fr1,fs1,ft1,fu1,fv1,fw1,fx1,fy1,fz1";
                            insertaSql += ",ga1,gb1 ,gc1 ,gd1,ge1,gf1,gg1,gh1,gi1,gj1,gk1,gl1,gm1,gn1,go1 ,gp1 ,gq1,gr1,gs1,gt1,gu1,gv1,gw1,gx1,gy1,gz1";
                            insertaSql += ",ha1,hb1 ,hc1 ,hd1,he1,hf1,hg1,hh1,hi1,hj1,hk1,hl1,hm1,hn1,ho1 ,hp1 ,hq1,hr1,hs1,ht1,hu1,hv1,hw1,hx1,hy1,hz1";
                            insertaSql += ",ia1,ib1 ,ic1 ,id1,ie1,if1,ig1,ih1,ii1,ij1,ik1,il1,im1,in1,io1 ,ip1 ,iq1,ir1,is1,it1,iu1,iv1,iw1,ix1,iy1,iz1";
                            insertaSql += ") VALUES(" + id_grupo + ", '" + a1 + "', '" + b1 + "','" + activo + "'";
                            insertaSql += " ,'" + c1 + "','" + d1 + "', '" + e1 + "','" + f1 + "', '" + g1 + "','" + h1 + "', '" + i1 + "','" + j1 + "', '" + k1 + "','" + l1 + "'";
                            insertaSql += " ,'" + m1 + "','" + n1 + "','" + o1 + "', '" + p1 + "','" + q1 + "', '" + r1 + "','" + s1 + "', '" + t1 + "','" + u1 + "'";
                            insertaSql += " ,'" + v1 + "','" + w1 + "', '" + x1 + "','" + y1 + "', '" + z1 + "','" + aa1 + "', '" + ab1 + "','" + ac1 + "', '" + ad1 + "','" + ae1 + "'";
                            insertaSql += " ,'" + af1 + "','" + ag1 + "', '" + ah1 + "','" + ai1 + "', '" + aj1 + "','" + ak1 + "', '" + al1 + "','" + am1 + "'";
                            insertaSql += " ,'" + an1 + "','" + ao1 + "', '" + ap1 + "','" + aq1 + "', '" + ar1 + "','" + as1 + "', '" + at1 + "','" + au1 + "','" + av1 + "', '" + aw1 + "','" + ax1 + "', '" + ay1 + "','" + az1 + "'";
                            insertaSql += ",'" + ba1 + "', '" + bb1 + "','" + bc1 + "', '" + bd1 + "','" + be1 + "' ,'" + bf1 + "','" + bg1 + "', '" + bh1 + "','" + bi1 + "', '" + bj1 + "','" + bk1 + "', '" + bl1 + "','" + bm1 + "'";
                            insertaSql += ",'" + bn1 + "', '" + bo1 + "','" + bp1 + "', '" + bq1 + "','" + br1 + "' ,'" + bs1 + "','" + bt1 + "', '" + bu1 + "','" + bv1 + "', '" + bw1 + "','" + bx1 + "', '" + by1 + "','" + bz1 + "'";
                            insertaSql += ",'" + ca1 + "', '" + cb1 + "','" + cc1 + "', '" + cd1 + "','" + ce1 + "' ,'" + cf1 + "','" + cg1 + "', '" + ch1 + "','" + ci1 + "', '" + cj1 + "','" + ck1 + "', '" + cl1 + "','" + cm1 + "'";
                            insertaSql += ",'" + cn1 + "', '" + co1 + "','" + cp1 + "', '" + cq1 + "','" + cr1 + "' ,'" + cs1 + "','" + ct1 + "', '" + cu1 + "','" + cv1 + "', '" + cw1 + "','" + cx1 + "', '" + cy1 + "','" + cz1 + "'";
                            insertaSql += ",'" + da1 + "', '" + db1 + "','" + dc1 + "', '" + dd1 + "','" + de1 + "' ,'" + df1 + "','" + dg1 + "', '" + dh1 + "','" + di1 + "', '" + dj1 + "','" + dk1 + "', '" + dl1 + "','" + dm1 + "'";
                            insertaSql += ",'" + dn1 + "', '" + do1 + "','" + dp1 + "', '" + dq1 + "','" + dr1 + "' ,'" + ds1 + "','" + dt1 + "', '" + du1 + "','" + dv1 + "', '" + dw1 + "','" + dx1 + "', '" + dy1 + "','" + dz1 + "'";
                            insertaSql += ",'" + ea1 + "', '" + eb1 + "','" + ec1 + "', '" + ed1 + "','" + ee1 + "' ,'" + ef1 + "','" + eg1 + "', '" + eh1 + "','" + ei1 + "', '" + ej1 + "','" + ek1 + "', '" + el1 + "','" + em1 + "'";
                            insertaSql += ",'" + en1 + "', '" + eo1 + "','" + ep1 + "', '" + eq1 + "','" + er1 + "' ,'" + es1 + "','" + et1 + "', '" + eu1 + "','" + ev1 + "', '" + ew1 + "','" + ex1 + "', '" + ey1 + "','" + ez1 + "'";
                            insertaSql += ",'" + fa1 + "', '" + fb1 + "','" + fc1 + "', '" + fd1 + "','" + fe1 + "' ,'" + ff1 + "','" + fg1 + "', '" + fh1 + "','" + fi1 + "', '" + fj1 + "','" + fk1 + "', '" + fl1 + "','" + fm1 + "'";
                            insertaSql += ",'" + fn1 + "', '" + fo1 + "','" + fp1 + "', '" + fq1 + "','" + fr1 + "' ,'" + fs1 + "','" + ft1 + "', '" + fu1 + "','" + fv1 + "', '" + fw1 + "','" + fx1 + "', '" + fy1 + "','" + fz1 + "'";
                            insertaSql += ",'" + ga1 + "', '" + gb1 + "','" + gc1 + "', '" + gd1 + "','" + ge1 + "' ,'" + gf1 + "','" + gg1 + "', '" + gh1 + "','" + gi1 + "', '" + gj1 + "','" + gk1 + "', '" + gl1 + "','" + gm1 + "'";
                            insertaSql += ",'" + gn1 + "', '" + go1 + "','" + gp1 + "', '" + gq1 + "','" + gr1 + "' ,'" + gs1 + "','" + gt1 + "', '" + gu1 + "','" + gv1 + "', '" + gw1 + "','" + gx1 + "', '" + gy1 + "','" + gz1 + "'";
                            insertaSql += ",'" + ha1 + "', '" + hb1 + "','" + hc1 + "', '" + hd1 + "','" + he1 + "' ,'" + hf1 + "','" + hg1 + "', '" + hh1 + "','" + hi1 + "', '" + hj1 + "','" + hk1 + "', '" + hl1 + "','" + hm1 + "'";
                            insertaSql += ",'" + hn1 + "', '" + ho1 + "','" + hp1 + "', '" + hq1 + "','" + hr1 + "' ,'" + hs1 + "','" + ht1 + "', '" + hu1 + "','" + hv1 + "', '" + hw1 + "','" + hx1 + "', '" + hy1 + "','" + hz1 + "'";
                            insertaSql += ",'" + ia1 + "', '" + ib1 + "','" + ic1 + "', '" + id1 + "','" + ie1 + "' ,'" + if1 + "','" + ig1 + "', '" + ih1 + "','" + ii1 + "', '" + ij1 + "','" + ik1 + "', '" + il1 + "','" + im1 + "'";
                            insertaSql += ",'" + in1 + "', '" + io1 + "','" + ip1 + "', '" + iq1 + "','" + ir1 + "' ,'" + is1 + "','" + it1 + "', '" + iu1 + "','" + iv1 + "', '" + iw1 + "','" + ix1 + "', '" + iy1 + "','" + iz1 + "'";

                            insertaSql += "  ) ";
                            if (inserta.ejecutorBase(insertaSql))
                            {
                                ingresado++;
                                this.Invoke(new DisplayEstado(Progreso), "Ingreso (" + (k + 1) + " de " + subreg + ") de   " + b1 + " - " + a1 + " fue correcto");
                            }
                            else
                            {
                                noconecta++;
                                this.Invoke(new DisplayEstado(Progreso), "Ingreso de  (" + (k + 1) + " de " + subreg + ")  " + b1 + " - " + a1 + " fallido");
                                string log = inserta.ejecutorBaseString(insertaSql);
                                log = log.Replace("'", "");
                                string sqlError = "INSERT INTO log_masivo (linea,id_temp,registro) VALUES ( ";
                                sqlError += (k + 1) + "," + id_temp + " ,'" + log + "')";
                                inserta.ejecutorBase(sqlError);
                            }


                            // }

                        }
                        #endregion

                    }
                    else
                    {
                        this.Invoke(new DisplayEstado(Progreso), "Formato no válido de " + a1 + " ");
                        noconecta++;
                    }
                    #region


                    if (k % 100 == 0 && k != 0)
                    {
                        int porcentaje = Convert.ToInt32((100 * k) / subreg);
                        //inserta.ejecutorBase("update grupo set porc=" + porcentaje + " where  id_grupo=" + id_grupo);
                        inserta.ejecutorBase("update temporales set porc=" + porcentaje + " where  id_temp=" + id_temp);



                    }



                    if (k % CadaxFilas == 0 && k != 0)
                    {
                        this.Invoke(new DisplayEstado(Progreso), "Detenido  " + segundos + " segundos");
                        Thread.Sleep(segundos * 1000);
                    }
                    if (k % 100 == 0 && k != 0)
                    {

                        try
                        {
                            this.Invoke(new DisplayLimpia(limpiaMsg));
                        }
                        catch (Exception)
                        { }
                    }
                    #endregion
                }
                else { k = subreg; }

            }

        }
        public void Procesa_Excel()
        {

            DataTable ParaProceso = ConexionCall.SqlDTable("select * from Excel where estado=1 ");
          //  DataTable ParaProceso = ConexionCall.SqlDTable("select * from Excel where id_mensaje=  35143");
            int reg = ParaProceso.Rows.Count;
            int correcto = 0;
            int segencuesta = 0;
            int evento = 0;
            int inexistentes = 0;
            
            int malo = 0;
            int segUrl = 0;
            int clickAqui = 0;
            if (reg > 0)
            {
                string id_mensaje = "", carpeta = "", id_excel = "", archivo = "", archivo2 = "", archivo3 = "", archivo4 = "", archivo6 = "", archivo7 = "", archivo8 = "", id_grupo = "", nombreGrupo = "";
                ConexionCall exe = new ConexionCall();
                for (int i = 0; i < reg; i++)
                {
                    #region
                    segUrl = 0;
                    clickAqui = 0;
                    segencuesta = 0;
                    id_mensaje = ParaProceso.Rows[i]["id_mensaje"].ToString();
                    carpeta = ParaProceso.Rows[i]["carpeta"].ToString();
                    id_excel = ParaProceso.Rows[i]["id_excel"].ToString();
                    archivo = ParaProceso.Rows[i]["archivo"].ToString();
                    id_grupo = "";
                    nombreGrupo = "";
                   
                    DataTable grupo = ConexionCall.SqlDTable("select nombre,id_grupo from Grupo where id_grupo=(SELECT Id_Grupo FROM Mensaje where id_mensaje=" + id_mensaje + ")");
                        if (grupo.Rows.Count>0)
                        {
                            id_grupo = grupo.Rows[i]["id_grupo"].ToString();
                            nombreGrupo = grupo.Rows[i]["nombre"].ToString();
                        }
                    #endregion
                    int valida = ConexionCall.devuelveValorINT("select count(*) from Excel where estado=1 and id_excel=" + id_excel);
                    if (valida > 0)
                    {
                        #region
                        exe.ejecutorBase("update Excel set estado=2 where id_excel=" + id_excel);
                        borra_anterior(carpeta + archivo);
                        string fecha = DateTime.Now.ToString().Trim();
                        this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
                        ///////carga datos
                      
                        #region carga tablas                      
                       
                        #region correos
                        DataTable emails = traspaso(id_mensaje,id_grupo);
                        int rr = emails.Rows.Count;
                        if (rr > 0)
                        {
                            this.Invoke(new DisplayEstado(Progreso), "Generando Informe general del mensaje "+id_mensaje);
                            archivo = "ID_" + id_mensaje + "_" + fecha + ".xls";
                            archivo = limpiaString(archivo);
                        }
                        #endregion
                        #region segimiento url
                        DataTable url = traspaso2(id_grupo, nombreGrupo, id_mensaje);
                        int ur = url.Rows.Count;

                        if (ur > 0)
                        {
                            this.Invoke(new DisplayEstado(Progreso), "Generando Informe URL del mensaje " + id_mensaje);
                            archivo2 = "ID_" + id_mensaje + "_" + fecha + "_Seguimiento.xls";
                            // archivo2 = "ID_" + id_mensaje + "_Seguimiento" + fecha + ".xls";
                            archivo2 = limpiaString(archivo2);
                            string ruta2 = carpeta + archivo2;
                            ExportarExcelDataTable(url, ruta2);
                            segUrl = 1;
                        }
                        #endregion                                              
                        #region click aqui
                        //DataTable click = traspaso3(id_mensaje);
                        //int cl = click.Rows.Count;

                        //if (cl > 0)
                        //{
                        //    this.Invoke(new DisplayEstado(Progreso), "Generando Informe Click del mensaje " + id_mensaje);
                        //    archivo3 = "ID_" + id_mensaje + "_" + fecha + "_Click.xls";
                        //   // archivo3 = "ID_" + id_mensaje + "_Click_" + fecha + ".xls";
                        //    archivo3 = limpiaString(archivo3);
                        //    string ruta3 = carpeta + archivo3;
                        //    ExportarExcelDataTable(click, ruta3);
                        //    clickAqui = 1;
                        //}
                        #endregion
                        #region desinscripcion
                        DataTable desincripcion = traspaso4(id_grupo,id_mensaje);
                        int des = desincripcion.Rows.Count;
                       // if (des > 0)  //JM Lo saque para que genere informe de desisncritos aunque sean 0
                       // {
                            this.Invoke(new DisplayEstado(Progreso), "Generando Informe Desinscritos del mensaje " + id_mensaje);
                            archivo4 = "ID_" + id_mensaje + "_" + fecha + "_Desinscritos.xls";
                            archivo4 = limpiaString(archivo4);
                            string ruta4 = carpeta + archivo4;
                            ExportarExcelDataTable(desincripcion, ruta4);
                      //  }
                        #endregion
                        #region
                        traspaso5( id_mensaje);
                        #endregion
                        #region
                        DataTable encuesta = traspaso6(id_mensaje);
                        int en = encuesta.Rows.Count;

                        if (en > 0)
                        {
                            this.Invoke(new DisplayEstado(Progreso), "Generando Informe Encuesta del mensaje " + id_mensaje);
                            archivo6 = "ID_" + id_mensaje + "_" + fecha + "_Encuesta.xls";
                           archivo6 = limpiaString(archivo6);
                            string ruta6 = carpeta + archivo6;
                            ExportarExcelDataTable(encuesta, ruta6);
                            segencuesta = 1;
                        }
                        #endregion
                        #region eventos

                        if (tieneevento(id_mensaje))
                        {
                            DataTable eventos = traspaso7(id_mensaje);
                            int ev = eventos.Rows.Count;

                            if (ev > 0)
                            {
                                this.Invoke(new DisplayEstado(Progreso), "Generando Informe Eventos del mensaje " + id_mensaje);
                                archivo7 = "ID_" + id_mensaje + "_" + fecha + "_Confirmacion.xls";
                                archivo7 = limpiaString(archivo7);
                                string ruta7 = carpeta + archivo7;
                                ExportarExcelDataTable(eventos, ruta7);
                                evento = 1;
                            }
                        }
                        else evento = 0;

                        #endregion
                        #region Inexistentes
                        //DataTable inexistent = traspaso8(id_mensaje);
                        //int inex = inexistent.Rows.Count;
                        // if (inex > 0)
                        //    {
                               
                        //this.Invoke(new DisplayEstado(Progreso), "Generando Informe Inexistentes del mensaje " + id_mensaje);
                        //archivo8 = "ID_" + id_mensaje + "_" + fecha + "_Inexistentes.xls";
                        //archivo8 = limpiaString(archivo8);
                        //string ruta8 = carpeta + archivo8;
                        //ExportarExcelDataTable(inexistent, ruta8);

                        //     inexistentes=1;
                                 
                        //}
                        //else inexistentes = 0;
                        #endregion 

                     
                        if (rr > 0)
                        {
                            exe.ejecutorBase("update Excel set url = " + segUrl + " ,clickaqui =" + clickAqui + ",encuesta=" + segencuesta + ",evento=" + evento + " ,inexistentes=" + inexistentes + "  where id_excel=" + id_excel);
                            string ruta1 = carpeta + archivo;

                            if (ExportarExcelDataTable(emails, ruta1))
                            {
                                exe.ejecutorBase("update Excel set archivo='" + archivo + "',estado=4,fecha=getdate() where id_excel=" + id_excel);
                                this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
                                correcto++;
                            }
                            else
                            {
                                malo++;
                                exe.ejecutorBase("update Excel set estado=1, archivo='' where id_excel=" + id_excel);
                            }

                        }
                        else {
                            exe.ejecutorBase("update Excel set estado=4,fecha=getdate() where id_excel=" + id_excel);                               
                        }                        
                        #endregion
                        #endregion
                    }
                    this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos excel");
                    this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
                }
            }
        }



        private void ejecutar_envioProgramado(object myObject, EventArgs myEventArgs)
        {
            hilo4.Stop();
            DataTable programado = ConexionCall.SqlDTable(" select * from Mensaje where id_estado in (7,10) order by ejecucion asc");
            int reg = programado.Rows.Count;
            string id_mensaje, ejecucion, asunto, id_estado;
            int anio, mes, dia, minuto, hora;
            string sqlHora = "select day(getdate())as dia, month(getdate())as mes,year(getdate())as año, ";
            sqlHora += "DATEPART(hour, getdate()) as hora,DATEPART(minute, getdate()) as minuto";
            if (reg != 0)
            {

                for (int i = 0; i < reg; i++)
                {

                    try
                    {

                    id_mensaje = programado.Rows[i]["id_mensaje"].ToString();
                    id_estado = programado.Rows[i]["id_estado"].ToString();
                    ejecucion = programado.Rows[i]["ejecucion"].ToString();
                    asunto = programado.Rows[i]["asunto"].ToString();

                    DateTime Hora_exe = Convert.ToDateTime(ejecucion);
                    DataTable horatabla = ConexionCall.SqlDTable(sqlHora);

                    anio = Convert.ToInt32(horatabla.Rows[0]["año"]);
                    mes = Convert.ToInt32(horatabla.Rows[0]["mes"]);
                    dia = Convert.ToInt32(horatabla.Rows[0]["dia"]);

                    hora = Convert.ToInt32(horatabla.Rows[0]["hora"]);
                    minuto = Convert.ToInt32(horatabla.Rows[0]["minuto"]);

                    if (Hora_exe.Year == anio && Hora_exe.Month == mes && Hora_exe.Day == dia)
                    {
                        ConexionCall exeHora = new ConexionCall();

                        try
                        {

                            if (Convert.ToInt32(id_estado) == 10)
                            {
                                if (Hora_exe.Hour == hora && Hora_exe.Minute <= minuto)
                                {
                                    exeHora.ejecutorBase("UPDATE Mensaje SET id_estado=9  WHERE id_mensaje=" + id_mensaje);
                                    this.Invoke(new DisplayEstado(Progreso), "Mensaje " + asunto + " se está procesando");

                                }
                                if (hora > Hora_exe.Hour)
                                {
                                    exeHora.ejecutorBase("UPDATE Mensaje SET id_estado=9  WHERE id_mensaje=" + id_mensaje);
                                    this.Invoke(new DisplayEstado(Progreso), "Mensaje " + asunto + " se está procesando");
                                }
                            }
                            else
                            {
                                if (Hora_exe.Hour == hora && Hora_exe.Minute <= minuto)
                                {
                                    exeHora.ejecutorBase("UPDATE Mensaje SET id_estado=1  WHERE id_mensaje=" + id_mensaje);
                                    this.Invoke(new DisplayEstado(Progreso), "Mensaje " + asunto + " se está procesando");

                                }
                                if (hora > Hora_exe.Hour)
                                {
                                    exeHora.ejecutorBase("UPDATE Mensaje SET id_estado=1  WHERE id_mensaje=" + id_mensaje);
                                    this.Invoke(new DisplayEstado(Progreso), "Mensaje " + asunto + " se está procesando");
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                           
                        }
                    }
                    }
                    catch (Exception)
                    {
                    }

                }

            }

            try
            {
                Thread.Sleep(60 * 1000);
                hilo4.Enabled = true;
                this.Invoke(new DisplayLimpia(limpiaMsg));
                //  ejecutar_envioProgramado();
            }
            catch (Exception)
            {
                hilo4.Enabled = true;
                //   ejecutar_envioProgramado();
            }
        }
        private void ejecutar_envioProgramadoSMS(object myObject, EventArgs myEventArgs)
        {
            hilo7.Stop();
            DataTable programado = ConexionCall.SqlDTable(" select * from Mensaje_SMS where id_estado in (7) order by ejecucion asc");
            int reg = programado.Rows.Count;
            string id_sms, ejecucion, asunto;
            int anio, mes, dia, minuto, hora;
            string sqlHora = "select day(getdate())as dia, month(getdate())as mes,year(getdate())as año, ";
            sqlHora += "DATEPART(hour, getdate()) as hora,DATEPART(minute, getdate()) as minuto";
            if (reg != 0)
            {

                for (int i = 0; i < reg; i++)
                {
                    id_sms = programado.Rows[i]["id_sms"].ToString();
                    ejecucion = programado.Rows[i]["ejecucion"].ToString();
                    asunto = programado.Rows[i]["asunto"].ToString();

                    DateTime Hora_exe = Convert.ToDateTime(ejecucion);
                    DataTable horatabla = ConexionCall.SqlDTable(sqlHora);

                    anio = Convert.ToInt32(horatabla.Rows[0]["año"]);
                    mes = Convert.ToInt32(horatabla.Rows[0]["mes"]);
                    dia = Convert.ToInt32(horatabla.Rows[0]["dia"]);

                    hora = Convert.ToInt32(horatabla.Rows[0]["hora"]);
                    minuto = Convert.ToInt32(horatabla.Rows[0]["minuto"]);

                    if (Hora_exe.Year == anio && Hora_exe.Month == mes && Hora_exe.Day == dia)
                    {
                        ConexionCall exeHora = new ConexionCall();

                        if (Hora_exe.Hour == hora && Hora_exe.Minute <= minuto)
                        {
                            exeHora.ejecutorBase("UPDATE Mensaje_SMS SET id_estado=1  WHERE id_sms=" + id_sms);
                            this.Invoke(new DisplayEstado(Progreso), "SMS " + asunto + " se está procesando");

                        }
                        if (hora > Hora_exe.Hour)
                        {
                            exeHora.ejecutorBase("UPDATE Mensaje_SMS SET id_estado=1  WHERE id_sms=" + id_sms);
                            this.Invoke(new DisplayEstado(Progreso), "SMS " + asunto + " se está procesando");
                        }

                    }

                }

            }

            try
            {
                Thread.Sleep(60 * 1000);
                hilo7.Enabled = true;
                this.Invoke(new DisplayLimpia(limpiaMsg));
                //  ejecutar_envioProgramado();
            }
            catch (Exception)
            {
                //   ejecutar_envioProgramado();
            }
        }
        private void limpiaMsg()
        {
            txb_msg.Text = "";
        }
        private void cargarRuta(string mss)
        {
            txb_direccion.Text = mss;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea cancelar el proceso de Mail List?", "Cancelación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    #region
                    hilo1.Stop();
                    hilo1.Dispose();

                    hilo2.Stop();
                    hilo2.Dispose();

                    hilo3.Stop();
                    hilo3.Dispose();

                    hilo4.Stop();
                    hilo4.Dispose();


                    hilo5.Stop();
                    hilo5.Dispose();

                    hilo6.Stop();
                    hilo6.Dispose();

                    hilo6.Stop();
                    hilo6.Dispose();

                    hilo7.Stop();
                    hilo7.Dispose();

                    hilo8.Stop();
                    hilo8.Dispose();


                    hilo9.Stop();
                    hilo9.Dispose();

                    hilo10.Stop();
                    hilo10.Dispose();

                    hilo11.Stop();
                    hilo11.Dispose();

                    hilo12.Stop();
                    hilo12.Dispose();

                    hilo13.Stop();
                    hilo13.Dispose();








                    hilo14.Stop();
                    hilo14.Dispose();

                    hilo15.Stop();
                    hilo15.Dispose();


                    hilo16.Stop();
                    hilo16.Dispose();

                    //hilo17.Stop();
                    //hilo17.Dispose();

                    //hilo18.Stop();
                    //hilo18.Dispose();

                    hilo19.Stop();
                    hilo19.Dispose();

                    hilo20.Stop();
                    hilo20.Dispose();






                    #endregion
                }
                catch (Exception) { }
                button1.Enabled = true;
                Application.Exit();
            }
        }
        private string[] separador(string valor)
        {
            int cont = 0;
            char[] Separa = new char[] { ',' };
            string[] valores = new string[40];
            foreach (string substr in valor.Split(Separa))
            {
                try
                {
                    System.Console.WriteLine(substr);
                    valores[cont] = substr;
                    cont++;
                }
                catch (Exception)
                { }

            }
            return valores;
        }
        private DataTable arreglaColumnas(DataTable dtAntigua)
        {
            int reg = dtAntigua.Rows.Count;
            int columnas = dtAntigua.Columns.Count;
            DataTable dt = new DataTable();

            if (reg > 0)
            {
                DataRow dr = null;
                dt.Columns.Clear();
                dt.Rows.Clear();
                dt.Columns.Add(new DataColumn("a1", typeof(string)));
                dt.Columns.Add(new DataColumn("b1", typeof(string)));
                dt.Columns.Add(new DataColumn("c1", typeof(string)));
                dt.Columns.Add(new DataColumn("d1", typeof(string)));
                dt.Columns.Add(new DataColumn("e1", typeof(string)));
                dt.Columns.Add(new DataColumn("f1", typeof(string)));
                dt.Columns.Add(new DataColumn("g1", typeof(string)));
                dt.Columns.Add(new DataColumn("h1", typeof(string)));
                dt.Columns.Add(new DataColumn("i1", typeof(string)));
                dt.Columns.Add(new DataColumn("j1", typeof(string)));
                dt.Columns.Add(new DataColumn("k1", typeof(string)));
                dt.Columns.Add(new DataColumn("l1", typeof(string)));
                dt.Columns.Add(new DataColumn("m1", typeof(string)));
                dt.Columns.Add(new DataColumn("n1", typeof(string)));
                dt.Columns.Add(new DataColumn("o1", typeof(string)));
                dt.Columns.Add(new DataColumn("p1", typeof(string)));
                dt.Columns.Add(new DataColumn("q1", typeof(string)));
                dt.Columns.Add(new DataColumn("r1", typeof(string)));
                dt.Columns.Add(new DataColumn("s1", typeof(string)));
                dt.Columns.Add(new DataColumn("t1", typeof(string)));
                dt.Columns.Add(new DataColumn("u1", typeof(string)));
                dt.Columns.Add(new DataColumn("v1", typeof(string)));
                dt.Columns.Add(new DataColumn("w1", typeof(string)));
                dt.Columns.Add(new DataColumn("x1", typeof(string)));
                dt.Columns.Add(new DataColumn("y1", typeof(string)));
                dt.Columns.Add(new DataColumn("z1", typeof(string)));

                dt.Columns.Add(new DataColumn("aa1", typeof(string)));
                dt.Columns.Add(new DataColumn("ab1", typeof(string)));
                dt.Columns.Add(new DataColumn("ac1", typeof(string)));
                dt.Columns.Add(new DataColumn("ad1", typeof(string)));
                dt.Columns.Add(new DataColumn("ae1", typeof(string)));
                dt.Columns.Add(new DataColumn("af1", typeof(string)));
                dt.Columns.Add(new DataColumn("ag1", typeof(string)));
                dt.Columns.Add(new DataColumn("ah1", typeof(string)));
                dt.Columns.Add(new DataColumn("ai1", typeof(string)));
                dt.Columns.Add(new DataColumn("aj1", typeof(string)));
                dt.Columns.Add(new DataColumn("ak1", typeof(string)));
                dt.Columns.Add(new DataColumn("al1", typeof(string)));
                dt.Columns.Add(new DataColumn("am1", typeof(string)));
                dt.Columns.Add(new DataColumn("an1", typeof(string)));
                dt.Columns.Add(new DataColumn("ao1", typeof(string)));
                dt.Columns.Add(new DataColumn("ap1", typeof(string)));
                dt.Columns.Add(new DataColumn("aq1", typeof(string)));
                dt.Columns.Add(new DataColumn("ar1", typeof(string)));
                dt.Columns.Add(new DataColumn("as1", typeof(string)));
                dt.Columns.Add(new DataColumn("at1", typeof(string)));
                dt.Columns.Add(new DataColumn("au1", typeof(string)));
                dt.Columns.Add(new DataColumn("av1", typeof(string)));
                dt.Columns.Add(new DataColumn("aw1", typeof(string)));
                dt.Columns.Add(new DataColumn("ax1", typeof(string)));
                dt.Columns.Add(new DataColumn("ay1", typeof(string)));
                dt.Columns.Add(new DataColumn("az1", typeof(string)));

                dt.Columns.Add(new DataColumn("ba1", typeof(string)));
                dt.Columns.Add(new DataColumn("bb1", typeof(string)));
                dt.Columns.Add(new DataColumn("bc1", typeof(string)));
                dt.Columns.Add(new DataColumn("bd1", typeof(string)));
                dt.Columns.Add(new DataColumn("be1", typeof(string)));
                dt.Columns.Add(new DataColumn("bf1", typeof(string)));
                dt.Columns.Add(new DataColumn("bg1", typeof(string)));
                dt.Columns.Add(new DataColumn("bh1", typeof(string)));
                dt.Columns.Add(new DataColumn("bi1", typeof(string)));
                dt.Columns.Add(new DataColumn("bj1", typeof(string)));
                dt.Columns.Add(new DataColumn("bk1", typeof(string)));
                dt.Columns.Add(new DataColumn("bl1", typeof(string)));
                dt.Columns.Add(new DataColumn("bm1", typeof(string)));
                dt.Columns.Add(new DataColumn("bn1", typeof(string)));
                dt.Columns.Add(new DataColumn("bo1", typeof(string)));
                dt.Columns.Add(new DataColumn("bp1", typeof(string)));
                dt.Columns.Add(new DataColumn("bq1", typeof(string)));
                dt.Columns.Add(new DataColumn("br1", typeof(string)));
                dt.Columns.Add(new DataColumn("bs1", typeof(string)));
                dt.Columns.Add(new DataColumn("bt1", typeof(string)));
                dt.Columns.Add(new DataColumn("bu1", typeof(string)));
                dt.Columns.Add(new DataColumn("bv1", typeof(string)));
                dt.Columns.Add(new DataColumn("bw1", typeof(string)));
                dt.Columns.Add(new DataColumn("bx1", typeof(string)));
                dt.Columns.Add(new DataColumn("by1", typeof(string)));
                dt.Columns.Add(new DataColumn("bz1", typeof(string)));

                dt.Columns.Add(new DataColumn("ca1", typeof(string)));
                dt.Columns.Add(new DataColumn("cb1", typeof(string)));
                dt.Columns.Add(new DataColumn("cc1", typeof(string)));
                dt.Columns.Add(new DataColumn("cd1", typeof(string)));
                dt.Columns.Add(new DataColumn("ce1", typeof(string)));
                dt.Columns.Add(new DataColumn("cf1", typeof(string)));
                dt.Columns.Add(new DataColumn("cg1", typeof(string)));
                dt.Columns.Add(new DataColumn("ch1", typeof(string)));
                dt.Columns.Add(new DataColumn("ci1", typeof(string)));
                dt.Columns.Add(new DataColumn("cj1", typeof(string)));
                dt.Columns.Add(new DataColumn("ck1", typeof(string)));
                dt.Columns.Add(new DataColumn("cl1", typeof(string)));
                dt.Columns.Add(new DataColumn("cm1", typeof(string)));
                dt.Columns.Add(new DataColumn("cn1", typeof(string)));
                dt.Columns.Add(new DataColumn("co1", typeof(string)));
                dt.Columns.Add(new DataColumn("cp1", typeof(string)));
                dt.Columns.Add(new DataColumn("cq1", typeof(string)));
                dt.Columns.Add(new DataColumn("cr1", typeof(string)));
                dt.Columns.Add(new DataColumn("cs1", typeof(string)));
                dt.Columns.Add(new DataColumn("ct1", typeof(string)));
                dt.Columns.Add(new DataColumn("cu1", typeof(string)));
                dt.Columns.Add(new DataColumn("cv1", typeof(string)));
                dt.Columns.Add(new DataColumn("cw1", typeof(string)));
                dt.Columns.Add(new DataColumn("cx1", typeof(string)));
                dt.Columns.Add(new DataColumn("cy1", typeof(string)));
                dt.Columns.Add(new DataColumn("cz1", typeof(string)));

                dt.Columns.Add(new DataColumn("da1", typeof(string)));
                dt.Columns.Add(new DataColumn("db1", typeof(string)));
                dt.Columns.Add(new DataColumn("dc1", typeof(string)));
                dt.Columns.Add(new DataColumn("dd1", typeof(string)));
                dt.Columns.Add(new DataColumn("de1", typeof(string)));
                dt.Columns.Add(new DataColumn("df1", typeof(string)));
                dt.Columns.Add(new DataColumn("dg1", typeof(string)));
                dt.Columns.Add(new DataColumn("dh1", typeof(string)));
                dt.Columns.Add(new DataColumn("di1", typeof(string)));
                dt.Columns.Add(new DataColumn("dj1", typeof(string)));
                dt.Columns.Add(new DataColumn("dk1", typeof(string)));
                dt.Columns.Add(new DataColumn("dl1", typeof(string)));
                dt.Columns.Add(new DataColumn("dm1", typeof(string)));
                dt.Columns.Add(new DataColumn("dn1", typeof(string)));
                dt.Columns.Add(new DataColumn("do1", typeof(string)));
                dt.Columns.Add(new DataColumn("dp1", typeof(string)));
                dt.Columns.Add(new DataColumn("dq1", typeof(string)));
                dt.Columns.Add(new DataColumn("dr1", typeof(string)));
                dt.Columns.Add(new DataColumn("ds1", typeof(string)));
                dt.Columns.Add(new DataColumn("dt1", typeof(string)));
                dt.Columns.Add(new DataColumn("du1", typeof(string)));
                dt.Columns.Add(new DataColumn("dv1", typeof(string)));
                dt.Columns.Add(new DataColumn("dw1", typeof(string)));
                dt.Columns.Add(new DataColumn("dx1", typeof(string)));
                dt.Columns.Add(new DataColumn("dy1", typeof(string)));
                dt.Columns.Add(new DataColumn("dz1", typeof(string)));

                dt.Columns.Add(new DataColumn("ea1", typeof(string)));
                dt.Columns.Add(new DataColumn("eb1", typeof(string)));
                dt.Columns.Add(new DataColumn("ec1", typeof(string)));
                dt.Columns.Add(new DataColumn("ed1", typeof(string)));
                dt.Columns.Add(new DataColumn("ee1", typeof(string)));
                dt.Columns.Add(new DataColumn("ef1", typeof(string)));
                dt.Columns.Add(new DataColumn("eg1", typeof(string)));
                dt.Columns.Add(new DataColumn("eh1", typeof(string)));
                dt.Columns.Add(new DataColumn("ei1", typeof(string)));
                dt.Columns.Add(new DataColumn("ej1", typeof(string)));
                dt.Columns.Add(new DataColumn("ek1", typeof(string)));
                dt.Columns.Add(new DataColumn("el1", typeof(string)));
                dt.Columns.Add(new DataColumn("em1", typeof(string)));
                dt.Columns.Add(new DataColumn("en1", typeof(string)));
                dt.Columns.Add(new DataColumn("eo1", typeof(string)));
                dt.Columns.Add(new DataColumn("ep1", typeof(string)));
                dt.Columns.Add(new DataColumn("eq1", typeof(string)));
                dt.Columns.Add(new DataColumn("er1", typeof(string)));
                dt.Columns.Add(new DataColumn("es1", typeof(string)));
                dt.Columns.Add(new DataColumn("et1", typeof(string)));
                dt.Columns.Add(new DataColumn("eu1", typeof(string)));
                dt.Columns.Add(new DataColumn("ev1", typeof(string)));
                dt.Columns.Add(new DataColumn("ew1", typeof(string)));
                dt.Columns.Add(new DataColumn("ex1", typeof(string)));
                dt.Columns.Add(new DataColumn("ey1", typeof(string)));
                dt.Columns.Add(new DataColumn("ez1", typeof(string)));

                dt.Columns.Add(new DataColumn("fa1", typeof(string)));
                dt.Columns.Add(new DataColumn("fb1", typeof(string)));
                dt.Columns.Add(new DataColumn("fc1", typeof(string)));
                dt.Columns.Add(new DataColumn("fd1", typeof(string)));
                dt.Columns.Add(new DataColumn("fe1", typeof(string)));
                dt.Columns.Add(new DataColumn("ff1", typeof(string)));
                dt.Columns.Add(new DataColumn("fg1", typeof(string)));
                dt.Columns.Add(new DataColumn("fh1", typeof(string)));
                dt.Columns.Add(new DataColumn("fi1", typeof(string)));
                dt.Columns.Add(new DataColumn("fj1", typeof(string)));
                dt.Columns.Add(new DataColumn("fk1", typeof(string)));
                dt.Columns.Add(new DataColumn("fl1", typeof(string)));
                dt.Columns.Add(new DataColumn("fm1", typeof(string)));
                dt.Columns.Add(new DataColumn("fn1", typeof(string)));
                dt.Columns.Add(new DataColumn("fo1", typeof(string)));
                dt.Columns.Add(new DataColumn("fp1", typeof(string)));
                dt.Columns.Add(new DataColumn("fq1", typeof(string)));
                dt.Columns.Add(new DataColumn("fr1", typeof(string)));
                dt.Columns.Add(new DataColumn("fs1", typeof(string)));
                dt.Columns.Add(new DataColumn("ft1", typeof(string)));
                dt.Columns.Add(new DataColumn("fu1", typeof(string)));
                dt.Columns.Add(new DataColumn("fv1", typeof(string)));
                dt.Columns.Add(new DataColumn("fw1", typeof(string)));
                dt.Columns.Add(new DataColumn("fx1", typeof(string)));
                dt.Columns.Add(new DataColumn("fy1", typeof(string)));
                dt.Columns.Add(new DataColumn("fz1", typeof(string)));

                dt.Columns.Add(new DataColumn("ga1", typeof(string)));
                dt.Columns.Add(new DataColumn("gb1", typeof(string)));
                dt.Columns.Add(new DataColumn("gc1", typeof(string)));
                dt.Columns.Add(new DataColumn("gd1", typeof(string)));
                dt.Columns.Add(new DataColumn("ge1", typeof(string)));
                dt.Columns.Add(new DataColumn("gf1", typeof(string)));
                dt.Columns.Add(new DataColumn("gg1", typeof(string)));
                dt.Columns.Add(new DataColumn("gh1", typeof(string)));
                dt.Columns.Add(new DataColumn("gi1", typeof(string)));
                dt.Columns.Add(new DataColumn("gj1", typeof(string)));
                dt.Columns.Add(new DataColumn("gk1", typeof(string)));
                dt.Columns.Add(new DataColumn("gl1", typeof(string)));
                dt.Columns.Add(new DataColumn("gm1", typeof(string)));
                dt.Columns.Add(new DataColumn("gn1", typeof(string)));
                dt.Columns.Add(new DataColumn("go1", typeof(string)));
                dt.Columns.Add(new DataColumn("gp1", typeof(string)));
                dt.Columns.Add(new DataColumn("gq1", typeof(string)));
                dt.Columns.Add(new DataColumn("gr1", typeof(string)));
                dt.Columns.Add(new DataColumn("gs1", typeof(string)));
                dt.Columns.Add(new DataColumn("gt1", typeof(string)));
                dt.Columns.Add(new DataColumn("gu1", typeof(string)));
                dt.Columns.Add(new DataColumn("gv1", typeof(string)));
                dt.Columns.Add(new DataColumn("gw1", typeof(string)));
                dt.Columns.Add(new DataColumn("gx1", typeof(string)));
                dt.Columns.Add(new DataColumn("gy1", typeof(string)));
                dt.Columns.Add(new DataColumn("gz1", typeof(string)));


                dt.Columns.Add(new DataColumn("ha1", typeof(string)));
                dt.Columns.Add(new DataColumn("hb1", typeof(string)));
                dt.Columns.Add(new DataColumn("hc1", typeof(string)));
                dt.Columns.Add(new DataColumn("hd1", typeof(string)));
                dt.Columns.Add(new DataColumn("he1", typeof(string)));
                dt.Columns.Add(new DataColumn("hf1", typeof(string)));
                dt.Columns.Add(new DataColumn("hg1", typeof(string)));
                dt.Columns.Add(new DataColumn("hh1", typeof(string)));
                dt.Columns.Add(new DataColumn("hi1", typeof(string)));
                dt.Columns.Add(new DataColumn("hj1", typeof(string)));
                dt.Columns.Add(new DataColumn("hk1", typeof(string)));
                dt.Columns.Add(new DataColumn("hl1", typeof(string)));
                dt.Columns.Add(new DataColumn("hm1", typeof(string)));
                dt.Columns.Add(new DataColumn("hn1", typeof(string)));
                dt.Columns.Add(new DataColumn("ho1", typeof(string)));
                dt.Columns.Add(new DataColumn("hp1", typeof(string)));
                dt.Columns.Add(new DataColumn("hq1", typeof(string)));
                dt.Columns.Add(new DataColumn("hr1", typeof(string)));
                dt.Columns.Add(new DataColumn("hs1", typeof(string)));
                dt.Columns.Add(new DataColumn("ht1", typeof(string)));
                dt.Columns.Add(new DataColumn("hu1", typeof(string)));
                dt.Columns.Add(new DataColumn("hv1", typeof(string)));
                dt.Columns.Add(new DataColumn("hw1", typeof(string)));
                dt.Columns.Add(new DataColumn("hx1", typeof(string)));
                dt.Columns.Add(new DataColumn("hy1", typeof(string)));
                dt.Columns.Add(new DataColumn("hz1", typeof(string)));


                dt.Columns.Add(new DataColumn("ia1", typeof(string)));
                dt.Columns.Add(new DataColumn("ib1", typeof(string)));
                dt.Columns.Add(new DataColumn("ic1", typeof(string)));
                dt.Columns.Add(new DataColumn("id1", typeof(string)));
                dt.Columns.Add(new DataColumn("ie1", typeof(string)));
                dt.Columns.Add(new DataColumn("if1", typeof(string)));
                dt.Columns.Add(new DataColumn("ig1", typeof(string)));
                dt.Columns.Add(new DataColumn("ih1", typeof(string)));
                dt.Columns.Add(new DataColumn("ii1", typeof(string)));
                dt.Columns.Add(new DataColumn("ij1", typeof(string)));
                dt.Columns.Add(new DataColumn("ik1", typeof(string)));
                dt.Columns.Add(new DataColumn("il1", typeof(string)));
                dt.Columns.Add(new DataColumn("im1", typeof(string)));
                dt.Columns.Add(new DataColumn("in1", typeof(string)));
                dt.Columns.Add(new DataColumn("io1", typeof(string)));
                dt.Columns.Add(new DataColumn("ip1", typeof(string)));
                dt.Columns.Add(new DataColumn("iq1", typeof(string)));
                dt.Columns.Add(new DataColumn("ir1", typeof(string)));
                dt.Columns.Add(new DataColumn("is1", typeof(string)));
                dt.Columns.Add(new DataColumn("it1", typeof(string)));
                dt.Columns.Add(new DataColumn("iu1", typeof(string)));
                dt.Columns.Add(new DataColumn("iv1", typeof(string)));
                dt.Columns.Add(new DataColumn("iw1", typeof(string)));
                dt.Columns.Add(new DataColumn("ix1", typeof(string)));
                dt.Columns.Add(new DataColumn("iy1", typeof(string)));
                dt.Columns.Add(new DataColumn("iz1", typeof(string)));


                for (int i = 0; i < reg; i++)
                {
                    // string qqq = dtAntigua.Rows[i][0].ToString();
                    //  string[] vecT = separador(dtAntigua.Rows[i][0].ToString());
                    dr = dt.NewRow();
                    try
                    {
                        
                        dr["a1"] = dtAntigua.Rows[i][0].ToString();
                        dr["b1"] = dtAntigua.Rows[i][1].ToString();
                        dr["c1"] = dtAntigua.Rows[i][2].ToString();
                        dr["d1"] = dtAntigua.Rows[i][3].ToString();
                        dr["e1"] = dtAntigua.Rows[i][4].ToString();
                        dr["f1"] = dtAntigua.Rows[i][5].ToString();
                        dr["g1"] = dtAntigua.Rows[i][6].ToString();
                        dr["h1"] = dtAntigua.Rows[i][7].ToString();
                        dr["i1"] = dtAntigua.Rows[i][8].ToString();
                        dr["j1"] = dtAntigua.Rows[i][9].ToString();
                        dr["k1"] = dtAntigua.Rows[i][10].ToString();
                        dr["l1"] = dtAntigua.Rows[i][11].ToString();
                        dr["m1"] = dtAntigua.Rows[i][12].ToString();
                        dr["n1"] = dtAntigua.Rows[i][13].ToString();
                        dr["o1"] = dtAntigua.Rows[i][14].ToString();
                        dr["p1"] = dtAntigua.Rows[i][15].ToString();
                        dr["q1"] = dtAntigua.Rows[i][16].ToString();
                        dr["r1"] = dtAntigua.Rows[i][17].ToString();
                        dr["s1"] = dtAntigua.Rows[i][18].ToString();
                        dr["t1"] = dtAntigua.Rows[i][19].ToString();
                        dr["u1"] = dtAntigua.Rows[i][20].ToString();
                        dr["v1"] = dtAntigua.Rows[i][21].ToString();
                        dr["w1"] = dtAntigua.Rows[i][22].ToString();
                        dr["x1"] = dtAntigua.Rows[i][23].ToString();
                        dr["y1"] = dtAntigua.Rows[i][24].ToString();
                        dr["z1"] = dtAntigua.Rows[i][25].ToString();
                        dr["aa1"] = dtAntigua.Rows[i][26].ToString();
                        dr["ab1"] = dtAntigua.Rows[i][27].ToString();
                        dr["ac1"] = dtAntigua.Rows[i][28].ToString();
                        dr["ad1"] = dtAntigua.Rows[i][29].ToString();
                        dr["ae1"] = dtAntigua.Rows[i][30].ToString();
                        dr["af1"] = dtAntigua.Rows[i][31].ToString();
                        dr["ag1"] = dtAntigua.Rows[i][32].ToString();
                        dr["ah1"] = dtAntigua.Rows[i][33].ToString();
                        dr["ai1"] = dtAntigua.Rows[i][34].ToString();
                        dr["aj1"] = dtAntigua.Rows[i][35].ToString();
                        dr["ak1"] = dtAntigua.Rows[i][36].ToString();
                        dr["al1"] = dtAntigua.Rows[i][37].ToString();
                        dr["am1"] = dtAntigua.Rows[i][38].ToString();
                        dr["an1"] = dtAntigua.Rows[i][39].ToString();
                        dr["ao1"] = dtAntigua.Rows[i][40].ToString();
                        dr["ap1"] = dtAntigua.Rows[i][41].ToString();
                        dr["aq1"] = dtAntigua.Rows[i][42].ToString();
                        dr["ar1"] = dtAntigua.Rows[i][43].ToString();
                        dr["as1"] = dtAntigua.Rows[i][44].ToString();
                        dr["at1"] = dtAntigua.Rows[i][45].ToString();
                        dr["au1"] = dtAntigua.Rows[i][46].ToString();
                        dr["av1"] = dtAntigua.Rows[i][47].ToString();
                        dr["aw1"] = dtAntigua.Rows[i][48].ToString();
                        dr["ax1"] = dtAntigua.Rows[i][49].ToString();
                        dr["ay1"] = dtAntigua.Rows[i][50].ToString();
                        dr["az1"] = dtAntigua.Rows[i][51].ToString();
                        dr["ba1"] = dtAntigua.Rows[i][52].ToString();
                        dr["bb1"] = dtAntigua.Rows[i][53].ToString();
                        dr["bc1"] = dtAntigua.Rows[i][54].ToString();
                        dr["bd1"] = dtAntigua.Rows[i][55].ToString();
                        dr["be1"] = dtAntigua.Rows[i][56].ToString();
                        dr["bf1"] = dtAntigua.Rows[i][57].ToString();
                        dr["bg1"] = dtAntigua.Rows[i][58].ToString();
                        dr["bh1"] = dtAntigua.Rows[i][59].ToString();
                        dr["bi1"] = dtAntigua.Rows[i][60].ToString();
                        dr["bj1"] = dtAntigua.Rows[i][61].ToString();
                        dr["bk1"] = dtAntigua.Rows[i][62].ToString();
                        dr["bl1"] = dtAntigua.Rows[i][63].ToString();
                        dr["bm1"] = dtAntigua.Rows[i][64].ToString();
                        dr["bn1"] = dtAntigua.Rows[i][65].ToString();
                        dr["bo1"] = dtAntigua.Rows[i][66].ToString();
                        dr["bp1"] = dtAntigua.Rows[i][67].ToString();
                        dr["bq1"] = dtAntigua.Rows[i][68].ToString();
                        dr["br1"] = dtAntigua.Rows[i][69].ToString();
                        dr["bs1"] = dtAntigua.Rows[i][70].ToString();
                        dr["bt1"] = dtAntigua.Rows[i][71].ToString();
                        dr["bu1"] = dtAntigua.Rows[i][72].ToString();
                        dr["bv1"] = dtAntigua.Rows[i][73].ToString();
                        dr["bw1"] = dtAntigua.Rows[i][74].ToString();
                        dr["bx1"] = dtAntigua.Rows[i][75].ToString();
                        dr["by1"] = dtAntigua.Rows[i][76].ToString();
                        dr["bz1"] = dtAntigua.Rows[i][77].ToString();
                        dr["ca1"] = dtAntigua.Rows[i][78].ToString();
                        dr["cb1"] = dtAntigua.Rows[i][79].ToString();
                        dr["cc1"] = dtAntigua.Rows[i][80].ToString();
                        dr["cd1"] = dtAntigua.Rows[i][81].ToString();
                        dr["ce1"] = dtAntigua.Rows[i][82].ToString();
                        dr["cf1"] = dtAntigua.Rows[i][83].ToString();
                        dr["cg1"] = dtAntigua.Rows[i][84].ToString();
                        dr["ch1"] = dtAntigua.Rows[i][85].ToString();
                        dr["ci1"] = dtAntigua.Rows[i][86].ToString();
                        dr["cj1"] = dtAntigua.Rows[i][87].ToString();
                        dr["ck1"] = dtAntigua.Rows[i][88].ToString();
                        dr["cl1"] = dtAntigua.Rows[i][89].ToString();
                        dr["cm1"] = dtAntigua.Rows[i][90].ToString();
                        dr["cn1"] = dtAntigua.Rows[i][91].ToString();
                        dr["co1"] = dtAntigua.Rows[i][92].ToString();
                        dr["cp1"] = dtAntigua.Rows[i][93].ToString();
                        dr["cq1"] = dtAntigua.Rows[i][94].ToString();
                        dr["cr1"] = dtAntigua.Rows[i][95].ToString();
                        dr["cs1"] = dtAntigua.Rows[i][96].ToString();
                        dr["ct1"] = dtAntigua.Rows[i][97].ToString();
                        dr["cu1"] = dtAntigua.Rows[i][98].ToString();
                        dr["cv1"] = dtAntigua.Rows[i][99].ToString();
                        dr["cw1"] = dtAntigua.Rows[i][100].ToString();
                        dr["cx1"] = dtAntigua.Rows[i][101].ToString();
                        dr["cy1"] = dtAntigua.Rows[i][102].ToString();
                        dr["cz1"] = dtAntigua.Rows[i][103].ToString();
                        dr["da1"] = dtAntigua.Rows[i][104].ToString();
                        dr["db1"] = dtAntigua.Rows[i][105].ToString();
                        dr["dc1"] = dtAntigua.Rows[i][106].ToString();
                        dr["dd1"] = dtAntigua.Rows[i][107].ToString();
                        dr["de1"] = dtAntigua.Rows[i][108].ToString();
                        dr["df1"] = dtAntigua.Rows[i][109].ToString();
                        dr["dg1"] = dtAntigua.Rows[i][110].ToString();
                        dr["dh1"] = dtAntigua.Rows[i][111].ToString();
                        dr["di1"] = dtAntigua.Rows[i][112].ToString();
                        dr["dj1"] = dtAntigua.Rows[i][113].ToString();
                        dr["dk1"] = dtAntigua.Rows[i][114].ToString();
                        dr["dl1"] = dtAntigua.Rows[i][115].ToString();
                        dr["dm1"] = dtAntigua.Rows[i][116].ToString();
                        dr["dn1"] = dtAntigua.Rows[i][117].ToString();
                        dr["do1"] = dtAntigua.Rows[i][118].ToString();
                        dr["dp1"] = dtAntigua.Rows[i][119].ToString();
                        dr["dq1"] = dtAntigua.Rows[i][120].ToString();
                        dr["dr1"] = dtAntigua.Rows[i][121].ToString();
                        dr["ds1"] = dtAntigua.Rows[i][122].ToString();
                        dr["dt1"] = dtAntigua.Rows[i][123].ToString();
                        dr["du1"] = dtAntigua.Rows[i][124].ToString();
                        dr["dv1"] = dtAntigua.Rows[i][125].ToString();
                        dr["dw1"] = dtAntigua.Rows[i][126].ToString();
                        dr["dx1"] = dtAntigua.Rows[i][127].ToString();
                        dr["dy1"] = dtAntigua.Rows[i][128].ToString();
                        dr["dz1"] = dtAntigua.Rows[i][129].ToString();
                        dr["ea1"] = dtAntigua.Rows[i][130].ToString();
                        dr["eb1"] = dtAntigua.Rows[i][131].ToString();
                        dr["ec1"] = dtAntigua.Rows[i][132].ToString();
                        dr["ed1"] = dtAntigua.Rows[i][133].ToString();
                        dr["ee1"] = dtAntigua.Rows[i][134].ToString();
                        dr["ef1"] = dtAntigua.Rows[i][135].ToString();
                        dr["eg1"] = dtAntigua.Rows[i][136].ToString();
                        dr["eh1"] = dtAntigua.Rows[i][137].ToString();
                        dr["ei1"] = dtAntigua.Rows[i][138].ToString();
                        dr["ej1"] = dtAntigua.Rows[i][139].ToString();
                        dr["ek1"] = dtAntigua.Rows[i][140].ToString();
                        dr["el1"] = dtAntigua.Rows[i][141].ToString();
                        dr["em1"] = dtAntigua.Rows[i][142].ToString();
                        dr["en1"] = dtAntigua.Rows[i][143].ToString();
                        dr["eo1"] = dtAntigua.Rows[i][144].ToString();
                        dr["ep1"] = dtAntigua.Rows[i][145].ToString();
                        dr["eq1"] = dtAntigua.Rows[i][146].ToString();
                        dr["er1"] = dtAntigua.Rows[i][147].ToString();
                        dr["es1"] = dtAntigua.Rows[i][148].ToString();
                        dr["et1"] = dtAntigua.Rows[i][149].ToString();
                        dr["eu1"] = dtAntigua.Rows[i][150].ToString();
                        dr["ev1"] = dtAntigua.Rows[i][151].ToString();
                        dr["ew1"] = dtAntigua.Rows[i][152].ToString();
                        dr["ex1"] = dtAntigua.Rows[i][153].ToString();
                        dr["ey1"] = dtAntigua.Rows[i][154].ToString();
                        dr["ez1"] = dtAntigua.Rows[i][155].ToString();
                        dr["fa1"] = dtAntigua.Rows[i][156].ToString();
                        dr["fb1"] = dtAntigua.Rows[i][157].ToString();
                        dr["fc1"] = dtAntigua.Rows[i][158].ToString();
                        dr["fd1"] = dtAntigua.Rows[i][159].ToString();
                        dr["fe1"] = dtAntigua.Rows[i][160].ToString();
                        dr["ff1"] = dtAntigua.Rows[i][161].ToString();
                        dr["fg1"] = dtAntigua.Rows[i][162].ToString();
                        dr["fh1"] = dtAntigua.Rows[i][163].ToString();
                        dr["fi1"] = dtAntigua.Rows[i][164].ToString();
                        dr["fj1"] = dtAntigua.Rows[i][165].ToString();
                        dr["fk1"] = dtAntigua.Rows[i][166].ToString();
                        dr["fl1"] = dtAntigua.Rows[i][167].ToString();
                        dr["fm1"] = dtAntigua.Rows[i][168].ToString();
                        dr["fn1"] = dtAntigua.Rows[i][169].ToString();
                        dr["fo1"] = dtAntigua.Rows[i][170].ToString();
                        dr["fp1"] = dtAntigua.Rows[i][171].ToString();
                        dr["fq1"] = dtAntigua.Rows[i][172].ToString();
                        dr["fr1"] = dtAntigua.Rows[i][173].ToString();
                        dr["fs1"] = dtAntigua.Rows[i][174].ToString();
                        dr["ft1"] = dtAntigua.Rows[i][175].ToString();
                        dr["fu1"] = dtAntigua.Rows[i][176].ToString();
                        dr["fv1"] = dtAntigua.Rows[i][177].ToString();
                        dr["fw1"] = dtAntigua.Rows[i][178].ToString();
                        dr["fx1"] = dtAntigua.Rows[i][179].ToString();
                        dr["fy1"] = dtAntigua.Rows[i][180].ToString();
                        dr["fz1"] = dtAntigua.Rows[i][181].ToString();
                        dr["ga1"] = dtAntigua.Rows[i][182].ToString();
                        dr["gb1"] = dtAntigua.Rows[i][183].ToString();
                        dr["gc1"] = dtAntigua.Rows[i][184].ToString();
                        dr["gd1"] = dtAntigua.Rows[i][185].ToString();
                        dr["ge1"] = dtAntigua.Rows[i][186].ToString();
                        dr["gf1"] = dtAntigua.Rows[i][187].ToString();
                        dr["gg1"] = dtAntigua.Rows[i][188].ToString();
                        dr["gh1"] = dtAntigua.Rows[i][189].ToString();
                        dr["gi1"] = dtAntigua.Rows[i][190].ToString();
                        dr["gj1"] = dtAntigua.Rows[i][191].ToString();
                        dr["gk1"] = dtAntigua.Rows[i][192].ToString();
                        dr["gl1"] = dtAntigua.Rows[i][193].ToString();
                        dr["gm1"] = dtAntigua.Rows[i][194].ToString();
                        dr["gn1"] = dtAntigua.Rows[i][195].ToString();
                        dr["go1"] = dtAntigua.Rows[i][196].ToString();
                        dr["gp1"] = dtAntigua.Rows[i][197].ToString();
                        dr["gq1"] = dtAntigua.Rows[i][198].ToString();
                        dr["gr1"] = dtAntigua.Rows[i][199].ToString();
                        dr["gs1"] = dtAntigua.Rows[i][200].ToString();
                        dr["gt1"] = dtAntigua.Rows[i][201].ToString();
                        dr["gu1"] = dtAntigua.Rows[i][202].ToString();
                        dr["gv1"] = dtAntigua.Rows[i][203].ToString();
                        dr["gw1"] = dtAntigua.Rows[i][204].ToString();
                        dr["gx1"] = dtAntigua.Rows[i][205].ToString();
                        dr["gy1"] = dtAntigua.Rows[i][206].ToString();
                        dr["gz1"] = dtAntigua.Rows[i][207].ToString();
                        dr["ha1"] = dtAntigua.Rows[i][208].ToString();
                        dr["hb1"] = dtAntigua.Rows[i][209].ToString();
                        dr["hc1"] = dtAntigua.Rows[i][210].ToString();
                        dr["hd1"] = dtAntigua.Rows[i][211].ToString();
                        dr["he1"] = dtAntigua.Rows[i][212].ToString();
                        dr["hf1"] = dtAntigua.Rows[i][213].ToString();
                        dr["hg1"] = dtAntigua.Rows[i][214].ToString();
                        dr["hh1"] = dtAntigua.Rows[i][215].ToString();
                        dr["hi1"] = dtAntigua.Rows[i][216].ToString();
                        dr["hj1"] = dtAntigua.Rows[i][217].ToString();
                        dr["hk1"] = dtAntigua.Rows[i][218].ToString();
                        dr["hl1"] = dtAntigua.Rows[i][219].ToString();
                        dr["hm1"] = dtAntigua.Rows[i][220].ToString();
                        dr["hn1"] = dtAntigua.Rows[i][221].ToString();
                        dr["ho1"] = dtAntigua.Rows[i][222].ToString();
                        dr["hp1"] = dtAntigua.Rows[i][223].ToString();
                        dr["hq1"] = dtAntigua.Rows[i][224].ToString();
                        dr["hr1"] = dtAntigua.Rows[i][225].ToString();
                        dr["hs1"] = dtAntigua.Rows[i][226].ToString();
                        dr["ht1"] = dtAntigua.Rows[i][227].ToString();
                        dr["hu1"] = dtAntigua.Rows[i][228].ToString();
                        dr["hv1"] = dtAntigua.Rows[i][229].ToString();
                        dr["hw1"] = dtAntigua.Rows[i][230].ToString();
                        dr["hx1"] = dtAntigua.Rows[i][231].ToString();
                        dr["hy1"] = dtAntigua.Rows[i][232].ToString();
                        dr["hz1"] = dtAntigua.Rows[i][233].ToString();
                        dr["ia1"] = dtAntigua.Rows[i][234].ToString();
                        dr["ib1"] = dtAntigua.Rows[i][235].ToString();
                        dr["ic1"] = dtAntigua.Rows[i][236].ToString();
                        dr["id1"] = dtAntigua.Rows[i][237].ToString();
                        dr["ie1"] = dtAntigua.Rows[i][238].ToString();
                        dr["if1"] = dtAntigua.Rows[i][239].ToString();
                        dr["ig1"] = dtAntigua.Rows[i][240].ToString();
                        dr["ih1"] = dtAntigua.Rows[i][241].ToString();
                        dr["ii1"] = dtAntigua.Rows[i][242].ToString();
                        dr["ij1"] = dtAntigua.Rows[i][243].ToString();
                        dr["ik1"] = dtAntigua.Rows[i][244].ToString();
                        dr["il1"] = dtAntigua.Rows[i][245].ToString();
                        dr["im1"] = dtAntigua.Rows[i][246].ToString();
                        dr["in1"] = dtAntigua.Rows[i][247].ToString();
                        dr["io1"] = dtAntigua.Rows[i][248].ToString();
                        dr["ip1"] = dtAntigua.Rows[i][249].ToString();
                        dr["iq1"] = dtAntigua.Rows[i][250].ToString();
                        dr["ir1"] = dtAntigua.Rows[i][251].ToString();
                        dr["is1"] = dtAntigua.Rows[i][252].ToString();
                        dr["it1"] = dtAntigua.Rows[i][253].ToString();
                        dr["iu1"] = dtAntigua.Rows[i][254].ToString();
                        dr["iv1"] = dtAntigua.Rows[i][255].ToString();
                        dr["iw1"] = dtAntigua.Rows[i][256].ToString();
                        dr["ix1"] = dtAntigua.Rows[i][257].ToString();
                        dr["iy1"] = dtAntigua.Rows[i][258].ToString();
                        dr["iz1"] = dtAntigua.Rows[i][259].ToString();

                    }
                    catch (Exception) { }
                    dt.Rows.Add(dr);
                }
            }


            return dt;
        }
        public void Progreso(string msg)
        {
            string enter = char.ConvertFromUtf32(13) + char.ConvertFromUtf32(10);
            string ant_Text = txb_msg.Text;
            txb_msg.Text = ant_Text + msg + enter;
        }
        public enum TipoDeArchivoPlano { Delimited, Fixed }
        public  DataTable LeerArchivoPlano(FileInfo archivo, bool tieneEncabezado)
        {
            string nomTabla = "Hoja1";
            if (!archivo.Exists)
                throw new FileNotFoundException("No se encontró el archivo especificado");

           // string conEncabezado = tieneEncabezado ? "YES" : "NO";

            string connectionString =   @"Provider=Microsoft.ACE.OLEDB.12.0;" +     
                                     "Data Source="+ archivo +";Extended Properties='Excel 8.0;HDR=YES;IMAX=1'";
                
                
                //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + archivo.DirectoryName + "\\;Extended Properties=\"Text;HDR=Yes\"";
            OleDbConnection dbConn = null;
            DataTable resultTable = new DataTable(nomTabla); 
            try
            {
                dbConn = new OleDbConnection(connectionString);
                dbConn.Open();
                DataTable shema = dbConn.GetSchema("Tables");
                string Hoja = shema.Rows[0]["TABLE_NAME"].ToString();
                string query = string.Format("SELECT * FROM ["+Hoja+"]");
                


                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, dbConn))
                {
                    adapter.Fill(resultTable);      // tabla con los datos del excel ya cargado 
                }
            }
            catch (Exception ex)
            {
                this.Invoke(new DisplayEstado(Progreso), "Error: "+ ex.Message );
               // Thread.Sleep(4000);
            }
              
  
           
                //@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};" +
                //"Extended Properties='text;HDR={1};FMT=Delimited(;);'",
                // archivo.DirectoryName, "NO", "Delimited(;)");
          
         //   DataTable dt = new DataTable("Hoja1");
         ////   DataTable dt2 = new DataTable("miTabla");
            
         //   string select = "SELECT * FROM [" + archivo.Name + "]";
         //   try
         //   {
         //       using (OleDbConnection conn = new OleDbConnection(connectionString))
         //       using (OleDbDataAdapter da = new OleDbDataAdapter(select, conn))
         //       { 
         //           da.Fill(dt);
         //       }
         //   }
         //   catch (Exception)
         //   { dt = null; }

            //string dd = dt.Rows[0][0].ToString();
            //string[] separar = dd.Split(';');

            //for (int j = 0; j < separar.Length; j++)
            //{
            //    dt2.Columns.Add("");
            //}

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    try
            //    {

            //        string temp = dt.Rows[i][0].ToString();
            //        temp.Replace(",","&");
            //        string[] separar2 = temp.Split(';');
            //        for (int k = 0; k < separar2.Length; k++)
            //        {
            //            separar2[k] = separar2[k].Replace("&",",");
            //        }

            //        dt2.Rows.Add(separar2);



            //        //string[] separar2 = dt.Rows[i][0].ToString().Split(';');
            //        //dt2.Rows.Add(separar2);
            //    }
            //    catch (Exception)
            //    {

            //    }
            //}
            return resultTable;
        }

      
        public void numProcesado(string msg)
        {
            lbl_mensajes.Text = msg;
        }
        public bool validarEmail(string email)
        {
            // string expresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            string expresion = @"^[\w!#$%&'*+\-/=?\^_`{|}~]+(\.[\w!#$%&'*+\-/=?\^_`{|}~]+)*"
                                    + "@"
                                    + @"((([\-\w]+\.)+[a-zA-Z]{2,7})|(([0-9]{1,3}\.){3}[0-9]{1,3}))$";
            if (Regex.IsMatch(email, expresion))
            {
                if (Regex.Replace(email, expresion, String.Empty).Length == 0)
                {

                    if (email.EndsWith(".con") || email.EndsWith(".CON"))
                    {
                        return false;
                    }
                    else
                    {
                        if (email.Contains("ñ") || email.Contains("Ñ") || email.Contains("á") || email.Contains("'Á") || email.Contains("É") || email.Contains("é") || email.Contains("í") || email.Contains("Í") || email.Contains("ó") || email.Contains("Ó") || email.Contains("Ú") || email.Contains("ú") ||
                            email.Contains("!") || email.Contains("#") || email.Contains("$") || email.Contains("%") || email.Contains("&") || email.Contains("/") || email.Contains("(") || email.Contains(")") || email.Contains("&") || email.Contains("&") || email.Contains("=") || email.Contains("?") ||
                            email.Contains("¡") || email.Contains("¿") || email.Contains("<") || email.Contains(">") || email.Contains(".@") || email.Contains("@.") || email.Contains(" "))
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
                else
                { return false; }
            }
            else
            { return false; }
        }
        public bool validarNumSMS(string email)
        {
            string expresion = "^[0-9]+$";
            if (Regex.IsMatch(email, expresion))
            {

                if (email.Length == 8)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            else
            { return false; }
        }
        public bool enviar_Correo(string para, string from, string nombre, string asunto, string cuerpo)
        {
            bool estado = false;
            try
            {
                MailMessage msg = new MailMessage();
                msg.To.Add(para);
                msg.From = new MailAddress(from, nombre, System.Text.Encoding.UTF8);
                msg.Subject = asunto;

                //El tipo de codificacion del Asunto
                msg.SubjectEncoding = System.Text.Encoding.UTF8;

                //Escribo el mensaje Y su codificacion
                //Especifico si va ha ser interpertado con HTML
                msg.Body = cuerpo;
                msg.IsBodyHtml = true;

                //Creo un objeto de tipo cliente de correo (Por donde se enviara el correo)
                SmtpClient client = new SmtpClient();

                //Si no voy a usar credenciales pongo false, Pero la mayoria de servidores exigen las credenciales para evitar el spam
                //client.UseDefaultCredentials = false;

                //Como voy a utilizar credenciales las pongo
                string credencial = System.Configuration.ConfigurationSettings.AppSettings["credencial"].ToString();
                client.Credentials = new System.Net.NetworkCredential(from, credencial);
                //Si fuera gmail seria 587 el puerto, si es un servidor outlook casi siempre el puerto 25, yo utilizo un servidor propio de correo
                //client.Port = 587;
                string puerto = System.Configuration.ConfigurationSettings.AppSettings["puerto"].ToString();
                client.Port = int.Parse(puerto);

                //identifico el cliente que voy a utilizar
                string Host = System.Configuration.ConfigurationSettings.AppSettings["host"].ToString();
                client.Host = Host;

                //Si fuera a utilizar gmail esto deberia ir en true, esto es un certificado de seguridad
                //client.EnableSsl = true;
                client.EnableSsl = false;
                //Envio el mensaje

                client.Send(msg);
                estado = true;
            }
            catch (Exception)
            {
                estado = false;
            }

            return estado;
        }
        public string buscaCorreo(string id_cliente)
        {
            string correo = "";
            correo = ConexionCall.devuelveValor("select email from Usuario where id_cliente=" + id_cliente);
            if (string.IsNullOrEmpty(correo))
            {
                correo = ConexionCall.devuelveValor("select emailcontacto from Cliente where id_cliente= " + id_cliente);
            }
            return correo;
        }
        public static bool ExportarExcelDataTable(DataTable dt, string RutaExcel)
        {
            try
            {
                const string FIELDSEPARATOR = "\t";
                const string ROWSEPARATOR = "\n";
                StringBuilder output = new StringBuilder();
                // Escribir encabezados    
                foreach (DataColumn dc in dt.Columns)
                {
                    output.Append(dc.ColumnName);
                    output.Append(FIELDSEPARATOR);
                }
                output.Append(ROWSEPARATOR);
                foreach (DataRow item in dt.Rows)
                {
                    foreach (object value in item.ItemArray)
                    {

                        //JM asgregado para evitar que los saltos de lineas destruyan el formato de los excel generados
                        string v=value.ToString();
                        if (value.ToString().Contains(FIELDSEPARATOR)) {
                            
                            v = v.Replace(FIELDSEPARATOR, " "); }



                        if (value.ToString().Contains("\r"))
                        {

                            v = v.Replace("\r", " ");
                        }



                        if (value.ToString().Contains(ROWSEPARATOR))
                        {

                            v = v.Replace(ROWSEPARATOR, " ");
                        }


                        output.Append(v);

                       // output.Append(value.ToString());
                        output.Append(FIELDSEPARATOR);
                    }
                    // Escribir una línea de registro        
                    output.Append(ROWSEPARATOR);
                }
                // Valor de retorno    
                // output.ToString();
                
                FileStream fs = new FileStream(RutaExcel, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF32);

                sw.Write(output.ToString());
                sw.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        protected DataTable traspasoTotal(string id_mensaje, string Id_grupo)
        {
            string sqlEm = "";
            string sqlverifica = "select id_padre from Grupo where Id_Grupo=" + Id_grupo;
            string id_padre = ConexionCall.devuelveValor(sqlverifica);
           
            if (id_padre == "0")
            {
                string nombre_grupo = ConexionCall.devuelveValor("SELECT Nombre FROM [Grupo] where Id_Grupo ="+Id_grupo);

                sqlEm = "select (Case WHEN Id_Mensaje is null THEN '"+id_mensaje+ "' else Id_Mensaje END) as 'ID MENSAJE',Id_Grupo as 'ID Grupo', '"+nombre_grupo+"' as Grupo, id_enviado as 'ID ENVIADO',a1 as 'EMAIL',ec.fecha as 'FECHA ',(Case WHEN enviado is null THEN '0' else enviado END) as 'ENVIADO',(Case WHEN abierto > 0 THEN 1 WHEN abierto = 0 THEN 0 END) as 'LECTURA UNICA', FechaLectura as 'Fecha APERTURA', abierto as 'LECTUAR TOTAL', error as 'REBOTE', em.nombre as 'TIPO REBOTE'";
                sqlEm += " ,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1] ,[s1]";
                sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
                sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] from email e";
                sqlEm += " left join envio_correo ec on e.id_email=ec.id_email and ec.Id_Mensaje = "+id_mensaje+"  left join Estado_mail em on em.id_estado= ec.error";
                sqlEm += " where e.Id_Grupo = "+Id_grupo+" order by fecha desc";
            }
            else
            {
                string SQlSelect = "Select query from Sub_grupos where id_subgrupo = (select id_subgrupo from grupo where id_grupo = (Select Id_Grupo from Mensaje where Id_Mensaje =" + id_mensaje + "))";
                string segmento = ConexionCall.devuelveValor(SQlSelect);

                segmento = segmento.Replace("SELECT *  FROM Email where id_grupo =" + id_padre + " and", "");

                sqlEm = "SELECT a1,ec.*,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1],[s1]";
                sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
                sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] FROM Email e";
                sqlEm += " left join envio_correo ec on e.Id_Email=ec.id_email  and ec.Id_Mensaje = " + id_mensaje + " where id_grupo =" + id_padre + " and " + segmento + "";
            }
          
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            DataTable tablaCos = new DataTable();

            if (Convert.ToInt32(id_padre) > 0)
            {
                #region recorre cuando tiene sub grupo
                int id_cliente = ConexionCall.devuelveValorINT("SELECT Top 1 id_cliente FROM Usuario where Id_Usuario in (Select id_usuario from Mensaje where id_mensaje =" + id_mensaje + ")");
                string error = "";
                int totLec = 0;
                int enviados = 0;
                int errores = 0;
                int unicosss = 0;
                int hdr = 0;
               
                DataRow dr = null;
                for (int i = 0; i < TablaPrin.Rows.Count; i++)
                {
                    if (hdr == 0)
                    {
                        tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ID_Envio", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Fecha", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Enviado", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Lectura Unica", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Fecha Apertura", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Lectura Total", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Rebote", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Tipo Rebote", typeof(string)));


                        tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("L1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("M1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("N1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("O1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("P1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Q1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("R1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("S1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("T1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("U1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("V1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("W1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("X1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Y1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Z1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("AA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AZ1", typeof(string)));

                        #region NO utilizado
                        //tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));

                        //tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));

                        //tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));

                        //tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));

                        //tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));

                        //tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));
                        #endregion

                        hdr = 1;
                    }
                    string abierto = "";
                    if (!string.IsNullOrEmpty(TablaPrin.Rows[i]["enviado"].ToString()))
                    {
                        if (Convert.ToInt32(TablaPrin.Rows[i]["enviado"]) == 1)
                        {
                            enviados++;

                        }
                        abierto = Convert.ToInt32(TablaPrin.Rows[i]["abierto"]).ToString();
                    }
                    else
                    {
                        abierto = "0";

                    }

                    //string abierto = Convert.ToInt32(TablaPrin.Rows[i]["abierto"]).ToString();

                    totLec += Convert.ToInt32(abierto);
                    string unico = "1";
                    if (abierto == "0")
                    {
                        unico = "0";
                    }
                    else
                    {
                        unicosss++;
                    }
                    if (i != 0 && i % 5000 == 0)
                    {
                        Thread.Sleep(20 * 1000);
                    }

                    dr = tablaCos.NewRow();
                    dr["ID_Mensaje"] = id_mensaje;

                    dr["ID_Envio"] = TablaPrin.Rows[i]["id_enviado"];
                    dr["Email"] = TablaPrin.Rows[i]["a1"];
                    dr["Fecha"] = TablaPrin.Rows[i]["fecha"];
                    error = TablaPrin.Rows[i]["error"].ToString();
                    if (string.IsNullOrEmpty(error)) error = "0";
                    if (error == "0")
                    {
                        dr["Rebote"] = "0";
                        dr["Tipo Rebote"] = "";
                    }
                    else
                    {
                        errores++;
                        dr["Rebote"] = "1";
                        dr["Tipo Rebote"] = ConexionCall.devuelveValor("SELECT nombre FROM Estado_mail where id_Estado=" + error);
                    }
                    string enviado = TablaPrin.Rows[i]["enviado"].ToString();
                    if (enviado == "1")
                    {
                        dr["Enviado"] = TablaPrin.Rows[i]["enviado"];
                    }
                    else
                    {
                        dr["Enviado"] = "0";
                        string email = TablaPrin.Rows[i]["a1"].ToString();
                        if (!validarEmail(email))
                        {
                            dr["Tipo Rebote"] = "Sintaxis incorrecta";
                        }
                        else
                        {
                            int descincrito = ConexionCall.devuelveValorINT("SELECT count(*) FROM Email_desincritos where mail='" + email + "' and id_cliente=" + id_cliente);
                            if (descincrito > 0)
                            {
                                dr["Tipo Rebote"] = "Email Desuscrito";
                            }
                            else
                            {
                                int inexistente = ConexionCall.devuelveValorINT("SELECT count(*) FROM email_errores where email='" + email + "'");
                                if (inexistente > 0)
                                {
                                    dr["Tipo Rebote"] = "No existe o inactivo";
                                }
                                else
                                {
                                    dr["Tipo Rebote"] = "No enviado";
                                }

                            }
                        }
                    }
                    dr["Lectura Unica"] = unico;
                    dr["Fecha Apertura"] = TablaPrin.Rows[i]["FechaLectura"];
                    dr["Lectura Total"] = abierto.ToString();


                    dr["B1"] = TablaPrin.Rows[i]["b1"];
                    dr["C1"] = TablaPrin.Rows[i]["c1"];
                    dr["D1"] = TablaPrin.Rows[i]["d1"];
                    dr["e1"] = TablaPrin.Rows[i]["e1"];
                    dr["f1"] = TablaPrin.Rows[i]["f1"];
                    dr["g1"] = TablaPrin.Rows[i]["g1"];
                    dr["h1"] = TablaPrin.Rows[i]["h1"];
                    dr["i1"] = TablaPrin.Rows[i]["i1"];
                    dr["j1"] = TablaPrin.Rows[i]["j1"];
                    dr["k1"] = TablaPrin.Rows[i]["k1"];
                    dr["L1"] = TablaPrin.Rows[i]["L1"];
                    dr["M1"] = TablaPrin.Rows[i]["M1"];
                    dr["N1"] = TablaPrin.Rows[i]["N1"];
                    dr["O1"] = TablaPrin.Rows[i]["O1"];
                    dr["P1"] = TablaPrin.Rows[i]["P1"];
                    dr["Q1"] = TablaPrin.Rows[i]["Q1"];
                    dr["R1"] = TablaPrin.Rows[i]["R1"];
                    dr["S1"] = TablaPrin.Rows[i]["S1"];
                    dr["T1"] = TablaPrin.Rows[i]["T1"];
                    dr["U1"] = TablaPrin.Rows[i]["U1"];
                    dr["V1"] = TablaPrin.Rows[i]["V1"];
                    dr["W1"] = TablaPrin.Rows[i]["W1"];
                    dr["X1"] = TablaPrin.Rows[i]["X1"];
                    dr["Y1"] = TablaPrin.Rows[i]["Y1"];
                    dr["Z1"] = TablaPrin.Rows[i]["Z1"];


                    dr["AA1"] = TablaPrin.Rows[i]["AA1"];
                    dr["AB1"] = TablaPrin.Rows[i]["AB1"];
                    dr["AC1"] = TablaPrin.Rows[i]["AC1"];
                    dr["AD1"] = TablaPrin.Rows[i]["AD1"];
                    dr["AE1"] = TablaPrin.Rows[i]["AE1"];
                    dr["AF1"] = TablaPrin.Rows[i]["AF1"];
                    dr["AG1"] = TablaPrin.Rows[i]["AG1"];
                    dr["AH1"] = TablaPrin.Rows[i]["AH1"];
                    dr["AI1"] = TablaPrin.Rows[i]["AI1"];
                    dr["AJ1"] = TablaPrin.Rows[i]["AJ1"];
                    dr["AK1"] = TablaPrin.Rows[i]["AK1"];
                    dr["AL1"] = TablaPrin.Rows[i]["AL1"];
                    dr["AM1"] = TablaPrin.Rows[i]["AM1"];
                    dr["AN1"] = TablaPrin.Rows[i]["AN1"];
                    dr["AO1"] = TablaPrin.Rows[i]["AO1"];
                    dr["AP1"] = TablaPrin.Rows[i]["AP1"];
                    dr["AQ1"] = TablaPrin.Rows[i]["AQ1"];
                    dr["AR1"] = TablaPrin.Rows[i]["AR1"];
                    dr["AS1"] = TablaPrin.Rows[i]["AS1"];
                    dr["AT1"] = TablaPrin.Rows[i]["AT1"];
                    dr["AU1"] = TablaPrin.Rows[i]["AU1"];
                    dr["AV1"] = TablaPrin.Rows[i]["AV1"];
                    dr["AW1"] = TablaPrin.Rows[i]["AW1"];
                    dr["AX1"] = TablaPrin.Rows[i]["AX1"];
                    dr["AY1"] = TablaPrin.Rows[i]["AY1"];
                    dr["AZ1"] = TablaPrin.Rows[i]["AZ1"];


                    #region NO ulilizado
                    //dr["BA1"] = TablaPrin.Rows[i]["BA1"];
                    //dr["BB1"] = TablaPrin.Rows[i]["BB1"];
                    //dr["BC1"] = TablaPrin.Rows[i]["BC1"];
                    //dr["BD1"] = TablaPrin.Rows[i]["BD1"];
                    //dr["BE1"] = TablaPrin.Rows[i]["BE1"];
                    //dr["BF1"] = TablaPrin.Rows[i]["BF1"];
                    //dr["BG1"] = TablaPrin.Rows[i]["BG1"];
                    //dr["BH1"] = TablaPrin.Rows[i]["BH1"];
                    //dr["BI1"] = TablaPrin.Rows[i]["BI1"];
                    //dr["BJ1"] = TablaPrin.Rows[i]["BJ1"];
                    //dr["BK1"] = TablaPrin.Rows[i]["BK1"];
                    //dr["BL1"] = TablaPrin.Rows[i]["BL1"];
                    //dr["BM1"] = TablaPrin.Rows[i]["BM1"];
                    //dr["BN1"] = TablaPrin.Rows[i]["BN1"];
                    //dr["BO1"] = TablaPrin.Rows[i]["BO1"];
                    //dr["BP1"] = TablaPrin.Rows[i]["BP1"];
                    //dr["BQ1"] = TablaPrin.Rows[i]["BQ1"];
                    //dr["BR1"] = TablaPrin.Rows[i]["BR1"];
                    //dr["BS1"] = TablaPrin.Rows[i]["BS1"];
                    //dr["BT1"] = TablaPrin.Rows[i]["BT1"];
                    //dr["BU1"] = TablaPrin.Rows[i]["BU1"];
                    //dr["BV1"] = TablaPrin.Rows[i]["BV1"];
                    //dr["BW1"] = TablaPrin.Rows[i]["BW1"];
                    //dr["BX1"] = TablaPrin.Rows[i]["BX1"];
                    //dr["BY1"] = TablaPrin.Rows[i]["BY1"];
                    //dr["BZ1"] = TablaPrin.Rows[i]["BZ1"];


                    //dr["CA1"] = TablaPrin.Rows[i]["CA1"];
                    //dr["CB1"] = TablaPrin.Rows[i]["CB1"];
                    //dr["CC1"] = TablaPrin.Rows[i]["CC1"];
                    //dr["CD1"] = TablaPrin.Rows[i]["CD1"];
                    //dr["CE1"] = TablaPrin.Rows[i]["CE1"];
                    //dr["CF1"] = TablaPrin.Rows[i]["CF1"];
                    //dr["CG1"] = TablaPrin.Rows[i]["CG1"];
                    //dr["CH1"] = TablaPrin.Rows[i]["CH1"];
                    //dr["CI1"] = TablaPrin.Rows[i]["CI1"];
                    //dr["CJ1"] = TablaPrin.Rows[i]["CJ1"];
                    //dr["CK1"] = TablaPrin.Rows[i]["CK1"];
                    //dr["CL1"] = TablaPrin.Rows[i]["CL1"];
                    //dr["CM1"] = TablaPrin.Rows[i]["CM1"];
                    //dr["CN1"] = TablaPrin.Rows[i]["CN1"];
                    //dr["CO1"] = TablaPrin.Rows[i]["CO1"];
                    //dr["CP1"] = TablaPrin.Rows[i]["CP1"];
                    //dr["CQ1"] = TablaPrin.Rows[i]["CQ1"];
                    //dr["CR1"] = TablaPrin.Rows[i]["CR1"];
                    //dr["CS1"] = TablaPrin.Rows[i]["CS1"];
                    //dr["CT1"] = TablaPrin.Rows[i]["CT1"];
                    //dr["CU1"] = TablaPrin.Rows[i]["CU1"];
                    //dr["CV1"] = TablaPrin.Rows[i]["CV1"];
                    //dr["CW1"] = TablaPrin.Rows[i]["CW1"];
                    //dr["CX1"] = TablaPrin.Rows[i]["CX1"];
                    //dr["CY1"] = TablaPrin.Rows[i]["CY1"];
                    //dr["CZ1"] = TablaPrin.Rows[i]["CZ1"];

                    //dr["DA1"] = TablaPrin.Rows[i]["DA1"];
                    //dr["DB1"] = TablaPrin.Rows[i]["DB1"];
                    //dr["DC1"] = TablaPrin.Rows[i]["DC1"];
                    //dr["DD1"] = TablaPrin.Rows[i]["DD1"];
                    //dr["DE1"] = TablaPrin.Rows[i]["DE1"];
                    //dr["DF1"] = TablaPrin.Rows[i]["DF1"];
                    //dr["DG1"] = TablaPrin.Rows[i]["DG1"];
                    //dr["DH1"] = TablaPrin.Rows[i]["DH1"];
                    //dr["DI1"] = TablaPrin.Rows[i]["DI1"];
                    //dr["DJ1"] = TablaPrin.Rows[i]["DJ1"];
                    //dr["DK1"] = TablaPrin.Rows[i]["DK1"];
                    //dr["DL1"] = TablaPrin.Rows[i]["DL1"];
                    //dr["DM1"] = TablaPrin.Rows[i]["DM1"];
                    //dr["DN1"] = TablaPrin.Rows[i]["DN1"];
                    //dr["DO1"] = TablaPrin.Rows[i]["DO1"];
                    //dr["DP1"] = TablaPrin.Rows[i]["DP1"];
                    //dr["DQ1"] = TablaPrin.Rows[i]["DQ1"];
                    //dr["DR1"] = TablaPrin.Rows[i]["DR1"];
                    //dr["DS1"] = TablaPrin.Rows[i]["DS1"];
                    //dr["DT1"] = TablaPrin.Rows[i]["DT1"];
                    //dr["DU1"] = TablaPrin.Rows[i]["DU1"];
                    //dr["DV1"] = TablaPrin.Rows[i]["DV1"];
                    //dr["DW1"] = TablaPrin.Rows[i]["DW1"];
                    //dr["DX1"] = TablaPrin.Rows[i]["DX1"];
                    //dr["DY1"] = TablaPrin.Rows[i]["DY1"];
                    //dr["DZ1"] = TablaPrin.Rows[i]["DZ1"];


                    //dr["EA1"] = TablaPrin.Rows[i]["EA1"];
                    //dr["EB1"] = TablaPrin.Rows[i]["EB1"];
                    //dr["EC1"] = TablaPrin.Rows[i]["EC1"];
                    //dr["ED1"] = TablaPrin.Rows[i]["ED1"];
                    //dr["EE1"] = TablaPrin.Rows[i]["EE1"];
                    //dr["EF1"] = TablaPrin.Rows[i]["EF1"];
                    //dr["EG1"] = TablaPrin.Rows[i]["EG1"];
                    //dr["EH1"] = TablaPrin.Rows[i]["EH1"];
                    //dr["EI1"] = TablaPrin.Rows[i]["EI1"];
                    //dr["EJ1"] = TablaPrin.Rows[i]["EJ1"];
                    //dr["EK1"] = TablaPrin.Rows[i]["EK1"];
                    //dr["EL1"] = TablaPrin.Rows[i]["EL1"];
                    //dr["EM1"] = TablaPrin.Rows[i]["EM1"];
                    //dr["EN1"] = TablaPrin.Rows[i]["EN1"];
                    //dr["EO1"] = TablaPrin.Rows[i]["EO1"];
                    //dr["EP1"] = TablaPrin.Rows[i]["EP1"];
                    //dr["EQ1"] = TablaPrin.Rows[i]["EQ1"];
                    //dr["ER1"] = TablaPrin.Rows[i]["ER1"];
                    //dr["ES1"] = TablaPrin.Rows[i]["ES1"];
                    //dr["ET1"] = TablaPrin.Rows[i]["ET1"];
                    //dr["EU1"] = TablaPrin.Rows[i]["EU1"];
                    //dr["EV1"] = TablaPrin.Rows[i]["EV1"];
                    //dr["EW1"] = TablaPrin.Rows[i]["EW1"];
                    //dr["EX1"] = TablaPrin.Rows[i]["EX1"];
                    //dr["EY1"] = TablaPrin.Rows[i]["EY1"];
                    //dr["EZ1"] = TablaPrin.Rows[i]["EZ1"];



                    //dr["FA1"] = TablaPrin.Rows[i]["FA1"];
                    //dr["FB1"] = TablaPrin.Rows[i]["FB1"];
                    //dr["FC1"] = TablaPrin.Rows[i]["FC1"];
                    //dr["FD1"] = TablaPrin.Rows[i]["FD1"];
                    //dr["FE1"] = TablaPrin.Rows[i]["FE1"];
                    //dr["FF1"] = TablaPrin.Rows[i]["FF1"];
                    //dr["FG1"] = TablaPrin.Rows[i]["FG1"];
                    //dr["FH1"] = TablaPrin.Rows[i]["FH1"];
                    //dr["FI1"] = TablaPrin.Rows[i]["FI1"];
                    //dr["FJ1"] = TablaPrin.Rows[i]["FJ1"];
                    //dr["FK1"] = TablaPrin.Rows[i]["FK1"];
                    //dr["FL1"] = TablaPrin.Rows[i]["FL1"];
                    //dr["FM1"] = TablaPrin.Rows[i]["FM1"];
                    //dr["FN1"] = TablaPrin.Rows[i]["FN1"];
                    //dr["FO1"] = TablaPrin.Rows[i]["FO1"];
                    //dr["FP1"] = TablaPrin.Rows[i]["FP1"];
                    //dr["FQ1"] = TablaPrin.Rows[i]["FQ1"];
                    //dr["FR1"] = TablaPrin.Rows[i]["FR1"];
                    //dr["FS1"] = TablaPrin.Rows[i]["FS1"];
                    //dr["FT1"] = TablaPrin.Rows[i]["FT1"];
                    //dr["FU1"] = TablaPrin.Rows[i]["FU1"];
                    //dr["FV1"] = TablaPrin.Rows[i]["FV1"];
                    //dr["FW1"] = TablaPrin.Rows[i]["FW1"];
                    //dr["FX1"] = TablaPrin.Rows[i]["FX1"];
                    //dr["FY1"] = TablaPrin.Rows[i]["FY1"];
                    //dr["FZ1"] = TablaPrin.Rows[i]["FZ1"];


                    //dr["GA1"] = TablaPrin.Rows[i]["GA1"];
                    //dr["GB1"] = TablaPrin.Rows[i]["GB1"];
                    //dr["GC1"] = TablaPrin.Rows[i]["GC1"];
                    //dr["GD1"] = TablaPrin.Rows[i]["GD1"];
                    //dr["GE1"] = TablaPrin.Rows[i]["GE1"];
                    //dr["GF1"] = TablaPrin.Rows[i]["GF1"];
                    //dr["GG1"] = TablaPrin.Rows[i]["GG1"];
                    //dr["GH1"] = TablaPrin.Rows[i]["GH1"];
                    //dr["GI1"] = TablaPrin.Rows[i]["GI1"];
                    //dr["GJ1"] = TablaPrin.Rows[i]["GJ1"];
                    //dr["GK1"] = TablaPrin.Rows[i]["GK1"];
                    //dr["GL1"] = TablaPrin.Rows[i]["GL1"];
                    //dr["GM1"] = TablaPrin.Rows[i]["GM1"];
                    //dr["GN1"] = TablaPrin.Rows[i]["GN1"];
                    //dr["GO1"] = TablaPrin.Rows[i]["GO1"];
                    //dr["GP1"] = TablaPrin.Rows[i]["GP1"];
                    //dr["GQ1"] = TablaPrin.Rows[i]["GQ1"];
                    //dr["GR1"] = TablaPrin.Rows[i]["GR1"];
                    //dr["GS1"] = TablaPrin.Rows[i]["GS1"];
                    //dr["GT1"] = TablaPrin.Rows[i]["GT1"];
                    //dr["GU1"] = TablaPrin.Rows[i]["GU1"];
                    //dr["GV1"] = TablaPrin.Rows[i]["GV1"];
                    //dr["GW1"] = TablaPrin.Rows[i]["GW1"];
                    //dr["GX1"] = TablaPrin.Rows[i]["GX1"];
                    //dr["GY1"] = TablaPrin.Rows[i]["GY1"];
                    //dr["GZ1"] = TablaPrin.Rows[i]["GZ1"];


                    //dr["HA1"] = TablaPrin.Rows[i]["HA1"];
                    //dr["HB1"] = TablaPrin.Rows[i]["HB1"];
                    //dr["HC1"] = TablaPrin.Rows[i]["HC1"];
                    //dr["HD1"] = TablaPrin.Rows[i]["HD1"];
                    //dr["HE1"] = TablaPrin.Rows[i]["HE1"];
                    //dr["HF1"] = TablaPrin.Rows[i]["HF1"];
                    //dr["HG1"] = TablaPrin.Rows[i]["HG1"];
                    //dr["HH1"] = TablaPrin.Rows[i]["HH1"];
                    //dr["HI1"] = TablaPrin.Rows[i]["HI1"];
                    //dr["HJ1"] = TablaPrin.Rows[i]["HJ1"];
                    //dr["HK1"] = TablaPrin.Rows[i]["HK1"];
                    //dr["HL1"] = TablaPrin.Rows[i]["HL1"];
                    //dr["HM1"] = TablaPrin.Rows[i]["HM1"];
                    //dr["HN1"] = TablaPrin.Rows[i]["HN1"];
                    //dr["HO1"] = TablaPrin.Rows[i]["HO1"];
                    //dr["HP1"] = TablaPrin.Rows[i]["HP1"];
                    //dr["HQ1"] = TablaPrin.Rows[i]["HQ1"];
                    //dr["HR1"] = TablaPrin.Rows[i]["HR1"];
                    //dr["HS1"] = TablaPrin.Rows[i]["HS1"];
                    //dr["HT1"] = TablaPrin.Rows[i]["HT1"];
                    //dr["HU1"] = TablaPrin.Rows[i]["HU1"];
                    //dr["HV1"] = TablaPrin.Rows[i]["HV1"];
                    //dr["HW1"] = TablaPrin.Rows[i]["HW1"];
                    //dr["HX1"] = TablaPrin.Rows[i]["HX1"];
                    //dr["HY1"] = TablaPrin.Rows[i]["HY1"];
                    //dr["HZ1"] = TablaPrin.Rows[i]["HZ1"];


                    //dr["IA1"] = TablaPrin.Rows[i]["IA1"];
                    //dr["IB1"] = TablaPrin.Rows[i]["IB1"];
                    //dr["IC1"] = TablaPrin.Rows[i]["IC1"];
                    //dr["ID1"] = TablaPrin.Rows[i]["ID1"];
                    //dr["IE1"] = TablaPrin.Rows[i]["IE1"];
                    //dr["IF1"] = TablaPrin.Rows[i]["IF1"];
                    //dr["IG1"] = TablaPrin.Rows[i]["IG1"];
                    //dr["IH1"] = TablaPrin.Rows[i]["IH1"];
                    //dr["II1"] = TablaPrin.Rows[i]["II1"];
                    //dr["IJ1"] = TablaPrin.Rows[i]["IJ1"];
                    //dr["IK1"] = TablaPrin.Rows[i]["IK1"];
                    //dr["IL1"] = TablaPrin.Rows[i]["IL1"];
                    //dr["IM1"] = TablaPrin.Rows[i]["IM1"];
                    //dr["IN1"] = TablaPrin.Rows[i]["IN1"];
                    //dr["IO1"] = TablaPrin.Rows[i]["IO1"];
                    //dr["IP1"] = TablaPrin.Rows[i]["IP1"];
                    //dr["IQ1"] = TablaPrin.Rows[i]["IQ1"];
                    //dr["IR1"] = TablaPrin.Rows[i]["IR1"];
                    //dr["IS1"] = TablaPrin.Rows[i]["IS1"];
                    //dr["IT1"] = TablaPrin.Rows[i]["IT1"];
                    //dr["IU1"] = TablaPrin.Rows[i]["IU1"];
                    //dr["IV1"] = TablaPrin.Rows[i]["IV1"];
                    //dr["IW1"] = TablaPrin.Rows[i]["IW1"];
                    //dr["IX1"] = TablaPrin.Rows[i]["IX1"];
                    //dr["IY1"] = TablaPrin.Rows[i]["IY1"];
                    //dr["IZ1"] = TablaPrin.Rows[i]["IZ1"];
                    #endregion
                    tablaCos.Rows.Add(dr);

                    this.Invoke(new DisplayEstado(Progreso), "Transpaso N°=" + i + "");
                }
                #endregion
            }

            if (Convert.ToInt32(id_padre) > 0)
            {
                return tablaCos;
            }
            else
            {
                return TablaPrin;
            }
           
                
        }
        protected DataTable traspaso(string id_mensaje, string Id_grupo)
        {

            string sqlEm = "";
            string sqlverifica = "select id_padre from Grupo where Id_Grupo=" + Id_grupo;
            string id_padre = ConexionCall.devuelveValor(sqlverifica);
            int validaAuto = 0;
            int totLec = 0;
            int enviados = 0;
            int errores = 0;
            int unicosss = 0;

            if (id_padre == "0")
            {
                string SqlvalidaAuto = "SELECT COUNT(*) FROM [Mensaje] where Id_Mensaje = "+id_mensaje+" and Id_Estado = 8";
                validaAuto = ConexionCall.devuelveValorINT(SqlvalidaAuto);

                if(validaAuto > 0)
                {
                    string sqlselect = " select sum(enviado) as 'Enviado' , sum(abierto) as 'Total', sum((case when abierto > 0 then 1 end)) as 'Unica' , sum(error) as 'Rebote'";
                           sqlselect += " FROM [envio_correo]";
                           sqlselect += " where Id_Mensaje = "+id_mensaje+"";
                    DataTable totales = ConexionCall.SqlDTable(sqlselect);

                    if(totales.Rows.Count > 0)
                    {
                        try
                        {

                        totLec = Convert.ToInt32(totales.Rows[0]["Total"].ToString());
                        enviados = Convert.ToInt32(totales.Rows[0]["Enviado"].ToString());
                        errores = Convert.ToInt32(totales.Rows[0]["Rebote"].ToString());
                        unicosss = Convert.ToInt32(totales.Rows[0]["Unica"].ToString());
                        }
                        catch (Exception)
                        {
                            totLec = 0;
                            enviados = 0;
                            errores = 0;
                            unicosss = 0;
                        }

                    }

                    sqlEm = " select Id_Mensaje as 'ID MENSAJE',id_enviado as 'ID ENVIADO',a1 as 'EMAIL',ec.fecha as 'FECHA ',enviado as 'ENVIADO',(Case WHEN abierto > 0 THEN 1 WHEN abierto = 0 THEN 0 END) as 'LECTURA UNICA', FechaLectura as 'Fecha APERTURA', abierto as 'LECTUAR TOTAL', error as 'REBOTE', em.nombre as 'TIPO REBOTE'";
                    sqlEm += " ,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1] ,[s1]";
                    sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
                    sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] from envio_correo ec";
                    sqlEm += " join Email e on e.id_email=ec.id_email left join Estado_mail em on em.id_estado= ec.error  where id_mensaje = "+id_mensaje+" and id_grupo = "+Id_grupo+" order by fecha desc";
                }
                else
                {
                sqlEm = "Select id_grupo,Activo,a1,ec.*,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1],[s1] ";
                sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
                sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] ";
                sqlEm += " from Email e left join envio_correo ec on e.Id_Email=ec.id_email  and ec.Id_Mensaje = " + id_mensaje + "";
                sqlEm += " where id_grupo = "+Id_grupo+" order by ec.id_enviado desc";
                }
            }
            else
            {
              //  int Id_grupo_padre = ConexionCall.devuelveValorINT("select id_padre from grupo where id_grupo = (Select Id_Grupo from Mensaje where Id_Mensaje =" + id_mensaje + ")");

                string SQlSelect = "Select query from Sub_grupos where id_subgrupo = (select id_subgrupo from grupo where id_grupo = (Select Id_Grupo from Mensaje where Id_Mensaje ="+id_mensaje+"))";
                string segmento = ConexionCall.devuelveValor(SQlSelect);

                segmento = segmento.Replace("SELECT *  FROM Email where id_grupo ="+id_padre+" and","");

                sqlEm = "SELECT a1,ec.*,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1],[s1]";
                sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
                sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] FROM Email e";
                sqlEm += " left join envio_correo ec on e.Id_Email=ec.id_email  and ec.Id_Mensaje = " + id_mensaje + " where id_grupo =" + id_padre + " and " + segmento + "";


               // sqlEm = "select ec.*,e.id_grupo,a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,[k1],[l1],[m1],[n1],[ñ1],[o1],[p1],[q1],[r1]  ,[s1]";
               // sqlEm += " ,[t1],[u1],[v1],[w1],[x1],[y1],[z1],[aa1],[ab1],[ac1],[ad1],[ae1],[af1],[ag1],[ah1],[ai1],[aj1],[ak1],[al1],[am1],[an1]";
               // sqlEm += " ,[ao1],[ap1],[aq1],[ar1],[as1],[at1],[au1],[av1],[aw1],[ax1],[ay1],[az1] from envio_correo ec ";
               // sqlEm += " join Email e on e.id_email=ec.id_email where id_mensaje=" + id_mensaje + " order by ec.id_enviado desc";

               // sqlEm = query;
            }
            string error = "";
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            int registros = TablaPrin.Rows.Count;

            //JM agergado para oponer registros de la bd cargada. aca recupera elñ total del grupo independiente de si estan buenos o dseinscritos
            //int registros_grupo=ConexionCall.devuelveValorINT("select count (*) from email where id_grupo=( select id_grupo from mensaje where id_mensaje=" + id_mensaje + ")");


            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();

            int id_cliente = ConexionCall.devuelveValorINT("SELECT Top 1 id_cliente FROM Usuario where Id_Usuario in (Select id_usuario from Mensaje where id_mensaje =" + id_mensaje + ")");
            string usuarioAutomatico = ConfigurationSettings.AppSettings["usuarios"].ToString();
            int estAutomatico = ConexionCall.devuelveValorINT("SELECT count(*) FROM Mensaje where id_usuario in ("+usuarioAutomatico+") and id_mensaje="+id_mensaje+" and estatico=1");
            
         
            int hdr = 0;
            if (registros > 0)
            {
                try
                {
                    #region
                    // se borra estadistica anterior
                    string sqlBorr = "delete from Estadisticas_Errores where id_mensaje=" + id_mensaje;
                    objEst.ejecutorBase(sqlBorr);
                    // se ingresa estadistica nueva
                    string sqlinsert = "  insert into Estadisticas_Errores ";
                    sqlinsert += " SELECT Id_Mensaje ,veces,error FROM Moda_Error where id_mensaje=" + id_mensaje;
                    objEst.ejecutorBase(sqlinsert);

                    if (validaAuto < 1)
                    {
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                            if (estAutomatico > 0)
                            {
                                tablaCos.Columns.Add(new DataColumn("Id_Original", typeof(string)));
                            }
                            tablaCos.Columns.Add(new DataColumn("ID_Envio", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Fecha", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Enviado", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Lectura Unica", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Fecha Apertura", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Lectura Total", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Rebote", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Tipo Rebote", typeof(string)));


                            tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("L1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("M1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("N1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("O1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("P1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Q1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("R1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("S1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("T1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("U1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("V1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("W1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("X1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Y1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Z1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("AA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("AZ1", typeof(string)));

                            #region NO utilizado
                            //tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));
                            #endregion

                            hdr = 1;
                        }
                        string abierto = "";
                        if (!string.IsNullOrEmpty(TablaPrin.Rows[i]["enviado"].ToString()))
                        {
                            if (Convert.ToInt32(TablaPrin.Rows[i]["enviado"]) == 1)
                            {
                                enviados++;

                            }
                            abierto = Convert.ToInt32(TablaPrin.Rows[i]["abierto"]).ToString();
                        }
                        else
                        {
                            abierto = "0";

                        }

                        //string abierto = Convert.ToInt32(TablaPrin.Rows[i]["abierto"]).ToString();

                        totLec += Convert.ToInt32(abierto);
                        string unico = "1";
                        if (abierto == "0")
                        {
                            unico = "0";
                        }
                        else
                        {
                            unicosss++;
                        }
                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(20 * 1000);
                        }

                        dr = tablaCos.NewRow();
                        dr["ID_Mensaje"] = id_mensaje;
                        if (estAutomatico > 0)
                        {
                            dr["Id_Original"] = ConexionCall.devuelveValorINT("SELECT top 1 MEN_ID FROM Respaldo_Email_WS where id_email=" + TablaPrin.Rows[i]["id_email"].ToString());
                        }
                        dr["ID_Envio"] = TablaPrin.Rows[i]["id_enviado"];
                        dr["Email"] = TablaPrin.Rows[i]["a1"];
                        dr["Fecha"] = TablaPrin.Rows[i]["fecha"];
                        error = TablaPrin.Rows[i]["error"].ToString();
                        if (string.IsNullOrEmpty(error)) error = "0";
                        if (error == "0")
                        {
                            dr["Rebote"] = "0";
                            dr["Tipo Rebote"] = "";
                        }
                        else
                        {
                            errores++;
                            dr["Rebote"] = "1";
                            dr["Tipo Rebote"] = ConexionCall.devuelveValor("SELECT nombre FROM Estado_mail where id_Estado=" + error);
                        }
                        string enviado = TablaPrin.Rows[i]["enviado"].ToString();
                        if (enviado == "1")
                        {
                            dr["Enviado"] = TablaPrin.Rows[i]["enviado"];
                        }
                        else
                        {
                            dr["Enviado"] = "0";
                            string email = TablaPrin.Rows[i]["a1"].ToString();
                            if (!validarEmail(email))
                            {
                                dr["Tipo Rebote"] = "Sintaxis incorrecta";
                            }
                            else
                            {
                                int descincrito = ConexionCall.devuelveValorINT("SELECT count(*) FROM Email_desincritos where mail='" + email + "' and id_cliente=" + id_cliente);
                                if (descincrito > 0)
                                {
                                    dr["Tipo Rebote"] = "Email Desuscrito";
                                }
                                else
                                {
                                    int inexistente = ConexionCall.devuelveValorINT("SELECT count(*) FROM email_errores where email='" + email + "'");
                                    if (inexistente > 0)
                                    {
                                        dr["Tipo Rebote"] = "No existe o inactivo";
                                    }
                                    else
                                    {
                                        dr["Tipo Rebote"] = "No enviado";
                                    }

                                }
                            }
                        }
                        dr["Lectura Unica"] = unico;
                        dr["Fecha Apertura"] = TablaPrin.Rows[i]["FechaLectura"];
                        dr["Lectura Total"] = abierto.ToString();


                        dr["B1"] = TablaPrin.Rows[i]["b1"];
                        dr["C1"] = TablaPrin.Rows[i]["c1"];
                        dr["D1"] = TablaPrin.Rows[i]["d1"];
                        dr["e1"] = TablaPrin.Rows[i]["e1"];
                        dr["f1"] = TablaPrin.Rows[i]["f1"];
                        dr["g1"] = TablaPrin.Rows[i]["g1"];
                        dr["h1"] = TablaPrin.Rows[i]["h1"];
                        dr["i1"] = TablaPrin.Rows[i]["i1"];
                        dr["j1"] = TablaPrin.Rows[i]["j1"];
                        dr["k1"] = TablaPrin.Rows[i]["k1"];
                        dr["L1"] = TablaPrin.Rows[i]["L1"];
                        dr["M1"] = TablaPrin.Rows[i]["M1"];
                        dr["N1"] = TablaPrin.Rows[i]["N1"];
                        dr["O1"] = TablaPrin.Rows[i]["O1"];
                        dr["P1"] = TablaPrin.Rows[i]["P1"];
                        dr["Q1"] = TablaPrin.Rows[i]["Q1"];
                        dr["R1"] = TablaPrin.Rows[i]["R1"];
                        dr["S1"] = TablaPrin.Rows[i]["S1"];
                        dr["T1"] = TablaPrin.Rows[i]["T1"];
                        dr["U1"] = TablaPrin.Rows[i]["U1"];
                        dr["V1"] = TablaPrin.Rows[i]["V1"];
                        dr["W1"] = TablaPrin.Rows[i]["W1"];
                        dr["X1"] = TablaPrin.Rows[i]["X1"];
                        dr["Y1"] = TablaPrin.Rows[i]["Y1"];
                        dr["Z1"] = TablaPrin.Rows[i]["Z1"];


                        dr["AA1"] = TablaPrin.Rows[i]["AA1"];
                        dr["AB1"] = TablaPrin.Rows[i]["AB1"];
                        dr["AC1"] = TablaPrin.Rows[i]["AC1"];
                        dr["AD1"] = TablaPrin.Rows[i]["AD1"];
                        dr["AE1"] = TablaPrin.Rows[i]["AE1"];
                        dr["AF1"] = TablaPrin.Rows[i]["AF1"];
                        dr["AG1"] = TablaPrin.Rows[i]["AG1"];
                        dr["AH1"] = TablaPrin.Rows[i]["AH1"];
                        dr["AI1"] = TablaPrin.Rows[i]["AI1"];
                        dr["AJ1"] = TablaPrin.Rows[i]["AJ1"];
                        dr["AK1"] = TablaPrin.Rows[i]["AK1"];
                        dr["AL1"] = TablaPrin.Rows[i]["AL1"];
                        dr["AM1"] = TablaPrin.Rows[i]["AM1"];
                        dr["AN1"] = TablaPrin.Rows[i]["AN1"];
                        dr["AO1"] = TablaPrin.Rows[i]["AO1"];
                        dr["AP1"] = TablaPrin.Rows[i]["AP1"];
                        dr["AQ1"] = TablaPrin.Rows[i]["AQ1"];
                        dr["AR1"] = TablaPrin.Rows[i]["AR1"];
                        dr["AS1"] = TablaPrin.Rows[i]["AS1"];
                        dr["AT1"] = TablaPrin.Rows[i]["AT1"];
                        dr["AU1"] = TablaPrin.Rows[i]["AU1"];
                        dr["AV1"] = TablaPrin.Rows[i]["AV1"];
                        dr["AW1"] = TablaPrin.Rows[i]["AW1"];
                        dr["AX1"] = TablaPrin.Rows[i]["AX1"];
                        dr["AY1"] = TablaPrin.Rows[i]["AY1"];
                        dr["AZ1"] = TablaPrin.Rows[i]["AZ1"];


                        #region NO ulilizado
                        //dr["BA1"] = TablaPrin.Rows[i]["BA1"];
                        //dr["BB1"] = TablaPrin.Rows[i]["BB1"];
                        //dr["BC1"] = TablaPrin.Rows[i]["BC1"];
                        //dr["BD1"] = TablaPrin.Rows[i]["BD1"];
                        //dr["BE1"] = TablaPrin.Rows[i]["BE1"];
                        //dr["BF1"] = TablaPrin.Rows[i]["BF1"];
                        //dr["BG1"] = TablaPrin.Rows[i]["BG1"];
                        //dr["BH1"] = TablaPrin.Rows[i]["BH1"];
                        //dr["BI1"] = TablaPrin.Rows[i]["BI1"];
                        //dr["BJ1"] = TablaPrin.Rows[i]["BJ1"];
                        //dr["BK1"] = TablaPrin.Rows[i]["BK1"];
                        //dr["BL1"] = TablaPrin.Rows[i]["BL1"];
                        //dr["BM1"] = TablaPrin.Rows[i]["BM1"];
                        //dr["BN1"] = TablaPrin.Rows[i]["BN1"];
                        //dr["BO1"] = TablaPrin.Rows[i]["BO1"];
                        //dr["BP1"] = TablaPrin.Rows[i]["BP1"];
                        //dr["BQ1"] = TablaPrin.Rows[i]["BQ1"];
                        //dr["BR1"] = TablaPrin.Rows[i]["BR1"];
                        //dr["BS1"] = TablaPrin.Rows[i]["BS1"];
                        //dr["BT1"] = TablaPrin.Rows[i]["BT1"];
                        //dr["BU1"] = TablaPrin.Rows[i]["BU1"];
                        //dr["BV1"] = TablaPrin.Rows[i]["BV1"];
                        //dr["BW1"] = TablaPrin.Rows[i]["BW1"];
                        //dr["BX1"] = TablaPrin.Rows[i]["BX1"];
                        //dr["BY1"] = TablaPrin.Rows[i]["BY1"];
                        //dr["BZ1"] = TablaPrin.Rows[i]["BZ1"];


                        //dr["CA1"] = TablaPrin.Rows[i]["CA1"];
                        //dr["CB1"] = TablaPrin.Rows[i]["CB1"];
                        //dr["CC1"] = TablaPrin.Rows[i]["CC1"];
                        //dr["CD1"] = TablaPrin.Rows[i]["CD1"];
                        //dr["CE1"] = TablaPrin.Rows[i]["CE1"];
                        //dr["CF1"] = TablaPrin.Rows[i]["CF1"];
                        //dr["CG1"] = TablaPrin.Rows[i]["CG1"];
                        //dr["CH1"] = TablaPrin.Rows[i]["CH1"];
                        //dr["CI1"] = TablaPrin.Rows[i]["CI1"];
                        //dr["CJ1"] = TablaPrin.Rows[i]["CJ1"];
                        //dr["CK1"] = TablaPrin.Rows[i]["CK1"];
                        //dr["CL1"] = TablaPrin.Rows[i]["CL1"];
                        //dr["CM1"] = TablaPrin.Rows[i]["CM1"];
                        //dr["CN1"] = TablaPrin.Rows[i]["CN1"];
                        //dr["CO1"] = TablaPrin.Rows[i]["CO1"];
                        //dr["CP1"] = TablaPrin.Rows[i]["CP1"];
                        //dr["CQ1"] = TablaPrin.Rows[i]["CQ1"];
                        //dr["CR1"] = TablaPrin.Rows[i]["CR1"];
                        //dr["CS1"] = TablaPrin.Rows[i]["CS1"];
                        //dr["CT1"] = TablaPrin.Rows[i]["CT1"];
                        //dr["CU1"] = TablaPrin.Rows[i]["CU1"];
                        //dr["CV1"] = TablaPrin.Rows[i]["CV1"];
                        //dr["CW1"] = TablaPrin.Rows[i]["CW1"];
                        //dr["CX1"] = TablaPrin.Rows[i]["CX1"];
                        //dr["CY1"] = TablaPrin.Rows[i]["CY1"];
                        //dr["CZ1"] = TablaPrin.Rows[i]["CZ1"];

                        //dr["DA1"] = TablaPrin.Rows[i]["DA1"];
                        //dr["DB1"] = TablaPrin.Rows[i]["DB1"];
                        //dr["DC1"] = TablaPrin.Rows[i]["DC1"];
                        //dr["DD1"] = TablaPrin.Rows[i]["DD1"];
                        //dr["DE1"] = TablaPrin.Rows[i]["DE1"];
                        //dr["DF1"] = TablaPrin.Rows[i]["DF1"];
                        //dr["DG1"] = TablaPrin.Rows[i]["DG1"];
                        //dr["DH1"] = TablaPrin.Rows[i]["DH1"];
                        //dr["DI1"] = TablaPrin.Rows[i]["DI1"];
                        //dr["DJ1"] = TablaPrin.Rows[i]["DJ1"];
                        //dr["DK1"] = TablaPrin.Rows[i]["DK1"];
                        //dr["DL1"] = TablaPrin.Rows[i]["DL1"];
                        //dr["DM1"] = TablaPrin.Rows[i]["DM1"];
                        //dr["DN1"] = TablaPrin.Rows[i]["DN1"];
                        //dr["DO1"] = TablaPrin.Rows[i]["DO1"];
                        //dr["DP1"] = TablaPrin.Rows[i]["DP1"];
                        //dr["DQ1"] = TablaPrin.Rows[i]["DQ1"];
                        //dr["DR1"] = TablaPrin.Rows[i]["DR1"];
                        //dr["DS1"] = TablaPrin.Rows[i]["DS1"];
                        //dr["DT1"] = TablaPrin.Rows[i]["DT1"];
                        //dr["DU1"] = TablaPrin.Rows[i]["DU1"];
                        //dr["DV1"] = TablaPrin.Rows[i]["DV1"];
                        //dr["DW1"] = TablaPrin.Rows[i]["DW1"];
                        //dr["DX1"] = TablaPrin.Rows[i]["DX1"];
                        //dr["DY1"] = TablaPrin.Rows[i]["DY1"];
                        //dr["DZ1"] = TablaPrin.Rows[i]["DZ1"];


                        //dr["EA1"] = TablaPrin.Rows[i]["EA1"];
                        //dr["EB1"] = TablaPrin.Rows[i]["EB1"];
                        //dr["EC1"] = TablaPrin.Rows[i]["EC1"];
                        //dr["ED1"] = TablaPrin.Rows[i]["ED1"];
                        //dr["EE1"] = TablaPrin.Rows[i]["EE1"];
                        //dr["EF1"] = TablaPrin.Rows[i]["EF1"];
                        //dr["EG1"] = TablaPrin.Rows[i]["EG1"];
                        //dr["EH1"] = TablaPrin.Rows[i]["EH1"];
                        //dr["EI1"] = TablaPrin.Rows[i]["EI1"];
                        //dr["EJ1"] = TablaPrin.Rows[i]["EJ1"];
                        //dr["EK1"] = TablaPrin.Rows[i]["EK1"];
                        //dr["EL1"] = TablaPrin.Rows[i]["EL1"];
                        //dr["EM1"] = TablaPrin.Rows[i]["EM1"];
                        //dr["EN1"] = TablaPrin.Rows[i]["EN1"];
                        //dr["EO1"] = TablaPrin.Rows[i]["EO1"];
                        //dr["EP1"] = TablaPrin.Rows[i]["EP1"];
                        //dr["EQ1"] = TablaPrin.Rows[i]["EQ1"];
                        //dr["ER1"] = TablaPrin.Rows[i]["ER1"];
                        //dr["ES1"] = TablaPrin.Rows[i]["ES1"];
                        //dr["ET1"] = TablaPrin.Rows[i]["ET1"];
                        //dr["EU1"] = TablaPrin.Rows[i]["EU1"];
                        //dr["EV1"] = TablaPrin.Rows[i]["EV1"];
                        //dr["EW1"] = TablaPrin.Rows[i]["EW1"];
                        //dr["EX1"] = TablaPrin.Rows[i]["EX1"];
                        //dr["EY1"] = TablaPrin.Rows[i]["EY1"];
                        //dr["EZ1"] = TablaPrin.Rows[i]["EZ1"];



                        //dr["FA1"] = TablaPrin.Rows[i]["FA1"];
                        //dr["FB1"] = TablaPrin.Rows[i]["FB1"];
                        //dr["FC1"] = TablaPrin.Rows[i]["FC1"];
                        //dr["FD1"] = TablaPrin.Rows[i]["FD1"];
                        //dr["FE1"] = TablaPrin.Rows[i]["FE1"];
                        //dr["FF1"] = TablaPrin.Rows[i]["FF1"];
                        //dr["FG1"] = TablaPrin.Rows[i]["FG1"];
                        //dr["FH1"] = TablaPrin.Rows[i]["FH1"];
                        //dr["FI1"] = TablaPrin.Rows[i]["FI1"];
                        //dr["FJ1"] = TablaPrin.Rows[i]["FJ1"];
                        //dr["FK1"] = TablaPrin.Rows[i]["FK1"];
                        //dr["FL1"] = TablaPrin.Rows[i]["FL1"];
                        //dr["FM1"] = TablaPrin.Rows[i]["FM1"];
                        //dr["FN1"] = TablaPrin.Rows[i]["FN1"];
                        //dr["FO1"] = TablaPrin.Rows[i]["FO1"];
                        //dr["FP1"] = TablaPrin.Rows[i]["FP1"];
                        //dr["FQ1"] = TablaPrin.Rows[i]["FQ1"];
                        //dr["FR1"] = TablaPrin.Rows[i]["FR1"];
                        //dr["FS1"] = TablaPrin.Rows[i]["FS1"];
                        //dr["FT1"] = TablaPrin.Rows[i]["FT1"];
                        //dr["FU1"] = TablaPrin.Rows[i]["FU1"];
                        //dr["FV1"] = TablaPrin.Rows[i]["FV1"];
                        //dr["FW1"] = TablaPrin.Rows[i]["FW1"];
                        //dr["FX1"] = TablaPrin.Rows[i]["FX1"];
                        //dr["FY1"] = TablaPrin.Rows[i]["FY1"];
                        //dr["FZ1"] = TablaPrin.Rows[i]["FZ1"];


                        //dr["GA1"] = TablaPrin.Rows[i]["GA1"];
                        //dr["GB1"] = TablaPrin.Rows[i]["GB1"];
                        //dr["GC1"] = TablaPrin.Rows[i]["GC1"];
                        //dr["GD1"] = TablaPrin.Rows[i]["GD1"];
                        //dr["GE1"] = TablaPrin.Rows[i]["GE1"];
                        //dr["GF1"] = TablaPrin.Rows[i]["GF1"];
                        //dr["GG1"] = TablaPrin.Rows[i]["GG1"];
                        //dr["GH1"] = TablaPrin.Rows[i]["GH1"];
                        //dr["GI1"] = TablaPrin.Rows[i]["GI1"];
                        //dr["GJ1"] = TablaPrin.Rows[i]["GJ1"];
                        //dr["GK1"] = TablaPrin.Rows[i]["GK1"];
                        //dr["GL1"] = TablaPrin.Rows[i]["GL1"];
                        //dr["GM1"] = TablaPrin.Rows[i]["GM1"];
                        //dr["GN1"] = TablaPrin.Rows[i]["GN1"];
                        //dr["GO1"] = TablaPrin.Rows[i]["GO1"];
                        //dr["GP1"] = TablaPrin.Rows[i]["GP1"];
                        //dr["GQ1"] = TablaPrin.Rows[i]["GQ1"];
                        //dr["GR1"] = TablaPrin.Rows[i]["GR1"];
                        //dr["GS1"] = TablaPrin.Rows[i]["GS1"];
                        //dr["GT1"] = TablaPrin.Rows[i]["GT1"];
                        //dr["GU1"] = TablaPrin.Rows[i]["GU1"];
                        //dr["GV1"] = TablaPrin.Rows[i]["GV1"];
                        //dr["GW1"] = TablaPrin.Rows[i]["GW1"];
                        //dr["GX1"] = TablaPrin.Rows[i]["GX1"];
                        //dr["GY1"] = TablaPrin.Rows[i]["GY1"];
                        //dr["GZ1"] = TablaPrin.Rows[i]["GZ1"];


                        //dr["HA1"] = TablaPrin.Rows[i]["HA1"];
                        //dr["HB1"] = TablaPrin.Rows[i]["HB1"];
                        //dr["HC1"] = TablaPrin.Rows[i]["HC1"];
                        //dr["HD1"] = TablaPrin.Rows[i]["HD1"];
                        //dr["HE1"] = TablaPrin.Rows[i]["HE1"];
                        //dr["HF1"] = TablaPrin.Rows[i]["HF1"];
                        //dr["HG1"] = TablaPrin.Rows[i]["HG1"];
                        //dr["HH1"] = TablaPrin.Rows[i]["HH1"];
                        //dr["HI1"] = TablaPrin.Rows[i]["HI1"];
                        //dr["HJ1"] = TablaPrin.Rows[i]["HJ1"];
                        //dr["HK1"] = TablaPrin.Rows[i]["HK1"];
                        //dr["HL1"] = TablaPrin.Rows[i]["HL1"];
                        //dr["HM1"] = TablaPrin.Rows[i]["HM1"];
                        //dr["HN1"] = TablaPrin.Rows[i]["HN1"];
                        //dr["HO1"] = TablaPrin.Rows[i]["HO1"];
                        //dr["HP1"] = TablaPrin.Rows[i]["HP1"];
                        //dr["HQ1"] = TablaPrin.Rows[i]["HQ1"];
                        //dr["HR1"] = TablaPrin.Rows[i]["HR1"];
                        //dr["HS1"] = TablaPrin.Rows[i]["HS1"];
                        //dr["HT1"] = TablaPrin.Rows[i]["HT1"];
                        //dr["HU1"] = TablaPrin.Rows[i]["HU1"];
                        //dr["HV1"] = TablaPrin.Rows[i]["HV1"];
                        //dr["HW1"] = TablaPrin.Rows[i]["HW1"];
                        //dr["HX1"] = TablaPrin.Rows[i]["HX1"];
                        //dr["HY1"] = TablaPrin.Rows[i]["HY1"];
                        //dr["HZ1"] = TablaPrin.Rows[i]["HZ1"];


                        //dr["IA1"] = TablaPrin.Rows[i]["IA1"];
                        //dr["IB1"] = TablaPrin.Rows[i]["IB1"];
                        //dr["IC1"] = TablaPrin.Rows[i]["IC1"];
                        //dr["ID1"] = TablaPrin.Rows[i]["ID1"];
                        //dr["IE1"] = TablaPrin.Rows[i]["IE1"];
                        //dr["IF1"] = TablaPrin.Rows[i]["IF1"];
                        //dr["IG1"] = TablaPrin.Rows[i]["IG1"];
                        //dr["IH1"] = TablaPrin.Rows[i]["IH1"];
                        //dr["II1"] = TablaPrin.Rows[i]["II1"];
                        //dr["IJ1"] = TablaPrin.Rows[i]["IJ1"];
                        //dr["IK1"] = TablaPrin.Rows[i]["IK1"];
                        //dr["IL1"] = TablaPrin.Rows[i]["IL1"];
                        //dr["IM1"] = TablaPrin.Rows[i]["IM1"];
                        //dr["IN1"] = TablaPrin.Rows[i]["IN1"];
                        //dr["IO1"] = TablaPrin.Rows[i]["IO1"];
                        //dr["IP1"] = TablaPrin.Rows[i]["IP1"];
                        //dr["IQ1"] = TablaPrin.Rows[i]["IQ1"];
                        //dr["IR1"] = TablaPrin.Rows[i]["IR1"];
                        //dr["IS1"] = TablaPrin.Rows[i]["IS1"];
                        //dr["IT1"] = TablaPrin.Rows[i]["IT1"];
                        //dr["IU1"] = TablaPrin.Rows[i]["IU1"];
                        //dr["IV1"] = TablaPrin.Rows[i]["IV1"];
                        //dr["IW1"] = TablaPrin.Rows[i]["IW1"];
                        //dr["IX1"] = TablaPrin.Rows[i]["IX1"];
                        //dr["IY1"] = TablaPrin.Rows[i]["IY1"];
                        //dr["IZ1"] = TablaPrin.Rows[i]["IZ1"];
                        #endregion
                        tablaCos.Rows.Add(dr);

                        this.Invoke(new DisplayEstado(Progreso), "Transpaso N°=" + i + "");
                    }
                }
                    objEst.ejecutorBase("delete from Estadisticas where id_mensaje=" + id_mensaje);
                    string sqlEst = "INSERT INTO Estadisticas (id_mensaje ,unico ,errores,enviados,total) VALUES ";
                    sqlEst += "  (" + id_mensaje + " ," + unicosss + "," + errores + "," + enviados + "," + totLec + ")";
                    objEst.ejecutorBase(sqlEst);
                    //////actualiza mensaje

                    string sqlMsg = "UPDATE Mensaje ";
                    sqlMsg += " SET CodigoError = " + errores;
                    sqlMsg += " ,Enviados = " + enviados;
                    sqlMsg += " ,Leidos = " + totLec;
                    sqlMsg += " ,unico = " + unicosss;
                   // sqlMsg += " ,reg = " + registros; //JM cambiado 09042015
                   // sqlMsg += " ,reg = " + registros_grupo; //JM sacado para que no se actualice, la idea es guardar la cantidad de registros del grupo AL MOMENTO del envío
                    sqlMsg += " WHERE id_mensaje=" + id_mensaje;
                    objEst.ejecutorBase(sqlMsg);

                    #endregion
                }
                catch( Exception ex)
                {}
            }
            if (validaAuto > 0)
            {
                return TablaPrin;
            }
            else
            {
                return tablaCos;
            }
        }
        protected DataTable traspaso2(string id_grupo, string nombreGrupo, string id_mensaje)
        {
            int id_padre = ConexionCall.devuelveValorINT("SELECT id_padre FROM Grupo  where id_grupo="+id_grupo);
            ConexionCall ojjURL = new ConexionCall();
            string sqlEm = " select a1,url,abierto,columna ";
            sqlEm += " from Url_abiertas ua ";
            sqlEm += " join Email e on e.id_email=ua.id_email ";
            sqlEm += " where  e.id_grupo=" + id_grupo + " and id_mensaje="+id_mensaje;
            sqlEm += " order by a1,url ";

            if (id_padre != 0)
            {               
                sqlEm = " select a1,url,abierto,columna ";
                sqlEm += " from Url_abiertas ua ";
                sqlEm += " join Email e on e.id_email=ua.id_email ";
                sqlEm += " where  e.id_grupo=" + id_padre + " and id_mensaje=" + id_mensaje;
                sqlEm += " order by a1,url ";
            }
            
            ojjURL.ejecutorBase("delete from estadisticas_url where id_mensaje="+id_mensaje);

            string inserta="INSERT INTO estadisticas_url (url,unico ,total,id_mensaje)";
            inserta += " select url,count(ua.id_email)unico,sum(abierto)total,'"+id_mensaje+"' id_mensaje ";
            inserta += " from Url_abiertas ua join Email e on e.id_email=ua.id_email ";
            inserta += " where ua.id_mensaje=" + id_mensaje + " group by url";

            ojjURL.ejecutorBase(inserta);
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();
            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {
                try
                {
                    #region
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("URL", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Abierto", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Variable", typeof(string)));
                            hdr = 1;
                        }

                        dr = tablaCos.NewRow();
                        dr["ID_Mensaje"] = id_mensaje;
                        dr["Email"] = TablaPrin.Rows[i]["a1"];
                        dr["URL"] = TablaPrin.Rows[i]["url"];
                        dr["Abierto"] = TablaPrin.Rows[i]["abierto"];
                        dr["Variable"] = TablaPrin.Rows[i]["columna"];
                        tablaCos.Rows.Add(dr);
                    }
                    #endregion
                }
                 catch(Exception ex){}
           }
            return tablaCos;
        }
        protected DataTable traspaso3(string id_mensaje)
        {
            string sqlEm = "  select id_click,a1,visto ";
            sqlEm += " from Click_aqui ca  ";
            sqlEm += " join Email e on ca.id_email=e.id_email  ";
            sqlEm += " where ca.id_mensaje=" + id_mensaje;
            sqlEm += " order by a1";
         
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();
            
            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {
                try
                {
                    #region
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Visto", typeof(string)));
                            /*    tablaCos.Columns.Add(new DataColumn("Fecha", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Enviado", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Lectura Unica", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Fecha Apertura", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Lectura Total", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Error", typeof(string)));
                                tablaCos.Columns.Add(new DataColumn("Tipo Error", typeof(string)));*/
                            hdr = 1;
                        }

                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(20 * 1000);
                        }

                        dr = tablaCos.NewRow();
                        dr["ID_Mensaje"] = id_mensaje;
                        dr["Email"] = TablaPrin.Rows[i]["a1"];
                        dr["Visto"] = TablaPrin.Rows[i]["visto"];
                        /*    dr["Enviado"] = TablaPrin.Rows[i]["enviado"];
                            dr["Lectura Unica"] = unico;
                            dr["Fecha Apertura"] = TablaPrin.Rows[i]["FechaLectura"];
                            dr["Lectura Total"] = abierto.ToString();*/


                        tablaCos.Rows.Add(dr);
                    }
                    #endregion
                }catch( Exception ex)
                {}
              }
            return tablaCos;
        }
        protected DataTable traspaso4(string id_grupo,string id_mensaje)
        {
            int id_padre = ConexionCall.devuelveValorINT("SELECT  id_padre FROM Grupo  where id_grupo=" + id_grupo);
            string sqlEm = " SELECT d.id_email,e.* FROM Email_desincritos d inner join Email e on d.id_email=e.Id_Email where d.id_grupo="+id_grupo+" order by d.mail";

            if (id_padre != 0)
            {
                sqlEm = " SELECT mail FROM Email_desincritos  where id_grupo=" + id_padre + "  order by mail";
            }
            ConexionCall actualiza=new ConexionCall();
            actualiza.ejecutorBase("delete FROM estadisticas_desincritos where id_mensaje="+id_mensaje);
         
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            int emails=TablaPrin.Rows.Count;
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();

            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {
                try
                {
                    #region
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("a1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("b1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("c1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("d1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("e1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("f1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("g1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("h1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("i1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("j1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("k1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("l1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("m1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("n1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("o1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("p1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("q1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("r1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("s1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("t1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("u1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("v1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("w1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("x1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("y1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("z1", typeof(string)));


                            tablaCos.Columns.Add(new DataColumn("aa1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ab1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ac1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ad1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ae1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("af1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ag1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ah1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ai1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("aj1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ak1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("al1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("am1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("an1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ao1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ap1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("aq1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ar1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("as1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("at1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("au1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("av1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("aw1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ax1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ay1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("az1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));

                            tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));


                            hdr = 1;
                        }

                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(20 * 2000);
                        }
                        dr = tablaCos.NewRow();

                        dr["ID_Mensaje"] = id_mensaje;
                        dr["a1"] = TablaPrin.Rows[i]["a1"];
                        dr["b1"] = TablaPrin.Rows[i]["b1"];
                        dr["c1"] = TablaPrin.Rows[i]["c1"];
                        dr["d1"] = TablaPrin.Rows[i]["d1"];
                        dr["e1"] = TablaPrin.Rows[i]["e1"];
                        dr["f1"] = TablaPrin.Rows[i]["f1"];
                        dr["g1"] = TablaPrin.Rows[i]["g1"];
                        dr["h1"] = TablaPrin.Rows[i]["h1"];
                        dr["h1"] = TablaPrin.Rows[i]["i1"];
                        dr["j1"] = TablaPrin.Rows[i]["j1"];
                        dr["k1"] = TablaPrin.Rows[i]["k1"];
                        dr["l1"] = TablaPrin.Rows[i]["l1"];
                        dr["m1"] = TablaPrin.Rows[i]["m1"];
                        dr["n1"] = TablaPrin.Rows[i]["n1"];
                        dr["o1"] = TablaPrin.Rows[i]["o1"];
                        dr["p1"] = TablaPrin.Rows[i]["p1"];
                        dr["q1"] = TablaPrin.Rows[i]["q1"];
                        dr["r1"] = TablaPrin.Rows[i]["r1"];
                        dr["s1"] = TablaPrin.Rows[i]["s1"];
                        dr["t1"] = TablaPrin.Rows[i]["t1"];
                        dr["u1"] = TablaPrin.Rows[i]["u1"];
                        dr["v1"] = TablaPrin.Rows[i]["v1"];
                        dr["w1"] = TablaPrin.Rows[i]["w1"];
                        dr["x1"] = TablaPrin.Rows[i]["x1"];
                        dr["y1"] = TablaPrin.Rows[i]["y1"];
                        dr["z1"] = TablaPrin.Rows[i]["z1"];


                        dr["aa1"] = TablaPrin.Rows[i]["aa1"];
                        dr["ab1"] = TablaPrin.Rows[i]["ab1"];
                        dr["ac1"] = TablaPrin.Rows[i]["ac1"];
                        dr["ad1"] = TablaPrin.Rows[i]["ad1"];
                        dr["ae1"] = TablaPrin.Rows[i]["ae1"];
                        dr["af1"] = TablaPrin.Rows[i]["af1"];
                        dr["ag1"] = TablaPrin.Rows[i]["ag1"];
                        dr["ah1"] = TablaPrin.Rows[i]["ah1"];
                        dr["ah1"] = TablaPrin.Rows[i]["ai1"];
                        dr["aj1"] = TablaPrin.Rows[i]["aj1"];
                        dr["ak1"] = TablaPrin.Rows[i]["ak1"];
                        dr["al1"] = TablaPrin.Rows[i]["al1"];
                        dr["am1"] = TablaPrin.Rows[i]["am1"];
                        dr["an1"] = TablaPrin.Rows[i]["an1"];
                        dr["ao1"] = TablaPrin.Rows[i]["ao1"];
                        dr["ap1"] = TablaPrin.Rows[i]["ap1"];
                        dr["aq1"] = TablaPrin.Rows[i]["aq1"];
                        dr["ar1"] = TablaPrin.Rows[i]["ar1"];
                        dr["as1"] = TablaPrin.Rows[i]["as1"];
                        dr["at1"] = TablaPrin.Rows[i]["at1"];
                        dr["au1"] = TablaPrin.Rows[i]["au1"];
                        dr["av1"] = TablaPrin.Rows[i]["av1"];
                        dr["aw1"] = TablaPrin.Rows[i]["aw1"];
                        dr["ax1"] = TablaPrin.Rows[i]["ax1"];
                        dr["ay1"] = TablaPrin.Rows[i]["ay1"];
                        dr["az1"] = TablaPrin.Rows[i]["az1"];

                        dr["BA1"] = TablaPrin.Rows[i]["BA1"];
                        dr["BB1"] = TablaPrin.Rows[i]["BB1"];
                        dr["BC1"] = TablaPrin.Rows[i]["BC1"];
                        dr["BD1"] = TablaPrin.Rows[i]["BD1"];
                        dr["BE1"] = TablaPrin.Rows[i]["BE1"];
                        dr["BF1"] = TablaPrin.Rows[i]["BF1"];
                        dr["BG1"] = TablaPrin.Rows[i]["BG1"];
                        dr["BH1"] = TablaPrin.Rows[i]["BH1"];
                        dr["BI1"] = TablaPrin.Rows[i]["BI1"];
                        dr["BJ1"] = TablaPrin.Rows[i]["BJ1"];
                        dr["BK1"] = TablaPrin.Rows[i]["BK1"];
                        dr["BL1"] = TablaPrin.Rows[i]["BL1"];
                        dr["BM1"] = TablaPrin.Rows[i]["BM1"];
                        dr["BN1"] = TablaPrin.Rows[i]["BN1"];
                        dr["BO1"] = TablaPrin.Rows[i]["BO1"];
                        dr["BP1"] = TablaPrin.Rows[i]["BP1"];
                        dr["BQ1"] = TablaPrin.Rows[i]["BQ1"];
                        dr["BR1"] = TablaPrin.Rows[i]["BR1"];
                        dr["BS1"] = TablaPrin.Rows[i]["BS1"];
                        dr["BT1"] = TablaPrin.Rows[i]["BT1"];
                        dr["BU1"] = TablaPrin.Rows[i]["BU1"];
                        dr["BV1"] = TablaPrin.Rows[i]["BV1"];
                        dr["BW1"] = TablaPrin.Rows[i]["BW1"];
                        dr["BX1"] = TablaPrin.Rows[i]["BX1"];
                        dr["BY1"] = TablaPrin.Rows[i]["BY1"];
                        dr["BZ1"] = TablaPrin.Rows[i]["BZ1"];


                        dr["CA1"] = TablaPrin.Rows[i]["CA1"];
                        dr["CB1"] = TablaPrin.Rows[i]["CB1"];
                        dr["CC1"] = TablaPrin.Rows[i]["CC1"];
                        dr["CD1"] = TablaPrin.Rows[i]["CD1"];
                        dr["CE1"] = TablaPrin.Rows[i]["CE1"];
                        dr["CF1"] = TablaPrin.Rows[i]["CF1"];
                        dr["CG1"] = TablaPrin.Rows[i]["CG1"];
                        dr["CH1"] = TablaPrin.Rows[i]["CH1"];
                        dr["CI1"] = TablaPrin.Rows[i]["CI1"];
                        dr["CJ1"] = TablaPrin.Rows[i]["CJ1"];
                        dr["CK1"] = TablaPrin.Rows[i]["CK1"];
                        dr["CL1"] = TablaPrin.Rows[i]["CL1"];
                        dr["CM1"] = TablaPrin.Rows[i]["CM1"];
                        dr["CN1"] = TablaPrin.Rows[i]["CN1"];
                        dr["CO1"] = TablaPrin.Rows[i]["CO1"];
                        dr["CP1"] = TablaPrin.Rows[i]["CP1"];
                        dr["CQ1"] = TablaPrin.Rows[i]["CQ1"];
                        dr["CR1"] = TablaPrin.Rows[i]["CR1"];
                        dr["CS1"] = TablaPrin.Rows[i]["CS1"];
                        dr["CT1"] = TablaPrin.Rows[i]["CT1"];
                        dr["CU1"] = TablaPrin.Rows[i]["CU1"];
                        dr["CV1"] = TablaPrin.Rows[i]["CV1"];
                        dr["CW1"] = TablaPrin.Rows[i]["CW1"];
                        dr["CX1"] = TablaPrin.Rows[i]["CX1"];
                        dr["CY1"] = TablaPrin.Rows[i]["CY1"];
                        dr["CZ1"] = TablaPrin.Rows[i]["CZ1"];

                        dr["DA1"] = TablaPrin.Rows[i]["DA1"];
                        dr["DB1"] = TablaPrin.Rows[i]["DB1"];
                        dr["DC1"] = TablaPrin.Rows[i]["DC1"];
                        dr["DD1"] = TablaPrin.Rows[i]["DD1"];
                        dr["DE1"] = TablaPrin.Rows[i]["DE1"];
                        dr["DF1"] = TablaPrin.Rows[i]["DF1"];
                        dr["DG1"] = TablaPrin.Rows[i]["DG1"];
                        dr["DH1"] = TablaPrin.Rows[i]["DH1"];
                        dr["DI1"] = TablaPrin.Rows[i]["DI1"];
                        dr["DJ1"] = TablaPrin.Rows[i]["DJ1"];
                        dr["DK1"] = TablaPrin.Rows[i]["DK1"];
                        dr["DL1"] = TablaPrin.Rows[i]["DL1"];
                        dr["DM1"] = TablaPrin.Rows[i]["DM1"];
                        dr["DN1"] = TablaPrin.Rows[i]["DN1"];
                        dr["DO1"] = TablaPrin.Rows[i]["DO1"];
                        dr["DP1"] = TablaPrin.Rows[i]["DP1"];
                        dr["DQ1"] = TablaPrin.Rows[i]["DQ1"];
                        dr["DR1"] = TablaPrin.Rows[i]["DR1"];
                        dr["DS1"] = TablaPrin.Rows[i]["DS1"];
                        dr["DT1"] = TablaPrin.Rows[i]["DT1"];
                        dr["DU1"] = TablaPrin.Rows[i]["DU1"];
                        dr["DV1"] = TablaPrin.Rows[i]["DV1"];
                        dr["DW1"] = TablaPrin.Rows[i]["DW1"];
                        dr["DX1"] = TablaPrin.Rows[i]["DX1"];
                        dr["DY1"] = TablaPrin.Rows[i]["DY1"];
                        dr["DZ1"] = TablaPrin.Rows[i]["DZ1"];


                        dr["EA1"] = TablaPrin.Rows[i]["EA1"];
                        dr["EB1"] = TablaPrin.Rows[i]["EB1"];
                        dr["EC1"] = TablaPrin.Rows[i]["EC1"];
                        dr["ED1"] = TablaPrin.Rows[i]["ED1"];
                        dr["EE1"] = TablaPrin.Rows[i]["EE1"];
                        dr["EF1"] = TablaPrin.Rows[i]["EF1"];
                        dr["EG1"] = TablaPrin.Rows[i]["EG1"];
                        dr["EH1"] = TablaPrin.Rows[i]["EH1"];
                        dr["EI1"] = TablaPrin.Rows[i]["EI1"];
                        dr["EJ1"] = TablaPrin.Rows[i]["EJ1"];
                        dr["EK1"] = TablaPrin.Rows[i]["EK1"];
                        dr["EL1"] = TablaPrin.Rows[i]["EL1"];
                        dr["EM1"] = TablaPrin.Rows[i]["EM1"];
                        dr["EN1"] = TablaPrin.Rows[i]["EN1"];
                        dr["EO1"] = TablaPrin.Rows[i]["EO1"];
                        dr["EP1"] = TablaPrin.Rows[i]["EP1"];
                        dr["EQ1"] = TablaPrin.Rows[i]["EQ1"];
                        dr["ER1"] = TablaPrin.Rows[i]["ER1"];
                        dr["ES1"] = TablaPrin.Rows[i]["ES1"];
                        dr["ET1"] = TablaPrin.Rows[i]["ET1"];
                        dr["EU1"] = TablaPrin.Rows[i]["EU1"];
                        dr["EV1"] = TablaPrin.Rows[i]["EV1"];
                        dr["EW1"] = TablaPrin.Rows[i]["EW1"];
                        dr["EX1"] = TablaPrin.Rows[i]["EX1"];
                        dr["EY1"] = TablaPrin.Rows[i]["EY1"];
                        dr["EZ1"] = TablaPrin.Rows[i]["EZ1"];



                        dr["FA1"] = TablaPrin.Rows[i]["FA1"];
                        dr["FB1"] = TablaPrin.Rows[i]["FB1"];
                        dr["FC1"] = TablaPrin.Rows[i]["FC1"];
                        dr["FD1"] = TablaPrin.Rows[i]["FD1"];
                        dr["FE1"] = TablaPrin.Rows[i]["FE1"];
                        dr["FF1"] = TablaPrin.Rows[i]["FF1"];
                        dr["FG1"] = TablaPrin.Rows[i]["FG1"];
                        dr["FH1"] = TablaPrin.Rows[i]["FH1"];
                        dr["FI1"] = TablaPrin.Rows[i]["FI1"];
                        dr["FJ1"] = TablaPrin.Rows[i]["FJ1"];
                        dr["FK1"] = TablaPrin.Rows[i]["FK1"];
                        dr["FL1"] = TablaPrin.Rows[i]["FL1"];
                        dr["FM1"] = TablaPrin.Rows[i]["FM1"];
                        dr["FN1"] = TablaPrin.Rows[i]["FN1"];
                        dr["FO1"] = TablaPrin.Rows[i]["FO1"];
                        dr["FP1"] = TablaPrin.Rows[i]["FP1"];
                        dr["FQ1"] = TablaPrin.Rows[i]["FQ1"];
                        dr["FR1"] = TablaPrin.Rows[i]["FR1"];
                        dr["FS1"] = TablaPrin.Rows[i]["FS1"];
                        dr["FT1"] = TablaPrin.Rows[i]["FT1"];
                        dr["FU1"] = TablaPrin.Rows[i]["FU1"];
                        dr["FV1"] = TablaPrin.Rows[i]["FV1"];
                        dr["FW1"] = TablaPrin.Rows[i]["FW1"];
                        dr["FX1"] = TablaPrin.Rows[i]["FX1"];
                        dr["FY1"] = TablaPrin.Rows[i]["FY1"];
                        dr["FZ1"] = TablaPrin.Rows[i]["FZ1"];


                        dr["GA1"] = TablaPrin.Rows[i]["GA1"];
                        dr["GB1"] = TablaPrin.Rows[i]["GB1"];
                        dr["GC1"] = TablaPrin.Rows[i]["GC1"];
                        dr["GD1"] = TablaPrin.Rows[i]["GD1"];
                        dr["GE1"] = TablaPrin.Rows[i]["GE1"];
                        dr["GF1"] = TablaPrin.Rows[i]["GF1"];
                        dr["GG1"] = TablaPrin.Rows[i]["GG1"];
                        dr["GH1"] = TablaPrin.Rows[i]["GH1"];
                        dr["GI1"] = TablaPrin.Rows[i]["GI1"];
                        dr["GJ1"] = TablaPrin.Rows[i]["GJ1"];
                        dr["GK1"] = TablaPrin.Rows[i]["GK1"];
                        dr["GL1"] = TablaPrin.Rows[i]["GL1"];
                        dr["GM1"] = TablaPrin.Rows[i]["GM1"];
                        dr["GN1"] = TablaPrin.Rows[i]["GN1"];
                        dr["GO1"] = TablaPrin.Rows[i]["GO1"];
                        dr["GP1"] = TablaPrin.Rows[i]["GP1"];
                        dr["GQ1"] = TablaPrin.Rows[i]["GQ1"];
                        dr["GR1"] = TablaPrin.Rows[i]["GR1"];
                        dr["GS1"] = TablaPrin.Rows[i]["GS1"];
                        dr["GT1"] = TablaPrin.Rows[i]["GT1"];
                        dr["GU1"] = TablaPrin.Rows[i]["GU1"];
                        dr["GV1"] = TablaPrin.Rows[i]["GV1"];
                        dr["GW1"] = TablaPrin.Rows[i]["GW1"];
                        dr["GX1"] = TablaPrin.Rows[i]["GX1"];
                        dr["GY1"] = TablaPrin.Rows[i]["GY1"];
                        dr["GZ1"] = TablaPrin.Rows[i]["GZ1"];


                        dr["HA1"] = TablaPrin.Rows[i]["HA1"];
                        dr["HB1"] = TablaPrin.Rows[i]["HB1"];
                        dr["HC1"] = TablaPrin.Rows[i]["HC1"];
                        dr["HD1"] = TablaPrin.Rows[i]["HD1"];
                        dr["HE1"] = TablaPrin.Rows[i]["HE1"];
                        dr["HF1"] = TablaPrin.Rows[i]["HF1"];
                        dr["HG1"] = TablaPrin.Rows[i]["HG1"];
                        dr["HH1"] = TablaPrin.Rows[i]["HH1"];
                        dr["HI1"] = TablaPrin.Rows[i]["HI1"];
                        dr["HJ1"] = TablaPrin.Rows[i]["HJ1"];
                        dr["HK1"] = TablaPrin.Rows[i]["HK1"];
                        dr["HL1"] = TablaPrin.Rows[i]["HL1"];
                        dr["HM1"] = TablaPrin.Rows[i]["HM1"];
                        dr["HN1"] = TablaPrin.Rows[i]["HN1"];
                        dr["HO1"] = TablaPrin.Rows[i]["HO1"];
                        dr["HP1"] = TablaPrin.Rows[i]["HP1"];
                        dr["HQ1"] = TablaPrin.Rows[i]["HQ1"];
                        dr["HR1"] = TablaPrin.Rows[i]["HR1"];
                        dr["HS1"] = TablaPrin.Rows[i]["HS1"];
                        dr["HT1"] = TablaPrin.Rows[i]["HT1"];
                        dr["HU1"] = TablaPrin.Rows[i]["HU1"];
                        dr["HV1"] = TablaPrin.Rows[i]["HV1"];
                        dr["HW1"] = TablaPrin.Rows[i]["HW1"];
                        dr["HX1"] = TablaPrin.Rows[i]["HX1"];
                        dr["HY1"] = TablaPrin.Rows[i]["HY1"];
                        dr["HZ1"] = TablaPrin.Rows[i]["HZ1"];


                        dr["IA1"] = TablaPrin.Rows[i]["IA1"];
                        dr["IB1"] = TablaPrin.Rows[i]["IB1"];
                        dr["IC1"] = TablaPrin.Rows[i]["IC1"];
                        dr["ID1"] = TablaPrin.Rows[i]["ID1"];
                        dr["IE1"] = TablaPrin.Rows[i]["IE1"];
                        dr["IF1"] = TablaPrin.Rows[i]["IF1"];
                        dr["IG1"] = TablaPrin.Rows[i]["IG1"];
                        dr["IH1"] = TablaPrin.Rows[i]["IH1"];
                        dr["II1"] = TablaPrin.Rows[i]["II1"];
                        dr["IJ1"] = TablaPrin.Rows[i]["IJ1"];
                        dr["IK1"] = TablaPrin.Rows[i]["IK1"];
                        dr["IL1"] = TablaPrin.Rows[i]["IL1"];
                        dr["IM1"] = TablaPrin.Rows[i]["IM1"];
                        dr["IN1"] = TablaPrin.Rows[i]["IN1"];
                        dr["IO1"] = TablaPrin.Rows[i]["IO1"];
                        dr["IP1"] = TablaPrin.Rows[i]["IP1"];
                        dr["IQ1"] = TablaPrin.Rows[i]["IQ1"];
                        dr["IR1"] = TablaPrin.Rows[i]["IR1"];
                        dr["IS1"] = TablaPrin.Rows[i]["IS1"];
                        dr["IT1"] = TablaPrin.Rows[i]["IT1"];
                        dr["IU1"] = TablaPrin.Rows[i]["IU1"];
                        dr["IV1"] = TablaPrin.Rows[i]["IV1"];
                        dr["IW1"] = TablaPrin.Rows[i]["IW1"];
                        dr["IX1"] = TablaPrin.Rows[i]["IX1"];
                        dr["IY1"] = TablaPrin.Rows[i]["IY1"];
                        dr["IZ1"] = TablaPrin.Rows[i]["IZ1"];
                        tablaCos.Rows.Add(dr);
                    }
                    #endregion
                }
                catch(Exception ex){}
                }
            string sqlInsert="INSERT INTO estadisticas_desincritos (id_mensaje,cont ,total)  VALUES ";
            sqlInsert += " (" + id_mensaje + " ," + emails + "  ," + (emails ) + ")";
            actualiza.ejecutorBase(sqlInsert);         
            return tablaCos;
        }
        protected void traspaso5(string id_mensaje)
        {


   //string lect= "SELECT TOP 1 fecha  FROM envio_correo where id_mensaje="+id_mensaje;
   string lect = "SELECT TOP 1 fecha  FROM envio_correo where id_mensaje=" + id_mensaje+ " order by fecha";
            DataTable Tabla22 = ConexionCall.SqlDTable(lect);

            if (Tabla22.Rows.Count > 0)
            {
                lect = Tabla22.Rows[0]["fecha"].ToString();
            }



            ConexionCall actualiza = new ConexionCall();
           // actualiza.ejecutorBase("delete FROM Estadistica_Aperturas where id_mensaje=" + id_mensaje);

            //aca sacaria la fecha de la ultima estadistica?




            string sqlra = "select top 1 * FROM Estadistica_Aperturas where id_mensaje =" + id_mensaje + " order by anio desc, mes desc,dia desc, hora desc, minutos desc";
            DataTable TablaP22 = ConexionCall.SqlDTable(sqlra);

            string anio1 = string.Empty, mes1 = string.Empty, dia1 = string.Empty, hora1 = string.Empty, minutos1 = string.Empty, fecha1 = string.Empty;
            if (TablaP22.Rows.Count > 0)
            {
             
                anio1 = TablaP22.Rows[0]["anio"].ToString();
                mes1 = TablaP22.Rows[0]["mes"].ToString();
                dia1 = TablaP22.Rows[0]["dia"].ToString();
                hora1 = TablaP22.Rows[0]["hora"].ToString();
                minutos1 = TablaP22.Rows[0]["minutos"].ToString();
            }

            if(!string.IsNullOrEmpty(anio1)&&!string.IsNullOrEmpty(mes1)&&!string.IsNullOrEmpty(dia1)&&!string.IsNullOrEmpty(hora1)&&!string.IsNullOrEmpty(minutos1)){
            //'2015-03-17 15:11:45.390'
           // fecha1="'"+anio1+"-"+mes1+"-"+dia1+" "+hora1+":"+minutos1+":00.000"+"'";
                fecha1 = dia1 + "-" + mes1 + "-" + anio1 + " " + hora1 + ":" + minutos1 + ":00.000";
            }



            //string sqlApertura = "insert into Estadistica_Aperturas SELECT count(abierto) abierto ,sum(abierto) total ,year(fecha)anio,month(fecha)mes,day(fecha) dia ,datepart(hour,fecha)hora,datepart(minute,fecha)minutos,'" + id_mensaje + "'id_mensaje,0 ";
            //sqlApertura+=" FROM envio_correo where id_mensaje="+id_mensaje+" and abierto>0 ";
            //sqlApertura+=" and DATEDIFF(hour,(SELECT TOP 1 fecha  FROM envio_correo where id_mensaje="+id_mensaje+"),fecha) <121 ";
            //sqlApertura += "group by year(fecha),month(fecha),  day(fecha)  ,datepart(hour,fecha),datepart(minute,fecha) ";
            //sqlApertura+="order by anio,mes,dia ,hora,minutos";

            

     string sqlApertura = "insert into Estadistica_Aperturas SELECT count(abierto) abierto ,sum(abierto) total ,year(FechaLectura) anio,month(FechaLectura) mes,day(FechaLectura) dia ,datepart(hour,FechaLectura) hora,datepart(minute,FechaLectura) minutos,'" + id_mensaje + "' id_mensaje,";
            //aca sacar la diferencia de horas
     sqlApertura += " max(DATEDIFF(hour,('" + lect + "'),FechaLectura)) "; 
            
            
            sqlApertura+=" FROM envio_correo where id_mensaje="+id_mensaje+" and abierto>0 ";
            if (!string.IsNullOrEmpty(fecha1))
            {
            sqlApertura += " and FechaLectura >'" + fecha1 + "'";
            }
            sqlApertura+=" and DATEDIFF(hour,('"+lect+"'),FechaLectura) <121 ";
            sqlApertura += "group by year(FechaLectura),month(FechaLectura),  day(FechaLectura)  ,datepart(hour,FechaLectura),datepart(minute,FechaLectura) ";
            sqlApertura+="order by anio,mes,dia ,hora,minutos";




            //JM se puede mejoirar separando por horas reales (60 minutos corridos) AHora está por hora reloj


            
//SELECT count(abierto) abierto ,sum(abierto) total ,year(FechaLectura)anio,
//month(FechaLectura)mes,day(FechaLectura) dia ,datepart(hour,FechaLectura)hora,datepart(minute,FechaLectura)minutos,'4051'id_mensaje,0 
//  FROM envio_correo where id_mensaje=4051 and abierto>0 
//   and DATEDIFF(hour,(SELECT TOP 1 fecha  FROM envio_correo where id_mensaje=4051),FechaLectura) <121 
//   group by year(FechaLectura),month(FechaLectura),  day(FechaLectura)  ,datepart(hour,FechaLectura),datepart(minute,FechaLectura) 
//    order by anio,mes,dia ,hora,minutos










            actualiza.ejecutorBase(sqlApertura);
          
            //JM ya le agregué la hora en el insert anterior
            //DataTable TablaPrin = ConexionCall.SqlDTable("SELECT anio,mes,dia,hora,id_mensaje FROM Estadistica_Aperturas where id_mensaje="+id_mensaje+"  group by id_mensaje,anio,mes,dia,hora");
            //int emails = TablaPrin.Rows.Count;
           
            //if (TablaPrin.Rows.Count > 0)
            //{
            //    try
            //    {
            //        string anio, mes, dia, hora;
            //        for (int i = 0; i < TablaPrin.Rows.Count; i++)
            //        {
            //            anio = TablaPrin.Rows[i]["anio"].ToString();
            //            mes = TablaPrin.Rows[i]["mes"].ToString();
            //            dia = TablaPrin.Rows[i]["dia"].ToString();
            //            hora = TablaPrin.Rows[i]["hora"].ToString();

            //            string sqlHora = "update Estadistica_Aperturas set ContH=" + (i + 1);
            //            sqlHora += "where id_mensaje=" + id_mensaje + " and anio=" + anio + " and mes=" + mes + " and dia=" + dia + " and hora=" + hora;
            //            actualiza.ejecutorBase(sqlHora);
            //        }
            //    }catch( Exception es){}
            //}
        }
        //protected DataTable traspaso6(string id_mensaje)
        //{
        //    ConexionCall ojjURL = new ConexionCall();
        //    string sqlusuarios = "  SELECT en.id_email,a1 ";
        //    sqlusuarios += " FROM Encuesta_terminada en ";
        //    sqlusuarios += " join email em on em.id_email=en.id_email ";
        //    sqlusuarios += " where id_mensaje= " + id_mensaje;
        //    sqlusuarios += " group by en.id_email,a1 ";
        //    sqlusuarios += " order by a1 ";

        //    string sqlRespuesta = "";
        //    string sqlPreg = "SELECT id_preg ,pregunta ";
        //    sqlPreg += " FROM preguntas ";
        //    sqlPreg += " where id_encuesta=(SELECT id_encuesta";
        //    sqlPreg += " FROM Mensaje  where id_mensaje=" + id_mensaje + ")";
        //    sqlPreg += " order by id_preg";
        //    DataTable preguntas = ConexionCall.SqlDTable(sqlPreg);
        //    string[] idPreg = new string[preguntas.Rows.Count];
        //    string[] nombrePreg = new string[preguntas.Rows.Count];
        //    string id_encuesta = ConexionCall.devuelveValor("select id_encuesta from mensaje where id_mensaje="+id_mensaje);

        //    DataTable TablaPrin = ConexionCall.SqlDTable(sqlusuarios);
        //    ConexionCall objEst = new ConexionCall();
        //    DataTable tablaCos = new DataTable();
        //    DataRow dr = null;
        //    tablaCos.Columns.Clear();
        //    tablaCos.Rows.Clear();

        //    int hdr = 0;
        //    if (TablaPrin.Rows.Count > 0)
        //    {
        //        try
        //        {
        //            #region
        //            for (int i = 0; i < TablaPrin.Rows.Count; i++)
        //            {
        //                if (hdr == 0)
        //                {
        //                    tablaCos.Columns.Add(new DataColumn("ID Mensaje", typeof(string)));
        //                    tablaCos.Columns.Add(new DataColumn("ID Encuesta", typeof(string)));
        //                    tablaCos.Columns.Add(new DataColumn("Usuario", typeof(string)));
        //                    for (int j = 0; j < preguntas.Rows.Count; j++)
        //                    {
        //                        idPreg[j] = preguntas.Rows[j]["id_preg"].ToString();
        //                        nombrePreg[j] = preguntas.Rows[j]["pregunta"].ToString();
        //                        tablaCos.Columns.Add(new DataColumn(nombrePreg[j], typeof(string)));
        //                    }
        //                    hdr = 1;
        //                }

        //                dr = tablaCos.NewRow();
        //                dr["ID Mensaje"] = id_mensaje;
        //                dr["ID Encuesta"] = id_encuesta;
        //                dr["Usuario"] = TablaPrin.Rows[i]["a1"].ToString();
                   
        //                for (int j = 0; j < idPreg.Length; j++)
        //                {
        //                    sqlRespuesta = "SELECT TOP 1 respuesta  ";
        //                    sqlRespuesta += " FROM Encuesta_terminada en ";
        //                    sqlRespuesta += " join TBL_Respuestas re on re.id_resp=en.id_resp ";
        //                    sqlRespuesta += "   where id_mensaje=" + id_mensaje + " and en.id_email =" + TablaPrin.Rows[i]["id_email"].ToString() + " and en.id_preg= " + idPreg[j];

        //                    dr[nombrePreg[j]] = ConexionCall.devuelveValor(sqlRespuesta);
        //                }
        //                tablaCos.Rows.Add(dr);
        //            }
        //            #endregion
        //        }
        //        catch (Exception ex)
        //        { }
        //     }
        //    return tablaCos;
        //}


        //protected DataTable traspaso6(string id_mensaje)
        //{
        //    ConexionCall ojjURL = new ConexionCall();
        //    string sqlusuarios = "  SELECT en.id_email,a1,en.id_envio ";
        //    sqlusuarios += " FROM Encuesta_terminada en ";
        //    sqlusuarios += " join email em on em.id_email=en.id_email ";
        //    sqlusuarios += " where id_mensaje= " + id_mensaje;
        //    sqlusuarios += " group by en.id_email,a1, en.id_envio  ";
        //    sqlusuarios += " order by a1 ";


        //    string sqltexto = "";
        //    string sqlRespuesta = "";
        //    string sqlPreg = "SELECT id_preg ,pregunta ";
        //    sqlPreg += " FROM preguntas ";
        //    sqlPreg += " where id_encuesta=(SELECT id_encuesta";
        //    sqlPreg += " FROM Mensaje  where id_mensaje=" + id_mensaje + ")";
        //    sqlPreg += " order by orden";
        //    DataTable preguntas = ConexionCall.SqlDTable(sqlPreg);
        //    string[] idPreg = new string[preguntas.Rows.Count];
        //    string[] nombrePreg = new string[preguntas.Rows.Count];
        //    string id_encuesta = ConexionCall.devuelveValor("select id_encuesta from mensaje where id_mensaje=" + id_mensaje);

        //    DataTable TablaPrin = ConexionCall.SqlDTable(sqlusuarios);
        //    ConexionCall objEst = new ConexionCall();
        //    DataTable tablaCos = new DataTable();
        //    DataRow dr = null;
        //    tablaCos.Columns.Clear();
        //    tablaCos.Rows.Clear();

        //    int hdr = 0;
        //    if (TablaPrin.Rows.Count > 0)
        //    {
        //        try
        //        {
        //            #region
        //            for (int i = 0; i < TablaPrin.Rows.Count; i++)
        //            {
        //                if (hdr == 0)
        //                {
        //                    tablaCos.Columns.Add(new DataColumn("ID Mensaje", typeof(string)));
        //                    tablaCos.Columns.Add(new DataColumn("ID Envio", typeof(string)));
        //                    tablaCos.Columns.Add(new DataColumn("ID Encuesta", typeof(string)));
        //                    tablaCos.Columns.Add(new DataColumn("Usuario", typeof(string)));
        //                    for (int j = 0; j < preguntas.Rows.Count; j++)
        //                    {
        //                        idPreg[j] = preguntas.Rows[j]["id_preg"].ToString();
        //                        nombrePreg[j] = preguntas.Rows[j]["pregunta"].ToString();
        //                        tablaCos.Columns.Add(new DataColumn(nombrePreg[j], typeof(string)));
        //                    }
        //                    hdr = 1;
        //                }
                       
        //                dr = tablaCos.NewRow();
        //                dr["ID Mensaje"] = id_mensaje;
        //                dr["ID Envio"] = TablaPrin.Rows[i]["id_envio"].ToString();
        //                dr["ID Encuesta"] = id_encuesta;
        //                dr["Usuario"] = TablaPrin.Rows[i]["a1"].ToString();

        //                for (int j = 0; j < idPreg.Length; j++)
        //                {
        //                    sqlRespuesta = "SELECT TOP 1 respuesta  ";
        //                    sqlRespuesta += " FROM Encuesta_terminada en ";
        //                    sqlRespuesta += " join TBL_Respuestas re on re.id_resp=en.id_resp ";
        //                    sqlRespuesta += "   where id_mensaje=" + id_mensaje + " and en.id_email =" + TablaPrin.Rows[i]["id_email"].ToString() + " and en.id_preg= " + idPreg[j];
        //                    DataTable respu = ConexionCall.SqlDTable(sqlRespuesta);
        //                    //     for(int p =0; p < respu.Rows.Count;p++){
        //                    //     string id = respu.Rows[p]["id_resp"].ToString();
        //                    //     }

        //                    sqltexto = "select respuesta from Respuesta_texto ";
        //                    sqltexto += " where id_pregunta= " + idPreg[j] + " and id_mail = " + TablaPrin.Rows[i]["id_email"].ToString() + " and  id_mensaje=" + id_mensaje;


        //                    string ress = ConexionCall.devuelveValor(sqlRespuesta);
        //                    string respuestexto = ConexionCall.devuelveValor2(sqltexto);
                       
        //                    if (ress == "")
        //                    {
        //                        dr[nombrePreg[j]] = respuestexto;
        //                    }
        //                    else
        //                    {
        //                        dr[nombrePreg[j]] = ress;
        //                    }
        //                }
        //                tablaCos.Rows.Add(dr);
        //            }
        //            #endregion
        //        }
        //        catch (Exception ex)
        //        { }
        //    }
        //    return tablaCos;
        //}


        protected DataTable traspaso6(string id_mensaje)
        {
            ConexionCall ojjURL = new ConexionCall();
            string sqlusuarios = "  SELECT en.id_email,em.a1,em.b1,em.c1,em.d1,em.e1,em.f1,em.g1,em.h1,em.i1,em.j1,em.k1, en.id_envio ";
            sqlusuarios += " FROM Encuesta_terminada en ";
            sqlusuarios += " join email em on em.id_email=en.id_email ";
            sqlusuarios += " where id_mensaje= " + id_mensaje;
            sqlusuarios += " group by en.id_email,a1,b1,c1,d1,e1,f1,g1,h1,i1,j1,k1, en.id_envio  ";
            sqlusuarios += " order by a1 ";


            string sqltexto = "";
            string sqlRespuesta = "";
            string sqlPreg = "SELECT id_preg ,pregunta ";
            sqlPreg += " FROM preguntas ";
            sqlPreg += " where id_encuesta=(SELECT id_encuesta";
            sqlPreg += " FROM Mensaje  where id_mensaje=" + id_mensaje + ")";
            sqlPreg += " order by id_preg";
            DataTable preguntas = ConexionCall.SqlDTable(sqlPreg);
            string[] idPreg = new string[preguntas.Rows.Count];
            string[] nombrePreg = new string[preguntas.Rows.Count];
            string id_encuesta = ConexionCall.devuelveValor("select id_encuesta from mensaje where id_mensaje=" + id_mensaje);

            DataTable TablaPrin = ConexionCall.SqlDTable(sqlusuarios);
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();

            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {
                try
                {
                    #region
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID Mensaje", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ID Envio", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ID Encuesta", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Usuario", typeof(string)));

                            for (int j = 0; j < preguntas.Rows.Count; j++)
                            {
                                idPreg[j] = preguntas.Rows[j]["id_preg"].ToString();
                                nombrePreg[j] = preguntas.Rows[j]["pregunta"].ToString();
                                tablaCos.Columns.Add(new DataColumn(nombrePreg[j], typeof(string)));
                            }
                            tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));


                            hdr = 1;
                        }

                        dr = tablaCos.NewRow();
                        dr["ID Mensaje"] = id_mensaje;
                        dr["ID Envio"] = TablaPrin.Rows[i]["id_envio"].ToString();
                        dr["ID Encuesta"] = id_encuesta;
                        dr["Usuario"] = TablaPrin.Rows[i]["a1"].ToString();

                        for (int j = 0; j < idPreg.Length; j++)
                        {
                            sqlRespuesta = "SELECT TOP 1 respuesta  ";
                            sqlRespuesta += " FROM Encuesta_terminada en ";
                            sqlRespuesta += " join TBL_Respuestas re on re.id_resp=en.id_resp ";
                            sqlRespuesta += "   where id_mensaje=" + id_mensaje + " and en.id_email =" + TablaPrin.Rows[i]["id_email"].ToString() + " and en.id_preg= " + idPreg[j];
                            DataTable respu = ConexionCall.SqlDTable(sqlRespuesta);
                            //     for(int p =0; p < respu.Rows.Count;p++){
                            //     string id = respu.Rows[p]["id_resp"].ToString();
                            //     }

                            sqltexto = "select respuesta from Respuesta_texto ";
                            sqltexto += " where id_pregunta= " + idPreg[j] + " and id_mail = " + TablaPrin.Rows[i]["id_email"].ToString() + " and  id_mensaje=" + id_mensaje;


                            string ress = ConexionCall.devuelveValor(sqlRespuesta);
                            string respuestexto = ConexionCall.devuelveValor(sqltexto);


                            if (ress == "")
                            {
                                //dr[nombrePreg[j]] = respuestexto;

                                //JM agregada funcion arreglacaracter
                                dr[nombrePreg[j]] = arreglacaracter( respuestexto);


                            }
                            else
                            {
                                dr[nombrePreg[j]] = ress;
                            }
                        }
                        dr["B1"] = TablaPrin.Rows[i]["b1"].ToString();
                        dr["C1"] = TablaPrin.Rows[i]["c1"].ToString();
                        dr["D1"] = TablaPrin.Rows[i]["d1"].ToString();
                        dr["E1"] = TablaPrin.Rows[i]["e1"].ToString();
                        dr["F1"] = TablaPrin.Rows[i]["f1"].ToString();
                        dr["G1"] = TablaPrin.Rows[i]["g1"].ToString();
                        dr["H1"] = TablaPrin.Rows[i]["h1"].ToString();
                        dr["I1"] = TablaPrin.Rows[i]["i1"].ToString();
                        dr["J1"] = TablaPrin.Rows[i]["j1"].ToString();
                        dr["K1"] = TablaPrin.Rows[i]["k1"].ToString();
                       
                        tablaCos.Rows.Add(dr);

                    }
                    #endregion
                }
                catch (Exception ex)
                { }
            }
            return tablaCos;
        }
        protected DataTable traspaso7(string id_mensaje)
        {
            ConexionCall ojjURL = new ConexionCall();
            #region consultas viejas
            //string sqlEm = "SELECT * FROM Confirmacion_evento  where id_mensaje="+id_mensaje+" order by nombre ";      
            //string sqlEm = "select c.nombre,c.fecha, em.a1 FROM Confirmacion_evento c inner join   envio_correo e  on e.id_enviado=c.id_enviado inner join   email em on e.id_email=em.id_email where c.id_mensaje="+id_mensaje+" order by nombre ";
//string sqlEm = " select c.nombre,c.fecha, em.a1 ,e.id_enviado, IIF (c.fecha is null, 'FALSE', 'TRUE')  AS confirmado FROM  envio_correo e inner join  email em on  ";
// sqlEm += " e.id_email=em.id_email left join  Confirmacion_evento c  on e.id_enviado=c.id_enviado  where  e.id_mensaje="+id_mensaje+"  order by c.nombre ";

////para la bd sql srvr 2000 no funciona la consulta anterior, para PRUEBAS usamos la siguiente:
            //string sqlEm = " select c.*, em.a1 ,e.id_enviado,  'TRUE'  AS confirmado FROM  envio_correo e inner join  email em on  ";
            //sqlEm += " e.id_email=em.id_email left join  Confirmacion_evento c  on e.id_enviado=c.id_enviado  where  e.id_mensaje=" + id_mensaje + "  order by c.nombre ";


              #endregion
            //esta consulta funciona bien en sql server 2012
           // string sqlEm = " select c.*, em.a1,e.id_enviado, IIF (c.fecha is null, 'FALSE', 'TRUE')  AS confirmado FROM  envio_correo e inner join  email em on  ";

            //JM modificado para leer todos los campos de tabla email

            string sqlEm = " select c.*, em.* ,e.id_enviado, IIF (c.fecha is null, 'FALSE', 'TRUE')  AS confirmado FROM  envio_correo e inner join  email em on  ";
           
            sqlEm += " e.id_email=em.id_email left join  Confirmacion_evento c  on e.id_enviado=c.id_enviado  where  e.id_mensaje=" + id_mensaje + "  order by c.nombre ";

            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();

            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {

                for (int i = 0; i < TablaPrin.Rows.Count; i++)
                {
                    if (hdr == 0)
                    {
                        tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("id_enviado", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Email_evento", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Email_evento2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Fecha_Confirmacion", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Confirmado", typeof(string)));                        //////
                        tablaCos.Columns.Add(new DataColumn("Nombre", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("apellidop", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("apellidom", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("edad", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sexo", typeof(string)));                      
                        tablaCos.Columns.Add(new DataColumn("RUT", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("telefono", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("celular", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("telefono2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("direccion", typeof(string)));                        
                        tablaCos.Columns.Add(new DataColumn("region", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("provincia", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("comuna", typeof(string)));                        
                        tablaCos.Columns.Add(new DataColumn("hora1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("hora2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("hora3", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("hora4", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("fecha1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("fecha2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("fechaano1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("fechaano2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Obs", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("observacion2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("acompanante1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("acompanante2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto3", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto4", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto5", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto6", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto7", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("texto8", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sololectura1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sololectura2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sololectura3", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("disabled1", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("disabled2", typeof(string)));
                        //tablaCos.Columns.Add(new DataColumn("disabled3", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("codarea1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("codarea2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sino", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sino1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sino2", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sino3", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("sino4", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("numerico1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("carrera", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("industria", typeof(string)));

                        //campos de tabla EMAIl
                        tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("L1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("M1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("N1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("O1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("P1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Q1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("R1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("S1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("T1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("U1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("V1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("W1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("X1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Y1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Z1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("AA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AZ1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));


                        tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));


                        tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));




                        hdr = 1;
                    }
                    //ACÁ cuidar de que no exista ningún campo que no corresponda a la tabla Confirmacion_evento
                    dr = tablaCos.NewRow();
                    dr["ID_Mensaje"] = id_mensaje;
                    dr["id_enviado"] = TablaPrin.Rows[i]["id_enviado"];
                    dr["Email"] = TablaPrin.Rows[i]["a1"];
                    dr["Email_evento"] = TablaPrin.Rows[i]["email"];
                    dr["Email_evento2"] = TablaPrin.Rows[i]["email2"];
                    dr["Fecha_Confirmacion"] = TablaPrin.Rows[i]["fecha"];
                    dr["Confirmado"] = TablaPrin.Rows[i]["confirmado"];
                    dr["Nombre"] = TablaPrin.Rows[i]["nombre"].ToString() ?? "";
                    dr["Apellidop"] = TablaPrin.Rows[i]["apellido"].ToString() ?? "";
                    dr["Apellidom"] = TablaPrin.Rows[i]["apellidom"].ToString() ?? "";
                    dr["RUT"] = TablaPrin.Rows[i]["RUT"] ?? "";
                    dr["telefono"] = TablaPrin.Rows[i]["telefono"] ?? "";
                    dr["celular"] = TablaPrin.Rows[i]["celular"] ?? "";
                    dr["telefono2"] = TablaPrin.Rows[i]["telefono2"] ?? "";
                    dr["region"] = TablaPrin.Rows[i]["region"].ToString() ?? "";
                    dr["provincia"] = TablaPrin.Rows[i]["provincia"].ToString() ?? "";
                    dr["comuna"] = TablaPrin.Rows[i]["comuna"].ToString() ?? "";
                    dr["direccion"] = TablaPrin.Rows[i]["direccion"].ToString() ?? "";
                    dr["sexo"] = TablaPrin.Rows[i]["sexo"] ?? "";
                    dr["edad"] = TablaPrin.Rows[i]["edad"] ?? "";
                    dr["hora1"] = TablaPrin.Rows[i]["hora1"] ?? "";
                    dr["hora2"] = TablaPrin.Rows[i]["hora2"] ?? "";
                    dr["hora3"] = TablaPrin.Rows[i]["hora3"] ?? "";
                    dr["hora4"] = TablaPrin.Rows[i]["hora4"] ?? "";
                    dr["Obs"] = TablaPrin.Rows[i]["Obs"].ToString() ?? "";
                    dr["observacion2"] = TablaPrin.Rows[i]["observacion2"].ToString() ?? "";
                    dr["fecha1"] = TablaPrin.Rows[i]["fecha1"] ?? "";
                    dr["fecha2"] = TablaPrin.Rows[i]["fecha2"] ?? "";
                    dr["fechaano1"] = TablaPrin.Rows[i]["fechaano1"] ?? "";
                    dr["fechaano2"] = TablaPrin.Rows[i]["fechaano2"] ?? "";
                    dr["acompanante1"] = TablaPrin.Rows[i]["acom1"].ToString() ?? "";
                    dr["acompanante2"] = TablaPrin.Rows[i]["acom2"].ToString() ?? "";
                    dr["texto1"] = TablaPrin.Rows[i]["texto1"].ToString() ?? "";
                    dr["texto2"] = TablaPrin.Rows[i]["texto2"].ToString() ?? "";
                    dr["texto3"] = TablaPrin.Rows[i]["texto3"].ToString() ?? "";
                    dr["texto4"] = TablaPrin.Rows[i]["texto4"].ToString() ?? "";
                    dr["texto5"] = TablaPrin.Rows[i]["texto5"].ToString() ?? "";
                    dr["texto6"] = TablaPrin.Rows[i]["texto6"].ToString() ?? "";
                    dr["texto7"] = TablaPrin.Rows[i]["texto7"].ToString() ?? "";
                    dr["texto8"] = TablaPrin.Rows[i]["texto8"].ToString() ?? "";      
                    dr["sololectura1"] = TablaPrin.Rows[i]["sololectura1"].ToString() ?? "";
                    dr["sololectura2"] = TablaPrin.Rows[i]["sololectura2"].ToString() ?? "";
                    dr["sololectura3"] = TablaPrin.Rows[i]["sololectura3"].ToString() ?? "";
                    //dr["disabled1"] = TablaPrin.Rows[i]["disabled1"].ToString() ?? "";
                    //dr["disabled2"] = TablaPrin.Rows[i]["disabled2"].ToString() ?? "";
                    //dr["disabled3"] = TablaPrin.Rows[i]["disabled3"].ToString() ?? "";
                    dr["codarea1"] = TablaPrin.Rows[i]["codarea1"].ToString() ?? "";
                    dr["codarea2"] = TablaPrin.Rows[i]["codarea2"].ToString() ?? "";
                    dr["sino"] = TablaPrin.Rows[i]["sino"].ToString() ?? "";
                    dr["sino1"] = TablaPrin.Rows[i]["sino1"].ToString() ?? "";
                    dr["sino2"] = TablaPrin.Rows[i]["sino2"].ToString() ?? "";
                    dr["sino3"] = TablaPrin.Rows[i]["sino3"].ToString() ?? "";
                    dr["sino4"] = TablaPrin.Rows[i]["sino4"].ToString() ?? "";
                    dr["numerico1"] = TablaPrin.Rows[i]["numerico1"].ToString() ?? "";
                    dr["carrera"] = TablaPrin.Rows[i]["carrera"].ToString() ?? "";
                    dr["industria"] = TablaPrin.Rows[i]["industria"].ToString() ?? "";
                    /////













                    dr["B1"] = TablaPrin.Rows[i]["b1"];
                    dr["C1"] = TablaPrin.Rows[i]["c1"];
                    dr["D1"] = TablaPrin.Rows[i]["d1"]; 
                    dr["E1"] = TablaPrin.Rows[i]["e1"]; 
                    dr["F1"] = TablaPrin.Rows[i]["f1"];
                    dr["G1"] = TablaPrin.Rows[i]["g1"];
                    dr["H1"] = TablaPrin.Rows[i]["h1"];
                    dr["I1"] = TablaPrin.Rows[i]["i1"];
                    dr["J1"] = TablaPrin.Rows[i]["j1"];
                    dr["K1"] = TablaPrin.Rows[i]["k1"];
                    dr["L1"] = TablaPrin.Rows[i]["l1"];
                    dr["M1"] = TablaPrin.Rows[i]["m1"];
                    dr["N1"] = TablaPrin.Rows[i]["n1"];
                    dr["O1"] = TablaPrin.Rows[i]["o1"];
                    dr["P1"] = TablaPrin.Rows[i]["p1"];
                    dr["Q1"] = TablaPrin.Rows[i]["q1"];
                    dr["R1"] = TablaPrin.Rows[i]["r1"];
                    dr["S1"] = TablaPrin.Rows[i]["s1"];
                    dr["T1"] = TablaPrin.Rows[i]["t1"];
                    dr["U1"] = TablaPrin.Rows[i]["u1"];
                    dr["V1"] = TablaPrin.Rows[i]["v1"];
                    dr["W1"] = TablaPrin.Rows[i]["w1"];
                    dr["X1"] = TablaPrin.Rows[i]["x1"];
                    dr["Y1"] = TablaPrin.Rows[i]["y1"];
                    dr["Z1"] = TablaPrin.Rows[i]["z1"];

                    dr["AA1"] = TablaPrin.Rows[i]["aa1"];
                    dr["AB1"] = TablaPrin.Rows[i]["ab1"];
                    dr["AC1"] = TablaPrin.Rows[i]["ac1"];
                    dr["AD1"] = TablaPrin.Rows[i]["ad1"];
                    dr["AE1"] = TablaPrin.Rows[i]["ae1"];
                    dr["AF1"] = TablaPrin.Rows[i]["af1"];
                    dr["AG1"] = TablaPrin.Rows[i]["ag1"];
                    dr["AH1"] = TablaPrin.Rows[i]["ah1"];
                    dr["AI1"] = TablaPrin.Rows[i]["ai1"];
                    dr["AJ1"] = TablaPrin.Rows[i]["aj1"];
                    dr["AK1"] = TablaPrin.Rows[i]["ak1"];
                    dr["AL1"] = TablaPrin.Rows[i]["al1"];
                    dr["AM1"] = TablaPrin.Rows[i]["am1"];
                    dr["AN1"] = TablaPrin.Rows[i]["an1"];
                    dr["AO1"] = TablaPrin.Rows[i]["ao1"];
                    dr["AP1"] = TablaPrin.Rows[i]["ap1"];
                    dr["AQ1"] = TablaPrin.Rows[i]["aq1"];
                    dr["AR1"] = TablaPrin.Rows[i]["ar1"];
                    dr["AS1"] = TablaPrin.Rows[i]["as1"];
                    dr["AT1"] = TablaPrin.Rows[i]["at1"];
                    dr["AU1"] = TablaPrin.Rows[i]["au1"];
                    dr["AV1"] = TablaPrin.Rows[i]["av1"];
                    dr["AW1"] = TablaPrin.Rows[i]["aw1"];
                    dr["AX1"] = TablaPrin.Rows[i]["ax1"];
                    dr["AY1"] = TablaPrin.Rows[i]["ay1"];
                    dr["AZ1"] = TablaPrin.Rows[i]["az1"];

                    dr["BA1"] = TablaPrin.Rows[i]["ba1"];
                    dr["BB1"] = TablaPrin.Rows[i]["bb1"];
                    dr["BC1"] = TablaPrin.Rows[i]["bc1"];
                    dr["BD1"] = TablaPrin.Rows[i]["bd1"];
                    dr["BE1"] = TablaPrin.Rows[i]["be1"];
                    dr["BF1"] = TablaPrin.Rows[i]["bf1"];
                    dr["BG1"] = TablaPrin.Rows[i]["bg1"];
                    dr["BH1"] = TablaPrin.Rows[i]["bh1"];
                    dr["BI1"] = TablaPrin.Rows[i]["bi1"];
                    dr["BJ1"] = TablaPrin.Rows[i]["bj1"];
                    dr["BK1"] = TablaPrin.Rows[i]["bk1"];
                    dr["BL1"] = TablaPrin.Rows[i]["bl1"];
                    dr["BM1"] = TablaPrin.Rows[i]["bm1"];
                    dr["BN1"] = TablaPrin.Rows[i]["bn1"];
                    dr["BO1"] = TablaPrin.Rows[i]["bo1"];
                    dr["BP1"] = TablaPrin.Rows[i]["bp1"];
                    dr["BQ1"] = TablaPrin.Rows[i]["bq1"];
                    dr["BR1"] = TablaPrin.Rows[i]["br1"];
                    dr["BS1"] = TablaPrin.Rows[i]["bs1"];
                    dr["BT1"] = TablaPrin.Rows[i]["bt1"];
                    dr["BU1"] = TablaPrin.Rows[i]["bu1"];
                    dr["BV1"] = TablaPrin.Rows[i]["bv1"];
                    dr["BW1"] = TablaPrin.Rows[i]["bw1"];
                    dr["BX1"] = TablaPrin.Rows[i]["bx1"];
                    dr["BY1"] = TablaPrin.Rows[i]["by1"];
                    dr["BZ1"] = TablaPrin.Rows[i]["bz1"];


                    dr["CA1"] = TablaPrin.Rows[i]["ca1"];
                    dr["CB1"] = TablaPrin.Rows[i]["cb1"];
                    dr["CC1"] = TablaPrin.Rows[i]["cc1"];
                    dr["CD1"] = TablaPrin.Rows[i]["cd1"];
                    dr["CE1"] = TablaPrin.Rows[i]["ce1"];
                    dr["CF1"] = TablaPrin.Rows[i]["cf1"];
                    dr["CG1"] = TablaPrin.Rows[i]["cg1"];
                    dr["CH1"] = TablaPrin.Rows[i]["ch1"];
                    dr["CI1"] = TablaPrin.Rows[i]["ci1"];
                    dr["CJ1"] = TablaPrin.Rows[i]["cj1"];
                    dr["CK1"] = TablaPrin.Rows[i]["ck1"];
                    dr["CL1"] = TablaPrin.Rows[i]["cl1"];
                    dr["CM1"] = TablaPrin.Rows[i]["cm1"];
                    dr["CN1"] = TablaPrin.Rows[i]["cn1"];
                    dr["CO1"] = TablaPrin.Rows[i]["co1"];
                    dr["CP1"] = TablaPrin.Rows[i]["cp1"];
                    dr["CQ1"] = TablaPrin.Rows[i]["cq1"];
                    dr["CR1"] = TablaPrin.Rows[i]["cr1"];
                    dr["CS1"] = TablaPrin.Rows[i]["cs1"];
                    dr["CT1"] = TablaPrin.Rows[i]["ct1"];
                    dr["CU1"] = TablaPrin.Rows[i]["cu1"];
                    dr["CV1"] = TablaPrin.Rows[i]["cv1"];
                    dr["CW1"] = TablaPrin.Rows[i]["cw1"];
                    dr["CX1"] = TablaPrin.Rows[i]["cx1"];
                    dr["CY1"] = TablaPrin.Rows[i]["cy1"];
                    dr["CZ1"] = TablaPrin.Rows[i]["cz1"];


                    dr["DA1"] = TablaPrin.Rows[i]["da1"];
                    dr["DB1"] = TablaPrin.Rows[i]["db1"];
                    dr["DC1"] = TablaPrin.Rows[i]["dc1"];
                    dr["DD1"] = TablaPrin.Rows[i]["dd1"];
                    dr["DE1"] = TablaPrin.Rows[i]["de1"];
                    dr["DF1"] = TablaPrin.Rows[i]["df1"];
                    dr["DG1"] = TablaPrin.Rows[i]["dg1"];
                    dr["DH1"] = TablaPrin.Rows[i]["dh1"];
                    dr["DI1"] = TablaPrin.Rows[i]["di1"];
                    dr["DJ1"] = TablaPrin.Rows[i]["dj1"];
                    dr["DK1"] = TablaPrin.Rows[i]["dk1"];
                    dr["DL1"] = TablaPrin.Rows[i]["dl1"];
                    dr["DM1"] = TablaPrin.Rows[i]["dm1"];
                    dr["DN1"] = TablaPrin.Rows[i]["dn1"];
                    dr["DO1"] = TablaPrin.Rows[i]["do1"];
                    dr["DP1"] = TablaPrin.Rows[i]["dp1"];
                    dr["DQ1"] = TablaPrin.Rows[i]["dq1"];
                    dr["DR1"] = TablaPrin.Rows[i]["dr1"];
                    dr["DS1"] = TablaPrin.Rows[i]["ds1"];
                    dr["DT1"] = TablaPrin.Rows[i]["dt1"];
                    dr["DU1"] = TablaPrin.Rows[i]["du1"];
                    dr["DV1"] = TablaPrin.Rows[i]["dv1"];
                    dr["DW1"] = TablaPrin.Rows[i]["dw1"];
                    dr["DX1"] = TablaPrin.Rows[i]["dx1"];
                    dr["DY1"] = TablaPrin.Rows[i]["dy1"];
                    dr["DZ1"] = TablaPrin.Rows[i]["dz1"];



                    dr["EA1"] = TablaPrin.Rows[i]["ea1"];
                    dr["EB1"] = TablaPrin.Rows[i]["eb1"];
                    dr["EC1"] = TablaPrin.Rows[i]["ec1"];
                    dr["ED1"] = TablaPrin.Rows[i]["ed1"];
                    dr["EE1"] = TablaPrin.Rows[i]["ee1"];
                    dr["EF1"] = TablaPrin.Rows[i]["ef1"];
                    dr["EG1"] = TablaPrin.Rows[i]["eg1"];
                    dr["EH1"] = TablaPrin.Rows[i]["eh1"];
                    dr["EI1"] = TablaPrin.Rows[i]["ei1"];
                    dr["EJ1"] = TablaPrin.Rows[i]["ej1"];
                    dr["EK1"] = TablaPrin.Rows[i]["ek1"];
                    dr["EL1"] = TablaPrin.Rows[i]["el1"];
                    dr["EM1"] = TablaPrin.Rows[i]["em1"];
                    dr["EN1"] = TablaPrin.Rows[i]["en1"];
                    dr["EO1"] = TablaPrin.Rows[i]["eo1"];
                    dr["EP1"] = TablaPrin.Rows[i]["ep1"];
                    dr["EQ1"] = TablaPrin.Rows[i]["eq1"];
                    dr["ER1"] = TablaPrin.Rows[i]["er1"];
                    dr["ES1"] = TablaPrin.Rows[i]["es1"];
                    dr["ET1"] = TablaPrin.Rows[i]["et1"];
                    dr["EU1"] = TablaPrin.Rows[i]["eu1"];
                    dr["EV1"] = TablaPrin.Rows[i]["ev1"];
                    dr["EW1"] = TablaPrin.Rows[i]["ew1"];
                    dr["EX1"] = TablaPrin.Rows[i]["ex1"];
                    dr["EY1"] = TablaPrin.Rows[i]["ey1"];
                    dr["EZ1"] = TablaPrin.Rows[i]["ez1"];

                    dr["FA1"] = TablaPrin.Rows[i]["fa1"];
                    dr["FB1"] = TablaPrin.Rows[i]["fb1"];
                    dr["FC1"] = TablaPrin.Rows[i]["fc1"];
                    dr["FD1"] = TablaPrin.Rows[i]["fd1"];
                    dr["FE1"] = TablaPrin.Rows[i]["fe1"];
                    dr["FF1"] = TablaPrin.Rows[i]["ff1"];
                    dr["FG1"] = TablaPrin.Rows[i]["fg1"];
                    dr["FH1"] = TablaPrin.Rows[i]["fh1"];
                    dr["FI1"] = TablaPrin.Rows[i]["fi1"];
                    dr["FJ1"] = TablaPrin.Rows[i]["fj1"];
                    dr["FK1"] = TablaPrin.Rows[i]["fk1"];
                    dr["FL1"] = TablaPrin.Rows[i]["fl1"];
                    dr["FM1"] = TablaPrin.Rows[i]["fm1"];
                    dr["FN1"] = TablaPrin.Rows[i]["fn1"];
                    dr["FO1"] = TablaPrin.Rows[i]["fo1"];
                    dr["FP1"] = TablaPrin.Rows[i]["fp1"];
                    dr["FQ1"] = TablaPrin.Rows[i]["fq1"];
                    dr["FR1"] = TablaPrin.Rows[i]["fr1"];
                    dr["FS1"] = TablaPrin.Rows[i]["fs1"];
                    dr["FT1"] = TablaPrin.Rows[i]["ft1"];
                    dr["FU1"] = TablaPrin.Rows[i]["fu1"];
                    dr["FV1"] = TablaPrin.Rows[i]["fv1"];
                    dr["FW1"] = TablaPrin.Rows[i]["fw1"];
                    dr["FX1"] = TablaPrin.Rows[i]["fx1"];
                    dr["FY1"] = TablaPrin.Rows[i]["fy1"];
                    dr["FZ1"] = TablaPrin.Rows[i]["fz1"];


                    dr["GA1"] = TablaPrin.Rows[i]["ga1"];
                    dr["GB1"] = TablaPrin.Rows[i]["gb1"];
                    dr["GC1"] = TablaPrin.Rows[i]["gc1"];
                    dr["GD1"] = TablaPrin.Rows[i]["gd1"];
                    dr["GE1"] = TablaPrin.Rows[i]["ge1"];
                    dr["GF1"] = TablaPrin.Rows[i]["gf1"];
                    dr["GG1"] = TablaPrin.Rows[i]["gg1"];
                    dr["GH1"] = TablaPrin.Rows[i]["gh1"];
                    dr["GI1"] = TablaPrin.Rows[i]["gi1"];
                    dr["GJ1"] = TablaPrin.Rows[i]["gj1"];
                    dr["GK1"] = TablaPrin.Rows[i]["gk1"];
                    dr["GL1"] = TablaPrin.Rows[i]["gl1"];
                    dr["GM1"] = TablaPrin.Rows[i]["gm1"];
                    dr["GN1"] = TablaPrin.Rows[i]["gn1"];
                    dr["GO1"] = TablaPrin.Rows[i]["go1"];
                    dr["GP1"] = TablaPrin.Rows[i]["gp1"];
                    dr["GQ1"] = TablaPrin.Rows[i]["gq1"];
                    dr["GR1"] = TablaPrin.Rows[i]["gr1"];
                    dr["GS1"] = TablaPrin.Rows[i]["gs1"];
                    dr["GT1"] = TablaPrin.Rows[i]["gt1"];
                    dr["GU1"] = TablaPrin.Rows[i]["gu1"];
                    dr["GV1"] = TablaPrin.Rows[i]["gv1"];
                    dr["GW1"] = TablaPrin.Rows[i]["gw1"];
                    dr["GX1"] = TablaPrin.Rows[i]["gx1"];
                    dr["GY1"] = TablaPrin.Rows[i]["gy1"];
                    dr["GZ1"] = TablaPrin.Rows[i]["gz1"];


                    dr["HA1"] = TablaPrin.Rows[i]["ha1"];
                    dr["HB1"] = TablaPrin.Rows[i]["hb1"];
                    dr["HC1"] = TablaPrin.Rows[i]["hc1"];
                    dr["HD1"] = TablaPrin.Rows[i]["hd1"];
                    dr["HE1"] = TablaPrin.Rows[i]["he1"];
                    dr["HF1"] = TablaPrin.Rows[i]["hf1"];
                    dr["HG1"] = TablaPrin.Rows[i]["hg1"];
                    dr["HH1"] = TablaPrin.Rows[i]["hh1"];
                    dr["HI1"] = TablaPrin.Rows[i]["hi1"];
                    dr["HJ1"] = TablaPrin.Rows[i]["hj1"];
                    dr["HK1"] = TablaPrin.Rows[i]["hk1"];
                    dr["HL1"] = TablaPrin.Rows[i]["hl1"];
                    dr["HM1"] = TablaPrin.Rows[i]["hm1"];
                    dr["HN1"] = TablaPrin.Rows[i]["hn1"];
                    dr["HO1"] = TablaPrin.Rows[i]["ho1"];
                    dr["HP1"] = TablaPrin.Rows[i]["hp1"];
                    dr["HQ1"] = TablaPrin.Rows[i]["hq1"];
                    dr["HR1"] = TablaPrin.Rows[i]["hr1"];
                    dr["HS1"] = TablaPrin.Rows[i]["hs1"];
                    dr["HT1"] = TablaPrin.Rows[i]["ht1"];
                    dr["HU1"] = TablaPrin.Rows[i]["hu1"];
                    dr["HV1"] = TablaPrin.Rows[i]["hv1"];
                    dr["HW1"] = TablaPrin.Rows[i]["hw1"];
                    dr["HX1"] = TablaPrin.Rows[i]["hx1"];
                    dr["HY1"] = TablaPrin.Rows[i]["hy1"];
                    dr["HZ1"] = TablaPrin.Rows[i]["hz1"];




                    dr["IA1"] = TablaPrin.Rows[i]["ia1"];
                    dr["IB1"] = TablaPrin.Rows[i]["ib1"];
                    dr["IC1"] = TablaPrin.Rows[i]["ic1"];
                    dr["ID1"] = TablaPrin.Rows[i]["id1"];
                    dr["IE1"] = TablaPrin.Rows[i]["ie1"];
                    dr["IF1"] = TablaPrin.Rows[i]["if1"];
                    dr["IG1"] = TablaPrin.Rows[i]["ig1"];
                    dr["IH1"] = TablaPrin.Rows[i]["ih1"];
                    dr["II1"] = TablaPrin.Rows[i]["ii1"];
                    dr["IJ1"] = TablaPrin.Rows[i]["ij1"];
                    dr["IK1"] = TablaPrin.Rows[i]["ik1"];
                    dr["IL1"] = TablaPrin.Rows[i]["il1"];
                    dr["IM1"] = TablaPrin.Rows[i]["im1"];
                    dr["IN1"] = TablaPrin.Rows[i]["in1"];
                    dr["IO1"] = TablaPrin.Rows[i]["io1"];
                    dr["IP1"] = TablaPrin.Rows[i]["ip1"];
                    dr["IQ1"] = TablaPrin.Rows[i]["iq1"];
                    dr["IR1"] = TablaPrin.Rows[i]["ir1"];
                    dr["IS1"] = TablaPrin.Rows[i]["is1"];
                    dr["IT1"] = TablaPrin.Rows[i]["it1"];
                    dr["IU1"] = TablaPrin.Rows[i]["iu1"];
                    dr["IV1"] = TablaPrin.Rows[i]["iv1"];
                    dr["IW1"] = TablaPrin.Rows[i]["iw1"];
                    dr["IX1"] = TablaPrin.Rows[i]["ix1"];
                    dr["IY1"] = TablaPrin.Rows[i]["iy1"];
                    dr["IZ1"] = TablaPrin.Rows[i]["iz1"];
                   


                    tablaCos.Rows.Add(dr);
                }

            }
            return tablaCos;
        }




        protected DataTable traspaso8(string id_mensaje)
        {
            ConexionCall ojjURL = new ConexionCall();
          

            string sqlEm = "  select email.* from email inner join email_errores on a1=email ";
            sqlEm += " where id_grupo=(select id_grupo from mensaje where id_mensaje=" + id_mensaje + " ) and activo=0 and Id_Email not in (Select Id_Email from envio_correo where id_mensaje ="+id_mensaje+") ";

            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            ConexionCall objEst = new ConexionCall();
            DataTable tablaCos = new DataTable();
            DataRow dr = null;
            tablaCos.Columns.Clear();
            tablaCos.Rows.Clear();

            int hdr = 0;
            if (TablaPrin.Rows.Count > 0)
            {

                for (int i = 0; i < TablaPrin.Rows.Count; i++)
                {
                    if (hdr == 0)
                    {
                        
                        tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));

                        //campos de tabla EMAIl
                        tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("L1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("M1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("N1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("O1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("P1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Q1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("R1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("S1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("T1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("U1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("V1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("W1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("X1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Y1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("Z1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("AA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("AZ1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));


                        tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));



                        tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));


                        tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));

                        tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                        tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));




                        hdr = 1;
                    }
                    //ACÁ cuidar de que no exista ningún campo que no corresponda a la tabla Confirmacion_evento
                    dr = tablaCos.NewRow();
                  
                    dr["Email"] = TablaPrin.Rows[i]["a1"];
                  



                    dr["B1"] = TablaPrin.Rows[i]["b1"];
                    dr["C1"] = TablaPrin.Rows[i]["c1"];
                    dr["D1"] = TablaPrin.Rows[i]["d1"];
                    dr["E1"] = TablaPrin.Rows[i]["e1"];
                    dr["F1"] = TablaPrin.Rows[i]["f1"];
                    dr["G1"] = TablaPrin.Rows[i]["g1"];
                    dr["H1"] = TablaPrin.Rows[i]["h1"];
                    dr["I1"] = TablaPrin.Rows[i]["i1"];
                    dr["J1"] = TablaPrin.Rows[i]["j1"];
                    dr["K1"] = TablaPrin.Rows[i]["k1"];
                    dr["L1"] = TablaPrin.Rows[i]["l1"];
                    dr["M1"] = TablaPrin.Rows[i]["m1"];
                    dr["N1"] = TablaPrin.Rows[i]["n1"];
                    dr["O1"] = TablaPrin.Rows[i]["o1"];
                    dr["P1"] = TablaPrin.Rows[i]["p1"];
                    dr["Q1"] = TablaPrin.Rows[i]["q1"];
                    dr["R1"] = TablaPrin.Rows[i]["r1"];
                    dr["S1"] = TablaPrin.Rows[i]["s1"];
                    dr["T1"] = TablaPrin.Rows[i]["t1"];
                    dr["U1"] = TablaPrin.Rows[i]["u1"];
                    dr["V1"] = TablaPrin.Rows[i]["v1"];
                    dr["W1"] = TablaPrin.Rows[i]["w1"];
                    dr["X1"] = TablaPrin.Rows[i]["x1"];
                    dr["Y1"] = TablaPrin.Rows[i]["y1"];
                    dr["Z1"] = TablaPrin.Rows[i]["z1"];

                    dr["AA1"] = TablaPrin.Rows[i]["aa1"];
                    dr["AB1"] = TablaPrin.Rows[i]["ab1"];
                    dr["AC1"] = TablaPrin.Rows[i]["ac1"];
                    dr["AD1"] = TablaPrin.Rows[i]["ad1"];
                    dr["AE1"] = TablaPrin.Rows[i]["ae1"];
                    dr["AF1"] = TablaPrin.Rows[i]["af1"];
                    dr["AG1"] = TablaPrin.Rows[i]["ag1"];
                    dr["AH1"] = TablaPrin.Rows[i]["ah1"];
                    dr["AI1"] = TablaPrin.Rows[i]["ai1"];
                    dr["AJ1"] = TablaPrin.Rows[i]["aj1"];
                    dr["AK1"] = TablaPrin.Rows[i]["ak1"];
                    dr["AL1"] = TablaPrin.Rows[i]["al1"];
                    dr["AM1"] = TablaPrin.Rows[i]["am1"];
                    dr["AN1"] = TablaPrin.Rows[i]["an1"];
                    dr["AO1"] = TablaPrin.Rows[i]["ao1"];
                    dr["AP1"] = TablaPrin.Rows[i]["ap1"];
                    dr["AQ1"] = TablaPrin.Rows[i]["aq1"];
                    dr["AR1"] = TablaPrin.Rows[i]["ar1"];
                    dr["AS1"] = TablaPrin.Rows[i]["as1"];
                    dr["AT1"] = TablaPrin.Rows[i]["at1"];
                    dr["AU1"] = TablaPrin.Rows[i]["au1"];
                    dr["AV1"] = TablaPrin.Rows[i]["av1"];
                    dr["AW1"] = TablaPrin.Rows[i]["aw1"];
                    dr["AX1"] = TablaPrin.Rows[i]["ax1"];
                    dr["AY1"] = TablaPrin.Rows[i]["ay1"];
                    dr["AZ1"] = TablaPrin.Rows[i]["az1"];

                    dr["BA1"] = TablaPrin.Rows[i]["ba1"];
                    dr["BB1"] = TablaPrin.Rows[i]["bb1"];
                    dr["BC1"] = TablaPrin.Rows[i]["bc1"];
                    dr["BD1"] = TablaPrin.Rows[i]["bd1"];
                    dr["BE1"] = TablaPrin.Rows[i]["be1"];
                    dr["BF1"] = TablaPrin.Rows[i]["bf1"];
                    dr["BG1"] = TablaPrin.Rows[i]["bg1"];
                    dr["BH1"] = TablaPrin.Rows[i]["bh1"];
                    dr["BI1"] = TablaPrin.Rows[i]["bi1"];
                    dr["BJ1"] = TablaPrin.Rows[i]["bj1"];
                    dr["BK1"] = TablaPrin.Rows[i]["bk1"];
                    dr["BL1"] = TablaPrin.Rows[i]["bl1"];
                    dr["BM1"] = TablaPrin.Rows[i]["bm1"];
                    dr["BN1"] = TablaPrin.Rows[i]["bn1"];
                    dr["BO1"] = TablaPrin.Rows[i]["bo1"];
                    dr["BP1"] = TablaPrin.Rows[i]["bp1"];
                    dr["BQ1"] = TablaPrin.Rows[i]["bq1"];
                    dr["BR1"] = TablaPrin.Rows[i]["br1"];
                    dr["BS1"] = TablaPrin.Rows[i]["bs1"];
                    dr["BT1"] = TablaPrin.Rows[i]["bt1"];
                    dr["BU1"] = TablaPrin.Rows[i]["bu1"];
                    dr["BV1"] = TablaPrin.Rows[i]["bv1"];
                    dr["BW1"] = TablaPrin.Rows[i]["bw1"];
                    dr["BX1"] = TablaPrin.Rows[i]["bx1"];
                    dr["BY1"] = TablaPrin.Rows[i]["by1"];
                    dr["BZ1"] = TablaPrin.Rows[i]["bz1"];


                    dr["CA1"] = TablaPrin.Rows[i]["ca1"];
                    dr["CB1"] = TablaPrin.Rows[i]["cb1"];
                    dr["CC1"] = TablaPrin.Rows[i]["cc1"];
                    dr["CD1"] = TablaPrin.Rows[i]["cd1"];
                    dr["CE1"] = TablaPrin.Rows[i]["ce1"];
                    dr["CF1"] = TablaPrin.Rows[i]["cf1"];
                    dr["CG1"] = TablaPrin.Rows[i]["cg1"];
                    dr["CH1"] = TablaPrin.Rows[i]["ch1"];
                    dr["CI1"] = TablaPrin.Rows[i]["ci1"];
                    dr["CJ1"] = TablaPrin.Rows[i]["cj1"];
                    dr["CK1"] = TablaPrin.Rows[i]["ck1"];
                    dr["CL1"] = TablaPrin.Rows[i]["cl1"];
                    dr["CM1"] = TablaPrin.Rows[i]["cm1"];
                    dr["CN1"] = TablaPrin.Rows[i]["cn1"];
                    dr["CO1"] = TablaPrin.Rows[i]["co1"];
                    dr["CP1"] = TablaPrin.Rows[i]["cp1"];
                    dr["CQ1"] = TablaPrin.Rows[i]["cq1"];
                    dr["CR1"] = TablaPrin.Rows[i]["cr1"];
                    dr["CS1"] = TablaPrin.Rows[i]["cs1"];
                    dr["CT1"] = TablaPrin.Rows[i]["ct1"];
                    dr["CU1"] = TablaPrin.Rows[i]["cu1"];
                    dr["CV1"] = TablaPrin.Rows[i]["cv1"];
                    dr["CW1"] = TablaPrin.Rows[i]["cw1"];
                    dr["CX1"] = TablaPrin.Rows[i]["cx1"];
                    dr["CY1"] = TablaPrin.Rows[i]["cy1"];
                    dr["CZ1"] = TablaPrin.Rows[i]["cz1"];


                    dr["DA1"] = TablaPrin.Rows[i]["da1"];
                    dr["DB1"] = TablaPrin.Rows[i]["db1"];
                    dr["DC1"] = TablaPrin.Rows[i]["dc1"];
                    dr["DD1"] = TablaPrin.Rows[i]["dd1"];
                    dr["DE1"] = TablaPrin.Rows[i]["de1"];
                    dr["DF1"] = TablaPrin.Rows[i]["df1"];
                    dr["DG1"] = TablaPrin.Rows[i]["dg1"];
                    dr["DH1"] = TablaPrin.Rows[i]["dh1"];
                    dr["DI1"] = TablaPrin.Rows[i]["di1"];
                    dr["DJ1"] = TablaPrin.Rows[i]["dj1"];
                    dr["DK1"] = TablaPrin.Rows[i]["dk1"];
                    dr["DL1"] = TablaPrin.Rows[i]["dl1"];
                    dr["DM1"] = TablaPrin.Rows[i]["dm1"];
                    dr["DN1"] = TablaPrin.Rows[i]["dn1"];
                    dr["DO1"] = TablaPrin.Rows[i]["do1"];
                    dr["DP1"] = TablaPrin.Rows[i]["dp1"];
                    dr["DQ1"] = TablaPrin.Rows[i]["dq1"];
                    dr["DR1"] = TablaPrin.Rows[i]["dr1"];
                    dr["DS1"] = TablaPrin.Rows[i]["ds1"];
                    dr["DT1"] = TablaPrin.Rows[i]["dt1"];
                    dr["DU1"] = TablaPrin.Rows[i]["du1"];
                    dr["DV1"] = TablaPrin.Rows[i]["dv1"];
                    dr["DW1"] = TablaPrin.Rows[i]["dw1"];
                    dr["DX1"] = TablaPrin.Rows[i]["dx1"];
                    dr["DY1"] = TablaPrin.Rows[i]["dy1"];
                    dr["DZ1"] = TablaPrin.Rows[i]["dz1"];



                    dr["EA1"] = TablaPrin.Rows[i]["ea1"];
                    dr["EB1"] = TablaPrin.Rows[i]["eb1"];
                    dr["EC1"] = TablaPrin.Rows[i]["ec1"];
                    dr["ED1"] = TablaPrin.Rows[i]["ed1"];
                    dr["EE1"] = TablaPrin.Rows[i]["ee1"];
                    dr["EF1"] = TablaPrin.Rows[i]["ef1"];
                    dr["EG1"] = TablaPrin.Rows[i]["eg1"];
                    dr["EH1"] = TablaPrin.Rows[i]["eh1"];
                    dr["EI1"] = TablaPrin.Rows[i]["ei1"];
                    dr["EJ1"] = TablaPrin.Rows[i]["ej1"];
                    dr["EK1"] = TablaPrin.Rows[i]["ek1"];
                    dr["EL1"] = TablaPrin.Rows[i]["el1"];
                    dr["EM1"] = TablaPrin.Rows[i]["em1"];
                    dr["EN1"] = TablaPrin.Rows[i]["en1"];
                    dr["EO1"] = TablaPrin.Rows[i]["eo1"];
                    dr["EP1"] = TablaPrin.Rows[i]["ep1"];
                    dr["EQ1"] = TablaPrin.Rows[i]["eq1"];
                    dr["ER1"] = TablaPrin.Rows[i]["er1"];
                    dr["ES1"] = TablaPrin.Rows[i]["es1"];
                    dr["ET1"] = TablaPrin.Rows[i]["et1"];
                    dr["EU1"] = TablaPrin.Rows[i]["eu1"];
                    dr["EV1"] = TablaPrin.Rows[i]["ev1"];
                    dr["EW1"] = TablaPrin.Rows[i]["ew1"];
                    dr["EX1"] = TablaPrin.Rows[i]["ex1"];
                    dr["EY1"] = TablaPrin.Rows[i]["ey1"];
                    dr["EZ1"] = TablaPrin.Rows[i]["ez1"];

                    dr["FA1"] = TablaPrin.Rows[i]["fa1"];
                    dr["FB1"] = TablaPrin.Rows[i]["fb1"];
                    dr["FC1"] = TablaPrin.Rows[i]["fc1"];
                    dr["FD1"] = TablaPrin.Rows[i]["fd1"];
                    dr["FE1"] = TablaPrin.Rows[i]["fe1"];
                    dr["FF1"] = TablaPrin.Rows[i]["ff1"];
                    dr["FG1"] = TablaPrin.Rows[i]["fg1"];
                    dr["FH1"] = TablaPrin.Rows[i]["fh1"];
                    dr["FI1"] = TablaPrin.Rows[i]["fi1"];
                    dr["FJ1"] = TablaPrin.Rows[i]["fj1"];
                    dr["FK1"] = TablaPrin.Rows[i]["fk1"];
                    dr["FL1"] = TablaPrin.Rows[i]["fl1"];
                    dr["FM1"] = TablaPrin.Rows[i]["fm1"];
                    dr["FN1"] = TablaPrin.Rows[i]["fn1"];
                    dr["FO1"] = TablaPrin.Rows[i]["fo1"];
                    dr["FP1"] = TablaPrin.Rows[i]["fp1"];
                    dr["FQ1"] = TablaPrin.Rows[i]["fq1"];
                    dr["FR1"] = TablaPrin.Rows[i]["fr1"];
                    dr["FS1"] = TablaPrin.Rows[i]["fs1"];
                    dr["FT1"] = TablaPrin.Rows[i]["ft1"];
                    dr["FU1"] = TablaPrin.Rows[i]["fu1"];
                    dr["FV1"] = TablaPrin.Rows[i]["fv1"];
                    dr["FW1"] = TablaPrin.Rows[i]["fw1"];
                    dr["FX1"] = TablaPrin.Rows[i]["fx1"];
                    dr["FY1"] = TablaPrin.Rows[i]["fy1"];
                    dr["FZ1"] = TablaPrin.Rows[i]["fz1"];


                    dr["GA1"] = TablaPrin.Rows[i]["ga1"];
                    dr["GB1"] = TablaPrin.Rows[i]["gb1"];
                    dr["GC1"] = TablaPrin.Rows[i]["gc1"];
                    dr["GD1"] = TablaPrin.Rows[i]["gd1"];
                    dr["GE1"] = TablaPrin.Rows[i]["ge1"];
                    dr["GF1"] = TablaPrin.Rows[i]["gf1"];
                    dr["GG1"] = TablaPrin.Rows[i]["gg1"];
                    dr["GH1"] = TablaPrin.Rows[i]["gh1"];
                    dr["GI1"] = TablaPrin.Rows[i]["gi1"];
                    dr["GJ1"] = TablaPrin.Rows[i]["gj1"];
                    dr["GK1"] = TablaPrin.Rows[i]["gk1"];
                    dr["GL1"] = TablaPrin.Rows[i]["gl1"];
                    dr["GM1"] = TablaPrin.Rows[i]["gm1"];
                    dr["GN1"] = TablaPrin.Rows[i]["gn1"];
                    dr["GO1"] = TablaPrin.Rows[i]["go1"];
                    dr["GP1"] = TablaPrin.Rows[i]["gp1"];
                    dr["GQ1"] = TablaPrin.Rows[i]["gq1"];
                    dr["GR1"] = TablaPrin.Rows[i]["gr1"];
                    dr["GS1"] = TablaPrin.Rows[i]["gs1"];
                    dr["GT1"] = TablaPrin.Rows[i]["gt1"];
                    dr["GU1"] = TablaPrin.Rows[i]["gu1"];
                    dr["GV1"] = TablaPrin.Rows[i]["gv1"];
                    dr["GW1"] = TablaPrin.Rows[i]["gw1"];
                    dr["GX1"] = TablaPrin.Rows[i]["gx1"];
                    dr["GY1"] = TablaPrin.Rows[i]["gy1"];
                    dr["GZ1"] = TablaPrin.Rows[i]["gz1"];


                    dr["HA1"] = TablaPrin.Rows[i]["ha1"];
                    dr["HB1"] = TablaPrin.Rows[i]["hb1"];
                    dr["HC1"] = TablaPrin.Rows[i]["hc1"];
                    dr["HD1"] = TablaPrin.Rows[i]["hd1"];
                    dr["HE1"] = TablaPrin.Rows[i]["he1"];
                    dr["HF1"] = TablaPrin.Rows[i]["hf1"];
                    dr["HG1"] = TablaPrin.Rows[i]["hg1"];
                    dr["HH1"] = TablaPrin.Rows[i]["hh1"];
                    dr["HI1"] = TablaPrin.Rows[i]["hi1"];
                    dr["HJ1"] = TablaPrin.Rows[i]["hj1"];
                    dr["HK1"] = TablaPrin.Rows[i]["hk1"];
                    dr["HL1"] = TablaPrin.Rows[i]["hl1"];
                    dr["HM1"] = TablaPrin.Rows[i]["hm1"];
                    dr["HN1"] = TablaPrin.Rows[i]["hn1"];
                    dr["HO1"] = TablaPrin.Rows[i]["ho1"];
                    dr["HP1"] = TablaPrin.Rows[i]["hp1"];
                    dr["HQ1"] = TablaPrin.Rows[i]["hq1"];
                    dr["HR1"] = TablaPrin.Rows[i]["hr1"];
                    dr["HS1"] = TablaPrin.Rows[i]["hs1"];
                    dr["HT1"] = TablaPrin.Rows[i]["ht1"];
                    dr["HU1"] = TablaPrin.Rows[i]["hu1"];
                    dr["HV1"] = TablaPrin.Rows[i]["hv1"];
                    dr["HW1"] = TablaPrin.Rows[i]["hw1"];
                    dr["HX1"] = TablaPrin.Rows[i]["hx1"];
                    dr["HY1"] = TablaPrin.Rows[i]["hy1"];
                    dr["HZ1"] = TablaPrin.Rows[i]["hz1"];




                    dr["IA1"] = TablaPrin.Rows[i]["ia1"];
                    dr["IB1"] = TablaPrin.Rows[i]["ib1"];
                    dr["IC1"] = TablaPrin.Rows[i]["ic1"];
                    dr["ID1"] = TablaPrin.Rows[i]["id1"];
                    dr["IE1"] = TablaPrin.Rows[i]["ie1"];
                    dr["IF1"] = TablaPrin.Rows[i]["if1"];
                    dr["IG1"] = TablaPrin.Rows[i]["ig1"];
                    dr["IH1"] = TablaPrin.Rows[i]["ih1"];
                    dr["II1"] = TablaPrin.Rows[i]["ii1"];
                    dr["IJ1"] = TablaPrin.Rows[i]["ij1"];
                    dr["IK1"] = TablaPrin.Rows[i]["ik1"];
                    dr["IL1"] = TablaPrin.Rows[i]["il1"];
                    dr["IM1"] = TablaPrin.Rows[i]["im1"];
                    dr["IN1"] = TablaPrin.Rows[i]["in1"];
                    dr["IO1"] = TablaPrin.Rows[i]["io1"];
                    dr["IP1"] = TablaPrin.Rows[i]["ip1"];
                    dr["IQ1"] = TablaPrin.Rows[i]["iq1"];
                    dr["IR1"] = TablaPrin.Rows[i]["ir1"];
                    dr["IS1"] = TablaPrin.Rows[i]["is1"];
                    dr["IT1"] = TablaPrin.Rows[i]["it1"];
                    dr["IU1"] = TablaPrin.Rows[i]["iu1"];
                    dr["IV1"] = TablaPrin.Rows[i]["iv1"];
                    dr["IW1"] = TablaPrin.Rows[i]["iw1"];
                    dr["IX1"] = TablaPrin.Rows[i]["ix1"];
                    dr["IY1"] = TablaPrin.Rows[i]["iy1"];
                    dr["IZ1"] = TablaPrin.Rows[i]["iz1"];



                    tablaCos.Rows.Add(dr);
                }

            }
            return tablaCos;
        }





        protected bool tieneevento(string id_mensaje)
        {
            ConexionCall ojjURL = new ConexionCall();
            string sqlag = " select id_ag from Mensaje  ";          
            sqlag += " where id_ag<>0 and id_mensaje= " + id_mensaje;
          

            DataTable TablaPrin = ConexionCall.SqlDTable(sqlag);
      
            if (TablaPrin.Rows.Count > 0)
            {
                return true;
            }else
            return false;
        }

        public string limpiaString(string valor)
        {
            valor = valor.Replace("/", "");
            valor = valor.Replace("-", "");
            valor = valor.Replace(" ", "_");
            valor = valor.Replace(":", "");
            valor = valor.Replace(">", "");
            valor = valor.Replace("<", "");

            string comillas = char.ConvertFromUtf32(34);
            string back_slash = char.ConvertFromUtf32(92);

            valor = valor.Replace(comillas, "");
            valor = valor.Replace(back_slash, "");

            return valor;
        }
        protected bool borra_anterior(string ruta)
        {

            try
            {
                if (File.Exists(ruta))
                {
                    File.Delete(ruta);
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                return false;
            }


        }
        public void actualizaSMS(object myObject, EventArgs myEventArgs)
        {

            try
            {
                hilo8.Stop();
                DataTable mensajes = ConexionCall.SqlDTable("select Id_sms from Mensaje_SMS order by Id_sms desc");
                string Id_sms, unico, query, enviados, conerror;
                int reg = mensajes.Rows.Count;
                if (reg > 0)
                {
                    for (int i = 0; i < reg; i++)
                    {

                        Id_sms = mensajes.Rows[i]["Id_sms"].ToString();

                        query = " select count(*) from envio_sms  where Id_sms=" + Id_sms + " and enviado=1";
                        enviados = ConexionCall.devuelveValor(query);

                        query = " select count(*) from envio_sms where Id_sms=" + Id_sms + " and error>0";
                        conerror = ConexionCall.devuelveValor(query);

                        this.Invoke(new DisplayEstado(Progreso), "Actualizando Mensaje " + Id_sms);
                        ConexionCall actualiza = new ConexionCall();
                        actualiza.ejecutorBase(" update Mensaje_SMS set enviados=" + enviados + ",codigoerror=" + conerror + " where Id_sms=" + Id_sms);
                    }

                }
                this.Invoke(new DisplayEstado(Progreso), "Actualizados " + reg + " mensajes");
                hilo8.Enabled = true;
            }
            catch (Exception)
            {
                hilo8.Enabled = true;
            }


        }
        public void crearCarpetasUsuarios(object myObject, EventArgs myEventArgs)
        {
            string root = ConfigurationSettings.AppSettings["root"];

            DataTable Tab = ConexionCall.SqlDTable("SELECT Id_Cliente FROM Usuario");
            int reg = Tab.Rows.Count;
            if (reg > 0)
            {
                for (int i = 0; i < reg; i++)
                {
                    try
                    {
                        if (!Directory.Exists(root + Tab.Rows[i]["Id_Cliente"].ToString() + ""))
                        {
                            Directory.CreateDirectory(root + Tab.Rows[i]["Id_Cliente"].ToString());
                            Directory.CreateDirectory(root + Tab.Rows[i]["Id_Cliente"].ToString() + "/banco");
                           //Directory.CreateDirectory(root + Tab.Rows[i]["Id_Cliente"].ToString() + "/adjunto");
                        }
                    }
                    catch (Exception es)
                    { }
                }
            }
        }
        public void Procesa_Excel_1(object myObject, EventArgs myEventArgs)
        {
            hilo3.Stop();
            try
            {
                Procesa_Excel();
                Thread.Sleep(60 * 1000);
                hilo3.Enabled = true;
            }
            catch (Exception ec)
            { hilo3.Enabled = true; }

        }
        public void Procesa_Excel_2(object myObject, EventArgs myEventArgs)
        {
            hilo10.Stop();
            try
            {
                Procesa_Excel();
                Thread.Sleep(60 * 1000);
                hilo10.Enabled = true;
            }
            catch (Exception ec)
            { hilo10.Enabled = true; }

        }













        /// <summary>
        /// /////////////////////////
        /// 
        /// 
        /// 
     //public void Procesa_Excel_3(object myObject, EventArgs myEventArgs)
     //   {
     //       hilo17.Stop();
     //       try
     //       {
     //           Procesa_Excel();
     //           Thread.Sleep(60 * 1000);
     //           hilo17.Enabled = true;
     //       }
     //       catch (Exception ec)
     //       { hilo17.Enabled = true; }

     //   }
     //   public void Procesa_Excel_4(object myObject, EventArgs myEventArgs)
     //   {
     //       hilo18.Stop();
     //       try
     //       {
     //           Procesa_Excel();
     //           Thread.Sleep(60 * 1000);
     //           hilo18.Enabled = true;
     //       }
     //       catch (Exception ec)
     //       { hilo18.Enabled = true; }

     //   }
        /// ////////////////////////
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            txb_aviso.Text = "";
            //string sqlCSV = "select t.archivos ,u.nombre+' '+u.appaterno as nombre from temporales t ";
            //sqlCSV += " join Usuario u on u.id_cliente=t.id_cliente where estado in (1,4)";

            string sqlCSV = "select t.archivos ,u.nombre+' '+u.appaterno as nombre from temporales t   join Grupo g on g.Id_Grupo=t.id_grupo ";
            sqlCSV += "  join Usuario u on u.Id_Usuario=g.Id_Usuario   where estado in (1,4)";





            DataTable tabCSV = ConexionCall.SqlDTable(sqlCSV);
            int csv = tabCSV.Rows.Count;

            //string sqlExcel = "select u.nombre+' '+u.appaterno as nombre ,m.asunto from Excel e ";
            //sqlExcel += " join Usuario u on u.id_cliente=e.id_cliente join mensaje m on m.id_mensaje=e.id_mensaje where estado=2";


            string sqlExcel = "select u.nombre+' '+u.appaterno as nombre ,m.asunto from Excel e  join Mensaje m on m.id_mensaje=e.id_mensaje  ";
            sqlExcel += " join Usuario u on u.Id_Usuario=m.Id_Usuario  where estado=2 ";






            DataTable tabExcel = ConexionCall.SqlDTable(sqlExcel);
            int exel = tabExcel.Rows.Count;

            //string sqlTXT = "SELECT Archivo,u.nombre+' '+u.appaterno as nombre  FROM TBL_Archivo a join grupo g on g.id_grupo=a.id_grupo ";
            //sqlTXT += "   join usuario u on u.id_usuario=g.id_usuario  where estado=1";


            string sqlTXT = "SELECT Archivo,u.nombre+' '+u.appaterno as nombre  FROM TBL_Archivo a join grupo g on g.id_grupo=a.id_grupo ";
            sqlTXT += "   join usuario u on u.id_usuario=g.id_usuario  where estado=1";





            DataTable tabTXT = ConexionCall.SqlDTable(sqlTXT);
            int txt = tabTXT.Rows.Count;
            #region
            string enter = char.ConvertFromUtf32(13) + char.ConvertFromUtf32(10);
            if (csv > 0)
            {
                for (int i = 0; i < csv; i++)
                {
                    txb_aviso.Text += tabCSV.Rows[i]["nombre"].ToString() + " está procesando " + tabCSV.Rows[i]["archivos"].ToString() + "";
                    txb_aviso.Text += enter;
                }
            }
            else
            {
                txb_aviso.Text += "No hay procesos de lectura de CSV ejecutándose";
            }
            #endregion
            #region
            txb_aviso.Text += enter;
            if (exel > 0)
            {
                for (int i = 0; i < exel; i++)
                {
                    txb_aviso.Text += tabExcel.Rows[i]["nombre"].ToString() + " está procesando mensaje " + tabExcel.Rows[i]["asunto"].ToString() + "";
                    txb_aviso.Text += enter;
                }


            }
            else
            {
                txb_aviso.Text += "No hay procesos de Excel ejecutándose";
            }
            #endregion
            #region
            txb_aviso.Text += enter;           
            if (txt > 0)
            {
                for (int i = 0; i < txt; i++)
                {
                    txb_aviso.Text += tabTXT.Rows[i]["nombre"].ToString() + " está exportando documento " + tabTXT.Rows[i]["Archivo"].ToString() + "";
                    txb_aviso.Text += enter;
                }
            }
            else
            {
                txb_aviso.Text += "No hay procesos de exportación ejecutándose";
            }
            #endregion
        }
        private void precarga1(object sender, EventArgs e)
        {
            hilo2.Stop();
            try
            {
                Precarga();
                Thread.Sleep(60 * 1000);
                hilo2.Enabled = true;
            }
            catch (Exception ec)
            { hilo2.Enabled = true; }
        }
        private void precarga2(object sender, EventArgs e)
        {
            hilo9.Stop();
            try
            {
                Thread.Sleep(100 * 1000);
                Precarga();
                hilo9.Enabled = true;
            }
            catch (Exception ec)
            { hilo9.Enabled = true; }
        }

        private void precarga3(object sender, EventArgs e)
        {
            hilo21.Stop();
            try
            {
                Thread.Sleep(180 * 1000);
                Precarga();
                hilo21.Enabled = true;
            }
            catch (Exception ec)
            { hilo21.Enabled = true; }
        }

        private void precarga4(object sender, EventArgs e)
        {
            hilo22.Stop();
            try
            {
                Thread.Sleep(280 * 1000);
                Precarga();
                hilo22.Enabled = true;
            }
            catch (Exception ec)
            { hilo22.Enabled = true; }
        }

        private void traspasar1(object sender, EventArgs e)
        {
            hilo1.Stop();
            try
            {
                Ejecuta_Traspaso();
                Thread.Sleep(60 * 1000);
                hilo1.Enabled = true;
            }
            catch (Exception ec)
            { hilo1.Enabled = true; }
        }
        private void traspasar2(object sender, EventArgs e)
        {
            hilo11.Stop();
            try
            {
                Thread.Sleep(80 * 1000);
                Ejecuta_Traspaso();
                hilo11.Enabled = true;
            }
            catch (Exception ec)
            { hilo11.Enabled = true; }
        }





        private void traspasar3(object sender, EventArgs e)
        {
            hilo20.Stop();
            try
            {
                Thread.Sleep(100 * 1000);
                Ejecuta_Traspaso();
                hilo20.Enabled = true;
            }
            catch (Exception ec)
            { hilo20.Enabled = true; }
        }





        private void traspasar4(object sender, EventArgs e)
        {
            hilo19.Stop();
            try
            {
                Thread.Sleep(120 * 1000);
                Ejecuta_Traspaso();
                hilo19.Enabled = true;
            }
            catch (Exception ec)
            { hilo19.Enabled = true; }
        }

        private void traspasar5(object sender, EventArgs e)
        {
            hilo23.Stop();
            try
            {
                Thread.Sleep(140 * 1000);
                Ejecuta_Traspaso();
                hilo23.Enabled = true;
            }
            catch (Exception ec)
            { hilo23.Enabled = true; }
        }

        private void traspasar6(object sender, EventArgs e)
        {
            hilo24.Stop();
            try
            {
                Thread.Sleep(160 * 1000);
                Ejecuta_Traspaso();
                hilo24.Enabled = true;
            }
            catch (Exception ec)
            { hilo24.Enabled = true; }
        }

        private void Unir_bases(object sender, EventArgs e)
        {
            hilo25.Stop();
            try
            {
                Procesa_Union_bases();
                Thread.Sleep(60 * 1000 * 2);
                hilo25.Enabled = true;
            }
            catch (Exception ec)
            { hilo25.Enabled = true; }
        }

        private void Envio_CargaFinalizada(object sender, EventArgs e)
        {
            hilo26.Stop();
            try
            {
                Procesa_Envio_cargabase();
                Thread.Sleep(60 * 1000);
                hilo26.Enabled = true;
            }
            catch (Exception ec)
            { hilo26.Enabled = true; }
        }


        /* private void actUnic()
        {
            try
            {
                int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
                if (validaProcesos == 0)
                {
                    #region
                    try
                    {
                        ConexionCall actualiza = new ConexionCall();
                        string Id_Mensaje, unico, query, enviados, conerror;

                       

                      DataTable  mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje  where id_estado=4 order by Id_Mensaje desc");
                        int    reg = mensajes.Rows.Count;
                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                query = " select count(*) from envio_correo  where id_mensaje=" + Id_Mensaje + " and abierto>0";
                                unico = ConexionCall.devuelveValor(query);
                                actualiza.ejecutorBase(" update Mensaje set unico=" + unico + " where Id_Mensaje=" + Id_Mensaje);

                                int abiertos = ConexionCall.devuelveValorINT("SELECT abierto FROM suma_env_abiertos where id_mensaje=" + Id_Mensaje);
                                actualiza.ejecutorBase(" update Mensaje set leidos=" + abiertos + " where Id_Mensaje=" + Id_Mensaje);

                                int registros = ConexionCall.devuelveValorINT("SELECT reg  FROM Reg_por_Men where id_mensaje=" + Id_Mensaje);
                                actualiza.ejecutorBase(" update Mensaje set reg=" + registros + " where Id_Mensaje=" + Id_Mensaje);
                                Thread.Sleep(10 * 1000);
                            }
                        }

                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                query = " select count(*) from envio_correo  where id_mensaje=" + Id_Mensaje + " and error>0";
                                conerror = ConexionCall.devuelveValor(query);
                                actualiza.ejecutorBase(" update Mensaje set codigoerror=" + conerror + "  where Id_Mensaje=" + Id_Mensaje);
                                Thread.Sleep(10 * 1000);
                            }

                        }


                    }
                    catch (Exception)
                    {

                    }
                    #endregion
                }
                else
                {
                    Thread.Sleep(120 * 1000);
                    actUnic();
                }
            }
            catch(Exception es)
            {}
        }
      */

        public void Procesa_reporte_1(object myObject, EventArgs myEventArgs)
        {
            hilo14.Stop();
            try
            {
                Procesa_reporte();
                Thread.Sleep(60 * 1000);
                hilo14.Enabled = true;
            }
            catch (Exception ec)
            { hilo14.Enabled = true; }

        }
        public void Procesa_reporte()
        {
            DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes where id_estado=1 ");
            int reg = ParaProceso.Rows.Count;
            //    bool anterior_Borrado = true;
            int correcto = 0;
            int malo = 0;
            if (reg > 0)
            {
                string  carpeta = "", id_rep = "", archivo;
                ConexionCall exe = new ConexionCall();

                for (int i = 0; i < reg; i++)
                {
                  
                    carpeta = ParaProceso.Rows[i]["carpeta"].ToString();
                    id_rep = ParaProceso.Rows[i]["id_rep"].ToString();
                    archivo = ParaProceso.Rows[i]["archivo"].ToString();

                    int valida = ConexionCall.devuelveValorINT("select count(*) from Reportes where id_estado=1 and id_rep=" + id_rep);

                    if (valida > 0)
                    {

                        #region
                        exe.ejecutorBase("update Reportes set id_estado=2 where id_rep=" + id_rep);

                        string fecha = DateTime.Now.ToString().Trim();
                        this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
                        ///////carga datos
                        DataTable emails = traspasoReporte(id_rep);
                        int rr = emails.Rows.Count;

                        //  archivo = ConexionCall.devuelveValor("select asunto from Mensaje where id_mensaje=" + id_mensaje);
                        archivo = "ID_rep_" + id_rep + "_" + fecha + ".xls";

                        archivo = limpiaString(archivo);
                        string ruta = carpeta + archivo;
                        if (ExportarExcelDataTable(emails, ruta))
                        {

                            exe.ejecutorBase("update Reportes set archivo='" + archivo + "',id_estado=4,fecha=getdate() where id_rep=" + id_rep);
                            this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
                            correcto++;
                        }
                        else
                        {
                            malo++;
                            exe.ejecutorBase("update Reportes set id_estado=1, archivo='' where id_rep=" + id_rep);

                        }
                        #endregion
                    }
                    this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos excel");
                    this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
                }
            }
        }







        public void Procesa_reporte_total(object myObject, EventArgs myEventArgs)
        {
            hilo15.Stop();
            try
            {
                Procesa_reporte_total();
                Thread.Sleep(60 * 1000);
                hilo15.Enabled = true;
            }
            catch (Exception ec)
            { hilo15.Enabled = true; }

        }



        public void Procesa_reporte_regunico(object myObject, EventArgs myEventArgs)
        {
            hilo16.Stop();
            try
            {
                Procesa_reporte_regunico();
                Thread.Sleep(60 * 1000);
                hilo16.Enabled = true;
            }
            catch (Exception ec)
            { hilo16.Enabled = true; }

        }



        //public void Procesa_reporte_total()
        //{
        //    DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes_total where id_estado=1 ");
        //    int reg = ParaProceso.Rows.Count;
        //    //    bool anterior_Borrado = true;
        //    int correcto = 0;
        //    int malo = 0;
        //    if (reg > 0)
        //    {
        //        string carpeta = "", id_rep = "", archivo;
        //        ConexionCall exe = new ConexionCall();

        //        for (int i = 0; i < reg; i++)
        //        {

        //           // carpeta = ParaProceso.Rows[i]["carpeta"].ToString();
        //            carpeta = ConfigurationSettings.AppSettings["rootExportaciones"] + "\\temp\\" + ParaProceso.Rows[i]["id_cliente"].ToString() + "\\";

        //           string   carpetalink = "http://clientes.hugoo.com/temp/" + ParaProceso.Rows[i]["id_cliente"].ToString() + "/";



        //            id_rep = ParaProceso.Rows[i]["id_rep"].ToString();
        //            archivo = ParaProceso.Rows[i]["archivo"].ToString();

        //            int valida = ConexionCall.devuelveValorINT("select count(*) from Reportes_total where id_estado=1 and id_rep=" + id_rep);

        //            if (valida > 0)
        //            {

        //                #region
        //                exe.ejecutorBase("update Reportes_total set id_estado=2 where id_rep=" + id_rep);

        //                string fecha = DateTime.Now.ToString().Trim();
        //                this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
        //                ///////carga datos
        //                DataTable emails = traspasoReporte_total(id_rep);
        //                int rr = emails.Rows.Count;
        //                //si hay 0 ya se perdieron los datos de envio_correo, hay q ver si existe archivo previo. SI existe renombrar? si no existe crear excel vacio
        //                //  archivo = ConexionCall.devuelveValor("select asunto from Mensaje where id_mensaje=" + id_mensaje);


        //                if ((rr > 0) || (archivo==""))
        //                { 


        //                    archivo = "Resumen_total_" + id_rep + "_" + fecha + ".xls";

        //                    archivo = limpiaString(archivo);
        //                    string ruta = carpeta + archivo;
        //                    string rutalink = carpetalink + archivo;


        //                    if (ExportarExcelDataTable(emails, ruta))
        //                    {

        //                        exe.ejecutorBase("update Reportes_total set archivo='" + archivo + "',carpeta='" + rutalink + "',id_grupo='" + rr + "',id_estado=4,fecha=getdate() where id_rep=" + id_rep);
        //                        this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
        //                        correcto++;
        //                    }
        //                    else
        //                    {
        //                        malo++;
        //                        exe.ejecutorBase("update Reportes_total set id_estado=1, archivo='' where id_rep=" + id_rep);

        //                    }
        //                #endregion





        //                }


        //                else {

        //                    exe.ejecutorBase("update Reportes_total set id_estado=4 where id_rep=" + id_rep);
                        
                        
        //                }
        //            }
        //            this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos excel");
        //            this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
        //        }
        //    }
        //}



        public void Procesa_reporte_total()
        {
            this.Invoke(new DisplayEstado(Progreso), "Comienzo generación de Reporte Total");
            DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes_total where id_estado=1 ");
         //   DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes_total where id_rep=320 ");
            int reg = ParaProceso.Rows.Count;
            //    bool anterior_Borrado = true;
            int correcto = 0;
            int malo = 0;
            if (reg > 0)
            {
                string carpeta = "", id_rep = "", archivo;
                ConexionCall exe = new ConexionCall();

                for (int i = 0; i < reg; i++)
                {

                    // carpeta = ParaProceso.Rows[i]["carpeta"].ToString();
                    carpeta = ConfigurationSettings.AppSettings["rootExportaciones"] + "\\temp\\" + ParaProceso.Rows[i]["id_cliente"].ToString() + "\\";

                    string carpetalink = "http://clientes.puntonet.com/temp/" + ParaProceso.Rows[i]["id_cliente"].ToString() + "/";



                    id_rep = ParaProceso.Rows[i]["id_rep"].ToString();
                    archivo = ParaProceso.Rows[i]["archivo"].ToString();

                    int valida = ConexionCall.devuelveValorINT("select count(*) from Reportes_total where id_estado=1 and id_rep=" + id_rep);
                    string rutalink = "";
                    if (valida > 0)
                    {

                        #region
                        exe.ejecutorBase("update Reportes_total set id_estado=2 where id_rep=" + id_rep);

                        string fecha = DateTime.Now.ToString().Trim();
                       // this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
                        ///////carga datos


                        // DataTable emails = traspasoReporte_total(id_rep);
                        DataTable emails = generareportetotal(id_rep);
                        int rr = emails.Rows.Count;
                        int ss = 0;
                        int sufijo = 1;
                        //si hay 0 ya se perdieron los datos de envio_correo, hay q ver si existe archivo previo. SI existe renombrar? si no existe crear excel vacio
                     
                        DataTable emailsub=emails.Clone();
                        //si es mas de 300000 separar en varios

                        foreach (DataRow dd in emails.Rows) {


                            emailsub.ImportRow(dd);



                            ss++;


                            if (ss % 300000 == 0||ss==rr) { 
                            


                                
                            archivo = "Resumen_total_" + id_rep + "_" + fecha + "_"+sufijo+".xls";

                            archivo = limpiaString(archivo);
                            string ruta = carpeta + archivo;
                             rutalink = carpetalink + archivo;


                            ExportarExcelDataTable(emailsub, ruta);



                            this.Invoke(new DisplayEstado(Progreso), "Procesadas " + sufijo + " partes de reporte total");


                            //aca crear el archivo
                            emailsub.Clear();
                                sufijo++;

                            }
                        
                        }


                        //if ((rr > 0) || (archivo == ""))
                        //{


                        //    archivo = "Resumen_total_" + id_rep + "_" + fecha + ".xls";

                        //    archivo = limpiaString(archivo);
                        //    string ruta = carpeta + archivo;
                        //    string rutalink = carpetalink + archivo;


                            //if (ExportarExcelDataTable(emails, ruta))
                            //{

                                exe.ejecutorBase("update Reportes_total set archivo='" + archivo + "',carpeta='" + rutalink + "',id_grupo='" + rr + "',id_estado=4,fecha=getdate() where id_rep=" + id_rep);
                               // this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
                                correcto++;
                            //}
                            //else
                            //{
                            //    malo++;
                            //    exe.ejecutorBase("update Reportes_total set id_estado=1, archivo='' where id_rep=" + id_rep);

                            //}
                        #endregion





                        //}


                        //else
                        //{

                        //    exe.ejecutorBase("update Reportes_total set id_estado=4 where id_rep=" + id_rep);


                        //}
                    }
                   // this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos excel");
                    this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
                }
            }
        }



        protected DataTable generareportetotal(string id_rep)
        {
            // id_rep = "320";
            this.Invoke(new DisplayEstado(Progreso), "Generando informe Total");
            DataTable repo = ConexionCall.SqlDTable("SELECT bus, id_grupo, mes, anio,id_cliente FROM reportes_total WHERE id_rep =" + id_rep);
            DataTable tablaCos = new DataTable();
            if (repo.Rows.Count > 0)
            {
                string bus = repo.Rows[0]["bus"].ToString();
                //string id_grupo = repo.Rows[0]["id_grupo"].ToString();
                string mes = repo.Rows[0]["mes"].ToString();
                string anio = repo.Rows[0]["anio"].ToString();
                string id_cliente = repo.Rows[0]["id_cliente"].ToString();
               // int id_usuario = ConexionCall.devuelveValorINT("select id_usuario from usuario where id_cliente=" + id_cliente);

                string sqlselect = "select id_mensaje, id_grupo from Mensaje where Id_Usuario in (select id_usuario from usuario where id_cliente=" + id_cliente+") and MONTH(Ejecucion) =" + mes + " and YEAR(Ejecucion) =" + anio+" and id_estado = 4";
                DataTable tabla_Mensajes = ConexionCall.SqlDTable(sqlselect);
                int cantidad = tabla_Mensajes.Rows.Count;


                for (int i = 0; i < cantidad; i++)
                {
                    this.Invoke(new DisplayEstado(Progreso), "Generando informe Total - informe "+(i+1)+"/"+ cantidad + "");
                    string id_mensaje = tabla_Mensajes.Rows[i]["id_mensaje"].ToString();
                    string id_grupo = tabla_Mensajes.Rows[i]["id_grupo"].ToString();
                    tablaCos.Merge(traspasoTotal(id_mensaje, id_grupo));
                }

            }

            return tablaCos;
        }

        protected DataTable traspasoReporte_total(string id_rep)
        {

            DataTable repo = ConexionCall.SqlDTable("SELECT bus, id_grupo, mes, anio,id_cliente FROM reportes_total WHERE id_rep =" + id_rep);
            DataTable tablaCos = new DataTable();
            if (repo.Rows.Count > 0)
            {
                string bus = repo.Rows[0]["bus"].ToString();
                string id_grupo = repo.Rows[0]["id_grupo"].ToString();
                string mes = repo.Rows[0]["mes"].ToString();
                string anio = repo.Rows[0]["anio"].ToString();
                string id_cliente = repo.Rows[0]["id_cliente"].ToString();
                int id_usuario = ConexionCall.devuelveValorINT("select id_usuario from usuario where id_cliente=" + id_cliente);
                #region

   //             string sqlEm = "  select isnull(em.nombre,'') as det_error,case abierto when 0 then 0 else 1 end as leido,  case error when 0 then 0 else 1 end as erro,";
   //sqlEm = sqlEm + " ec.*,e.* from envio_correo ec join Email e on e.id_email=ec.id_email  join mensaje m on m.id_mensaje= ec.id_mensaje";
   // sqlEm = sqlEm + " join Usuario u on u.Id_Usuario=m.Id_Usuario left join Estado_mail em on em.id_estado=ec.error";




                string sqlEm = "  select isnull(em.nombre,'') as det_error,case abierto when 0 then 0 else 1 end as leido,  case error when 0 then 0 else 1 end as erro";
               // sqlEm = sqlEm + " ,ec.id_enviado,ec.abierto ,ec.enviado,ec.fecha ,ec.id_mensaje ,ec.id_email ,ec.error ,ec.FechaLectura ,e.a1,e.e1,e.b1,e.c1,e.d1,e.f1,e.g1,e.h1,e.i1,e.j1,e.k1,e.l1,e.m1,e.Id_Grupo from envio_correo ec join Email e on e.id_email=ec.id_email  join mensaje m on m.id_mensaje= ec.id_mensaje";

                sqlEm = sqlEm + " ,ec.id_enviado,ec.abierto ,ec.enviado,ec.fecha ,ec.id_mensaje ,ec.id_email ,ec.error ,ec.FechaLectura ,e.a1,";

                sqlEm = sqlEm + " e.e1,e.b1,e.c1,e.d1,e.f1,e.g1,e.h1,e.i1,e.j1,e.k1,e.l1,e.m1,e.n1,e.o1,e.p1,e.q1,e.r1,e.s1,e.t1,e.u1,e.v1,e.w1,e.x1,e.y1,e.z1,"; // e.aa1,e.ab1,e.ac1,e.ad1,e.ae1,e.af1,e.ag1,e.ah1,e.ai1,e.aj1,e.ak1,e.al1,e.am1,e.an1,e.ao1,e.ap1,e.aq1,e.ar1,e.as1,e.at1,e.au1,e.av1,e.aw1,e.ax1,e.ay1,e.az1,e.ba1,e.bb1,e.bc1,e.bd1,e.be1,e.bf1,e.bg1,e.bh1,e.bi1,e.bj1,e.bk1,e.bl1,e.bm1,e.bn1,e.bo1,e.bp1,e.bq1,e.br1,e.bs1,e.bt1,e.bu1,e.bv1,e.bw1,e.bx1,e.by1,e.bz1,e.ca1,e.cb1,e.cc1,e.cd1,e.ce1,e.cf1,e.cg1,e.ch1,e.ci1,e.cj1,e.ck1,e.cl1,e.cm1,e.cn1,e.co1,e.cp1,e.cq1,e.cr1,e.cs1,e.ct1,e.cu1,e.cv1,e.cw1,e.cx1,e.cy1,e.cz1,e.da1,e.db1,e.dc1,e.dd1,e.de1,e.df1,e.dg1,e.dh1,e.di1,e.dj1,e.dk1,e.dl1,e.dm1,e.dn1,e.do1,e.dp1,e.dq1,e.dr1,e.ds1,e.dt1,e.du1,e.dv1,e.dw1,e.dx1,e.dy1,e.dz1,e.ea1,e.eb1,e.ec1,e.ed1,e.ee1,e.ef1,e.eg1,e.eh1,e.ei1,e.ej1,e.ek1,e.el1,e.em1,e.en1,e.eo1,e.ep1,e.eq1,e.er1,e.es1,e.et1,e.eu1,e.ev1,e.ew1,e.ex1,e.ey1,e.ez1,e.fa1,e.fb1,e.fc1,e.fd1,e.fe1,e.ff1,e.fg1,e.fh1,e.fi1,e.fj1,e.fk1,e.fl1,e.fm1,e.fn1,e.fo1,e.fp1,e.fq1,e.fr1,e.fs1,e.ft1,e.fu1,e.fv1,e.fw1,e.fx1,e.fy1,e.fz1,e.ga1,e.gb1,e.gc1,e.gd1,e.ge1,e.gf1,e.gg1,e.gh1,e.gi1,e.gj1,e.gk1,e.gl1,e.gm1,e.gn1,e.go1,e.gp1,e.gq1,e.gr1,e.gs1,e.gt1,e.gu1,e.gv1,e.gw1,e.gx1,e.gy1,e.gz1,e.ha1,e.hb1,e.hc1,e.hd1,e.he1,e.hf1,e.hg1,e.hh1,e.hi1,e.hj1,e.hk1,e.hl1,e.hm1,e.hn1,e.ho1,e.hp1,e.hq1,e.hr1,e.hs1,e.ht1,e.hu1,e.hv1,e.hw1,e.hx1,e.hy1,e.hz1,e.ia1,e.ib1,e.ic1,e.id1,e.ie1,e.if1,e.ig1,e.ih1,e.ii1,e.ij1,e.ik1,e.il1,e.im1,e.in1,e.io1,e.ip1,e.iq1,e.ir1,e.is1,e.it1,e.iu1,e.iv1,e.iw1,e.ix1,e.iy1,e.iz1, ";       
                
               sqlEm = sqlEm + "  e.Id_Grupo, g.Nombre from envio_correo ec join Email e on e.id_email=ec.id_email  join mensaje m on m.id_mensaje= ec.id_mensaje";
               
                
                sqlEm = sqlEm + " join Usuario u on u.Id_Usuario=m.Id_Usuario left join Estado_mail em on em.id_estado=ec.error left join Grupo g on m.Id_Grupo=g.Id_Grupo";
    

	sqlEm = sqlEm + " where u.Id_Cliente="+id_cliente+" and m.id_estado=4 and m.estatico=0 ";

 


                if (!string.IsNullOrEmpty(anio) && anio != "0")
                {
                    sqlEm += " and year(m.ejecucion)= " + anio;
                }
                if (!string.IsNullOrEmpty(mes) && mes != "0")
                {
                    sqlEm += " and month(m.ejecucion)=" + mes;
                }
                if (!string.IsNullOrEmpty(bus) && bus != "0")
                {
                    sqlEm += " and m.Id_Mensaje in ("+bus+") ";
                }
                sqlEm += " order by m.id_mensaje desc ";

                DataTable TablaPrin = ConexionCall.SqlDTable2(sqlEm);
                ConexionCall objEst = new ConexionCall();

                DataRow dr = null;
                tablaCos.Columns.Clear();
                tablaCos.Rows.Clear();

                int hdr = 0;
                if (TablaPrin.Rows.Count > 0)
                {


                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
 //ID_Mensaje;	ID_Envio;	Email;	Fecha;	Enviado;	Lectura Unica;	Fecha Apertura;	Lectura Total;	Error;	Tipo Error;	B1;	C1;D1;E1;F1;G1;H1;I1;J1;K1;L1;M1;"

                            tablaCos.Columns.Add(new DataColumn("ID_Mensaje", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Id_Grupo", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Nombre", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("ID_Enviado", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Email", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Fecha", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Fecha Apertura", typeof(string)));   
                            tablaCos.Columns.Add(new DataColumn("Enviado", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Lectura Unica", typeof(string)));
                                                     
                            tablaCos.Columns.Add(new DataColumn("Lectura Total", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Rebote", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Tipo Rebote", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("B1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("C1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("D1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("E1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("F1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("G1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("H1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("I1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("J1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("K1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("L1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("M1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("N1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("O1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("P1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Q1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("R1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("S1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("T1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("U1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("V1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("W1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("X1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Y1", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Z1", typeof(string)));


                            //tablaCos.Columns.Add(new DataColumn("AA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("AZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("BA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("BZ1", typeof(string)));


                            //tablaCos.Columns.Add(new DataColumn("CA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("CZ1", typeof(string)));


                            //tablaCos.Columns.Add(new DataColumn("DA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("DZ1", typeof(string)));

                            //tablaCos.Columns.Add(new DataColumn("EA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ED1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ER1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ES1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ET1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("EZ1", typeof(string)));



                            //tablaCos.Columns.Add(new DataColumn("FA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("FZ1", typeof(string)));







                            //tablaCos.Columns.Add(new DataColumn("GA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("GZ1", typeof(string)));






                            //tablaCos.Columns.Add(new DataColumn("HA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HD1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HI1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("HZ1", typeof(string)));









                            //tablaCos.Columns.Add(new DataColumn("IA1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IB1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IC1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("ID1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IE1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IF1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IG1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IH1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("II1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IJ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IK1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IL1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IM1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IN1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IO1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IP1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IQ1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IR1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IS1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IT1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IU1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IV1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IW1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IX1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IY1", typeof(string)));
                            //tablaCos.Columns.Add(new DataColumn("IZ1", typeof(string)));





                            hdr = 1;
                        }
                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(2 * 1000);
                        }

                        dr = tablaCos.NewRow();
                        dr["ID_Mensaje"] = TablaPrin.Rows[i]["id_mensaje"];
                        dr["Id_Grupo"] = TablaPrin.Rows[i]["Id_Grupo"];
                        dr["Nombre"] = TablaPrin.Rows[i]["nombre"];
                        dr["ID_Enviado"] = TablaPrin.Rows[i]["ID_Enviado"];
                        dr["Email"] = TablaPrin.Rows[i]["A1"];
                        dr["Fecha"] = TablaPrin.Rows[i]["Fecha"];
                        dr["Fecha Apertura"] = TablaPrin.Rows[i]["fechalectura"];
                        dr["Enviado"] = TablaPrin.Rows[i]["Enviado"];
                        dr["Lectura Unica"] = TablaPrin.Rows[i]["leido"];
                  
                        dr["Lectura Total"] = TablaPrin.Rows[i]["abierto"];
                        dr["Rebote"] = TablaPrin.Rows[i]["erro"];
                        dr["Tipo Rebote"] = TablaPrin.Rows[i]["det_error"];
                        dr["B1"] = TablaPrin.Rows[i]["B1"];
                        dr["C1"] = TablaPrin.Rows[i]["C1"];
                        dr["D1"] = TablaPrin.Rows[i]["D1"];
                        dr["E1"] = TablaPrin.Rows[i]["E1"];
                        dr["F1"] = TablaPrin.Rows[i]["F1"];
                        dr["G1"] = TablaPrin.Rows[i]["G1"];
                        dr["H1"] = TablaPrin.Rows[i]["H1"];
                        dr["I1"] = TablaPrin.Rows[i]["I1"];
                        dr["J1"] = TablaPrin.Rows[i]["J1"];
                        dr["K1"] = TablaPrin.Rows[i]["K1"];
                        dr["L1"] = TablaPrin.Rows[i]["L1"];
                        dr["M1"] = TablaPrin.Rows[i]["M1"];
                        dr["N1"] = TablaPrin.Rows[i]["N1"];
                        dr["O1"] = TablaPrin.Rows[i]["O1"];
                        dr["P1"] = TablaPrin.Rows[i]["P1"];
                        dr["Q1"] = TablaPrin.Rows[i]["Q1"];
                        dr["R1"] = TablaPrin.Rows[i]["R1"];
                        dr["S1"] = TablaPrin.Rows[i]["S1"];
                        dr["T1"] = TablaPrin.Rows[i]["T1"];
                        dr["U1"] = TablaPrin.Rows[i]["U1"];
                        dr["V1"] = TablaPrin.Rows[i]["V1"];
                        dr["W1"] = TablaPrin.Rows[i]["W1"];
                        dr["X1"] = TablaPrin.Rows[i]["X1"];
                        dr["Y1"] = TablaPrin.Rows[i]["Y1"];
                        dr["Z1"] = TablaPrin.Rows[i]["Z1"];

                        //dr["AB1"] = TablaPrin.Rows[i]["AB1"];
                        //dr["AB1"] = TablaPrin.Rows[i]["AB1"];
                        //dr["AC1"] = TablaPrin.Rows[i]["AC1"];
                        //dr["AD1"] = TablaPrin.Rows[i]["AD1"];
                        //dr["AE1"] = TablaPrin.Rows[i]["AE1"];
                        //dr["AF1"] = TablaPrin.Rows[i]["AF1"];
                        //dr["AG1"] = TablaPrin.Rows[i]["AG1"];
                        //dr["AH1"] = TablaPrin.Rows[i]["AH1"];
                        //dr["AI1"] = TablaPrin.Rows[i]["AI1"];
                        //dr["AJ1"] = TablaPrin.Rows[i]["AJ1"];
                        //dr["AK1"] = TablaPrin.Rows[i]["AK1"];
                        //dr["AL1"] = TablaPrin.Rows[i]["AL1"];
                        //dr["AM1"] = TablaPrin.Rows[i]["AM1"];
                        //dr["AN1"] = TablaPrin.Rows[i]["AN1"];
                        //dr["AO1"] = TablaPrin.Rows[i]["AO1"];
                        //dr["AP1"] = TablaPrin.Rows[i]["AP1"];
                        //dr["AQ1"] = TablaPrin.Rows[i]["AQ1"];
                        //dr["AR1"] = TablaPrin.Rows[i]["AR1"];
                        //dr["AS1"] = TablaPrin.Rows[i]["AS1"];
                        //dr["AT1"] = TablaPrin.Rows[i]["AT1"];
                        //dr["AU1"] = TablaPrin.Rows[i]["AU1"];
                        //dr["AV1"] = TablaPrin.Rows[i]["AV1"];
                        //dr["AW1"] = TablaPrin.Rows[i]["AW1"];
                        //dr["AX1"] = TablaPrin.Rows[i]["AX1"];
                        //dr["AY1"] = TablaPrin.Rows[i]["AY1"];
                        //dr["AZ1"] = TablaPrin.Rows[i]["AZ1"];



                        //dr["BA1"] = TablaPrin.Rows[i]["BA1"];
                        //dr["BB1"] = TablaPrin.Rows[i]["BB1"];
                        //dr["BC1"] = TablaPrin.Rows[i]["BC1"];
                        //dr["BD1"] = TablaPrin.Rows[i]["BD1"];
                        //dr["BE1"] = TablaPrin.Rows[i]["BE1"];
                        //dr["BF1"] = TablaPrin.Rows[i]["BF1"];
                        //dr["BG1"] = TablaPrin.Rows[i]["BG1"];
                        //dr["BH1"] = TablaPrin.Rows[i]["BH1"];
                        //dr["BI1"] = TablaPrin.Rows[i]["BI1"];
                        //dr["BJ1"] = TablaPrin.Rows[i]["BJ1"];
                        //dr["BK1"] = TablaPrin.Rows[i]["BK1"];
                        //dr["BL1"] = TablaPrin.Rows[i]["BL1"];
                        //dr["BM1"] = TablaPrin.Rows[i]["BM1"];
                        //dr["BN1"] = TablaPrin.Rows[i]["BN1"];
                        //dr["BO1"] = TablaPrin.Rows[i]["BO1"];
                        //dr["BP1"] = TablaPrin.Rows[i]["BP1"];
                        //dr["BQ1"] = TablaPrin.Rows[i]["BQ1"];
                        //dr["BR1"] = TablaPrin.Rows[i]["BR1"];
                        //dr["BS1"] = TablaPrin.Rows[i]["BS1"];
                        //dr["BT1"] = TablaPrin.Rows[i]["BT1"];
                        //dr["BU1"] = TablaPrin.Rows[i]["BU1"];
                        //dr["BV1"] = TablaPrin.Rows[i]["BV1"];
                        //dr["BW1"] = TablaPrin.Rows[i]["BW1"];
                        //dr["BX1"] = TablaPrin.Rows[i]["BX1"];
                        //dr["BY1"] = TablaPrin.Rows[i]["BY1"];
                        //dr["BZ1"] = TablaPrin.Rows[i]["BZ1"];



                        //dr["CA1"] = TablaPrin.Rows[i]["CA1"];
                        //dr["CB1"] = TablaPrin.Rows[i]["CB1"];
                        //dr["CC1"] = TablaPrin.Rows[i]["CC1"];
                        //dr["CD1"] = TablaPrin.Rows[i]["CD1"];
                        //dr["CE1"] = TablaPrin.Rows[i]["CE1"];
                        //dr["CF1"] = TablaPrin.Rows[i]["CF1"];
                        //dr["CG1"] = TablaPrin.Rows[i]["CG1"];
                        //dr["CH1"] = TablaPrin.Rows[i]["CH1"];
                        //dr["CI1"] = TablaPrin.Rows[i]["CI1"];
                        //dr["CJ1"] = TablaPrin.Rows[i]["CJ1"];
                        //dr["CK1"] = TablaPrin.Rows[i]["CK1"];
                        //dr["CL1"] = TablaPrin.Rows[i]["CL1"];
                        //dr["CM1"] = TablaPrin.Rows[i]["CM1"];
                        //dr["CN1"] = TablaPrin.Rows[i]["CN1"];
                        //dr["CO1"] = TablaPrin.Rows[i]["CO1"];
                        //dr["CP1"] = TablaPrin.Rows[i]["CP1"];
                        //dr["CQ1"] = TablaPrin.Rows[i]["CQ1"];
                        //dr["CR1"] = TablaPrin.Rows[i]["CR1"];
                        //dr["CS1"] = TablaPrin.Rows[i]["CS1"];
                        //dr["CT1"] = TablaPrin.Rows[i]["CT1"];
                        //dr["CU1"] = TablaPrin.Rows[i]["CU1"];
                        //dr["CV1"] = TablaPrin.Rows[i]["CV1"];
                        //dr["CW1"] = TablaPrin.Rows[i]["CW1"];
                        //dr["CX1"] = TablaPrin.Rows[i]["CX1"];
                        //dr["CY1"] = TablaPrin.Rows[i]["CY1"];
                        //dr["CZ1"] = TablaPrin.Rows[i]["CZ1"];



                        //dr["DA1"] = TablaPrin.Rows[i]["DA1"];
                        //dr["DB1"] = TablaPrin.Rows[i]["DB1"];
                        //dr["DC1"] = TablaPrin.Rows[i]["DC1"];
                        //dr["DD1"] = TablaPrin.Rows[i]["DD1"];
                        //dr["DE1"] = TablaPrin.Rows[i]["DE1"];
                        //dr["DF1"] = TablaPrin.Rows[i]["DF1"];
                        //dr["DG1"] = TablaPrin.Rows[i]["DG1"];
                        //dr["DH1"] = TablaPrin.Rows[i]["DH1"];
                        //dr["DI1"] = TablaPrin.Rows[i]["DI1"];
                        //dr["DJ1"] = TablaPrin.Rows[i]["DJ1"];
                        //dr["DK1"] = TablaPrin.Rows[i]["DK1"];
                        //dr["DL1"] = TablaPrin.Rows[i]["DL1"];
                        //dr["DM1"] = TablaPrin.Rows[i]["DM1"];
                        //dr["DN1"] = TablaPrin.Rows[i]["DN1"];
                        //dr["DO1"] = TablaPrin.Rows[i]["DO1"];
                        //dr["DP1"] = TablaPrin.Rows[i]["DP1"];
                        //dr["DQ1"] = TablaPrin.Rows[i]["DQ1"];
                        //dr["DR1"] = TablaPrin.Rows[i]["DR1"];
                        //dr["DS1"] = TablaPrin.Rows[i]["DS1"];
                        //dr["DT1"] = TablaPrin.Rows[i]["DT1"];
                        //dr["DU1"] = TablaPrin.Rows[i]["DU1"];
                        //dr["DV1"] = TablaPrin.Rows[i]["DV1"];
                        //dr["DW1"] = TablaPrin.Rows[i]["DW1"];
                        //dr["DX1"] = TablaPrin.Rows[i]["DX1"];
                        //dr["DY1"] = TablaPrin.Rows[i]["DY1"];
                        //dr["DZ1"] = TablaPrin.Rows[i]["DZ1"];



                        //dr["EA1"] = TablaPrin.Rows[i]["EA1"];
                        //dr["EB1"] = TablaPrin.Rows[i]["EB1"];
                        //dr["EC1"] = TablaPrin.Rows[i]["EC1"];
                        //dr["ED1"] = TablaPrin.Rows[i]["ED1"];
                        //dr["EE1"] = TablaPrin.Rows[i]["EE1"];
                        //dr["EF1"] = TablaPrin.Rows[i]["EF1"];
                        //dr["EG1"] = TablaPrin.Rows[i]["EG1"];
                        //dr["EH1"] = TablaPrin.Rows[i]["EH1"];
                        //dr["EI1"] = TablaPrin.Rows[i]["EI1"];
                        //dr["EJ1"] = TablaPrin.Rows[i]["EJ1"];
                        //dr["EK1"] = TablaPrin.Rows[i]["EK1"];
                        //dr["EL1"] = TablaPrin.Rows[i]["EL1"];
                        //dr["EM1"] = TablaPrin.Rows[i]["EM1"];
                        //dr["EN1"] = TablaPrin.Rows[i]["EN1"];
                        //dr["EO1"] = TablaPrin.Rows[i]["EO1"];
                        //dr["EP1"] = TablaPrin.Rows[i]["EP1"];
                        //dr["EQ1"] = TablaPrin.Rows[i]["EQ1"];
                        //dr["ER1"] = TablaPrin.Rows[i]["ER1"];
                        //dr["ES1"] = TablaPrin.Rows[i]["ES1"];
                        //dr["ET1"] = TablaPrin.Rows[i]["ET1"];
                        //dr["EU1"] = TablaPrin.Rows[i]["EU1"];
                        //dr["EV1"] = TablaPrin.Rows[i]["EV1"];
                        //dr["EW1"] = TablaPrin.Rows[i]["EW1"];
                        //dr["EX1"] = TablaPrin.Rows[i]["EX1"];
                        //dr["EY1"] = TablaPrin.Rows[i]["EY1"];
                        //dr["EZ1"] = TablaPrin.Rows[i]["EZ1"];




                        //dr["FA1"] = TablaPrin.Rows[i]["FA1"];
                        //dr["FB1"] = TablaPrin.Rows[i]["FB1"];
                        //dr["FC1"] = TablaPrin.Rows[i]["FC1"];
                        //dr["FD1"] = TablaPrin.Rows[i]["FD1"];
                        //dr["FE1"] = TablaPrin.Rows[i]["FE1"];
                        //dr["FF1"] = TablaPrin.Rows[i]["FF1"];
                        //dr["FG1"] = TablaPrin.Rows[i]["FG1"];
                        //dr["FH1"] = TablaPrin.Rows[i]["FH1"];
                        //dr["FI1"] = TablaPrin.Rows[i]["FI1"];
                        //dr["FJ1"] = TablaPrin.Rows[i]["FJ1"];
                        //dr["FK1"] = TablaPrin.Rows[i]["FK1"];
                        //dr["FL1"] = TablaPrin.Rows[i]["FL1"];
                        //dr["FM1"] = TablaPrin.Rows[i]["FM1"];
                        //dr["FN1"] = TablaPrin.Rows[i]["FN1"];
                        //dr["FO1"] = TablaPrin.Rows[i]["FO1"];
                        //dr["FP1"] = TablaPrin.Rows[i]["FP1"];
                        //dr["FQ1"] = TablaPrin.Rows[i]["FQ1"];
                        //dr["FR1"] = TablaPrin.Rows[i]["FR1"];
                        //dr["FS1"] = TablaPrin.Rows[i]["FS1"];
                        //dr["FT1"] = TablaPrin.Rows[i]["FT1"];
                        //dr["FU1"] = TablaPrin.Rows[i]["FU1"];
                        //dr["FV1"] = TablaPrin.Rows[i]["FV1"];
                        //dr["FW1"] = TablaPrin.Rows[i]["FW1"];
                        //dr["FX1"] = TablaPrin.Rows[i]["FX1"];
                        //dr["FY1"] = TablaPrin.Rows[i]["FY1"];
                        //dr["FZ1"] = TablaPrin.Rows[i]["FZ1"];




                        //dr["GA1"] = TablaPrin.Rows[i]["GA1"];
                        //dr["GB1"] = TablaPrin.Rows[i]["GB1"];
                        //dr["GC1"] = TablaPrin.Rows[i]["GC1"];
                        //dr["GD1"] = TablaPrin.Rows[i]["GD1"];
                        //dr["GE1"] = TablaPrin.Rows[i]["GE1"];
                        //dr["GF1"] = TablaPrin.Rows[i]["GF1"];
                        //dr["GG1"] = TablaPrin.Rows[i]["GG1"];
                        //dr["GH1"] = TablaPrin.Rows[i]["GH1"];
                        //dr["GI1"] = TablaPrin.Rows[i]["GI1"];
                        //dr["GJ1"] = TablaPrin.Rows[i]["GJ1"];
                        //dr["GK1"] = TablaPrin.Rows[i]["GK1"];
                        //dr["GL1"] = TablaPrin.Rows[i]["GL1"];
                        //dr["GM1"] = TablaPrin.Rows[i]["GM1"];
                        //dr["GN1"] = TablaPrin.Rows[i]["GN1"];
                        //dr["GO1"] = TablaPrin.Rows[i]["GO1"];
                        //dr["GP1"] = TablaPrin.Rows[i]["GP1"];
                        //dr["GQ1"] = TablaPrin.Rows[i]["GQ1"];
                        //dr["GR1"] = TablaPrin.Rows[i]["GR1"];
                        //dr["GS1"] = TablaPrin.Rows[i]["GS1"];
                        //dr["GT1"] = TablaPrin.Rows[i]["GT1"];
                        //dr["GU1"] = TablaPrin.Rows[i]["GU1"];
                        //dr["GV1"] = TablaPrin.Rows[i]["GV1"];
                        //dr["GW1"] = TablaPrin.Rows[i]["GW1"];
                        //dr["GX1"] = TablaPrin.Rows[i]["GX1"];
                        //dr["GY1"] = TablaPrin.Rows[i]["GY1"];
                        //dr["GZ1"] = TablaPrin.Rows[i]["GZ1"];





                        //dr["HA1"] = TablaPrin.Rows[i]["HA1"];
                        //dr["HB1"] = TablaPrin.Rows[i]["HB1"];
                        //dr["HC1"] = TablaPrin.Rows[i]["HC1"];
                        //dr["HD1"] = TablaPrin.Rows[i]["HD1"];
                        //dr["HE1"] = TablaPrin.Rows[i]["HE1"];
                        //dr["HF1"] = TablaPrin.Rows[i]["HF1"];
                        //dr["HG1"] = TablaPrin.Rows[i]["HG1"];
                        //dr["HH1"] = TablaPrin.Rows[i]["HH1"];
                        //dr["HI1"] = TablaPrin.Rows[i]["HI1"];
                        //dr["HJ1"] = TablaPrin.Rows[i]["HJ1"];
                        //dr["HK1"] = TablaPrin.Rows[i]["HK1"];
                        //dr["HL1"] = TablaPrin.Rows[i]["HL1"];
                        //dr["HM1"] = TablaPrin.Rows[i]["HM1"];
                        //dr["HN1"] = TablaPrin.Rows[i]["HN1"];
                        //dr["HO1"] = TablaPrin.Rows[i]["HO1"];
                        //dr["HP1"] = TablaPrin.Rows[i]["HP1"];
                        //dr["HQ1"] = TablaPrin.Rows[i]["HQ1"];
                        //dr["HR1"] = TablaPrin.Rows[i]["HR1"];
                        //dr["HS1"] = TablaPrin.Rows[i]["HS1"];
                        //dr["HT1"] = TablaPrin.Rows[i]["HT1"];
                        //dr["HU1"] = TablaPrin.Rows[i]["HU1"];
                        //dr["HV1"] = TablaPrin.Rows[i]["HV1"];
                        //dr["HW1"] = TablaPrin.Rows[i]["HW1"];
                        //dr["HX1"] = TablaPrin.Rows[i]["HX1"];
                        //dr["HY1"] = TablaPrin.Rows[i]["HY1"];
                        //dr["HZ1"] = TablaPrin.Rows[i]["HZ1"];





                        //dr["IA1"] = TablaPrin.Rows[i]["IA1"];
                        //dr["IB1"] = TablaPrin.Rows[i]["IB1"];
                        //dr["IC1"] = TablaPrin.Rows[i]["IC1"];
                        //dr["ID1"] = TablaPrin.Rows[i]["ID1"];
                        //dr["IE1"] = TablaPrin.Rows[i]["IE1"];
                        //dr["IF1"] = TablaPrin.Rows[i]["IF1"];
                        //dr["IG1"] = TablaPrin.Rows[i]["IG1"];
                        //dr["IH1"] = TablaPrin.Rows[i]["IH1"];
                        //dr["II1"] = TablaPrin.Rows[i]["II1"];
                        //dr["IJ1"] = TablaPrin.Rows[i]["IJ1"];
                        //dr["IK1"] = TablaPrin.Rows[i]["IK1"];
                        //dr["IL1"] = TablaPrin.Rows[i]["IL1"];
                        //dr["IM1"] = TablaPrin.Rows[i]["IM1"];
                        //dr["IN1"] = TablaPrin.Rows[i]["IN1"];
                        //dr["IO1"] = TablaPrin.Rows[i]["IO1"];
                        //dr["IP1"] = TablaPrin.Rows[i]["IP1"];
                        //dr["IQ1"] = TablaPrin.Rows[i]["IQ1"];
                        //dr["IR1"] = TablaPrin.Rows[i]["IR1"];
                        //dr["IS1"] = TablaPrin.Rows[i]["IS1"];
                        //dr["IT1"] = TablaPrin.Rows[i]["IT1"];
                        //dr["IU1"] = TablaPrin.Rows[i]["IU1"];
                        //dr["IV1"] = TablaPrin.Rows[i]["IV1"];
                        //dr["IW1"] = TablaPrin.Rows[i]["IW1"];
                        //dr["IX1"] = TablaPrin.Rows[i]["IX1"];
                        //dr["IY1"] = TablaPrin.Rows[i]["IY1"];
                        //dr["IZ1"] = TablaPrin.Rows[i]["IZ1"];







                        tablaCos.Rows.Add(dr);
                    }



                }

                #endregion

            }


            return tablaCos;
        }













        public void Procesa_reporte_regunico()
        {
             DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes_regunico where id_estado=1 "); 
           // DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM Reportes_regunico where id_repu=158 ");
            int reg = ParaProceso.Rows.Count;
   
            int correcto = 0;
            int malo = 0;
            if (reg > 0)
            {
                string carpeta = "", id_repu = "", archivo;
                ConexionCall exe = new ConexionCall();

                for (int i = 0; i < reg; i++)
                {

                    carpeta = ConfigurationSettings.AppSettings["rootExportaciones"] + "\\temp\\" + ParaProceso.Rows[i]["id_cliente"].ToString() + "\\";

                    string carpetalink = "http://clientes.hugoo.com/temp/" + ParaProceso.Rows[i]["id_cliente"].ToString() + "/";



                    id_repu = ParaProceso.Rows[i]["id_repu"].ToString();
                    archivo = ParaProceso.Rows[i]["archivo"].ToString();

                    int valida = ConexionCall.devuelveValorINT("select count(*) from Reportes_regunico where id_estado=1 and id_repu=" + id_repu);
                    string rutalink = "";
                    if (valida > 0)
                    {

                        #region
                        exe.ejecutorBase("update Reportes_regunico set id_estado=2 where id_repu=" + id_repu);

                        string fecha = DateTime.Now.ToString().Trim();
                        this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
                        ///////carga datos

                        DataTable emails = traspasoReporte_regunico(id_repu);
                        int rr = emails.Rows.Count;
                        int ss = 0;
                        int sufijo = 1;
                       DataTable emailsub = emails.Clone();
                        //si es mas de 500000 separar en varios

                        foreach (DataRow dd in emails.Rows)
                        {
                            emailsub.ImportRow(dd);

                            ss++;


                            if (ss % 300000 == 0 || ss == rr)
                            {
                                archivo = "Registros_unicos_" + id_repu + "_" + fecha + "_" + sufijo + ".xls";




                                archivo = limpiaString(archivo);
                                string ruta = carpeta + archivo;
                                rutalink = carpetalink + archivo;
                                
                                ExportarExcelDataTable(emailsub, ruta);
                                
                                this.Invoke(new DisplayEstado(Progreso), "Procesadas " + sufijo + " partes de reporte registros unicos");
                                
                                //aca crear el archivo
                                emailsub.Clear();
                                sufijo++;

                            }

                        }



                        exe.ejecutorBase("update Reportes_regunico set archivo='" + archivo + "',carpeta='" + rutalink + "',total='" + rr + "',id_estado=4,fecha=getdate() where id_repu=" + id_repu);
                        this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
                        correcto++;
                     
                        #endregion



                    }
                    this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos excel");
                    this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
                }
            }
        }






        protected DataTable traspasoReporte_regunico(string id_repu)
        {

            DataTable repo = ConexionCall.SqlDTable("SELECT total, mes, anio,id_cliente FROM Reportes_regunico WHERE id_repu =" + id_repu);
            DataTable tablaCos = new DataTable();
            if (repo.Rows.Count > 0)
            {
               // string bus = repo.Rows[0]["bus"].ToString();
                string total = repo.Rows[0]["total"].ToString();
                string mes = repo.Rows[0]["mes"].ToString();
                string anio = repo.Rows[0]["anio"].ToString();
                string id_cliente = repo.Rows[0]["id_cliente"].ToString();
                int id_usuario = ConexionCall.devuelveValorINT("select id_usuario from usuario where id_cliente=" + id_cliente);
                #region

           

                string sqlEm = "  select distinct a1 ,count (distinct(envio_correo.id_enviado) ) as enviados from envio_correo	inner join email w ";
   sqlEm = sqlEm + "  on w.Id_Email=envio_correo.id_email inner join mensaje m on m.Id_Mensaje=envio_correo.Id_Mensaje inner join usuario ";
   sqlEm = sqlEm + " u on u.id_usuario=m.id_usuario	where 1=1 and id_cliente=" + id_cliente;// + " and year(envio_correo.fecha)=2015 and month(envio_correo.fecha)=6 
     sqlEm = sqlEm + "  and m.estatico=0 ";
            

                if (!string.IsNullOrEmpty(anio) && anio != "0")
                {
                    sqlEm += " and year(envio_correo.fecha)= " + anio;
                }
                if (!string.IsNullOrEmpty(mes) && mes != "0")
                {
                    sqlEm += " and month(envio_correo.fecha)=" + mes;
                }

                sqlEm += " group by a1  order by a1 ";

                DataTable TablaPrin = ConexionCall.SqlDTable2(sqlEm);
                ConexionCall objEst = new ConexionCall();

                DataRow dr = null;
                tablaCos.Columns.Clear();
                tablaCos.Rows.Clear();

                int hdr = 0;
                if (TablaPrin.Rows.Count > 0)
                {


                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            //ID_Mensaje;	ID_Envio;	Email;	Fecha;	Enviado;	Lectura Unica;	Fecha Apertura;	Lectura Total;	Error;	Tipo Error;	B1;	C1;D1;E1;F1;G1;H1;I1;J1;K1;L1;M1;"

                            tablaCos.Columns.Add(new DataColumn("Correo", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Envios", typeof(string)));
                   


                            hdr = 1;
                        }
                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(2 * 1000);
                        }

                        dr = tablaCos.NewRow();
                        dr["Correo"] = TablaPrin.Rows[i]["a1"];
                        dr["Envios"] = TablaPrin.Rows[i]["enviados"];
                       



                        tablaCos.Rows.Add(dr);
                    }



                }

                #endregion

            }


            return tablaCos;
        }

















        protected DataTable traspasoReporte(string id_rep)
        {

            DataTable repo = ConexionCall.SqlDTable("SELECT bus, id_grupo, mes, anio,id_cliente FROM reportes WHERE id_rep =" + id_rep);
            DataTable tablaCos = new DataTable();
            if (repo.Rows.Count > 0)
            {
                string bus = repo.Rows[0]["bus"].ToString();
                string id_grupo = repo.Rows[0]["id_grupo"].ToString();
                string mes = repo.Rows[0]["mes"].ToString();
                string anio = repo.Rows[0]["anio"].ToString();
                string id_cliente = repo.Rows[0]["id_cliente"].ToString();
                int id_usuario = ConexionCall.devuelveValorINT("select id_usuario from usuario where id_cliente="+id_cliente);               #region

                string sqlEm = "select m.id_mensaje,m.fecha,m.ejecucion,m.asunto,r.reg,s.enviado,s.abierto,isnull(e.error,0)error,isnull(u.unico,0)unico ";
                sqlEm += " from Mensaje m ";
                sqlEm += "join Reg_por_Men r on r.id_mensaje=m.id_mensaje ";
                sqlEm += "join suma_env_abiertos s on s.id_mensaje=m.id_mensaje ";
                sqlEm += "left join  Errores e on e.id_mensaje=m.id_mensaje ";
                sqlEm += "left join  Unico u on u.id_mensaje=m.id_mensaje ";
                sqlEm += "where id_estado=4 and estatico=0 ";
                sqlEm += "and id_usuario= " + id_usuario;
                if (!string.IsNullOrEmpty(anio) && anio!="0")
                {
                    sqlEm += " and year(fecha)= "+anio;
                }
                if (!string.IsNullOrEmpty(mes) && mes != "0")
                {
                    sqlEm += " and month(fecha)="+mes;
                }
                if (!string.IsNullOrEmpty(bus) && bus != "0")
                {
                sqlEm += " and asunto like '%"+bus+"%' ";
                 }
                sqlEm += " order by m.id_mensaje desc ";

                DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
                ConexionCall objEst = new ConexionCall();
               
                DataRow dr = null;
                tablaCos.Columns.Clear();
                tablaCos.Rows.Clear();

                int hdr = 0;
                if (TablaPrin.Rows.Count > 0)
                {
                   

                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        if (hdr == 0)
                        {
                            tablaCos.Columns.Add(new DataColumn("ID", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Ingreso", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Envio", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Asunto", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Registros", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Enviados", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Rebotes", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Unica", typeof(string)));
                            tablaCos.Columns.Add(new DataColumn("Lectura Total", typeof(string)));
                            hdr = 1;
                        }
                        

                        
                        if (i != 0 && i % 5000 == 0)
                        {
                            Thread.Sleep(20 * 1000);
                        }

                        dr = tablaCos.NewRow();
                        dr["ID"] = TablaPrin.Rows[i]["id_mensaje"];
                        dr["Ingreso"] = TablaPrin.Rows[i]["fecha"];
                        dr["Envio"] = TablaPrin.Rows[i]["ejecucion"];
                        dr["Asunto"] = TablaPrin.Rows[i]["asunto"];
                        dr["Registros"] = TablaPrin.Rows[i]["reg"];
                        dr["Enviados"] = TablaPrin.Rows[i]["enviado"];
                        dr["Rebotes"] = TablaPrin.Rows[i]["error"];
                        dr["Unica"] = TablaPrin.Rows[i]["unico"];
                        dr["Lectura Total"] = TablaPrin.Rows[i]["abierto"];


                        tablaCos.Rows.Add(dr);
                    }
                 


                }

                #endregion

            }

            
            return tablaCos;
        }
        public void Elimina_Desinscritos(object myObject, EventArgs myEventArgs)
        {
            hilo5.Stop();
            DataTable tab = ConexionCall.SqlDTable("SELECT id_email,mail,id_grupo,id_cliente FROM Email_desincritos where eliminado=0");
            int reg = tab.Rows.Count;
            string id_email, mail, id_grupo, cliente, actualiza = "";
            if (reg > 0)
            {
                ConexionCall del = new ConexionCall();
                for (int i = 0; i < reg; i++)
                {
                    id_email = tab.Rows[i]["id_email"].ToString();
                    mail = tab.Rows[i]["mail"].ToString();
                    id_grupo = tab.Rows[i]["id_grupo"].ToString();
                    cliente = tab.Rows[i]["id_cliente"].ToString();

                    if (del.ejecutorBase("update Email set activo=0 where id_email=" + id_email + " and id_grupo=" + id_grupo))

                    //   if (del.ejecutorBase("DELETE FROM Email where id_email=" + id_email))
                    {
                        this.Invoke(new DisplayEstado(Progreso), "Correo " + mail + " del grupo " + id_grupo + " eliminado correctamente ");
                        del.ejecutorBase("UPDATE Email_desincritos   SET eliminado = 1 WHERE id_email=" + id_email);
                        //PAra evitar que el cliente continúe enviando mail desde BD YA CARGADAS: (Solo el cliente. Es problemático opara agencias de publicidad)
                        string datossql = "select id_email, id_grupo from email where id_grupo in (select id_grupo from grupo where Id_Usuario in (select id_usuario from Usuario  where id_cliente=" + cliente + ")) and a1 like '" + mail + "'";
                        DataTable datos = ConexionCall.SqlDTable(datossql);

                        if (datos.Rows.Count > 0)
                        {
                            for (int i2 = 0; datos.Rows.Count > i2; i2++)
                            {
                                string id_email2 = datos.Rows[i2]["id_email"].ToString();
                                string id_grupo2 = datos.Rows[i2]["id_grupo"].ToString();

                                del.ejecutorBase(" update email set activo=0 where id_email=" + id_email2 + " and id_grupo=" + id_grupo2);
                            }
                        }




                    }
                    else
                    {
                        this.Invoke(new DisplayEstado(Progreso), "Correo " + mail + " no pudo ser eliminado");
                    }

                    if (i % 50 == 0 && i != 0)
                    {
                        this.Invoke(new DisplayLimpia(limpiaMsg));

                    }

                }
                this.Invoke(new DisplayEstado(Progreso), "Se han eliminado " + reg + " correos ");


            }
            else
            {
                this.Invoke(new DisplayEstado(Progreso), "No hay correos para desinscripción ");
            }

            hilo5.Enabled = true;

        }

        public void AutomaticosEnviados_Abiertos()
        {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                try
                {
                    ConexionCall actualiza = new ConexionCall();
                    string Id_Mensaje, unico, query="", enviados, conerror;

                    DataTable mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje where id_estado=8 order by Id_Mensaje desc");
                    int reg = mensajes.Rows.Count;
                    if (reg > 0)
                    {
                        for (int i = 0; i < reg; i++)
                        {
                            Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                            DataTable env_abiertos = ConexionCall.SqlDTable("SELECT enviado,abierto FROM suma_env_abiertos  where id_mensaje=" + Id_Mensaje);
                            if (env_abiertos.Rows.Count > 0)
                            {
                                query = " update Mensaje set enviados=" + env_abiertos.Rows[0]["enviado"].ToString() + ", leidos=" + env_abiertos.Rows[0]["abierto"].ToString() + "  where Id_Mensaje=" + Id_Mensaje;
                            }

                            actualiza.ejecutorBase(query);

                        }
                    }
                }
                catch (Exception ex)
                {
                    AutomaticosEnviados_Abiertos();
                }
            }
            else
            {
                Thread.Sleep(60 * 1000);
                AutomaticosEnviados_Abiertos();
            }

        }

        public void AutomaticosRegistros()
        {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                try
                {
                    ConexionCall actualiza = new ConexionCall();
                    string Id_Mensaje, unico, query, enviados, conerror;

                    DataTable mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje where id_estado=8 order by Id_Mensaje desc");
                    int reg = mensajes.Rows.Count;
                    if (reg > 0)
                    {
                        for (int i = 0; i < reg; i++)
                        {
                            Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                            int registros = ConexionCall.devuelveValorINT("SELECT reg  FROM Reg_por_Men where id_mensaje=" + Id_Mensaje);
                            actualiza.ejecutorBase(" update Mensaje set reg=" + registros + " where Id_Mensaje=" + Id_Mensaje);
                            Thread.Sleep(10 * 1000);
                        }
                    }
                }
                catch (Exception ex)
                { }
            }
            else
            {
                Thread.Sleep(60 * 1000);
                AutomaticosRegistros();
            }

        }

        public void AutomaticosErrores()
        {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                #region
                try
                {
                    ConexionCall actualiza = new ConexionCall();
                    string Id_Mensaje, unico, query, enviados;
                    int conerror = 0;
                    DataTable mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje where id_estado=8 order by Id_Mensaje desc");
                    int reg = mensajes.Rows.Count;
                    if (reg > 0)
                    {
                        for (int i = 0; i < reg; i++)
                        {
                            Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                            query = "SELECT error FROM Errores  where id_mensaje=" + Id_Mensaje;
                            conerror = ConexionCall.devuelveValorINT(query);
                            actualiza.ejecutorBase(" update Mensaje set codigoerror=" + conerror + "  where Id_Mensaje=" + Id_Mensaje);
                            Thread.Sleep(10 * 1000);
                        }
                    }
                }
                catch (Exception)
                {
                    AutomaticosErrores();
                }
                #endregion
            }
            else
            {
                Thread.Sleep(60 * 1000);
                AutomaticosErrores();
            }

        }

        public void AutomaticosUnicos()
        {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                #region
                try
                {
                    ConexionCall actualiza = new ConexionCall();
                    string Id_Mensaje,  query, enviados, conerror;
                    int unico = 0;
                    DataTable mensajes = ConexionCall.SqlDTable("select Id_Mensaje from Mensaje where id_estado=8 order by Id_Mensaje desc");
                    int reg = mensajes.Rows.Count;
                    if (reg > 0)
                    {
                        for (int i = 0; i < reg; i++)
                        {
                            Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                            query = " SELECT unico FROM Unico  where id_mensaje=" + Id_Mensaje;
                            unico = ConexionCall.devuelveValorINT(query);
                            actualiza.ejecutorBase(" update Mensaje set unico=" + unico + " where Id_Mensaje=" + Id_Mensaje);

                        }
                    }

                }
                catch (Exception)
                {
                    AutomaticosUnicos();
                }
                #endregion
            }
            else
            {
                Thread.Sleep(60 * 1000);
                AutomaticosUnicos();
            }

        }

        public void MensajesUnicos()
        {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                    #region
                    try
                    {
                        ConexionCall actualiza = new ConexionCall();
                        string Id_Mensaje,query;
                        int unico=0,  enviados, conerror;  
                        DataTable  mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje  where id_estado=4 order by Id_Mensaje desc");
                        int    reg = mensajes.Rows.Count;
                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                query = " SELECT unico FROM Unico  where id_mensaje=" + Id_Mensaje;
                                unico = ConexionCall.devuelveValorINT(query);
                                actualiza.ejecutorBase(" update Mensaje set unico=" + unico + " where Id_Mensaje=" + Id_Mensaje);

                            }
                        }

                    }
                    catch (Exception)
                    {
                    MensajesUnicos();
                    }
                    #endregion 
            }
            else
            {
                Thread.Sleep(60 * 1000);
                MensajesUnicos();
            }

        }

        public void MensajesEnviados_Abiertos()
         {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                    #region
                    try
                    {
                        ConexionCall actualiza = new ConexionCall();
                        string Id_Mensaje, unico, query="", enviados, conerror;  
                        DataTable  mensajes = ConexionCall.SqlDTable("select Id_Mensaje from Mensaje  where id_estado=4 order by Id_Mensaje desc");
                        int    reg = mensajes.Rows.Count;
                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                DataTable env_abiertos = ConexionCall.SqlDTable("SELECT enviado,abierto FROM suma_env_abiertos  where id_mensaje=" + Id_Mensaje);
                                if (env_abiertos.Rows.Count > 0)
                                {
                                    query = " update Mensaje set enviados=" + env_abiertos.Rows[0]["enviado"].ToString() + ", leidos=" + env_abiertos.Rows[0]["abierto"].ToString() + "  where Id_Mensaje=" + Id_Mensaje;
                                }

                                actualiza.ejecutorBase(query);
                            }
                        }   
                    }
                    catch (Exception)
                    {
                        MensajesEnviados_Abiertos();
                    }
                    #endregion 
            }
            else
            {
                Thread.Sleep(60 * 1000);
                MensajesEnviados_Abiertos();
            }

        }

        public void MensajesRegistros()
         {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                    #region
                    try
                    {
                        ConexionCall actualiza = new ConexionCall();
                        string Id_Mensaje, unico, query, enviados, conerror;  
                        DataTable  mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje  where id_estado=4 order by Id_Mensaje desc");
                        int    reg = mensajes.Rows.Count;
                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                int registros = ConexionCall.devuelveValorINT("SELECT reg  FROM Reg_por_Men where id_mensaje=" + Id_Mensaje);
                                actualiza.ejecutorBase(" update Mensaje set reg=" + registros + " where Id_Mensaje=" + Id_Mensaje);
                                Thread.Sleep(10 * 1000);
                            }
                        }
                    }
                    catch (Exception ex) {  AutomaticosRegistros();}
                    #endregion 
            }
            else
            {
                Thread.Sleep(60 * 1000);
                AutomaticosRegistros();
            }

        }
        public void MensajesErrores()
         {
            int validaProcesos = ConexionCall.devuelveValorINT("select count(*) as total from Mensaje where  id_estado in (1,2,9)");
            if (validaProcesos == 0)
            {
                    #region
                    try
                    {
                        ConexionCall actualiza = new ConexionCall();
                        string Id_Mensaje, unico, query, enviados;
                        int conerror = 0;
                        DataTable  mensajes = ConexionCall.SqlDTable(" select Id_Mensaje from Mensaje  where id_estado=4 order by Id_Mensaje desc");
                        int    reg = mensajes.Rows.Count;                       
                        if (reg > 0)
                        {
                            for (int i = 0; i < reg; i++)
                            {
                                Id_Mensaje = mensajes.Rows[i]["Id_Mensaje"].ToString();
                                query = "SELECT error FROM Errores  where id_mensaje=" + Id_Mensaje;
                                conerror = ConexionCall.devuelveValorINT(query);
                                actualiza.ejecutorBase(" update Mensaje set codigoerror=" + conerror + "  where Id_Mensaje=" + Id_Mensaje);
                                Thread.Sleep(10 * 1000);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MensajesErrores();
                    }
                    #endregion
            }
            else
            {
                Thread.Sleep(60 * 1000);
                MensajesErrores();
            }

        }
        protected void eliminaDuplicados(object myObject, EventArgs myEventArgs)
        {//REVISAR ESTO!!!!
            hilo12.Stop();
            string sqlRepetidos = " select id_enviado from envio_correo ";
            sqlRepetidos+=" where id_mensaje=513 and id_email in ";
            sqlRepetidos+=" (select id_email from  repetidos_por_envio ";
            sqlRepetidos+=" where id_mensaje=513 and veces>1) ";
            sqlRepetidos+=" order by id_email";
            string id_enviado="";

            int borrados = 0;
                int replicados=0;
            DataTable rep = ConexionCall.SqlDTable(sqlRepetidos);
            int reg = rep.Rows.Count;
            if (reg > 0)
            { 
                ConexionCall borra=new ConexionCall();
                for(int i=0;i<reg;i++ )
                {
                    id_enviado = rep.Rows[i][0].ToString();

                    if (i % 2 == 0)
                    {
                        if (borra.ejecutorBase("delete from envio_correo where id_enviado=" + id_enviado))
                        {
                            borrados++;
                        }
                        else
                        {
                            replicados++;
                        }
                    }

                }

                string msg = "registros traídos son " + reg + " borrados son " + borrados + " y duplicados son " + replicados;
            
            }
                 
        }

        protected void cambiaIDEMAIL(object myObject, EventArgs myEventArgs)
        {
            hilo12.Stop();
            string sqlIEMAIL = "SELECT   id_email,a1 FROM Email where id_grupo=125 and id_email not in (SELECT id_email  FROM envio_correo where id_mensaje=563)";
            DataTable remail= ConexionCall.SqlDTable(sqlIEMAIL);
            int registrosMail=remail.Rows.Count;

            string id_email = "";
            string id_enviado = "";
            string actualiza = "";
            int borrados = 0;
            int replicados = 0;
            
            string sqlRepetidos = "SELECT id_enviado FROM envio_correo where id_mensaje=567";
            DataTable rep = ConexionCall.SqlDTable(sqlRepetidos);
            int reg = rep.Rows.Count;
            if (reg > 0)
            {
                ConexionCall acualiza = new ConexionCall();
                for (int i = 0; i < reg; i++)
                {
                    id_enviado = rep.Rows[i][0].ToString();
                    
                    try{
                    id_email=remail.Rows[i]["id_email"].ToString();
                    actualiza = "update envio_correo set id_email="+id_email+" where id_enviado=" + id_enviado;

                    if (acualiza.ejecutorBase(actualiza))
                    {
                        replicados++;
                    }
                    else
                    {
                        borrados++;
                    }
                    }
                    catch(Exception ex)
                    {
                        borrados++;
                    }
                }

                string msg = "registros traídos son " + reg + " borrados son " + borrados + " y duplicados son " + replicados;

            }

        }

        private void ExportarTXT1(object sender, EventArgs e)
        {
            hilo12.Stop();
            try
            {
                Exportacion_datos_Excel();
                Thread.Sleep(60 * 1000);
                hilo12.Enabled = true;
            }
            catch (Exception ec)
            { hilo12.Enabled = true; }
        }
        private void ExportarTXT2(object sender, EventArgs e)
        {
            hilo13.Stop();
            try
            {               
                Thread.Sleep(180 * 1000);
                Exportacion_datos_Excel();
                hilo12.Enabled = true;
            }
            catch (Exception ec)
            { hilo13.Enabled = true; }
        }
        public void Exportacion_datos_Excel()
        {
            DataTable ParaProceso = ConexionCall.SqlDTable("SELECT * FROM TBL_Archivo where estado=0 order by id");
            int reg = ParaProceso.Rows.Count;
            //    bool anterior_Borrado = true;
            int correcto = 0;
            int malo = 0;
            if (reg > 0)
            {
                string id, id_temp  , id_grupo, id_subgrupo ,nombreGrupo,archivo,carpeta;
                ConexionCall exe = new ConexionCall();

                for (int i = 0; i < reg; i++)
                {
                    #region
                    id_temp = ParaProceso.Rows[i]["id_temp"].ToString();
                    carpeta = ParaProceso.Rows[i]["carpeta"].ToString();
                    id = ParaProceso.Rows[i]["id"].ToString();
                    archivo = ParaProceso.Rows[i]["Archivo"].ToString();
                    id_grupo = ParaProceso.Rows[i]["id_grupo"].ToString();
                    id_subgrupo = ParaProceso.Rows[i]["id_subgrupo"].ToString();

                    nombreGrupo = ConexionCall.devuelveValor("select nombre from grupo where id_grupo="+id_grupo);
               
                    #endregion
                    int valida = ConexionCall.devuelveValorINT("select count(*) from TBL_Archivo where estado=0 and id=" + id);
                    if (valida > 0)
                    {
                        #region
                        exe.ejecutorBase("update TBL_Archivo set estado=1 where id=" + id);
                       
                        this.Invoke(new DisplayEstado(Progreso), "Cargando excel N° " + (i + 1) + " de" + reg);
                        ///////carga datos
                        #region carga tablas

                        string ruta1 = ConfigurationSettings.AppSettings["rootExportaciones"] + carpeta + archivo+".txt";
                        borra_anterior(ruta1);
                        bool txtCreado = traspasoExporta(id_grupo, id_subgrupo, ruta1);
                     
                        #endregion

                        if (txtCreado)
                        {
                            exe.ejecutorBase("update TBL_Archivo set estado=2  where id=" + id);
                            this.Invoke(new DisplayEstado(Progreso), "Excel N° " + (i + 1) + " generado correctamente");
                            correcto++;
                        }
                        else
                        {
                            malo++;
                            exe.ejecutorBase("update TBL_Archivo set estado=0 where id=" + id);

                        }
                        #endregion
                    }
                    this.Invoke(new DisplayEstado(Progreso), "Procesados " + reg + " documentos texto");
                    this.Invoke(new DisplayEstado(Progreso), "Procesados correctamente " + correcto + " , errores " + malo);
                }
            }
        }
        protected bool traspasoExporta(string id_grupo, string id_subgrupo, string fileName)
        {
            string sqlEm = "";
            bool resultado=true;
            int subG=int.Parse(id_subgrupo);
            string cabecera = "a1;b1;c1;d1;e1;f1;g1;h1;i1;j1;k1;l1;m1;n1;ñ1;o1;p1;q1;r1;s1;t1;u1;v1;w1;x1;y1;z1;aa1;ab1;ac1;ad1;ae1;af1;ag1;ah1;ai1;aj1;ak1;al1;am1;an1;ao1;ap1;aq1;ar1;as1;at1;au1;av1;aw1;ax1;ay1;az1;ba1;bb1;bc1;bd1;be1;bf1;bg1;bh1;bi1;bj1;bk1;bl1;bm1;bn1;bo1;bp1;bq1;br1;bs1;bt1;bu1;bv1;bw1;bx1;by1;bz1;ca1;cb1;cc1;cd1;ce1;cf1;cg1;ch1;ci1;cj1;ck1;cl1;cm1;cn1;co1;cp1;cq1;cr1;cs1;ct1;cu1;cv1;cw1;cx1;cy1;cz1;da1;db1;dc1;dd1;de1;df1;dg1;dh1;di1;dj1;dk1;dl1;dm1;dn1;do1;dp1;dq1;dr1;ds1;dt1;du1;dv1;dw1;dx1;dy1;dz1;ea1;eb1;ec1;ed1;ee1;ef1;eg1;eh1;ei1;ej1;ek1;el1;em1;en1;eo1;ep1;eq1;er1;es1;et1;eu1;ev1;ew1;ex1;ey1;ez1;fa1;fb1;fc1;fd1;fe1;ff1;fg1;fh1;fi1;fj1;fk1;fl1;fm1;fn1;fo1;fp1;fq1;fr1;fs1;ft1;fu1;fv1;fw1;fx1;fy1;fz1;ga1;gb1;gc1;gd1;ge1;gf1;gg1;gh1;gi1;gj1;gk1;gl1;gm1;gn1;go1;gp1;gq1;gr1;gs1;gt1;gu1;gv1;gw1;gx1;gy1;gz1;ha1;hb1;hc1;hd1;he1;hf1;hg1;hh1;hi1;hj1;hk1;hl1;hm1;hn1;ho1;hp1;hq1;hr1;hs1;ht1;hu1;hv1;hw1;hx1;hy1;hz1;ia1;ib1;ic1;id1;ie1;if1;ig1;ih1;ii1;ij1;ik1;il1;im1;in1;io1;ip1;iq1;ir1;is1;it1;iu1;iv1;iw1;ix1;iy1;iz1;";
            if (subG == 0)
            {
                sqlEm = "SELECT * FROM Email where id_grupo="+id_grupo+" ";
            }
            else {

                sqlEm = ConexionCall.devuelveValor("SELECT query FROM Sub_grupos where id_subgrupo=" + id_subgrupo);
                DataTable columnas = ConexionCall.SqlDTable("SELECT nombre,columna FROM columna_personalizada  where id_grupo= (SELECT id_padre  FROM Sub_grupos where id_subgrupo="+id_subgrupo+")  order by columna");
                int cols = columnas.Rows.Count;
                if (cols > 0)
                {
                    cabecera = "E-mail";
                    for (int j = 0; j < cols; j++)
                    {
                        cabecera += ";" + columnas.Rows[j]["nombre"].ToString();

                    }
                    string[] columnasdeRelleno = { "a1", "b1", "c1", "d1", "e1", "f1", "g1", "h1", "i1", "j1", "k1", "l1", "m1", "n1", "ñ1", "o1", "p1", "q1", "r1", "s1", "t1", "u1", "v1", "w1", "x1", "y1", "z1", "aa1", "ab1", "ac1", "ad1", "ae1", "af1", "ag1", "ah1", "ai1", "aj1", "ak1", "al1", "am1", "an1", "ao1", "ap1", "aq1", "ar1", "as1", "at1", "au1", "av1", "aw1", "ax1", "ay1", "az1", "ba1", "bb1", "bc1", "bd1", "be1", "bf1", "bg1", "bh1", "bi1", "bj1", "bk1", "bl1", "bm1", "bn1", "bo1", "bp1", "bq1", "br1", "bs1", "bt1", "bu1", "bv1", "bw1", "bx1", "by1", "bz1", "ca1", "cb1", "cc1", "cd1", "ce1", "cf1", "cg1", "ch1", "ci1", "cj1", "ck1", "cl1", "cm1", "cn1", "co1", "cp1", "cq1", "cr1", "cs1", "ct1", "cu1", "cv1", "cw1", "cx1", "cy1", "cz1", "da1", "db1", "dc1", "dd1", "de1", "df1", "dg1", "dh1", "di1", "dj1", "dk1", "dl1", "dm1", "dn1", "do1", "dp1", "dq1", "dr1", "ds1", "dt1", "du1", "dv1", "dw1", "dx1", "dy1", "dz1", "ea1", "eb1", "ec1", "ed1", "ee1", "ef1", "eg1", "eh1", "ei1", "ej1", "ek1", "el1", "em1", "en1", "eo1", "ep1", "eq1", "er1", "es1", "et1", "eu1", "ev1", "ew1", "ex1", "ey1", "ez1", "fa1", "fb1", "fc1", "fd1", "fe1", "ff1", "fg1", "fh1", "fi1", "fj1", "fk1", "fl1", "fm1", "fn1", "fo1", "fp1", "fq1", "fr1", "fs1", "ft1", "fu1", "fv1", "fw1", "fx1", "fy1", "fz1", "ga1", "gb1", "gc1", "gd1", "ge1", "gf1", "gg1", "gh1", "gi1", "gj1", "gk1", "gl1", "gm1", "gn1", "go1", "gp1", "gq1", "gr1", "gs1", "gt1", "gu1", "gv1", "gw1", "gx1", "gy1", "gz1", "ha1", "hb1", "hc1", "hd1", "he1", "hf1", "hg1", "hh1", "hi1", "hj1", "hk1", "hl1", "hm1", "hn1", "ho1", "hp1", "hq1", "hr1", "hs1", "ht1", "hu1", "hv1", "hw1", "hx1", "hy1", "hz1", "ia1", "ib1", "ic1", "id1", "ie1", "if1", "ig1", "ih1", "ii1", "ij1", "ik1", "il1", "im1", "in1", "io1", "ip1", "iq1", "ir1", "is1", "it1", "iu1", "iv1", "iw1", "ix1", "iy1", "iz1" };
                    for (int j = (++cols); j < 260; j++)
                    {
                        cabecera += ";" + columnasdeRelleno[j];

                    }
                }
            }

            sqlEm += " order by id_email ";
            DataTable TablaPrin = ConexionCall.SqlDTable(sqlEm);
            int registros = TablaPrin.Rows.Count;
            StreamWriter writer = new StreamWriter(fileName, true, Encoding.UTF8);
                   
            if (registros > 0)
            {

                try
                {
                    writer.WriteLine(cabecera);
                    string datos = "";
                    for (int i = 0; i < TablaPrin.Rows.Count; i++)
                    {
                        #region
                        datos = TablaPrin.Rows[i]["a1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["b1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["c1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["d1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["e1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["f1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["g1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["h1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["i1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["j1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["k1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["l1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["m1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["n1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ñ1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["o1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["p1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["q1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["r1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["s1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["t1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["u1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["v1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["w1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["x1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["y1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["z1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["aa1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ab1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ac1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ad1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ae1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["af1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ag1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ah1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ai1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["aj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ak1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["al1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["am1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["an1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ao1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ap1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["aq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ar1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["as1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["at1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["au1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["av1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["aw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ax1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ay1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["az1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ba1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["be1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bf1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bi1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bk1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bo1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["br1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bs1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bt1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["by1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["bz1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ca1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ce1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cf1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ch1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ci1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ck1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["co1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cr1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cs1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ct1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["cz1"].ToString();


                        datos += ";" + TablaPrin.Rows[i]["da1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["db1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["de1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["df1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["di1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dk1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["do1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dr1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ds1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dt1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["du1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["dz1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ea1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ec1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ed1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ee1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ef1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ei1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ej1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ek1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["el1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["em1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["en1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eo1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ep1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["er1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["es1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["et1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["eu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ev1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ew1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ex1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ey1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ez1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["fa1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fe1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ff1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fi1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fk1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fo1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fr1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fs1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ft1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["fz1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ga1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ge1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gf1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gi1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gk1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["go1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gr1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gs1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gt1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["gz1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ha1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hb1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hc1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hd1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["he1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hf1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hg1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hh1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hi1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hj1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hk1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hl1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hm1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hn1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ho1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hp1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hr1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hs1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ht1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hx1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["hz1"].ToString();

                        datos += ";" + TablaPrin.Rows[i]["ia1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ib1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ic1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["id1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ie1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["if1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ig1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ih1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ii1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ij1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ik1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["il1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["im1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["in1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["io1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ip1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iq1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ir1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["is1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["it1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iu1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iv1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iw1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["ix1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iy1"].ToString();
                        datos += ";" + TablaPrin.Rows[i]["iz1"].ToString();





                        #endregion
                        writer.WriteLine(datos);
                        if (i != 0 && i % 5000 == 0)
                        {
                         //   Thread.Sleep(20 * 1000);
                        }

                    }
                    resultado=true;
                }
                catch (Exception ex)
                {
                    resultado=false;
                }

            }
            else
            {
                writer.WriteLine("Consulta no produjo resultados ");
                resultado=true;
            }

            writer.Dispose();
            writer.Close();

            return resultado;
        }

//        protected void fechasColumna(object sender, EventArgs e)
//        {
//            hilo14.Stop();
//            try
//            {
//                DateTime Hoy = DateTime.Now;
//                if (Hoy.Hour == 1)
//                {
//                    #region
//                    DateTime mañana = DateTime.Now.AddDays(1);
//                    int dia = mañana.Day;
//                    int mes = mañana.Month;

//                    DataTable columnas = ConexionCall.SqlDTable("SELECT id_grupo,columna FROM fechas");
//                    int reg = columnas.Rows.Count;
//                    if (reg > 0)
//                    {
//                        string id_grupo, columna, sqlCumple, ids, inserta, sqlValida;
//                        int validador = 0;
//                        for (int i = 0; i < reg; i++)
//                        {
//                            id_grupo = columnas.Rows[i]["id_grupo"].ToString();
//                            columna = columnas.Rows[i]["columna"].ToString();

/////agregado para cumpleaños
////lo hice asi para ir paso a paso, es comprimible el codigo

//                            //si es un subgrupo
//                       string qg = " SELECT id_subgrupo   FROM grupo where id_grupo=" + id_grupo; 

//                       int sg = ConexionCall.devuelveValorINT(qg);


////si el id_subgrupo es 0 entonces es un grupo padre, si es distinto de 0 hay q sacar la query del subgrupo 
//                       string qsg = "", wheresg="";

//                          if (sg != 0)
//                          {
//                          qsg = " SELECT query   FROM Sub_grupos where id_subgrupo="+sg;
//                          wheresg = ConexionCall.devuelveValor(qsg);
//                          }

//                          if (wheresg != "")
//                          {
//                              //por formato sabemos q query trae todo usando *, pero solo necesitamos id_email
//                              wheresg = wheresg.Replace("*", "id_email");
//                              wheresg += " and " + columna + " !='' ";
//                              wheresg += "   order by id_email";
//                          }

//                            sqlCumple = " SELECT id_email   FROM Email where id_grupo=" + id_grupo;
//                            sqlCumple += " and " + columna + " !='' ";
//                            sqlCumple += "   order by id_email";

//                            //si qsg tiene datos hay que llenar el dt con el en vez de con sqlcumple
//                            if (wheresg != "") sqlCumple = wheresg;

//                            DataTable cumple =ConexionCall.SqlDTable(sqlCumple);

///// fin agregado para cumpleaños



//                            int n_cumple = cumple.Rows.Count;
//                            if (n_cumple > 0)
//                            {
//                                ConexionCall objIns = new ConexionCall();
//                                objIns.ejecutorBase("delete cumpleanios where dia=" + dia + " and mes=" + mes + " and id_grupo= " + id_grupo);
//                                ids = "0";


//                                for (int j = 0; j < n_cumple; j++)
//                                {
//                                    sqlValida = "SELECT id_email FROM Email where id_email=" + cumple.Rows[j]["id_email"].ToString();
//                                    sqlValida += " and " + columna + " !='' and month(convert(datetime," + columna + ", 21))=" + mes + " and day(convert(datetime," + columna + ", 21))=" + dia;

//                                    validador = ConexionCall.devuelveValorINT(sqlValida);

//                                    if (validador > 0)
//                                    {
//                                        ids += "," + cumple.Rows[j]["id_email"].ToString();

//                                        this.Invoke(new DisplayEstado(Progreso), cumple.Rows[j]["id_email"].ToString() + " tiene aniversario el dia " + dia + " del  mes " + mes);
//                                    }

//                                }

//                                inserta = "INSERT INTO cumpleanios (dia,mes,id_grupo  ,id_mail)  VALUES ";
//                                inserta += "  (" + dia + " ," + mes + ", " + id_grupo + " ,'" + ids + "')";
//                                objIns.ejecutorBase(inserta);
//                            }

//                        }
//                    }
//                    #endregion
//                }

//                Thread.Sleep(3600 * 1000);
//            }
//            catch (Exception ex) { }
//            hilo14.Enabled = true;
//        }


        private void Procesa_Union_bases()
        {

            ConexionCall _CON = new ConexionCall();

            string SQLSelect = "Select ID_GRUPO_FINAL from Unir_Bases where ESTADO = 0 group by ID_GRUPO_FINAL";
            DataTable Pendientes = ConexionCall.SqlDTable(SQLSelect);

            for (int i = 0; i < Pendientes.Rows.Count; i++)
            {
                string _ID_FINAL = Pendientes.Rows[i][0].ToString();

                string SQLSELECT2 = "Select  DISTINCT a1 from email where activo = 1 and id_grupo in (Select ID_GRUPO from Unir_Bases where ID_GRUPO_FINAL = " + _ID_FINAL + ")";
                DataTable TBL_Email = ConexionCall.SqlDTable(SQLSELECT2);
                int Registros = TBL_Email.Rows.Count;

                string SQlCliente = "Select id_cliente from Usuario where Id_Usuario  = (select id_usuario from Grupo where id_grupo =" + _ID_FINAL + ")";
                int _ID_CLIENTE = ConexionCall.devuelveValorINT(SQlCliente);

                string strSQL = " select id_temp from temporales where id_grupo = "+_ID_FINAL+" and estado = 1";
                int _ID_TEMP = ConexionCall.devuelveValorINT(strSQL);

                this.Invoke(new DisplayEstado(Progreso), "Comienzo de Unión de Base de datos " + Registros + "");

                string SQLUPDATE0 = "UPDATE Unir_Bases set ESTADO = 1 where ID_GRUPO_FINAL= " + _ID_FINAL;
                _CON.ejecutorBase(SQLUPDATE0);

                for (int k = 0; k < Registros; k++)
                {
                   
                    string SQLINSERT = "Insert Into Email ([Id_Grupo],[Activo],[a1]) ";
                    SQLINSERT += " values (" + _ID_FINAL +"";
                    SQLINSERT += ", 1 ";
                    SQLINSERT += ",'" + TBL_Email.Rows[k]["a1"].ToString().Trim() + "'";
                    SQLINSERT += ")";

                    _CON.ejecutorBase(SQLINSERT);

                    this.Invoke(new DisplayEstado(Progreso), "Unir Base de datos "+Registros+"/"+(k+1));

                    if (k % 100 == 0 && k != 0)
                    {
                        int porcentaje = Convert.ToInt32((100 * k) / Registros);
                        _CON.ejecutorBase("update temporales set porc=" + porcentaje + " where  id_temp=" + _ID_TEMP);
                    }

                }

                string SQLUPDATE = "UPDATE Unir_Bases set ESTADO = 2 where ID_GRUPO_FINAL= " + _ID_FINAL;
                _CON.ejecutorBase(SQLUPDATE);

                string SQLUPDATE2 = "UPDATE temporales set estado = 2 where id_temp= " + _ID_TEMP;
                _CON.ejecutorBase(SQLUPDATE2);

            }

        }

        private void Procesa_Envio_cargabase()
        {

            ConexionCall _CON = new ConexionCall();

            string SQlSelect = "Select m.Id_Grupo,id_estado,id_mensaje from Mensaje m, temporales t where m.Id_Grupo=t.id_grupo and Id_Estado in (13,14) and t.estado = 2";
            DataTable TBL_mensajes = ConexionCall.SqlDTable(SQlSelect);

            this.Invoke(new DisplayEstado(Progreso), "Comienzo proceso envío Cargar pendientes");

            for (int i = 0; i < TBL_mensajes.Rows.Count; i++)
            {
                string _ID_GRUPO = TBL_mensajes.Rows[i]["id_grupo"].ToString();
                string _ID_ESTADO = TBL_mensajes.Rows[i]["id_estado"].ToString();
                string _ID_MENSAJE = TBL_mensajes.Rows[i]["id_mensaje"].ToString();
               
                    int _NuevoEstado = 0;

                    if (Convert.ToInt32(_ID_ESTADO) == 13)
                    {
                        _NuevoEstado = 1;
                    }

                    if (Convert.ToInt32(_ID_ESTADO) == 14)
                    {
                        _NuevoEstado = 7;
                    }

                    if (_NuevoEstado != 0)
                    {
                        string SqlUpdate = "update Mensaje set Id_Estado = " + _NuevoEstado + " where Id_Mensaje = " + _ID_MENSAJE + " and Id_estado = " + _ID_ESTADO;
                        _CON.ejecutorBase(SqlUpdate);

                        this.Invoke(new DisplayEstado(Progreso), "Estado de campaña actualizado ("+_ID_MENSAJE+")");
                    }

            }

            this.Invoke(new DisplayEstado(Progreso), "Fin proceso envío Cargar pendientes");

        
        }


        private void ordenaServidores(object sender, EventArgs e)
        {
            hilo5.Stop();
            try
            {
                DataTable ser = null;
                DataTable mnesajes = ConexionCall.SqlDTable("SELECT Id_Mensaje  FROM Mensaje  where Id_Estado in (1,9) and id_ser=0 ");
                if (mnesajes.Rows.Count > 0)
                {
                    for (int i = 0; i < mnesajes.Rows.Count; i++)
                    {
                        string id_mensaje = mnesajes.Rows[i]["Id_Mensaje"].ToString();
                        #region
                        string inserta;
                        string sqlS = "select top 1 * from Servidores ";
                        sqlS += " where id_ser not in ( ";
                        sqlS += " SELECT id_ser FROM log_procesos where estado=0 group by id_ser ";
                        sqlS += " )order by importancia ";

                        ser = ConexionCall.SqlDTable(sqlS);
                        ConexionCall obj = new ConexionCall();
                        if (ser.Rows.Count > 0)
                        {
                            #region
                            string id_ser = ser.Rows[0]["id_ser"].ToString();
                            inserta = "INSERT INTO log_procesos (id_mensaje,id_ser,estado)  VALUES ";
                            inserta += "  (" + id_mensaje + " ," + id_ser + ",0)";

                            if (obj.ejecutorBase(inserta))
                            {
                                obj.ejecutorBase("UPDATE Mensaje SET id_ser =" + id_ser + " WHERE id_mensaje=" + id_mensaje);
                            }
                            #endregion
                        }
                        else
                        {
                            #region
                            sqlS = "select top 1 * from Servidores ";
                            sqlS += " where id_ser = ( ";
                            sqlS += " select top 1 v.id_ser ";
                            sqlS += " from Vista_logServidores v ";
                            sqlS += " join Servidores s on s.id_ser=v.id_ser ";
                            sqlS += " order by veces,importancia ";
                            sqlS += " )order by importancia ";
                            ser = ConexionCall.SqlDTable(sqlS);
                            string id_ser = ser.Rows[0]["id_ser"].ToString();
                            inserta = "INSERT INTO log_procesos (id_mensaje,id_ser,estado)  VALUES ";
                            inserta += "  (" + id_mensaje + " ," + id_ser + ",0)";

                            if (obj.ejecutorBase(inserta))
                            {
                                obj.ejecutorBase("UPDATE Mensaje SET id_ser =" + id_ser + " WHERE id_mensaje=" + id_mensaje);
                            }
                            #endregion
                        }

                        #endregion
                    }
                }
                Thread.Sleep(60 * 1000);
            }
            catch (Exception es) { }
            hilo5.Enabled = true;
        }

        private void btn_borra_Click(object sender, EventArgs e)
        {
            btn_borra.Enabled = false;
            EliminaEmail_Envios();
            btn_borra.Enabled = true;
        }

        public void EliminaEmail_Envios()
        {
            try
            {

                int contador = 0;
                ConexionCall objBorra = new ConexionCall();
                string sqlFijo = "SELECT id_mensaje,id_grupo,fecha,ejecucion,id_usuario ";
                sqlFijo += " from mensaje where DATEDIFF(month, ejecucion, getdate())>=3 ";
                sqlFijo += " and id_mensaje not in (SELECT id_mensaje FROM Datos_borrados) order by ejecucion";
               #region fijo
                 DataTable tbmenGrup = ConexionCall.SqlDTable(sqlFijo);
               int reg = tbmenGrup.Rows.Count;
               if (reg > 0)
                {
                    string id_mensaje, id_grupo, fecha, ejecucion, id_usuario, sqInsert;
                    int n_enviados = 0;

                    for (int i = 0; i < reg; i++)
                    {
                        #region
                        id_mensaje = tbmenGrup.Rows[i]["id_mensaje"].ToString();
                        id_grupo = tbmenGrup.Rows[i]["id_grupo"].ToString();
                        fecha = tbmenGrup.Rows[i]["fecha"].ToString();
                        ejecucion = tbmenGrup.Rows[i]["ejecucion"].ToString();
                        id_usuario = tbmenGrup.Rows[i]["id_usuario"].ToString();
                        #endregion
                        #region
                        n_enviados = ConexionCall.devuelveValorINT("select count(*) n_enviados from envio_correo where id_mensaje=" + id_mensaje);
                        if (n_enviados > 0)
                        {
                            Progreso("Eliminando datos de mensaje " + id_mensaje);
                            sqInsert = "INSERT INTO Datos_borrados (id_mensaje,id_grupo ,id_usuario ,inicio ,ejecuacion,borrado ,n_email ,n_enviados) ";
                            sqInsert += "  VALUES ";
                            sqInsert += " ( " + id_mensaje;
                            sqInsert += " ," + id_grupo;
                            sqInsert += " ," + id_usuario;
                            sqInsert += " ,'" + fecha + "' ";
                            sqInsert += " ,'" + ejecucion + "' ";
                            sqInsert += " , getdate() ";
                            sqInsert += " ,0";
                            sqInsert += " ," + n_enviados;
                            sqInsert += " )";
                            objBorra.ejecutorBase(sqInsert);
                            objBorra.ejecutorBase("delete from envio_correo where id_mensaje=" + id_mensaje);// borra enviados
                            contador++;
                        }
                        else
                        {
                            Progreso("Mensaje " + id_mensaje + " no tiene envios");
                        }
                        #endregion
                    }
                    Progreso("Los mensajes con datos que superen los 3 meses fueron limpiados correctamente ");
                }
                else
                {
                    Progreso("No hay datos que superen los 3 meses ");
                }
                #endregion
              /// automatico
                contador = 0;
                #region automatico
                string sqlAutomatico = "select id_mensaje,id_grupo,fecha,ejecucion,id_usuario ";
                sqlAutomatico += " from Mensaje where estatico=1";

                DataTable tbmenAuto = ConexionCall.SqlDTable(sqlAutomatico);
                int regAuto = tbmenAuto.Rows.Count;
                if (regAuto > 0)
                {

                    string id_mensaje, id_grupo, fecha, ejecucion, id_usuario, sqInsert;
                    int n_enviados = 0;
                    for (int j = 0; j < regAuto; j++)
                    {
                        #region
                        id_mensaje = tbmenAuto.Rows[j]["id_mensaje"].ToString();
                        id_grupo = tbmenAuto.Rows[j]["id_grupo"].ToString();
                        fecha = tbmenAuto.Rows[j]["fecha"].ToString();
                        ejecucion = tbmenAuto.Rows[j]["ejecucion"].ToString();
                        id_usuario = tbmenAuto.Rows[j]["id_usuario"].ToString();
                        #endregion
                        #region
                        n_enviados = ConexionCall.devuelveValorINT("select count(*) n_enviados from envio_correo where id_mensaje=" + id_mensaje + " and DATEDIFF(month, fecha, getdate())>=3");

                        if (n_enviados > 0)
                        {
                            Progreso("Eliminando datos de mensaje " + id_mensaje);
                            sqInsert = "INSERT INTO Datos_borrados (id_mensaje,id_grupo ,id_usuario ,inicio ,ejecuacion,borrado ,n_email ,n_enviados) ";
                            sqInsert += "  VALUES ";
                            sqInsert += " ( " + id_mensaje;
                            sqInsert += " ," + id_grupo;
                            sqInsert += " ," + id_usuario;
                            sqInsert += " ,'" + fecha + "' ";
                            sqInsert += " ,'" + ejecucion + "' ";
                            sqInsert += " , getdate() ";
                            sqInsert += " ,0";
                            sqInsert += " ," + n_enviados;
                            sqInsert += " )";
                            objBorra.ejecutorBase(sqInsert);
                            objBorra.ejecutorBase("delete from envio_correo where id_mensaje=" + id_mensaje + " and DATEDIFF(month, fecha, getdate())>=3)");// borra enviados

                        }
                        else
                        {
                            Progreso("Mensaje Automatico " + id_mensaje + " no tiene envios");
                        }
                        #endregion
                    }
                    Progreso("Los mensajes automaticos con datos que superen los 3 meses fueron limpiados correctamente ");
                }
                #endregion

            }
            catch (Exception ex)
            {
                Progreso("Error :" + ex.Message);
            }
        }





        public string arreglacaracter(string palabra)
        {


            string palabra2 =String.Empty;


            palabra2 = palabra.Replace("Ã¡", "á");
            palabra2 = palabra2.Replace("Ã©", "é");
            palabra2 = palabra2.Replace("Ã­", "í");
            palabra2 = palabra2.Replace("Ã³", "ó");
            palabra2 = palabra2.Replace("Ãº", "ú");
            palabra2 = palabra2.Replace("Ã±", "ñ");

            palabra2 = palabra2.Replace("ß", "á");
            palabra2 = palabra2.Replace("±", "ñ");

            palabra2 = palabra2.Replace("Ð", "Ñ");
            palabra2 = palabra2.Replace("┴", "Á");
            palabra2 = palabra2.Replace("Ý", "í");


            palabra2 = palabra2.Replace("Ã¿", "Í");




            palabra2 = palabra2.Replace("▄", "Ü");
            



           // palabra2 = palabra2.Replace("", "");
            //palabra2 = palabra2.Replace("", ""); 

            return palabra2;

            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        
    }
}