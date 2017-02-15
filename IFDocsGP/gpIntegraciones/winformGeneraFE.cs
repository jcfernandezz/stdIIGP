using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

using System.IO;
using System.Threading;

using Comun;
using gp.ModeloDeDatos;
using MyGeneration.dOOdads;
using OfficeOpenXml;
using IntegradorDeGP;
using ManipulaArchivos;
using System.Threading.Tasks;

namespace gp.InterfacesPersonalizadas
{
    public partial class winformGeneraFE : Form
    {
        string ultimoMensaje = "";
        private short idxNomArchivo = 1;                                                        //columna Nombre del archivo

        private ConexionDB DatosConexionDB = new ConexionDB(Environment.GetCommandLineArgs());  //Lee la configuración del archivo xml y obtiene los datos de conexión.
        Parametros Compannia = null;
        private static FileSystemWatcher watcher = new FileSystemWatcher();
        delegate void actualizaListaETCallback(int i, string carpeta);
        delegate void actualizaListaFNCallback(int i, string carpeta);

        public winformGeneraFE()
        {
            InitializeComponent();
            try
            {
                this.Text = "Integra documentos GP - " + DatosConexionDB.NombreArchivoParametros.Replace(".xml", "").Substring(14).ToUpper();
                this.tsButtonGenerar.ToolTipText = "Procesar " + DatosConexionDB.NombreArchivoParametros.Replace(".xml", "").Substring(14).ToUpper();
            }
            catch (Exception ex)
            {
                txtbxMensajes.Text = "No se puede ingresar a la aplicación. Edite su acceso directo e incluya el nombre del archivo de parámetros. " + ex.Message;
                HabilitarVentana(false, false, false, false, false, true);
            }
        }

        private void winformGeneraFE_Load(object sender, EventArgs e)
        {

            //if (!cargaCompannias(!DatosConexionDB.Elemento.IntegratedSecurity, DatosConexionDB.Elemento.Intercompany))
            if (!cargaCompanniasDeArchivoParam())
            {
                txtbxMensajes.Text = ultimoMensaje;
                HabilitarVentana(false,false,false,false,false, true);
            }
            //dtPickerDesde.Value = DateTime.Now;
            //dtPickerHasta.Value = DateTime.Now;
            lblFecha.Text = DateTime.Now.ToString();
        }

        /// <summary>
        /// Aplica los criterios de filtro, actualiza la pantalla e inicializa los checkboxes del grid.
        /// </summary>
        /// <param name=""></param>
        /// <returns>bool</returns>
        private bool AplicaFiltroYActualizaPantalla()
        {
            txtbxMensajes.AppendText("...\r\n");
            txtbxMensajes.Refresh();

            Compannia = new Parametros(DatosConexionDB.NombreArchivoParametros, DatosConexionDB.Elemento.Intercompany);
            txtbxMensajes.AppendText(Compannia.ultimoMensaje);
            txtbxMensajes.Refresh();

            if (Compannia.iError != 0)
                return false;

            actualizaListas(0, Compannia.rutaCarpeta);

            // Monitorea la carpeta EnTrabajo
            watcher.Path = Compannia.rutaCarpeta +"\\EnTrabajo";
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            watcher.Filter = "*.xlsx";

            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.Created += new FileSystemEventHandler(OnChanged);
            watcher.Deleted += new FileSystemEventHandler(OnChanged);
            watcher.Renamed += new RenamedEventHandler(OnRenamed);

            watcher.EnableRaisingEvents = true;

            return true;
        }

        /// <summary>
        /// Evento del monitoreo de la carpeta EnTrabajo
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            actualizaListaEnTrabajo(0, Compannia.rutaCarpeta + "\\EnTrabajo"); 
            actualizaListaFinalizados(0, Compannia.rutaCarpeta + @"\Finalizado\"); 
        }

        /// <summary>
        /// Evento del monitoreo de la carpeta EnTrabajo
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnRenamed(object source, RenamedEventArgs e)
        {
            actualizaListaEnTrabajo(0, Compannia.rutaCarpeta + "\\EnTrabajo");
            actualizaListaFinalizados(0, Compannia.rutaCarpeta + @"\Finalizado\");
        }

        private bool cargaCompanniasDeArchivoParam()
        {
            
            try
            {
                    //Ocasiona que se dispare el trigger textChanged del combo box
                    cmbBxCompannia.DisplayMember = "Nombre";
                    cmbBxCompannia.ValueMember = "Id";
                    cmbBxCompannia.DataSource = DatosConexionDB.ListaCompannias;
                    return true;
            }
            catch (Exception eCia)
            {
                ultimoMensaje = "Contacte al administrador. No están configuradas las compañías. [CargaCompannias] " + DatosConexionDB.ultimoMensaje + " - " + eCia.Message;
            }
            return false;
        }

        private bool cargaCompannias(bool Filtro, string Unica)
        {
            //vwIfcCompannias Compannias = new vwIfcCompannias(DatosConexionDB.Elemento.ConnStrDyn);
            //try
            //{
            //    if (Compannias.Query.Load())
            //    {
            //        //Ocasiona que se dispare el trigger textChanged del combo box
            //        cmbBxCompannia.DisplayMember = vwIfcCompannias.ColumnNames.CMPNYNAM;
            //        cmbBxCompannia.ValueMember = vwIfcCompannias.ColumnNames.INTERID;
            //        cmbBxCompannia.DataSource = Compannias.DefaultView;
            //        return true;
            //    }
            //    else
            //        ultimoMensaje = "No tiene acceso a ninguna compañía. Revise los privilegios otorgados a su usuario. [cargaCompannias]";
            //}
            //catch (Exception eCia)
            //{
            //    ultimoMensaje = "Contacte al administrador. No se puede acceder a la base de datos. [CargaCompannias] " + DatosConexionDB.ultimoMensaje + " - " + eCia.Message;
            //}
            return false;
        }

        private void HabilitarVentana(bool emite, bool anula, bool imprime, bool publica, bool envia, bool cambiaCia)
        {
            cmbBxCompannia.Enabled = cambiaCia;
            tsButtonGenerar.Enabled = emite;      //Emite xml
            //toolStripPDF.Enabled = imprime;       //Imprime
            //toolStripImpresion.Enabled = imprime; //Imprime
            //toolStripEmail.Enabled = envia;       //Envía emails
            //toolStripEmailMas.Enabled = envia;

            toolStripConsulta.Enabled = emite || anula || imprime || publica || envia;
            //btnBuscar.Enabled = emite || anula || imprime || publica || envia;
        }

        private void ReActualizaDatosDeVentana()
        {
            DatosConexionDB.Elemento.Compannia = cmbBxCompannia.Text.ToString().Trim();
            DatosConexionDB.Elemento.Intercompany = cmbBxCompannia.SelectedValue.ToString().Trim();
            lblUsuario.Text = DatosConexionDB.Elemento.Usuario;
            ToolTip tTipCompannia = new ToolTip();
            tTipCompannia.AutoPopDelay = 5000;
            tTipCompannia.InitialDelay = 1000;
            tTipCompannia.UseFading = true;
            tTipCompannia.Show("Está conectado a: " + DatosConexionDB.Elemento.Compannia, this.cmbBxCompannia, 5000);

            txtbxMensajes.Text = "";
            //if (!cargaIdDocumento())
            //{
            //    txtbxMensajes.AppendText(ultimoMensaje);
            //    HabilitarVentana(false,false,false,false,false, true);
            //}

            Parametros configCfd = new Parametros(DatosConexionDB.NombreArchivoParametros, DatosConexionDB.Elemento.Intercompany);   //Carga configuración desde xml
            //estadoCompletadoCia = configCfd.intEstadoCompletado;

            if (configCfd.iError != 0)
            {
                txtbxMensajes.AppendText(configCfd.ultimoMensaje);
                HabilitarVentana(false,false,false,false,false, true);
                return;
            }

            HabilitarVentana(true, false, false, false, false, true);
            AplicaFiltroYActualizaPantalla();
        }

        private void actualizaListaEnTrabajo(int i, string carpeta)
        {
            try
            {
                DirectoryInfo enTrabajoDir = new DirectoryInfo(carpeta);       
                archivosExcel archivosEnTrabajo = new archivosExcel();
                DataTable listaArchivos = archivosEnTrabajo.getArchivosDeCarpeta(enTrabajoDir);

                //thread safe call to a windows form control
                if (this.dgvEnTrabajo.InvokeRequired)
                {
                    actualizaListaETCallback d = new actualizaListaETCallback(actualizaListaEnTrabajo);
                    this.Invoke(d, new object[] { i, carpeta });
                }
                else
                {
                    dgvEnTrabajo.DataSource = listaArchivos;
                    dgvEnTrabajo.Refresh();
                }
            }
            catch (Exception errGral)
            {
                txtbxMensajes.AppendText("Excepción encontrada al leer archivos En Trabajo. " + errGral.Message + " [actualizaListaEnTrabajo]" + "\r\n");
                txtbxMensajes.Refresh();
            }
        }

        private void actualizaListaFinalizados(int i, string carpeta)
        {
            try
            {
                string[] archivosFinalizados = Directory.GetFiles(carpeta, "*.xlsx");

                //thread safe call to a windows form control
                if (this.liBxFinalizado.InvokeRequired)
                {
                    actualizaListaFNCallback d = new actualizaListaFNCallback(actualizaListaFinalizados);
                    this.Invoke(d, new object[] { i, carpeta });
                }
                else
                {
                    liBxFinalizado.DataSource = archivosFinalizados;
                    liBxFinalizado.Refresh();
                }
            }
            catch (Exception errGral)
            {
                txtbxMensajes.AppendText("Excepción encontrada al leer los archivos Finalizados. " + errGral.Message + " [actualizaListaFinalizados]" + "\r\n");
                txtbxMensajes.Refresh();
            }
        }

        private void actualizaListas(int i, string carpeta)
        {
            try
            {
                actualizaListaEnTrabajo(i, carpeta + "\\EnTrabajo");
                actualizaListaFinalizados(i, carpeta + @"\Finalizado\");
            }
            catch (Exception errGral)
            {
                txtbxMensajes.AppendText("Excepción encontrada al leer los archivos. " + errGral.Message + " [actualizaListas]" + "\r\n");
                txtbxMensajes.Refresh();
            }
        }

        void reportaProgreso(int i, string s)
        {
            progressBar1.Increment(i);
            progressBar1.Refresh();

            if (progressBar1.Value == progressBar1.Maximum)
                progressBar1.Value = 0;

            txtbxMensajes.AppendText(s + "\r\n");
            txtbxMensajes.Refresh();
        }

        /// <summary>
        /// Procesa la carga de cualquier documento de compras. El nombre del archivo de parámetros está en DatosConexionDB
        /// </summary>
        void ProcesaComprasDeAcuerdoAParametros()   //object sender, DoWorkEventArgs doWorkEventArgs)
        {
            IntegraComprasGP ic = new IntegraComprasGP(DatosConexionDB);
            if (ic.iError == 0)
            {
                ic.Progreso += new IntegraComprasGP.LogHandler(reportaProgreso);    //suscribe a reporte de progreso
                ic.Actualiza += new IntegraComprasGP.LogHandler(actualizaListas);
                
                List<string> archivosSeleccionados = new List<string>();
                for (int r = 0; r < dgvEnTrabajo.RowCount; r++)
                    archivosSeleccionados.Add(dgvEnTrabajo[idxNomArchivo, r].Value.ToString());

                ic.procesaCarpetaEnTrabajo(archivosSeleccionados);
            }
        }

        void ProcesaFacturasVenta()   //object sender, DoWorkEventArgs doWorkEventArgs)
        {
            IntegraVentasGP ve = new IntegraVentasGP(DatosConexionDB);
            if (ve.IError == 0)
            {
                ve.Progreso += new IntegraVentasGP.LogHandler(reportaProgreso);    //suscribe a reporte de progreso
                //ve.Actualiza += new Action<int, string>(actualizaListas);
                ve.Actualiza += new IntegraVentasGP.LogHandler(actualizaListas);

                List<string> archivosSeleccionados = new List<string>();
                for (int r = 0; r < dgvEnTrabajo.RowCount; r++)
                    archivosSeleccionados.Add(dgvEnTrabajo[idxNomArchivo, r].Value.ToString());

                ve.ProcesaCarpetaEnTrabajo(archivosSeleccionados);
            }
        }

        /// <summary>
        /// Inicia proceso de carga de documentos GP
        /// </summary>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            txtbxMensajes.Text = "";
            AplicaFiltroYActualizaPantalla();
            HabilitarVentana(false, false, false, false, false, false);
            watcher.EnableRaisingEvents = false;

            ProcesaComprasDeAcuerdoAParametros();   //Procesa la carga de cualquier doc de compras

            //ProcesaFacturasVenta();

            HabilitarVentana(true, true, true, true, true, true);
            AplicaFiltroYActualizaPantalla();
            progressBar1.Value = 0;
            
            //pBarProcesoActivo.Show();
            //BackgroundWorker _bw = new BackgroundWorker();
            //_bw.DoWork += procesaFacturasPOP;
            //_bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_Completed);
            //_bw.ProgressChanged += new ProgressChangedEventHandler(bw_Progress);
            //object[] arguments = { 0 };     //enviar argumento
            //_bw.RunWorkerAsync();
        }

        void bw_Progress(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                progressBar1.Increment(e.ProgressPercentage - progressBar1.Value);
                progressBar1.Refresh();

                txtbxMensajes.AppendText(e.UserState.ToString() + "\r\n");
                txtbxMensajes.Refresh();
            }
            catch (Exception ePr)
            {
                txtbxMensajes.AppendText("bw Progress: " + ePr.Message);
            }
        }

        void bw_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
            if (e.Cancelled)
                progressBar1.Value = 0;
            else if (e.Error != null)
                txtbxMensajes.AppendText("[bw_Completed] " + e.Error.ToString());
            else
                txtbxMensajes.AppendText(e.Result.ToString());

            //Actualiza la pantalla
            HabilitarVentana(true, true, true, true, true, true);
            AplicaFiltroYActualizaPantalla();
            progressBar1.Value = 0;
            //pBarProcesoActivo.Hide();

            }
            catch (Exception eCm)
            {
                txtbxMensajes.AppendText("bw Completed: " + eCm.Message);
            }

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tsDDButtonFiltroF_TextChanged(object sender, EventArgs e)
        {
            txtbxMensajes.Text = "";
            AplicaFiltroYActualizaPantalla();
        }

        public void GuardaArchivoMensual()
        {
            try
            {

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                // Default file extension
                saveFileDialog1.DefaultExt = "txt";
                saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.Title = "Dónde desea guardar el Informe mensual?";
                saveFileDialog1.InitialDirectory = @"C:/";
                //saveFileDialog1.FileName = "1" + dtPickerDesde.Value.Month.ToString().PadLeft(2, '0') + dtPickerDesde.Value.Year.ToString() + ".TXT";

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Stream stm = new FileStream(saveFileDialog1.FileName, FileMode.Create);
                    TextWriter tw = new StreamWriter(stm);
                    //int i = 0;
                    progressBar1.Value = 0;
                    //infMes.Rewind();          //move to first record
                    //do
                    //{
                    //    tw.WriteLine(infMes.ComprobanteEmitido);
                    //    tw.Flush(); // Ensure the TextWriter buffer is empty

                    //    txtbxMensajes.AppendText("Doc:" + infMes.Sopnumbe + "\r\n");
                    //    txtbxMensajes.Refresh();
                    //    progressBar1.Value = i * 100 / infMes.RowCount;
                    //    i++;
                    //} while (infMes.MoveNext());
                    progressBar1.Value = 0;

                    stm.Close();
                    ultimoMensaje = "El informe mensual fue almacenado satisfactoriamente en: " + saveFileDialog1.FileName;
                }
                else
                {
                    ultimoMensaje = "Operación cancelada a pedido del usuario.";
                }
            }
            catch (Exception eFile)
            {
                ultimoMensaje = "Error al almacenar el archivo. " + eFile.Message;
            }

        }

        private void cmbBxCompannia_TextChanged(object sender, EventArgs e)
        {
            ReActualizaDatosDeVentana();
        }

        private void cmbBxCompannia_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReActualizaDatosDeVentana();
        }

        private void tsButtonActualiza_Click(object sender, EventArgs e)
        {
            AplicaFiltroYActualizaPantalla();
        }

        private void AbrirArchivo()
        {
            try
            {
                archivosExcel.AbrirArchivo(liBxFinalizado.SelectedItem.ToString());

            }
            catch (IOException io)
            {
                txtbxMensajes.AppendText(io.Message + " " + io.TargetSite.ToString() + "\r\n");
                txtbxMensajes.Refresh();
            }
            catch (Exception ex)
            {
                txtbxMensajes.AppendText("Excepción desconocida al abrir el archivo. " + ex.Message + " " + ex.TargetSite.ToString() + "\r\n");
                txtbxMensajes.Refresh();
            }
        }

        private void liBxFinalizado_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AbrirArchivo();
        }

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirArchivo();
        }

        private void moverToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                archivosExcel.MueveATrabajo(liBxFinalizado.SelectedItem.ToString(), Compannia.rutaCarpeta);

            }
            catch (IOException io)
            {
                txtbxMensajes.AppendText(io.Message + " " + io.TargetSite.ToString() + "\r\n");
                txtbxMensajes.Refresh();
            }
            catch (Exception ex)
            {
                txtbxMensajes.AppendText("Excepción desconocida al mover el archivo. Muévalo manualmente. " + ex.Message + " " + ex.TargetSite.ToString() + "\r\n");
                txtbxMensajes.Refresh();
            }

        }
    }
}
