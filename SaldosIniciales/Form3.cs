using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LibreriaDoctos;
using System.IO;


namespace SaldosIniciales
{
    public partial class Form3 : Form
    {
        ClassRNLOB lrn = new ClassRNLOB();

        public string Cadenaconexion = "";
        public Form3()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (lrn.mValidaSQLConexion(txtServer.Text, txtBD.Text, txtUser.Text, txtPass.Text) == 1)
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();

                this.Close();
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                mllenarcomboempresas();
                //y.Visible = true;
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }
        public void mllenarcomboempresas()
        {
            ciCompanyList11.Populate(Cadenaconexion);
            ciCompanyList12.Populate(Cadenaconexion);
        }

        private int mcargarEmpresa(ComboBox comboBox1)
        {

            string mensaje = "";
            comboBox1.Items.Clear();
            comboBox1.DataSource = lrn.mCargarEmpresas(out mensaje);

            comboBox1.DisplayMember = "Nombre";
            comboBox1.ValueMember = "Ruta";
            comboBox1.Update();
            try
            {
                comboBox1.SelectedIndex = -1;
            }
            catch (Exception ee)
            {
            }
            comboBox1.SelectedIndex = 0;
            return 0;
            //}
            //else
            //   MessageBox.Show (mensaje);
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());
            //mcargarEmpresa(cbOrigen);

            txtServer.Text = Properties.Settings.Default.server;
            txtBD.Text = Properties.Settings.Default.database;
            txtUser.Text = Properties.Settings.Default.user;
            txtPass.Text = Properties.Settings.Default.password;

            Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";

            if (lrn.mValidaSQLConexion(txtServer.Text, txtBD.Text, txtUser.Text, txtPass.Text) == 1)
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();

                //this.Close();
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                mllenarcomboempresas();
                //y.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (ciCompanyList11.aliasbdd == ciCompanyList12.aliasbdd)
            {
                MessageBox.Show("Empresa Origen y Destino no pueden ser la misma");
                return;
            }
            StringBuilder x = new StringBuilder(string.Empty);


            x.AppendLine("SELECT m8a.ciddocumento");
            x.AppendLine(",m8a.ciddocumentode");
            x.AppendLine(",m8a.cidconceptodocumento");
            x.AppendLine(",m8a.cseriedocumento");
            x.AppendLine(",m8a.cfolio");
            x.AppendLine(",m8a.cfecha");
            //if (radioButton9.Checked == true)
            x.AppendLine(",m2.ccodigocliente");
            //if (radioButton10.Checked == true)
            //x.AppendLine(",m2.cdencome01");
            //  x.AppendLine(",m2.ctextoex01 as ccodigoc01");
            //            x.AppendLine(",m8a.cidclien01");
            x.AppendLine(",m8a.crazonsocial");
            x.AppendLine(",m8a.crfc");
            x.AppendLine(",m8a.cidagente");
            x.AppendLine(",m8a.cfechavencimiento");
            x.AppendLine(",m8a.cfechaprontopago");
            x.AppendLine(",m8a.cfechaentregarecepcion");
            x.AppendLine(",m8a.cfechaultimointeres");
            x.AppendLine(",m8a.cidmoneda");
            x.AppendLine(",m8a.ctipocambio");
            x.AppendLine(",m8a.creferencia");
            x.AppendLine(",m8a.cobservaciones");
            x.AppendLine(",m8a.cnaturaleza");
            x.AppendLine(",m8a.ciddocumentoorigen");
            x.AppendLine(",m8a.cplantilla");
            x.AppendLine(",m8a.cusacliente");
            x.AppendLine(",m8a.cusaproveedor");
            x.AppendLine(",m8a.cafectado");
            x.AppendLine(",m8a.cimpreso");
            x.AppendLine(",m8a.ccancelado");
            x.AppendLine(",m8a.cdevuelto");
            x.AppendLine(",m8a.cidprepoliza");
            x.AppendLine(",m8a.cidprepolizacancelacion");
            x.AppendLine(",m8a.cestadocontable");
            x.AppendLine(",m8a.cneto");
            x.AppendLine(",m8a.cimpuesto1");
            x.AppendLine(",m8a.cimpuesto2");
            x.AppendLine(",m8a.cimpuesto3");
            x.AppendLine(",m8a.cretencion1");
            x.AppendLine(",m8a.cretencion2");
            x.AppendLine(",m8a.cdescuentomov");
            x.AppendLine(",m8a.cdescuentodoc1");
            x.AppendLine(",m8a.cdescuentodoc2");
            x.AppendLine(",m8a.cgasto1");
            x.AppendLine(",m8a.cgasto2");
            x.AppendLine(",m8a.cgasto3");
            x.AppendLine(",m8a.ctotal");
            x.AppendLine(",m8a.cpendiente");
            x.AppendLine(",m8a.ctotalunidades");
            x.AppendLine(",m8a.cdescuentoprontopago");
            x.AppendLine(",m8a.cporcentajeimpuesto1");
            x.AppendLine(",m8a.cporcentajeimpuesto2");
            x.AppendLine(",m8a.cporcentajeimpuesto3");
            x.AppendLine(",m8a.cporcentajeretencion1");
            x.AppendLine(",m8a.cporcentajeretencion2");
            x.AppendLine(",m8a.cporcentajeinteres");
            x.AppendLine(",m8a.ctextoextra1");
            x.AppendLine(",m8a.ctextoextra2");
            x.AppendLine(",m8a.ctextoextra3");
            x.AppendLine(",m8a.cfechaextra");
            x.AppendLine(",m8a.cimporteextra1");
            x.AppendLine(",m8a.cimporteextra2");
            x.AppendLine(",m8a.cimporteextra2");
            x.AppendLine(",m8a.cimporteextra3");
            x.AppendLine(",m8a.cimporteextra3");
            x.AppendLine(",m8a.cimporteextra4");
            x.AppendLine(",m8a.cdestinatario");
            x.AppendLine(",m8a.cnumeroguia");
            x.AppendLine(",m8a.cmensajeria");
            x.AppendLine(",m8a.ccuentamensajeria");
            x.AppendLine(",m8a.cnumerocajas");
            x.AppendLine(",m8a.cpeso");
            x.AppendLine(",m8a.cbanobservaciones");
            x.AppendLine(",m8a.cbandatosenvio");
            x.AppendLine(",m8a.cbancondicionescredito");
            x.AppendLine(",m8a.cbangastos");
            x.AppendLine(",m8a.cunidadespendientes");
            x.AppendLine(",m8a.ctimestamp");
            x.AppendLine(",m8a.cimpcheqpaq");
            x.AppendLine(",m8a.csistorig");
            x.AppendLine(",m8a.cidmonedca");
            x.AppendLine(",m8a.ctipocamca");
            x.AppendLine(",m8a.cescfd");
            x.AppendLine(",m8a.ctienecfd");
            x.AppendLine(",m8a.clugarexpe");
            x.AppendLine(",m8a.cmetodopag");
            x.AppendLine(",m8a.cnumparcia");
            x.AppendLine(",m8a.ccantparci");
            x.AppendLine(",m8a.ccondipago");
            x.AppendLine(",m8a.cnumctapag  ");
            //x.AppendLine("  from mgw10008 m8a join mgw10002 m2 on m8a.cidclien01 = m2.cidclien01 and m2.ctipocli01 <= 2 ");
            x.AppendLine("  from admDocumentos m8a join admclientes m2 on m8a.CIDCLIENTEPROVEEDOR = m2.CIDCLIENTEPROVEEDOR and m2.CTIPOCLIENTE <= 2 ");
            x.AppendLine("where m8a.cpendiente > 0 and m8a.ccancelado = 0 and m8a.cnaturaleza = 0 ");
            x.AppendLine("order by m8a.ciddocumento");


            //Properties.Settings.Default.RutaEmpresaADM = cbOrigen.SelectedValue.ToString().Trim();
            //Properties.Settings.Default.RutaEmpresaDestino= cbDestino.SelectedValue.ToString().Trim();
            Properties.Settings.Default.Save();
            int lporcodigo = 0;
            if (radioButton9.Checked == true)
                lporcodigo = 1;
            int z = lrn.mEjecutarComando3(x.ToString(), 1, lporcodigo, ciCompanyList11.aliasbdd, ciCompanyList12.aliasbdd);





            if (z == 0)
                MessageBox.Show("Proceso Terminado Clientes");

            x = new StringBuilder(string.Empty);

            x.AppendLine("SELECT m8a.ciddocumento");
            x.AppendLine(",m8a.ciddocumentode");
            x.AppendLine(",m8a.cidconceptodocumento");
            x.AppendLine(",m8a.cseriedocumento");
            x.AppendLine(",m8a.cfolio");
            x.AppendLine(",m8a.cfecha");
            //if (radioButton9.Checked == true)
            x.AppendLine(",m2.ccodigocliente");
            //if (radioButton10.Checked == true)
            //x.AppendLine(",m2.cdencome01");
            //  x.AppendLine(",m2.ctextoex01 as ccodigoc01");
            //            x.AppendLine(",m8a.cidclien01");
            x.AppendLine(",m8a.crazonsocial");
            x.AppendLine(",m8a.crfc");
            x.AppendLine(",m8a.cidagente");
            x.AppendLine(",m8a.cfechavencimiento");
            x.AppendLine(",m8a.cfechaprontopago");
            x.AppendLine(",m8a.cfechaentregarecepcion");
            x.AppendLine(",m8a.cfechaultimointeres");
            x.AppendLine(",m8a.cidmoneda");
            x.AppendLine(",m8a.ctipocambio");
            x.AppendLine(",m8a.creferencia");
            x.AppendLine(",m8a.cobservaciones");
            x.AppendLine(",m8a.cnaturaleza");
            x.AppendLine(",m8a.ciddocumentoorigen");
            x.AppendLine(",m8a.cplantilla");
            x.AppendLine(",m8a.cusacliente");
            x.AppendLine(",m8a.cusaproveedor");
            x.AppendLine(",m8a.cafectado");
            x.AppendLine(",m8a.cimpreso");
            x.AppendLine(",m8a.ccancelado");
            x.AppendLine(",m8a.cdevuelto");
            x.AppendLine(",m8a.cidprepoliza");
            x.AppendLine(",m8a.cidprepolizacancelacion");
            x.AppendLine(",m8a.cestadocontable");
            x.AppendLine(",m8a.cneto");
            x.AppendLine(",m8a.cimpuesto1");
            x.AppendLine(",m8a.cimpuesto2");
            x.AppendLine(",m8a.cimpuesto3");
            x.AppendLine(",m8a.cretencion1");
            x.AppendLine(",m8a.cretencion2");
            x.AppendLine(",m8a.cdescuentomov");
            x.AppendLine(",m8a.cdescuentodoc1");
            x.AppendLine(",m8a.cdescuentodoc2");
            x.AppendLine(",m8a.cgasto1");
            x.AppendLine(",m8a.cgasto2");
            x.AppendLine(",m8a.cgasto3");
            x.AppendLine(",m8a.ctotal");
            x.AppendLine(",m8a.cpendiente");
            x.AppendLine(",m8a.ctotalunidades");
            x.AppendLine(",m8a.cdescuentoprontopago");
            x.AppendLine(",m8a.cporcentajeimpuesto1");
            x.AppendLine(",m8a.cporcentajeimpuesto2");
            x.AppendLine(",m8a.cporcentajeimpuesto3");
            x.AppendLine(",m8a.cporcentajeretencion1");
            x.AppendLine(",m8a.cporcentajeretencion2");
            x.AppendLine(",m8a.cporcentajeinteres");
            x.AppendLine(",m8a.ctextoextra1");
            x.AppendLine(",m8a.ctextoextra2");
            x.AppendLine(",m8a.ctextoextra3");
            x.AppendLine(",m8a.cfechaextra");
            x.AppendLine(",m8a.cimporteextra1");
            x.AppendLine(",m8a.cimporteextra2");
            x.AppendLine(",m8a.cimporteextra2");
            x.AppendLine(",m8a.cimporteextra3");
            x.AppendLine(",m8a.cimporteextra3");
            x.AppendLine(",m8a.cimporteextra4");
            x.AppendLine(",m8a.cdestinatario");
            x.AppendLine(",m8a.cnumeroguia");
            x.AppendLine(",m8a.cmensajeria");
            x.AppendLine(",m8a.ccuentamensajeria");
            x.AppendLine(",m8a.cnumerocajas");
            x.AppendLine(",m8a.cpeso");
            x.AppendLine(",m8a.cbanobservaciones");
            x.AppendLine(",m8a.cbandatosenvio");
            x.AppendLine(",m8a.cbancondicionescredito");
            x.AppendLine(",m8a.cbangastos");
            x.AppendLine(",m8a.cunidadespendientes");
            x.AppendLine(",m8a.ctimestamp");
            x.AppendLine(",m8a.cimpcheqpaq");
            x.AppendLine(",m8a.csistorig");
            x.AppendLine(",m8a.cidmonedca");
            x.AppendLine(",m8a.ctipocamca");
            x.AppendLine(",m8a.cescfd");
            x.AppendLine(",m8a.ctienecfd");
            x.AppendLine(",m8a.clugarexpe");
            x.AppendLine(",m8a.cmetodopag");
            x.AppendLine(",m8a.cnumparcia");
            x.AppendLine(",m8a.ccantparci");
            x.AppendLine(",m8a.ccondipago");
            x.AppendLine(",m8a.cnumctapag  ");
           /* x.AppendLine("  from mgw10008 m8a join mgw10002 m2 on m8a.cidclien01 = m2.cidclien01 and m2.ctipocli01 >= 2 ");
            x.AppendLine("join mgw10007 m7 on m7.ciddocum01 = m8a.ciddocum02 and m7.cusaprov01 = 1");
            x.AppendLine("where m8a.cpendiente > 0 and m8a.ccancelado = 0 and m8a.cnatural01 = 1 ");
            x.AppendLine("order by m8a.ciddocum01");*/

            x.AppendLine("  from admDocumentos m8a join admclientes m2 on m8a.CIDCLIENTEPROVEEDOR = m2.CIDCLIENTEPROVEEDOR and m2.CTIPOCLIENTE >= 2 ");
            x.AppendLine("where m8a.cpendiente > 0 and m8a.ccancelado = 0 and m8a.cnaturaleza = 0 ");
            x.AppendLine("order by m8a.ciddocumento");


            z = lrn.mEjecutarComando3(x.ToString(), 0, lporcodigo, ciCompanyList11.aliasbdd,ciCompanyList12.aliasbdd);
            //int z = lrn.mEjecutarComando3(x.ToString(), 1, lporcodigo, ciCompanyList11.aliasbdd, ciCompanyList12.aliasbdd);
            if (z == 0)
                MessageBox.Show("Proceso Terminado Proveedores");

        }
    }
}
