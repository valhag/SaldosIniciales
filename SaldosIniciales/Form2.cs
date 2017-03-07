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
//using Interfaces;


namespace SaldosIniciales
{
    public partial class Form2 : Form
    {
        ClassRNLOB lrn = new ClassRNLOB();

        public string Cadenaconexion = "";
        public Form2()
        {
            InitializeComponent();
            if (Properties.Settings.Default.server != "")
            {
                //                Properties.Settings.Default.server = "Toshiba-pc";
                //              Properties.Settings.Default.database = "GeneralesSQL";
                //            Properties.Settings.Default.user = "sa";
                //          Properties.Settings.Default.password = "ady123";

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
               // Archivo = Properties.Settings.Default.archivo;
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            lrn.mSeteaDirectorio(Directory.GetCurrentDirectory());
            mcargarEmpresa(cbOrigen);

              txtServer.Text = Properties.Settings.Default.server;
             txtBD.Text = Properties.Settings.Default.database ;
             txtUser.Text = Properties.Settings.Default.user ;
            txtPass.Text = Properties.Settings.Default.password ;

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
                ///mllenarcomboempresas();
                //y.Visible = true;
            }

            //mcargarEmpresa(cbDestino);

            
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

        private void button1_Click(object sender, EventArgs e)
        {
            
            StringBuilder x = new StringBuilder(string.Empty);

            
            x.AppendLine("SELECT m8a.ciddocum01");
            x.AppendLine(",m8a.ciddocum02");
            x.AppendLine(",m8a.cidconce01");
            x.AppendLine(",m8a.cseriedo01");
            x.AppendLine(",m8a.cfolio");
            x.AppendLine(",m8a.cfecha");
            //if (radioButton9.Checked == true)
                x.AppendLine(",m2.ccodigoc01");
            //if (radioButton10.Checked == true)
                //x.AppendLine(",m2.cdencome01");
              //  x.AppendLine(",m2.ctextoex01 as ccodigoc01");
//            x.AppendLine(",m8a.cidclien01");
            x.AppendLine(",m8a.crazonso01");
            x.AppendLine(",m8a.crfc");
            x.AppendLine(",m8a.cidagente");
            x.AppendLine(",m8a.cfechave01");
            x.AppendLine(",m8a.cfechapr01");
            x.AppendLine(",m8a.cfechaen01");
            x.AppendLine(",m8a.cfechaul01");
            x.AppendLine(",m8a.cidmoneda");
            x.AppendLine(",m8a.ctipocam01");
            x.AppendLine(",m8a.creferen01");
            x.AppendLine(",m8a.cobserva01");
            x.AppendLine(",m8a.cnatural01");
            x.AppendLine(",m8a.ciddocum03");
            x.AppendLine(",m8a.cplantilla");
            x.AppendLine(",m8a.cusaclie01");
            x.AppendLine(",m8a.cusaprov01");
            x.AppendLine(",m8a.cafectado");
            x.AppendLine(",m8a.cimpreso");
            x.AppendLine(",m8a.ccancelado");
            x.AppendLine(",m8a.cdevuelto");
            x.AppendLine(",m8a.cidprepo01");
            x.AppendLine(",m8a.cidprepo02");
            x.AppendLine(",m8a.cestadoc01");
            x.AppendLine(",m8a.cneto");
            x.AppendLine(",m8a.cimpuesto1");
            x.AppendLine(",m8a.cimpuesto2");
            x.AppendLine(",m8a.cimpuesto3");
            x.AppendLine(",m8a.cretenci01");
            x.AppendLine(",m8a.cretenci02");
            x.AppendLine(",m8a.cdescuen01");
            x.AppendLine(",m8a.cdescuen02");
            x.AppendLine(",m8a.cdescuen03");
            x.AppendLine(",m8a.cgasto1");
            x.AppendLine(",m8a.cgasto2");
            x.AppendLine(",m8a.cgasto3");
            x.AppendLine(",m8a.ctotal");
            x.AppendLine(",m8a.cpendiente");
            x.AppendLine(",m8a.ctotalun01");
            x.AppendLine(",m8a.cdescuen04");
            x.AppendLine(",m8a.cporcent01");
            x.AppendLine(",m8a.cporcent02");
            x.AppendLine(",m8a.cporcent03");
            x.AppendLine(",m8a.cporcent04");
            x.AppendLine(",m8a.cporcent05");
            x.AppendLine(",m8a.cporcent06");
            x.AppendLine(",m8a.ctextoex01");
            x.AppendLine(",m8a.ctextoex02");
            x.AppendLine(",m8a.ctextoex03");
            x.AppendLine(",m8a.cfechaex01");
            x.AppendLine(",m8a.cimporte01");
            x.AppendLine(",m8a.cimporte02");
            x.AppendLine(",m8a.cimporte02");
            x.AppendLine(",m8a.cimporte03");
            x.AppendLine(",m8a.cimporte03");
            x.AppendLine(",m8a.cimporte04");
            x.AppendLine(",m8a.cdestina01");
            x.AppendLine(",m8a.cnumerog01");
            x.AppendLine(",m8a.cmensaje01");
            x.AppendLine(",m8a.ccuentam01");
            x.AppendLine(",m8a.cnumeroc01");
            x.AppendLine(",m8a.cpeso");
            x.AppendLine(",m8a.cbanobse01");
            x.AppendLine(",m8a.cbandato01");
            x.AppendLine(",m8a.cbancond01");
            x.AppendLine(",m8a.cbangastos");
            x.AppendLine(",m8a.cunidade01");
            x.AppendLine(",m8a.ctimestamp");
            x.AppendLine(",m8a.cimpcheq01");
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
            x.AppendLine("  from mgw10008 m8a join mgw10002 m2 on m8a.cidclien01 = m2.cidclien01 and m2.ctipocli01 <= 2 ");
            x.AppendLine("where m8a.cpendiente > 0 and m8a.ccancelado = 0 and m8a.cnatural01 = 0 ");
            x.AppendLine("order by m8a.ciddocum01");


            Properties.Settings.Default.RutaEmpresaADM = cbOrigen.SelectedValue.ToString().Trim();
            //Properties.Settings.Default.RutaEmpresaDestino= cbDestino.SelectedValue.ToString().Trim();
            Properties.Settings.Default.Save();
            int lporcodigo = 0;
            if (radioButton9.Checked == true)
                lporcodigo = 1;
            int z = lrn.mEjecutarComando2(x.ToString(),1,lporcodigo,ciCompanyList11.aliasbdd);





            if (z == 0)
                MessageBox.Show("Proceso Terminado Clientes");

            x = new StringBuilder(string.Empty);

            x.AppendLine("SELECT m8a.ciddocum01");
            x.AppendLine(",m8a.ciddocum02");
            x.AppendLine(",m8a.cidconce01");
            x.AppendLine(",m8a.cseriedo01");
            x.AppendLine(",m8a.cfolio");
            x.AppendLine(",m8a.cfecha");
            //if (radioButton9.Checked == true)
                x.AppendLine(",m2.ccodigoc01");
            //if (radioButton10.Checked == true)
              //  x.AppendLine(",m2.ctextoex01 as ccodigoc01");
            //            x.AppendLine(",m8a.cidclien01");
            x.AppendLine(",m8a.crazonso01");
            x.AppendLine(",m8a.crfc");
            x.AppendLine(",m8a.cidagente");
            x.AppendLine(",m8a.cfechave01");
            x.AppendLine(",m8a.cfechapr01");
            x.AppendLine(",m8a.cfechaen01");
            x.AppendLine(",m8a.cfechaul01");
            x.AppendLine(",m8a.cidmoneda");
            x.AppendLine(",m8a.ctipocam01");
            x.AppendLine(",m8a.creferen01");
            x.AppendLine(",m8a.cobserva01");
            x.AppendLine(",m8a.cnatural01");
            x.AppendLine(",m8a.ciddocum03");
            x.AppendLine(",m8a.cplantilla");
            x.AppendLine(",m8a.cusaclie01");
            x.AppendLine(",m8a.cusaprov01");
            x.AppendLine(",m8a.cafectado");
            x.AppendLine(",m8a.cimpreso");
            x.AppendLine(",m8a.ccancelado");
            x.AppendLine(",m8a.cdevuelto");
            x.AppendLine(",m8a.cidprepo01");
            x.AppendLine(",m8a.cidprepo02");
            x.AppendLine(",m8a.cestadoc01");
            x.AppendLine(",m8a.cneto");
            x.AppendLine(",m8a.cimpuesto1");
            x.AppendLine(",m8a.cimpuesto2");
            x.AppendLine(",m8a.cimpuesto3");
            x.AppendLine(",m8a.cretenci01");
            x.AppendLine(",m8a.cretenci02");
            x.AppendLine(",m8a.cdescuen01");
            x.AppendLine(",m8a.cdescuen02");
            x.AppendLine(",m8a.cdescuen03");
            x.AppendLine(",m8a.cgasto1");
            x.AppendLine(",m8a.cgasto2");
            x.AppendLine(",m8a.cgasto3");
            x.AppendLine(",m8a.ctotal");
            x.AppendLine(",m8a.cpendiente");
            x.AppendLine(",m8a.ctotalun01");
            x.AppendLine(",m8a.cdescuen04");
            x.AppendLine(",m8a.cporcent01");
            x.AppendLine(",m8a.cporcent02");
            x.AppendLine(",m8a.cporcent03");
            x.AppendLine(",m8a.cporcent04");
            x.AppendLine(",m8a.cporcent05");
            x.AppendLine(",m8a.cporcent06");
            x.AppendLine(",m8a.ctextoex01");
            x.AppendLine(",m8a.ctextoex02");
            x.AppendLine(",m8a.ctextoex03");
            x.AppendLine(",m8a.cfechaex01");
            x.AppendLine(",m8a.cimporte01");
            x.AppendLine(",m8a.cimporte02");
            x.AppendLine(",m8a.cimporte02");
            x.AppendLine(",m8a.cimporte03");
            x.AppendLine(",m8a.cimporte03");
            x.AppendLine(",m8a.cimporte04");
            x.AppendLine(",m8a.cdestina01");
            x.AppendLine(",m8a.cnumerog01");
            x.AppendLine(",m8a.cmensaje01");
            x.AppendLine(",m8a.ccuentam01");
            x.AppendLine(",m8a.cnumeroc01");
            x.AppendLine(",m8a.cpeso");
            x.AppendLine(",m8a.cbanobse01");
            x.AppendLine(",m8a.cbandato01");
            x.AppendLine(",m8a.cbancond01");
            x.AppendLine(",m8a.cbangastos");
            x.AppendLine(",m8a.cunidade01");
            x.AppendLine(",m8a.ctimestamp");
            x.AppendLine(",m8a.cimpcheq01");
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
            x.AppendLine("  from mgw10008 m8a join mgw10002 m2 on m8a.cidclien01 = m2.cidclien01 and m2.ctipocli01 >= 2 ");
            //x.AppendLine("join mgw10002 m2d on m2d.ccodigoc01 = m2.ccodigoc01 ");
            x.AppendLine("join mgw10007 m7 on m7.ciddocum01 = m8a.ciddocum02 and m7.cusaprov01 = 1");
            x.AppendLine("where m8a.cpendiente > 0 and m8a.ccancelado = 0 and m8a.cnatural01 = 1 ");
            x.AppendLine("order by m8a.ciddocum01");


            z = lrn.mEjecutarComando2(x.ToString(),0,lporcodigo,ciCompanyList11.aliasbdd);
            if (z == 0)
                MessageBox.Show("Proceso Terminado Proveedores");




        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            if (lrn.mValidaSQLConexion(txtServer.Text, txtBD.Text, txtUser.Text, txtPass.Text )==1)
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
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.server = txtServer.Text;
            Properties.Settings.Default.database = txtBD.Text;
            Properties.Settings.Default.user = txtUser.Text;
            Properties.Settings.Default.password = txtPass.Text;

            Properties.Settings.Default.Save();
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            if (Cadenaconexion != "")
            {
                ciCompanyList11.Populate(Cadenaconexion);
            }
            else
            {
             
            }
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }
    }
}
