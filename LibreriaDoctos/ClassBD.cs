using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
//using BarradeProgreso;
using System.Data.SqlClient ;
using Interfaces ;
using System.Collections;
using System.Data;


namespace LibreriaDoctos
{
    public class ClassBD: ISujeto
    {
        protected decimal lsubtotal;
        protected decimal limpuestos;
        protected string aRutaExe;
        public string productos;
        public string almacenes;
        public RegDocto primerdocto = new RegDocto();
        List<IObservador> lista = new List<IObservador>();
        public string Cadenaconexion;
        public string cserver;
        public string cbd;
        public string cusr;
        public string cpwd;

        public int mValidaSQLConexion(string server, string bd, string user, string psw)
        {
            Cadenaconexion = "data source =" + server + ";initial catalog =" + bd + ";user id = " + user + "; password = " + psw + ";";
            SqlConnection _con = new SqlConnection();
            cserver = server;
            cbd = bd;
            cusr = user;
            cpwd = psw;

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return 1;
            }
            catch (Exception ee)
            {
                return 0;
            }
        }


        public int mEjecutarComando3(string comando, int aClientes, int aporcodigo, string empresaorigen, string sempresadestino)
        {
            //miconexion.mAbrirConexionOrigen();
            SqlConnection _conOrigen = new SqlConnection();
            SqlConnection _conDestino = new SqlConnection();
            string sempresa = empresaorigen.Substring(empresaorigen.LastIndexOf("\\") + 1);
            string CadenaconexionOrigen = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _con = new SqlConnection();

            _conOrigen.ConnectionString = CadenaconexionOrigen;
            _conOrigen.Open();





            /*OleDbDataAdapter lda = new OleDbDataAdapter(comando, miconexion._conexion);*/
            SqlDataAdapter lda = new SqlDataAdapter(comando, _conOrigen);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);
            /*miconexion.mCerrarConexionOrigen();*/



            sempresa = sempresadestino.Substring(sempresadestino.LastIndexOf("\\") + 1);

            string Cadenaconexion1 = "data source =" + cserver + ";initial catalog = " + sempresa + ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _conDes = new SqlConnection();

            _conDes.ConnectionString = Cadenaconexion1;

            //miconexion.mAbrirConexionDestino();

            _conDes.Open();
            SqlCommand com = new SqlCommand("select ISNULL(max(ciddocumento),0) from admDocumentos");
            SqlDataReader ldr;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = _conDes;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;

            SqlCommand com1 = new SqlCommand("select ISNULL(max(ciddocumento),0) from admMovimientos");
            SqlDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = _conDes;
            ldr.Close();
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim = int.Parse(ldr1[0].ToString()) + 1;
            ldr1.Close();
            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclienteproveedor from admClientes where ccodigocliente = '" + zz["ccodigocliente"].ToString() + "'";
                else
                    ltexto = "select cidclienteproveedor from admClientes where ctextoextra1 = '" + zz["ccodigoccliente"].ToString() + "'";

                SqlCommand com2 = new SqlCommand(ltexto);

                SqlDataReader ldr2;
                com2.Connection = _conDes;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                ldr2.Close();
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into admDocumentos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedocumento"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);
                    int espacio = fecha.IndexOf(" ");
                    if (espacio > -1)
                        fecha = fecha.Substring(0, espacio);
                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                    string sFecha = dfecha.Year.ToString().PadLeft(4, '0') + dfecha.Month.ToString().PadLeft(2, '0') + dfecha.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFecha + "',");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonsocial"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechavencimiento"].ToString().Substring(0, 10);
                    int espacio1 = fechav.IndexOf(" ");
                    if (espacio1 > -1)
                        fechav = fechav.Substring(0, espacio1);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    string sFechav = dfechav.Year.ToString().PadLeft(4, '0') + dfechav.Month.ToString().PadLeft(2, '0') + dfechav.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFechav + "',");
                    //x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");

                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocambio"].ToString() + ",");
                    x.AppendLine("'" + zz["creferencia"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnaturaleza"].ToString() + ",");
                    x.AppendLine(zz["ciddocumentoorigen"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusacliente"].ToString() + ",");
                    x.AppendLine(zz["cusaproveedor"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepoliza"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepolizacancelacion"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadocontable"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretencion1"].ToString() + ",");
                    x.AppendLine(zz["cretencion2"].ToString() + ",");
                    x.AppendLine(zz["cdescuentomov"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc1"].ToString() + ",");
                    x.AppendLine(zz["cdescuentodoc2"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuentoprontopago"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion1"].ToString() + ",");
                    x.AppendLine(zz["cporcentajeretencion2"].ToString() + ",");
                    x.AppendLine(zz["CPORCENTAJEINTERES"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoextra1"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra2"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoextra3"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    //x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra1"].ToString() + ",");

                    x.AppendLine(zz["cimporteextra2"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra3"].ToString() + ",");
                    x.AppendLine(zz["cimporteextra4"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestinatario"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumeroguia"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensajeria"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentamensajeria"].ToString() + "',");
                    x.AppendLine(zz["cnumerocajas"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobservaciones"].ToString() + ",");
                    x.AppendLine(zz["cbandatosenvio"].ToString() + ",");
                    x.AppendLine(zz["cbancondicionescredito"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidadespendientes"].ToString() + ",");
                    //x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cimpcheqpaq"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "'");
                    x.AppendLine(",NEWID(),'',0)");
                    //x.AppendLine(",NEWID(),'')");


                    comando = x.ToString();
                    SqlCommand lsql3 = new SqlCommand(comando, _conDes);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into admmovimientos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcentajeimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    //x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    //x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    SqlCommand lsql4 = new SqlCommand(comando, _conDes);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






                /*
                 INSERT INTO [DESTINO]...[mgw10010]
               ([cidmovim01]
               ,[ciddocum01]
               ,[cnumerom01]
               ,[ciddocum02]
               ,[cidprodu01]
               ,[cidalmacen]
               ,[cunidades]
               ,[cunidade01]
               ,[cunidade02]
               ,[cidunidad]
               ,[cidunida01]
               ,[cprecio]
               ,[cprecioc01]
               ,[ccostoca01]
               ,[ccostoes01]
               ,[cneto]
               ,[cimpuesto1]
               ,[cporcent01]
               ,[cimpuesto2]
               ,[cporcent02]
               ,[cimpuesto3]
               ,[cporcent03]
               ,[cretenci01]
               ,[cporcent04]
               ,[cretenci02]
               ,[cporcent05]
               ,[cdescuen01]
               ,[cporcent06]
               ,[cdescuen02]
               ,[cporcent07]
               ,[cdescuen03]
               ,[cporcent08]
               ,[cdescuen04]
               ,[cporcent09]
               ,[cdescuen05]
               ,[cporcent10]
               ,[ctotal]
               ,[cporcent11]
               ,[creferen01]
               ,[cobserva01]
               ,[cafectae01]
               ,[cafectad01]
               ,[cafectad02]
               ,[cfecha]
               ,[cmovtooc01]
               ,[cidmovto01]
               ,[cidmovto02]
               ,[cunidade03]
               ,[cunidade04]
               ,[cunidade05]
               ,[cunidade06]
               ,[ctipotra01]
               ,[cidvalor01]
               ,[ctextoex01]
               ,[ctextoex02]
               ,[ctextoex03]
               ,[cfechaex01]
               ,[cimporte01]
               ,[cimporte02]
               ,[cimporte03]
               ,[cimporte04]
               ,[ctimestamp]
               ,[cgtomovto]
               ,[cscmovto]
               ,[ccomventa]
               ,[cidmovdest]
               ,[cnumconsol])

                 */


            }
            //miconexion.mCerrarConexionDestino();
            _conDes.Close();

            return 0;
        }

        public int mEjecutarComando2(string comando, int aClientes, int aporcodigo, string empresa)
        {
            miconexion.mAbrirConexionOrigen();



            //mValidaSQLConexion(txtServer.Text, txtBD.Text, txtUser.Text, txtPass.Text);


            OleDbDataAdapter lda = new OleDbDataAdapter(comando, miconexion._conexion);
            System.Data.DataSet lds = new System.Data.DataSet();
            lda.Fill(lds);
            miconexion.mCerrarConexionOrigen();

            

            string sempresa = empresa.Substring (empresa.LastIndexOf("\\")+1);

            string Cadenaconexion1 = "data source =" + cserver + ";initial catalog = " + sempresa+ ";user id = " + cusr + "; password = " + cpwd + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion1;

            //miconexion.mAbrirConexionDestino();

            _con.Open();
            SqlCommand com = new SqlCommand("select ISNULL(max(ciddocumento),0) from admDocumentos");
            SqlDataReader ldr;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = _con;
            ldr = com.ExecuteReader();
            ldr.Read();
            int liddocum = int.Parse(ldr[0].ToString()) + 1;

            SqlCommand com1 = new SqlCommand("select ISNULL(max(ciddocumento),0) from admMovimientos");
            SqlDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = _con;
            ldr.Close();
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim = int.Parse(ldr1[0].ToString()) + 1;
            ldr1.Close();
            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "";
                if (aporcodigo == 1)
                    ltexto = "select cidclienteproveedor from admClientes where ccodigocliente = '" + zz["ccodigoc01"].ToString() + "'";
                else
                    ltexto = "select cidclienteproveedor from admClientes where ctextoextra1 = '" + zz["ccodigoc01"].ToString() + "'";

                SqlCommand com2 = new SqlCommand(ltexto);

                SqlDataReader ldr2;
                com2.Connection = _con;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                ldr2.Close();
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into admDocumentos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedo01"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);
                    int espacio = fecha.IndexOf(" ");
                    if (espacio > -1)
                        fecha = fecha.Substring(0, espacio);
                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');
                    string sFecha = dfecha.Year.ToString().PadLeft(4, '0') + dfecha.Month.ToString().PadLeft(2, '0') + dfecha.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFecha + "',");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonso01"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechave01"].ToString().Substring(0, 10);
                    int espacio1 = fechav.IndexOf(" ");
                    if (espacio1 > -1)
                        fechav = fechav.Substring(0, espacio1);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    string sFechav = dfechav.Year.ToString().PadLeft(4, '0') + dfechav.Month.ToString().PadLeft(2, '0') + dfechav.Day.ToString().PadLeft(2, '0');
                    x.AppendLine("'" + sFechav + "',");
                    //x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    //x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");

                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocam01"].ToString() + ",");
                    x.AppendLine("'" + zz["creferen01"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnatural01"].ToString() + ",");
                    x.AppendLine(zz["ciddocum03"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusaclie01"].ToString() + ",");
                    x.AppendLine(zz["cusaprov01"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepo01"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepo02"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadoc01"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretenci01"].ToString() + ",");
                    x.AppendLine(zz["cretenci02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen01"].ToString() + ",");
                    x.AppendLine(zz["cdescuen02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen03"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuen04"].ToString() + ",");
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine(zz["cporcent02"].ToString() + ",");
                    x.AppendLine(zz["cporcent03"].ToString() + ",");
                    x.AppendLine(zz["cporcent04"].ToString() + ",");
                    x.AppendLine(zz["cporcent05"].ToString() + ",");
                    x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoex01"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex02"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex03"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    //x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporte01"].ToString() + ",");

                    x.AppendLine(zz["cimporte02"].ToString() + ",");
                    x.AppendLine(zz["cimporte03"].ToString() + ",");
                    x.AppendLine(zz["cimporte04"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestina01"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumerog01"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensaje01"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentam01"].ToString() + "',");
                    x.AppendLine(zz["cnumeroc01"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobse01"].ToString() + ",");
                    x.AppendLine(zz["cbandato01"].ToString() + ",");
                    x.AppendLine(zz["cbancond01"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidade01"].ToString() + ",");
                    //x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");

                    x.AppendLine(zz["cimpcheq01"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "'");
                    x.AppendLine(",NEWID(),'',0)");
                    

                    comando = x.ToString();
                    SqlCommand lsql3 = new SqlCommand(comando, _con);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into admmovimientos values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    //x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    //x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("'" + sFechav + "',");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    SqlCommand lsql4 = new SqlCommand(comando, _con);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






                /*
                 INSERT INTO [DESTINO]...[mgw10010]
               ([cidmovim01]
               ,[ciddocum01]
               ,[cnumerom01]
               ,[ciddocum02]
               ,[cidprodu01]
               ,[cidalmacen]
               ,[cunidades]
               ,[cunidade01]
               ,[cunidade02]
               ,[cidunidad]
               ,[cidunida01]
               ,[cprecio]
               ,[cprecioc01]
               ,[ccostoca01]
               ,[ccostoes01]
               ,[cneto]
               ,[cimpuesto1]
               ,[cporcent01]
               ,[cimpuesto2]
               ,[cporcent02]
               ,[cimpuesto3]
               ,[cporcent03]
               ,[cretenci01]
               ,[cporcent04]
               ,[cretenci02]
               ,[cporcent05]
               ,[cdescuen01]
               ,[cporcent06]
               ,[cdescuen02]
               ,[cporcent07]
               ,[cdescuen03]
               ,[cporcent08]
               ,[cdescuen04]
               ,[cporcent09]
               ,[cdescuen05]
               ,[cporcent10]
               ,[ctotal]
               ,[cporcent11]
               ,[creferen01]
               ,[cobserva01]
               ,[cafectae01]
               ,[cafectad01]
               ,[cafectad02]
               ,[cfecha]
               ,[cmovtooc01]
               ,[cidmovto01]
               ,[cidmovto02]
               ,[cunidade03]
               ,[cunidade04]
               ,[cunidade05]
               ,[cunidade06]
               ,[ctipotra01]
               ,[cidvalor01]
               ,[ctextoex01]
               ,[ctextoex02]
               ,[ctextoex03]
               ,[cfechaex01]
               ,[cimporte01]
               ,[cimporte02]
               ,[cimporte03]
               ,[cimporte04]
               ,[ctimestamp]
               ,[cgtomovto]
               ,[cscmovto]
               ,[ccomventa]
               ,[cidmovdest]
               ,[cnumconsol])

                 */


            }
            miconexion.mCerrarConexionDestino();

            return 0;
        }

        public int mEjecutarComando(string comando, int aClientes, int aporcodigo)
        {
            miconexion.mAbrirConexionOrigen();
            //string lcadena44 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + " where ciddocum01 = " + lIdDocumento;

            
            


            OleDbDataAdapter lda = new OleDbDataAdapter(comando,miconexion._conexion);
            System.Data.DataSet lds = new System.Data.DataSet ();
            lda.Fill(lds);
            miconexion.mCerrarConexionOrigen();

            miconexion.mAbrirConexionDestino();

            OleDbCommand com = new OleDbCommand("select NVL(max(ciddocum01),0) from mgw10008");
            OleDbDataReader ldr ;

            //com.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com.Connection = miconexion._conexion;
            ldr= com.ExecuteReader();
            ldr.Read();
            int liddocum= int.Parse(ldr[0].ToString())+1;

            OleDbCommand com1 = new OleDbCommand("select NVL(max(cidmovim01),0) from mgw10010");
            OleDbDataReader ldr1;

            //com1.CommandText = "select NVL(max(ciddocum01),1) from mgw10008";
            com1.Connection = miconexion._conexion;
            ldr1 = com1.ExecuteReader();
            ldr1.Read();
            int lidmovim= int.Parse(ldr1[0].ToString())+1;

            foreach (DataRow zz in lds.Tables[0].Rows)
            {

                string ltexto = "" ;
                if (aporcodigo == 1)
                   ltexto = "select cidclien01 from mgw10002 where ccodigoc01 = '" + zz["ccodigoc01"].ToString() + "'";
                else
                    ltexto = "select cidclien01 from mgw10002 where ctextoex01 = '" + zz["ccodigoc01"].ToString() + "'";

                OleDbCommand com2 = new OleDbCommand(ltexto);

                OleDbDataReader ldr2;
                com2.Connection = miconexion._conexion;
                ldr2 = com2.ExecuteReader();
                int lidclien = 0;
                if (ldr2.HasRows == true)
                {
                    ldr2.Read();
                    lidclien = int.Parse(ldr2[0].ToString());
                }
                if (lidclien > 0)
                {
                    StringBuilder x = new StringBuilder();
                    x.AppendLine("insert into mgw10008 values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(liddocum + ",");
                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                        x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        x.AppendLine("40,");
                    }
                    //x.AppendLine(zz["cidconce01"].ToString() + ",");


                    x.AppendLine("'" + zz["cseriedo01"].ToString() + "',");
                    //x.AppendLine("'ABCD',");

                    x.AppendLine(zz["cfolio"].ToString() + ",");
                    //x.AppendLine(liddocum.ToString() + ",");


                    string fecha = zz["cfecha"].ToString().Substring(0, 10);

                    DateTime dfecha = DateTime.Parse(fecha);
                    fecha = dfecha.Month.ToString().PadLeft(2, '0') + "/" + dfecha.Day.ToString().PadLeft(2, '0') + "/" + dfecha.Year.ToString().PadLeft(4, '0');

                    x.AppendLine("ctod('" + fecha + "'),");
                    //x.AppendLine(zz["cidclien01"].ToString() + ",");

                    x.AppendLine(lidclien.ToString() + ",");


                    x.AppendLine("'" + zz["crazonso01"].ToString() + "',");
                    x.AppendLine("'" + zz["crfc"].ToString() + "',");
                    x.AppendLine(zz["cidagente"].ToString() + ",");

                    string fechav = zz["cfechave01"].ToString().Substring(0, 10);

                    DateTime dfechav = DateTime.Parse(fechav);
                    fechav = dfechav.Month.ToString().PadLeft(2, '0') + "/" + dfechav.Day.ToString().PadLeft(2, '0') + "/" + dfechav.Year.ToString().PadLeft(4, '0');
                    x.AppendLine("ctod('" + fechav + "'),");
                    x.AppendLine("ctod('" + zz["cfechapr01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("ctod('" + zz["cfechaen01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("ctod('" + zz["cfechaul01"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine(zz["cidmoneda"].ToString() + ",");
                    x.AppendLine(zz["ctipocam01"].ToString() + ",");
                    x.AppendLine("'" + zz["creferen01"].ToString() + "',");
                    //x.AppendLine("'" + zz["cobserva01"].ToString() + "',");
                    x.AppendLine("'" + zz["cfolio"].ToString() + "',");

                    x.AppendLine(zz["cnatural01"].ToString() + ",");
                    x.AppendLine(zz["ciddocum03"].ToString() + ",");
                    x.AppendLine(zz["cplantilla"].ToString() + ",");
                    x.AppendLine(zz["cusaclie01"].ToString() + ",");
                    x.AppendLine(zz["cusaprov01"].ToString() + ",");
                    x.AppendLine(zz["cafectado"].ToString() + ",");
                    //x.AppendLine(zz["cimpreso"].ToString() + ",");
                    x.AppendLine("0,");

                    x.AppendLine(zz["ccancelado"].ToString() + ",");
                    x.AppendLine(zz["cdevuelto"].ToString() + ",");
                    x.AppendLine(zz["cidprepo01"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cidprepo02"].ToString() + ",");
                    //x.AppendLine("0,");
                    x.AppendLine(zz["cestadoc01"].ToString() + ",");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cimpuesto2"].ToString() + ",");
                    x.AppendLine(zz["cimpuesto3"].ToString() + ",");
                    x.AppendLine(zz["cretenci01"].ToString() + ",");
                    x.AppendLine(zz["cretenci02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen01"].ToString() + ",");
                    x.AppendLine(zz["cdescuen02"].ToString() + ",");
                    x.AppendLine(zz["cdescuen03"].ToString() + ",");
                    x.AppendLine(zz["cgasto1"].ToString() + ",");
                    x.AppendLine(zz["cgasto2"].ToString() + ",");
                    x.AppendLine(zz["cgasto3"].ToString() + ",");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    //x.AppendLine(zz["ctotalun01"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["cdescuen04"].ToString() + ",");
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine(zz["cporcent02"].ToString() + ",");
                    x.AppendLine(zz["cporcent03"].ToString() + ",");
                    x.AppendLine(zz["cporcent04"].ToString() + ",");
                    x.AppendLine(zz["cporcent05"].ToString() + ",");
                    x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine("'" + zz["ctextoex01"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex02"].ToString() + "',");
                    x.AppendLine("'" + zz["ctextoex03"].ToString() + "',");
                    //x.AppendLine("'" + zz["cfechaex01"].ToString() + "',");
                    x.AppendLine("ctod('" + zz["cfechaex01"].ToString().Substring(0, 10) + "'),");


                    //x.AppendLine(zz["cporcent06"].ToString() + ",");
                    x.AppendLine(zz["cimporte01"].ToString() + ",");

                    x.AppendLine(zz["cimporte02"].ToString() + ",");
                    x.AppendLine(zz["cimporte03"].ToString() + ",");
                    x.AppendLine(zz["cimporte04"].ToString() + ",");
                    x.AppendLine("'" + zz["cdestina01"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumerog01"].ToString() + "',");
                    x.AppendLine("'" + zz["cmensaje01"].ToString() + "',");
                    x.AppendLine("'" + zz["ccuentam01"].ToString() + "',");
                    x.AppendLine(zz["cnumeroc01"].ToString() + ",");

                    x.AppendLine(zz["cpeso"].ToString() + ",");
                    x.AppendLine(zz["cbanobse01"].ToString() + ",");
                    x.AppendLine(zz["cbandato01"].ToString() + ",");
                    x.AppendLine(zz["cbancond01"].ToString() + ",");
                    x.AppendLine(zz["cbangastos"].ToString() + ",");
                    x.AppendLine(zz["cunidade01"].ToString() + ",");
                    x.AppendLine("ctod('" + zz["ctimestamp"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine(zz["cimpcheq01"].ToString() + ",");
                    x.AppendLine(zz["csistorig"].ToString() + ",");
                    x.AppendLine(zz["cidmonedca"].ToString() + ",");
                    x.AppendLine(zz["ctipocamca"].ToString() + ",");
                    //x.AppendLine(zz["cescfd"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine(zz["ctienecfd"].ToString() + ",");
                    x.AppendLine("'" + zz["clugarexpe"].ToString() + "',");
                    x.AppendLine("'" + zz["cmetodopag"].ToString() + "',");
                    x.AppendLine(zz["cnumparcia"].ToString() + ",");
                    x.AppendLine(zz["ccantparci"].ToString() + ",");
                    x.AppendLine("'" + zz["ccondipago"].ToString() + "',");
                    x.AppendLine("'" + zz["cnumctapag"].ToString() + "')");

                    comando = x.ToString();
                    OleDbCommand lsql3 = new OleDbCommand(comando, miconexion._conexion);
                    lsql3.ExecuteNonQuery();


                    x = new StringBuilder();
                    x.AppendLine("insert into mgw10010 values(");
                    //lds.Tables[0].Rows["ciddocum01"].ToString();
                    x.AppendLine(lidmovim.ToString() + ",");
                    x.AppendLine(liddocum.ToString() + ",");
                    x.AppendLine("1,");
                    //x.AppendLine("13,");

                    if (aClientes == 1)
                    {
                        //x.AppendLine(zz["ciddocum02"].ToString() + ",");
                        x.AppendLine("13,");
                      //  x.AppendLine("39,");
                    }
                    else
                    {
                        x.AppendLine("27,");
                        //x.AppendLine("40,");
                    }

                    x.AppendLine("0,"); //producto
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //unidades2
                    x.AppendLine("0,"); //cidunidad
                    x.AppendLine("0,"); //cidunida01
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine(zz["cneto"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");

                    //x.AppendLine(zz["cimpuesto1"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto1
                    x.AppendLine(zz["cporcent01"].ToString() + ",");
                    x.AppendLine("0,"); //impuesto2
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //impuesto3
                    x.AppendLine("0,");
                    x.AppendLine("0,"); //retencion1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//retencion2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 1
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 2
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 3
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 4
                    x.AppendLine("0,");
                    x.AppendLine("0,");//descuento 5
                    x.AppendLine("0,");
                    //x.AppendLine(zz["ctotal"].ToString() + ",");
                    x.AppendLine(zz["cpendiente"].ToString() + ",");
                    x.AppendLine("0,");
                    x.AppendLine("'',");
                    x.AppendLine("'',"); //observaciones
                    x.AppendLine("3,");
                    x.AppendLine("0,");
                    x.AppendLine("0,");
                    //x.AppendLine("3,");

                    x.AppendLine("ctod('" + zz["cfecha"].ToString().Substring(0, 10) + "'),");
                    x.AppendLine("0,0,0,0,0,0,0,");
                    x.AppendLine("0,0,'','','',");
                    x.AppendLine("ctod('" + fecha + "'),");
                    x.AppendLine("0,0,0,0,");
                    x.AppendLine("'',0,'',0,0,0)");

                    comando = x.ToString();
                    OleDbCommand lsql4 = new OleDbCommand(comando, miconexion._conexion);
                    lsql4.ExecuteNonQuery();

                    liddocum++;
                    lidmovim++;
                }






            /*
             INSERT INTO [DESTINO]...[mgw10010]
           ([cidmovim01]
           ,[ciddocum01]
           ,[cnumerom01]
           ,[ciddocum02]
           ,[cidprodu01]
           ,[cidalmacen]
           ,[cunidades]
           ,[cunidade01]
           ,[cunidade02]
           ,[cidunidad]
           ,[cidunida01]
           ,[cprecio]
           ,[cprecioc01]
           ,[ccostoca01]
           ,[ccostoes01]
           ,[cneto]
           ,[cimpuesto1]
           ,[cporcent01]
           ,[cimpuesto2]
           ,[cporcent02]
           ,[cimpuesto3]
           ,[cporcent03]
           ,[cretenci01]
           ,[cporcent04]
           ,[cretenci02]
           ,[cporcent05]
           ,[cdescuen01]
           ,[cporcent06]
           ,[cdescuen02]
           ,[cporcent07]
           ,[cdescuen03]
           ,[cporcent08]
           ,[cdescuen04]
           ,[cporcent09]
           ,[cdescuen05]
           ,[cporcent10]
           ,[ctotal]
           ,[cporcent11]
           ,[creferen01]
           ,[cobserva01]
           ,[cafectae01]
           ,[cafectad01]
           ,[cafectad02]
           ,[cfecha]
           ,[cmovtooc01]
           ,[cidmovto01]
           ,[cidmovto02]
           ,[cunidade03]
           ,[cunidade04]
           ,[cunidade05]
           ,[cunidade06]
           ,[ctipotra01]
           ,[cidvalor01]
           ,[ctextoex01]
           ,[ctextoex02]
           ,[ctextoex03]
           ,[cfechaex01]
           ,[cimporte01]
           ,[cimporte02]
           ,[cimporte03]
           ,[cimporte04]
           ,[ctimestamp]
           ,[cgtomovto]
           ,[cscmovto]
           ,[ccomventa]
           ,[cidmovdest]
           ,[cnumconsol])

             */


            }
            miconexion.mCerrarConexionDestino();
            
            return 0;
        }


        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fEditaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaCteProv();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoCteProv(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaProducto();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaProducto();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoProducto(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")]        static extern int fInsertaAlmacen();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaAlmacen();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoAlmacen(string aCampo, string aValor);

        [DllImport("MGW_SDK.DLL")] static extern int fInsertarDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDocumento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]        static extern int fInsertarMovimiento();
        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaMovimiento();
        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoMovimiento(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]      static extern int fInsertaDireccion();

        [DllImport("MGW_SDK.DLL")]        static extern int fGuardaDireccion();
        [DllImport("MGW_SDK.DLL")]        static extern int fBorraDocumento();
        //[DllImport("MGW_SDK.DLL")]        static extern int fBorraMovimiento();


        [DllImport("MGW_SDK.DLL")]        static extern int fSetDatoDireccion(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")]        static extern int fAfectaDocto_Param(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);
        [DllImport("MGW_SDK.DLL")]        static extern long fError(long aNumErrror, string aError, long aLen);

        [DllImport("MGW_SDK.DLL")]
        static extern int fSiguienteFolio(string lCodigoConcepto, ref string lSerie, ref double lFolio);

        [DllImport("MGW_SDK.DLL")]
        static extern long fSaldarDocumento_Param(string lCodConcepto_Pagar, string lSerie_Pagar, double lFolio_Pagar,
string lCodConcepto_Pago, string lSerie_Pago, double lFolio_Pago, double lImporte, int lIdMoneda, string lFecha);

        // Need this DllImport statement to reset the floating point register below
        [DllImport("msvcr71.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int _controlfp(int n, int mask);

        //protected ClassConexion miconexion;
         public ClassConexion  miconexion = new ClassConexion();
        public RegDocto _RegDoctoOrigen = new RegDocto();
        private string _rfc;
        private string _razonsocial;
        public const string _NombreAplicacionCompleto = "SaldosIniciales.exe";
        public const string _NombreAplicacion = "SaldosIniciales";
        public List<RegDocto> _RegDoctos = new List<RegDocto>();
        //protected OleDbConnection _con;

        protected OleDbConnection  _con;


        public ClassBD()
        {
          //  miconexion = new ClassConexion();
           _con = new OleDbConnection ();
        }

        public void mAsignaRuta(string aRuta)
        {
            aRutaExe = aRuta;
            miconexion.aRutaExe = aRuta;
        }

        protected virtual string mRegresarConsultaMovimientos(string aFuente, string lfolio, int ltipo)
        {
            string lregresa = "";
            switch (aFuente)
            {
                case "Flex":
                    lregresa = "select f.itemcode as ccodigop01,FCUnitPrice  as cprecioc01, " +
                    " BillTaxPerc as cporcent01,  '1' as ccodigoa01, p.ItemDesc  as cnombrep01, f.priceunitcode  as Unidad, " +
                    " IVAxLin as cimpuesto1, TotxLinea as cneto,  TotalxLineaIVA as ctotal, Cantidad as unidades   " +
                    " , isnull(f.itemdesc,'') as ctextoextra2, isnull(f.itemcode,'') as ctextoextra3, isnull(f.SHIPPINGRE,'') as creferen01 , isnull(f.CUSTITEMREF,'') as ctextoextra1" +
                    " from facturacione f join PM_Item p " +
                    " on f.ItemCode = p.ItemCode " +
                    " where f.billnum = " + lfolio;
                    break;
                case "Mercado":
                    lregresa = " select vd.articulo as ccodigop01,  " +
                    " Precio as cprecioc01, Cantidad as cunidades, vd.Impuesto1 as cimpuesto1,  Almacen as ccodigoa01, a.Descripcion1 as cnombrep01, a.Unidad " +
                    " from VentaD VD  join Art a  " +
                    " on VD.Articulo = a.Articulo   " +
                    " where ID = " + lfolio;

                    lregresa = " select vd.articulo as ccodigop01,  " +
                    " 'cprecioc01' =  case  " +
                    " when vd.impuesto1 <> 0 then round((Precio / (1 + (vd.impuesto1/100) )),4)  " +
                    " when vd.impuesto1 = 0 then Precio  " +
                    " end   " +
                    " , Cantidad as cunidades, vd.Impuesto1 as cimpuesto1,  Almacen as ccodigoa01, a.Descripcion1 as cnombrep01, a.Unidad, vd.impuesto1 as cPorcent01 " +
                    " from VentaD VD  join Art a  " +
                    " on VD.Articulo = a.Articulo   " +
                    " where ID = " + lfolio;



                    break;

            }
            return lregresa;
        }

        protected virtual Boolean mchecarvalido()
        {
            return true;
            //if (_RegDoctoOrigen.cFecha > DateTime.Parse("2011/08/01"))
            //    return false;
        }

        protected virtual string mModificaDatosCliente()
        {
            return "";
        }
        protected  string mModificaDatosClienteFlexo()
        {
            //return "";
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            string lrespuesta = "";
            long lidcliente = 0;
            miconexion.mAbrirConexionDestino();

            lsql.CommandText = "select * from mgw10002 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                _RegDoctoOrigen.cCodigoCliente = lreader["ccodigoc01"].ToString();
                _RegDoctoOrigen.cRazonSocial = lreader["crazonso01"].ToString();
                _RegDoctoOrigen.cRFC = lreader["cRFC"].ToString();
                _RegDoctoOrigen.cCond = lreader["cdiascre01"].ToString();
                lidcliente = long.Parse(lreader["cidclien01"].ToString());



                lsql.CommandText = "select * from mgw10001 where cidagente = " + lreader["cidagent01"].ToString();
                lreader.Close();
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    _RegDoctoOrigen.cAgente = lreader["ccodigoa01"].ToString();
                }

                // ahora checar si tiene direccion fiscal si no la tiene avisar, si la tiene asignarla de adminpaq
                lreader.Close();


                lsql.CommandText = "select * from mgw10011 where ctipocat01 = 1 and cidcatal01 = " + lidcliente + " and ctipodir01 = 0";
                //lreader.Close();
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    _RegDoctoOrigen._RegDireccion.cNombreCalle = lreader["cnombrec01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cNumeroExterior = lreader["cnumeroe01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cNumeroInterior = lreader["cnumeroi01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cColonia = lreader["ccolonia"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cCodigoPostal = lreader["ccodigop01"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cEstado = lreader["cestado"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cPais = lreader["cpais"].ToString().Trim();
                    _RegDoctoOrigen._RegDireccion.cCiudad = lreader["cciudad"].ToString().Trim();
                    lrespuesta = "";
                }

                else
                    lrespuesta = "Cliente sin direccion fiscal en ADMINPAQ";
            }
            else
                lrespuesta = "Cliente no existe en ADMINPAQ";

            lreader.Close();
            miconexion.mCerrarConexionDestino();
            return lrespuesta;


            //else
            //{
            //    lRespuesta = "El cliente no se ha dado de alta en Adminpaq "; // documento no encontrado
            //}

            //_con.Close();

            return "";
        }

        protected  string mLlenarDoctos(OleDbDataReader aReader)
        {
            _RegDoctos.Clear() ;
            string lfolio = "";
            aReader.Read();
            int lbandera = 1;
            while (lbandera == 1 && aReader.HasRows )
            {
                RegDocto x = new RegDocto ();
                List<RegMovto> movtos = new List<RegMovto> ();
                
                x.cAgente = "(Ninguno)";
                try
                {
                    x.cReferencia = aReader["cReferen01"].ToString();
                }
                catch (Exception dddd)
                {
 
                }
                x.cCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento");
                x.cFecha = DateTime.Parse (aReader["cfecha"].ToString());
                x.sMensaje = "";
                x.cMoneda = "Pesos";
                x.cTextoExtra1 = aReader["cObserva01"].ToString(); 

                string sfoliodocto = aReader["cfolio"].ToString();
                long lfoliodocto = 0 ;
                string lserie = "";
                try
                {
                    lfoliodocto = long.Parse( aReader["cfolio"].ToString());

                }
                catch (Exception eee)
                {
                    lserie = sfoliodocto.Substring(sfoliodocto.Length-1);
                    sfoliodocto = sfoliodocto.Substring(0, sfoliodocto.Length - 1);
                }

                
 
                x.cFolio = long.Parse (sfoliodocto );
                x.cSerie = lserie; 
                lfolio = aReader["cfolio"].ToString();
                while (lfolio == aReader["cfolio"].ToString())
                {
                    RegMovto mov = new RegMovto();
                    mov.cCodigoProducto = aReader["ccodigop01"].ToString();
                    mov.cNombreProducto = aReader["cnombrep01"].ToString();
                    mov.cCodigoAlmacen = "1";
                    mov.cUnidad = "PZA";
                    //" o.id_productos as ccodigop01, o.importe as cprecioc01, pr.nombreproducto as cnombrep01, o.cantidad as cunidades " +
                    mov.cUnidades = decimal.Parse(aReader["cunidades"].ToString());
                    mov.cPrecio = decimal.Parse(aReader["cprecioc01"].ToString());
                    movtos.Add(mov);
                    if (aReader.Read() == false)
                    {
                        lbandera = 0;
                        break;
                    }
                    //else
                        //lfolio = aReader["cfolio"].ToString();
                }
                x._RegMovtos = movtos;
                _RegDoctos.Add(x);
            }
            return "";

        }

        protected virtual string mLlenarDocto(OleDbDataReader aReader, int atipo, string aFolio, string aFuente)
        {
            string lrespuesta = "";
            string lfolio= "0";
            if (atipo == 1 || atipo == 2)
            {
                lfolio = aReader["cfolio"].ToString();
                _RegDoctoOrigen.cFolio  = long.Parse (lfolio);
            }
            if (aReader["cliente"].ToString() == string.Empty )
                return "Falta Codigo de cliente en documento " + aFolio ;
            else
                _RegDoctoOrigen.cCodigoCliente = aReader["cliente"].ToString();

            _RegDoctoOrigen.cFecha = DateTime.Parse(aReader["cfecha"].ToString());
            _RegDoctoOrigen.cFecha = DateTime.Parse(DateTime.Today.ToString ());
            if (mchecarvalido() == false)
                return "";


            
            //_RegDoctoOrigen.cFolio = long.Parse (aReader["cfolio"].ToString()) ;
            if (aReader["cRFC"].ToString() == string.Empty )
                return "Cliente sin RFC en documento " + aFolio;
            else
                if (!(aReader["cRFC"].ToString().Length == 12 ||  aReader["cRFC"].ToString().Length == 13))
                    return "El RFC tiene una longitud incorrecta en el documento " + aFolio;
                else
                    _RegDoctoOrigen.cRFC = aReader["cRFC"].ToString();
            
            
            if (atipo == 1)
            {
                _RegDoctoOrigen.cAgente = aReader["Agente"].ToString();
                _RegDoctoOrigen.cCond  = aReader["condpago"].ToString();

            }
            if (aReader["cRazonso01"].ToString() == string.Empty)
                return "Cliente sin Razon Social en documento " + aFolio;
            else
                _RegDoctoOrigen.cRazonSocial = aReader["cRazonso01"].ToString();

            //IsDBNull(
            //aReader["cTextoExtra1"].isnull
               // if(!aReader.IsDBNull(18))
                 //   _RegDoctoOrigen.cTextoExtra1 = aReader[18].ToString();



            // UNA modificacion que aplica para flexo es que los datos del cliente se toman de adminpaq
            lrespuesta = mModificaDatosCliente();
            //lrespuesta = mModificaDatosClienteFlexo();
            if (lrespuesta != string.Empty)
                return lrespuesta;


            _RegDoctoOrigen.cMoneda = aReader["Moneda"].ToString();
            _RegDoctoOrigen.cTipoCambio  = decimal.Parse (aReader["TipoCambio"].ToString());

            if (atipo != 1)
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();
            else
                _RegDoctoOrigen.cReferencia = aReader["cReferen01"].ToString();



            if (aReader["cnombrec01"].ToString().Trim() == string.Empty )
                _RegDoctoOrigen._RegDireccion.cNombreCalle = "Ninguna";
            else
                _RegDoctoOrigen._RegDireccion.cNombreCalle = aReader["cnombrec01"].ToString().Trim();

            _RegDoctoOrigen._RegDireccion.cNumeroExterior = aReader["cnumeroe01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cNumeroInterior = aReader["cnumeroi01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cColonia = aReader["ccolonia"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCodigoPostal = aReader["ccodigop01"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cEstado = aReader["cestado"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cPais = aReader["cpais"].ToString().Trim();
            _RegDoctoOrigen._RegDireccion.cCiudad = aReader["cciudad"].ToString().Trim();
            if (atipo == 3 || atipo == 4)
            {
                _RegDoctoOrigen.cNeto = double.Parse(aReader["importe"].ToString());
                _RegDoctoOrigen.cImpuestos = double.Parse(aReader["impuestos"].ToString().Trim());
            }

            
            OleDbCommand  lsql = new OleDbCommand ();
            OleDbDataReader   lreader;

            lsql.CommandText = mRegresarConsultaMovimientos(aFuente, lfolio, atipo );

            
            lsql.Connection = (OleDbConnection  )_con;
            aReader.Close();
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                while (lreader.Read())
                {
                    RegMovto lRegmovto = new RegMovto();
                    lRegmovto.cCodigoProducto = lreader["ccodigop01"].ToString();
                    lRegmovto.cNombreProducto = lreader["cnombrep01"].ToString();
                    lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                    lRegmovto.cPrecio = decimal.Parse(lreader["cprecioc01"].ToString());
                    
                    lRegmovto.cImpuesto = decimal.Parse(lreader["cimpuesto1"].ToString());
                    lRegmovto.cPorcent01 = decimal.Parse(lreader["cPorcent01"].ToString());
                    if (aFuente != "Mercado")
                    {
                        lRegmovto.cUnidades = decimal.Parse(lreader["unidades"].ToString());
                        lRegmovto.cTotal = decimal.Parse(lreader["cTotal"].ToString());
                        lRegmovto.cneto = decimal.Parse(lreader["cneto"].ToString());
                        lRegmovto.cReferencia = lreader["creferen01"].ToString();
                        lRegmovto.ctextoextra1 = lreader["ctextoextra1"].ToString();
                        lRegmovto.ctextoextra2 = lreader["ctextoextra2"].ToString();
                        lRegmovto.ctextoextra3 = lreader["ctextoextra3"].ToString();
                        
                    }
                    else
                        lRegmovto.cUnidades = decimal.Parse(lreader["cunidades"].ToString());
                    lRegmovto.cCodigoAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cNombreAlmacen = lreader["ccodigoa01"].ToString();
                    lRegmovto.cUnidad = lreader["unidad"].ToString();
                    _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
                }

            }
            else
            { 
                
            }
            lreader.Close();
            return lrespuesta;
            //miconexion.mCerrarConexionOrigen(); 
        }




        //public boolean mBuscar(long aFolio, long aIdDocum02)
        public Boolean  mBuscar(long aFolio, string aConcepto, string aSerie, int aTipo)
        {
            Boolean lRespuesta = false;
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader ;
            OleDbParameter lparametrofolio = new OleDbParameter ("@p2",aFolio );
            OleDbParameter lparametrodocumentode = new OleDbParameter("@p1", aConcepto);

            lsql.CommandText = "Select m2.ccodigoc01 as cliente,m6.ccodigoc01 as concepto, m6.cidconce01, m8.cfecha,m8.cfolio, m8.ciddocum01 " +
                " from mgw10008 m8 join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
                " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
                " and m6.ccodigoc01 =  '" + aConcepto + "'" +
                " where cfolio = " + aFolio +
            " and cseriedo01 = '" + aSerie + "'";
            //lsql.Parameters.Add(lparametrodocumentode);
            //lsql.Parameters.Add(lparametrofolio);
            if (aTipo==0)
                lsql.Connection = miconexion.mAbrirConexionOrigen();
            else
                lsql.Connection = miconexion.mAbrirConexionDestino();
            
            
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows )
            {
                lreader.Read ();
                //mLlenarDocto(lreader);

                lRespuesta = true;

            }
            miconexion.mCerrarConexionOrigen();
            lreader.Close();
            return lRespuesta ;


 

            

        }
        public string mGrabarDestinos()
        {
            ClassConexion miconexion = new ClassConexion();
            string lregresa = "";
            miconexion.aRutaExe = aRutaExe;
            miconexion.mAbrirConexionDestino (1); 
            lregresa = mGrabarCompra();
            if (lregresa == "")
            {
                miconexion.mAbrirConexionOrigen(1);
                // Grabar Factura
                mGrabarFactura();
            }

            return lregresa;
        }

        private string mGrabarCompra()
        {
            //barra.Avanzar();
            long lret;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoCompra");
            long lIdDocumento;

            RegProveedor lRegProveedor = new RegProveedor();
            lRegProveedor = mBuscarCliente(GetSettingValueFromAppConfigForDLL("Proveedor").ToString().Trim(), 1, 1);

            fInsertarDocumento();
            
            

            // lret = fSetDatoDocumento("cFecha", DateTime.Today.ToString()); 
            

            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);

            lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim());
            string lproveedor = GetSettingValueFromAppConfigForDLL("Proveedor");
            lret = fSetDatoDocumento("cCodigoCteProv", lproveedor);

            //buscar el rfc y la razon social de proveedor
            lret = fSetDatoDocumento("cRazonSocial", lRegProveedor.RazonSocial);
            lret = fSetDatoDocumento("cRFC", lRegProveedor.RFC);

            //lret = fSetDatoDocumento("cRazonSocial", ldr["crazonso01"].ToString());
            //lret = fSetDatoDocumento("cRFC", ldr["crfc"].ToString());
            lret = fSetDatoDocumento("cIdMoneda", "1");
            //barra.Avanzar();
            //lret = fSetDatoDocumento("cTipoCambio", z.Cells[21].Value.ToString());
            //lret = fSetDatoDocumento("cReferencia", z.Cells[18].Value.ToString());
            lret = fSetDatoDocumento("cFolio", GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim());
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            //lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieCompra"));
            DateTime lFecha;
            lFecha = DateTime.Today;
            
            string lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFecha); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFecha", lfechavenc );
            DateTime lFechaVencimiento;
            lFechaVencimiento = DateTime.Today.AddDays(lRegProveedor.DiasCredito);
            lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            lret = fGuardaDocumento();
            //barra.Avanzar();

            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1);
                return "El documento de compra ya existe con el folio y serie de la compra por lo que no se grabara";
            }
            // buscar el id del documento generado
            lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 1, GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()));
            long lNumeroMov = 100;
            string lregresa = "";
            lret = fInsertarMovimiento();
            productos = "";
            almacenes = "";

            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                
                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                if (lret != 0)
                {
                    lregresa += "@" + x.cCodigoProducto.Trim();
                    productos += x.cCodigoProducto.Trim();
                }

                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
                if (lret != 0)
                {
                    lregresa += "!" + x.cCodigoAlmacen.Trim();
                    almacenes += x.cCodigoAlmacen.Trim();
                }

            }
            if (lregresa == "")
            {
                decimal lprecioconmargen = 0;
                foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
                {
                    //barra.Avanzar();
                    lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                    lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                    if (lret != 0)
                        lregresa += "#&" + x.cCodigoProducto;

                    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);


                    lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());

                    //lprecioconmargen = x.cPrecio * (1 + (x.cMargenUtilidad / 100));

                    lprecioconmargen = x.cMargenUtilidad;

                    lret = fSetDatoMovimiento("cPrecio", lprecioconmargen.ToString());
                    //lret = fSetDatoMovimiento("cPorcentajeImpuesto1", z.Cells[17].Value.ToString());
                    //w = decimal.Parse(z.Cells[4].Value.ToString()) * decimal.Parse(z.Cells[6].Value.ToString());
                    lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());

                    lret = fGuardaMovimiento();
                    lNumeroMov += 100;
                    lret = fInsertarMovimiento();

                }
                long lrespuesta = 0;
                if (lret == 0)
                    lrespuesta = fAfectaDocto_Param(lCodigoConcepto, GetSettingValueFromAppConfigForDLL("SerieCompra").ToString().Trim(), double.Parse(_RegDoctoOrigen.cFolio.ToString()), true);
            }
            else
                fBorraDocumento();
            miconexion.mCerrarConexionDestino();
            //barra.Asignar(50);
            return lregresa;
        }

        private string mGrabarFactura()
        {
            //barra.Avanzar();
            //return "";
            long lret;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoFactura").ToString().Trim() ;
            long lIdDocumento;
            RegProveedor lRegProveedor = new RegProveedor ();
            lRegProveedor = mBuscarCliente(GetSettingValueFromAppConfigForDLL("Cliente").ToString ().Trim() , 0, 0);
            

            fInsertarDocumento();
            lret = fSetDatoDocumento("cFecha", DateTime.Today.ToString ()   );
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim());
            lret = fSetDatoDocumento("cCodigoCteProv", lRegProveedor.Codigo );
            lret = fSetDatoDocumento("cRazonSocial", lRegProveedor.RazonSocial );
            lret = fSetDatoDocumento("cRFC", lRegProveedor.RFC );
            lret = fSetDatoDocumento("cIdMoneda", "1");
            lret = fSetDatoDocumento("cTipoCambio", "1");
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            lret = fSetDatoDocumento("cFolio", GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()) ; 
            //lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim ()   );
            lret = fGuardaDocumento();
            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1); 
                return "El documento de factura ya existe con el folio y serie mostrados en pantalla";
            }

            // buscar el id del documento generado
            lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim()));

            RegDireccion lRegDireccion = new RegDireccion();
            // la direccion del cliente pasarla a la direccion de la factura
            lRegDireccion = mBuscarDireccion(lRegProveedor.Id ,0);
            
            if (!string.IsNullOrEmpty (lRegDireccion.cNombreCalle))
            {
                lret = fInsertaDireccion();
                lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString ());
                lret = fSetDatoDireccion("cTipoCatalogo", "3");
                lret = fSetDatoDireccion("cTipoDireccion", "0");
                lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle );
                lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior );
                lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior );
                lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia  );
                lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal  );
                lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado );
                lret = fSetDatoDireccion("cPais", lRegDireccion.cPais );
                lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad );
                lret = fGuardaDireccion();
            }

            
            long lNumeroMov = 100;

            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                //barra.Avanzar();
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());

                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto );
                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen );
                lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString () );
                lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString () );
                //lret = fSetDatoMovimiento("cPorcentajeImpuesto1", z.Cells[17].Value.ToString());
                //w = decimal.Parse(z.Cells[4].Value.ToString()) * decimal.Parse(z.Cells[6].Value.ToString());
                lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto .ToString());

                lret = fGuardaMovimiento();
                lNumeroMov += 100;

            }
            long lrespuesta = 0;
            if (lret == 0)
                lrespuesta = fAfectaDocto_Param(lCodigoConcepto, GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim(), double.Parse (_RegDoctoOrigen.cFolio.ToString()), true);
            miconexion.mCerrarConexionOrigen(1);
            //barra.Asignar(100);
            return "";
                    

        }

        //public long mBuscarDocumento(string aConcepto, long aFolio)
        //{
            //_RegDoctoOrigen.cFolio = aFolio;
            //'return mBuscarIdDocumento(aConcepto, 0);
        //}

        private long mBuscarIdDocumento(string aConcepto, int aTipo, string aSerie, long afolio)
        {
            OleDbConnection lconexion= new OleDbConnection ();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();

            string lcadena = "select m8.ciddocum01,m2.crazonso01, m2.crfc from mgw10008 m8 " +
            " join mgw10002 m2 on m8.cidclien01 = m2.cidclien01 " +
            " join mgw10006 m6 on m8.cidconce01 = m6.cidconce01 " +
            " where m6.ccodigoc01 = '" + aConcepto + "' and m8.cfolio = " + afolio.ToString() +
            " and cseriedo01 = '" + aSerie + "'";

            OleDbCommand lsql = new OleDbCommand (lcadena ,lconexion );
            OleDbDataReader lreader;
            long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                lIdDocumento = long.Parse(lreader["ciddocum01"].ToString());
                _rfc = lreader["crfc"].ToString(); 
                _razonsocial = lreader["crazonso01"].ToString();
            }
            lreader.Close();

            return lIdDocumento ;
 
        }
        private RegDireccion  mBuscarDireccion(long  aCliente, int aTipo)
        {
            string sql;
            OleDbConnection lconexion = new OleDbConnection();
            RegDireccion lreg = new RegDireccion ();
            lconexion = miconexion.mAbrirConexionOrigen();
            sql = "select * from mgw10011 where cidcatal01 = " + aCliente + 
                        " and ctipocat01 = 1 and ctipodir01 = " + aTipo ;
            OleDbCommand lsql = new OleDbCommand(sql, lconexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lreader.Read();
                lreg.cNombreCalle = lreader["cnombrec01"].ToString().Trim();
                lreg.cNumeroExterior = lreader["cnumeroe01"].ToString().Trim();
                lreg.cNumeroInterior = lreader["cnumeroi01"].ToString().Trim();
                lreg.cColonia = lreader["ccolonia"].ToString().Trim();
                lreg.cCodigoPostal = lreader["ccodigop01"].ToString().Trim();
                lreg.cEstado = lreader["cestado"].ToString().Trim();
                lreg.cPais = lreader["cpais"].ToString().Trim();
                lreg.cCiudad = lreader["cciudad"].ToString().Trim();
            }
            lreader.Close();

            return lreg ;

        }

        protected virtual  string GetSettingValueFromAppConfigForDLL(string aNombreSetting)
        {
            string lrutadminpaq = Directory.GetCurrentDirectory();
            if (Directory.GetCurrentDirectory() != aRutaExe)
                Directory.SetCurrentDirectory(aRutaExe);

            string value = "";
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            ClientSettingsSection userSettingsSection = (ClientSettingsSection)config.SectionGroups["userSettings"].Sections[_NombreAplicacion + ".Properties.Settings"];
            //SettingElement elemToDelete = null;

            foreach (SettingElement connStr in userSettingsSection.Settings)
            {
                if (connStr.Name == aNombreSetting)
                {
                    value = connStr.Value.ValueXml.InnerText;
                    break;
                }
            }
            if (lrutadminpaq != aRutaExe)
                Directory.SetCurrentDirectory( lrutadminpaq);
            return value;
        }

        public List<RegConcepto> mCargarConceptos(long aIdDocumentoDe, int aTipo)
        {
            List<RegConcepto > _RegFacturas = new List<RegConcepto >(); 
            OleDbConnection lconexion = new OleDbConnection();
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo
                OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01,cverfacele from mgw10006 where ciddocum01 = " + aIdDocumentoDe , lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegFacturas.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegConcepto lRegConcepto = new RegConcepto();
                        lRegConcepto.Codigo = lreader[0].ToString();
                        lRegConcepto.Nombre = lreader[1].ToString();
                        lRegConcepto.Tipocfd = lreader[2].ToString();
                        _RegFacturas.Add(lRegConcepto);
                    }
                }
                lreader.Close();
            }
            
            return _RegFacturas;

                  

        }


        public List<RegProveedor> mCargarClientes()
        {
            List<RegProveedor> _RegProveedores = new List<RegProveedor>();
            OleDbConnection lconexion = new OleDbConnection();
            
            lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {

                //OleDbCommand lsql = new OleDbCommand("select ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = " + aIdDocumentoDe + " and cescfd = 1 and cnombrec01 = 'CFDI'", lconexion);
                // este es para flexo

                //string lstring =  "select ccodigoc01,rtrim(crazonso01)+' ('+rtrim(ccodigoc01) + ')'"  +
                //" from mgw10002 where ctipocli01 < 2 and cidclien01 > 0";

                OleDbCommand lsql = new OleDbCommand("select ccodigoc01,rtrim(ccodigoc01)+' ('+rtrim(crazonso01) + ')'" +
                " from mgw10002 where ctipocli01 < 2 and cidclien01 > 0 order by ccodigoc01 ", lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                _RegProveedores.Clear();
                if (lreader.HasRows)
                {
                    while (lreader.Read())
                    {
                        RegProveedor lRegCliente = new RegProveedor();
                        lRegCliente.Codigo = lreader[0].ToString();
                        lRegCliente.RazonSocial = lreader[1].ToString();
                        //lRegCliente.Tipocfd = lreader[2].ToString();
                        _RegProveedores .Add(lRegCliente);
                    }
                }
                lreader.Close();
            }

            return _RegProveedores;



        }

        public RegProveedor mBuscarCliente(string aCliente, int aTipo, int aTipoCliente )
        {
            OleDbConnection lconexion = new OleDbConnection();
            RegProveedor lReg = new RegProveedor ();
            string lcadena;
            if (aTipo == 0)
                lconexion = miconexion.mAbrirConexionOrigen();
            else
                lconexion = miconexion.mAbrirConexionDestino();
            if (lconexion != null)
            {
                if (aTipoCliente == 0)
                    lcadena = "select ccodigoc01,crazonso01, cidclien01, crfc, cdiascre01 from mgw10002 where ctipocli01 < 2 and ccodigoc01 = '" + aCliente + "'";
                else
                    lcadena = "select ccodigoc01,crazonso01, cidclien01, crfc, cdiascre02 from mgw10002 where ctipocli01 > 1 and ccodigoc01 = '" + aCliente + "'";


                OleDbCommand lsql = new OleDbCommand(lcadena, lconexion);
                OleDbDataReader lreader;
                //long lIdDocumento = 0;
                lreader = lsql.ExecuteReader();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    lReg.Codigo = lreader[0].ToString();
                    lReg.RazonSocial = lreader[1].ToString();
                    lReg.Id = long.Parse(lreader[2].ToString());
                    lReg.RFC = lreader[3].ToString();
                    lReg.DiasCredito = int.Parse ( lreader[4].ToString());
                }
                lreader.Close();
            }
            return lReg;
                
                
        }

        public List<RegEmpresas> mCargarEmpresasAccess(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirConexionAccess(out amensaje);
            
            List<RegEmpresas> _RegEmpresas = new List<RegEmpresas >();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("SELECT distinct(Empresa) from tbl_puntosdeventa order by Empresa ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresas lRegEmpresas= new RegEmpresas();
                            lRegEmpresas.cEmpresa = lreader[0].ToString();
                            
                            _RegEmpresas.Add(lRegEmpresas);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }


            
            return _RegEmpresas;




        }



        public List<RegPuntodeVenta> mCargarPuntoVenta(string aEmpresa, out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = miconexion.mAbrirConexionAccess (out amensaje);

            List<RegPuntodeVenta> _RegPUntosVenta = new List<RegPuntodeVenta>();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("SELECT Nombre from tbl_puntosdeventa  where Empresa ='" + aEmpresa + "'", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegPUntosVenta.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegPuntodeVenta lRePuntodeVenta = new RegPuntodeVenta();
                            //lRePuntodeVenta.cEmpresa = lreader[0].ToString();
                            lRePuntodeVenta.cNombre  = lreader[0].ToString();
                            _RegPUntosVenta.Add(lRePuntodeVenta );
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }



            return _RegPUntosVenta ;




        }


        public List<RegEmpresa> mCargarEmpresas(out string amensaje)
        {
            
            OleDbConnection lconexion = new OleDbConnection();
            
            lconexion = miconexion.mAbrirRutaGlobal (out amensaje);

            List<RegEmpresa> _RegEmpresas = new List<RegEmpresa>();
            //amensaje = lconexion.ConnectionString;
            
            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {
                    
                    OleDbCommand lsql = new OleDbCommand("select cnombree01,crutadatos from mgw00001 where cidempresa > 1 ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresa lRegEmpresa = new RegEmpresa();
                            lRegEmpresa.Nombre = lreader[0].ToString();
                            lRegEmpresa.Ruta = lreader[1].ToString();
                            _RegEmpresas.Add(lRegEmpresa);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }
                
            }

                
            
            return _RegEmpresas;




        }

        public virtual bool mValidarConexionIntell(string aServidor, string aBd, string ausu, string apwd)
        {
            string Cadenaconexion = "data source =" + aServidor+ ";initial catalog =" + aBd   + ";user id = " + ausu  + "; password = " +  apwd  + ";";
            
            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        public virtual bool mValidarConexionIntell(string aRuta)
        {
            //string Cadenaconexion = "data source =" + aServidor + ";initial catalog =" + aBd + ";user id = " + ausu + "; password = " + apwd + ";";

            ClassConexion x = new ClassConexion();

            string lmsg = "'";

            //_con = miconexion.mAbrirConexionAccess (out lmsg);
            return true;

            /*
            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
             */
        }

        

        public virtual string mBuscarDoctos(long aFolio, long afoliofinal, int aTipo, Boolean aRevisar)
        {

            

            string lrespuesta = "";
            _RegDoctos.Clear();
            lrespuesta = mBuscarDoctoArchivo(aRevisar);
            return lrespuesta;
        /*

            for (long i = aFolio; i <= afoliofinal; i++)
            {
                RegDocto lDocto = new RegDocto();
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                lrespuesta = mBuscarDoctoAccess( aRevisar);
                if (lrespuesta == string.Empty)
                {
                    _RegDoctoOrigen.sMensaje = "";
                    _RegDoctoOrigen.cFolio = i;
                }
                else
                {
                    _RegDoctoOrigen.sMensaje = lrespuesta;
                    _RegDoctoOrigen.cFolio = i;
                }
                lDocto = _RegDoctoOrigen;
                _RegDoctos.Add(lDocto );
            }
            return lrespuesta;*/
        }

        
        public string mBuscarDocto (string aFolio, int aTipo, Boolean aRevisar)
        {
            OleDbCommand  lcmd = new OleDbCommand ();
            OleDbDataReader lreader;
            string lRespuesta= "";
            if (aTipo == 0)
                return lRespuesta;
            _con.Open();
            lcmd.Connection = _con;
            if (aTipo == 1 || aTipo == 2)
            {
                lcmd.CommandText = "select v.cliente as cliente, FechaEmision as cfecha, " +
                " ID as cfolio, c.Direccion as cnombrec01, c.DireccionNumero as cnumeroe01, c.DireccionNumeroInt as cnumeroi01, c.Colonia as ccolonia, c.Poblacion as cciudad, c.Estado as cestado, c.Pais as cpais " +
                " , c.RFC as crfc, c.Nombre as crazonso01, c.CodigoPostal as ccodigop01, v.moneda, v.tipocambio ";

                if (aTipo == 1)

                    lcmd.CommandText +=  ", case " +
                                        " when  v.condicion = '' then '0' " +
                                        " when  v.condicion = 'Contado' then '0'" + 
                                        " when  isnull(v.condicion,0) = '0' then '0'" + 
                                        " else left(v.condicion, isnull(charindex(' DIAS CREDITO',v.condicion,1),0)) " + 
                                        " end as condpago, v.agente " ;

                lcmd.CommandText += " from venta v join Cte c " +
                " on v.Cliente = c.Cliente " +
                "where MovID = '" + aFolio + "'" + 
                " and v.Estatus <> 'CANCELADO'";
            }
            if (aTipo == 3 || aTipo == 4)
            {
                lcmd.CommandText = "select " +
                " v.cliente as cliente, FechaEmision as cfecha, " +
                " v.MovID as cfolio, c.Direccion as cnombrec01, c.DireccionNumero as cnumeroe01, c.DireccionNumeroInt as cnumeroi01, c.Colonia as ccolonia, c.Poblacion as cciudad, c.Estado as cestado, c.Pais as cpais " +
                " , c.RFC as crfc, c.Nombre as crazonso01, c.CodigoPostal as ccodigop01, v.moneda, v.tipocambio, v.importe, v.impuestos " +
                " from cxc v  join Cte c " +
                " on v.Cliente = c.Cliente " +
                " where v.Estatus  = 'CONCLUIDO' " +
                " and MovID = '" + aFolio + "'";
            }


            switch (aTipo)
            {
                case 1:
                    lcmd.CommandText += " and (Mov = 'Factura' or Mov = 'Factura Global')";
                    break;
                case 2:
                    lcmd.CommandText += " and Mov = 'Devolucion Venta'";
                    break;
                case 3:
                    lcmd.CommandText += " and Mov = 'Nota Cargo'";
                    break;
                case 4:
                    lcmd.CommandText += " and Mov = 'Nota Credito'";
                    lcmd.CommandText += " and origen <> 'Devolucion Venta'";
                    break;

            }
                    

            
            lreader = lcmd.ExecuteReader();
            if (lreader.HasRows)
            {
                if (aRevisar == true)
                {
                    if (mBuscarGeneradoADM(aFolio, aTipo) == true)
                    {
                        _con.Close();
                        return "Documento ya existe en Adminpaq";
                    }
                }
                lreader.Read();
                lRespuesta = mLlenarDocto(lreader,aTipo, aFolio,"Mercado"  );
//                if (lRespuesta != string.Empty)
  //                  lRespuesta = "";
            }

             else
                {
                    //lreader.Read();
                    //mLlenarDocto(lreader);
                    lRespuesta = "Documento No Existe";
                }

            lreader.Close();
            
            _con.Close();
            return lRespuesta;
        }

        protected virtual string mConsultaEncabezado(int aTipo, string aFolio)
        {

            string aEmpresa = GetSettingValueFromAppConfigForDLL("Empresa");
            
            string aNombre = GetSettingValueFromAppConfigForDLL("Nombre");
            string aFecha = GetSettingValueFromAppConfigForDLL("Fecha");
            


            string lregresa = "";
            
            /*
                    lregresa = " select top 1 isnull(CustCode,'') as cliente, BillDate as cfecha, " +
                    " isnull(BillNum,0) as cfolio, isnull(billaddrname,'') as cnombrec01, isnull(billaddress2,'') as cnumeroe01, '' as cnumeroi01, " +
                    " isnull(billaddress3,'') as ccolonia, isnull(billtown,'') as cciudad, isnull(billcounty,'') as cestado, isnull(billcountry,'') as cpais " +
                    " , isnull(VatRegNo,'') as crfc, isnull(CustName,'') as crazonso01, isnull(BillPostCode,0) as ccodigop01, " +
                    " 'moneda' = case when currCode = 'MN' then 'Pesos' else '0' end,  " +
                    " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, isnull(OCCLIENTE,'') as creferen01 " +
                    " from facturacione " +
                    " where billnum = '" + aFolio + "'";
            */
            // "SELECT C.id_cliente as cliente, o.id_punto_de_venta as cfecha, o.fechamov, o.puntodeventa, p.nombre, o.cantidad, o.importe, pr.nombreproducto, o.id_productos, c.nombrefiscal, c.direccion, c.colonia, m.nombre, e.nombreestados, c.codigopostal " +
            // , c.nombrefiscal
            DateTime lfecha = DateTime.Parse(aFecha);
            string sfecha = lfecha.Month.ToString().Trim().PadLeft(2, '0') + "/" + lfecha.Day.ToString().Trim().PadLeft(2, '0') + "/" + lfecha.Year;
            aFecha = sfecha;
            lregresa = "SELECT C.id_cliente as cliente, o.fechamov as cfecha, 1 as cfolio, c.direccion as cnombrec01, c.celular as cnumeroe01, '' as cnumeroi01, " + 
                    " c.colonia as ccolonia, m.nombre as cciudad, e.nombreestados as cestado, 'Mexico' as cpais, " +
                    " c.fax as crfc, c.nombrefiscal as crazonso01,   c.codigopostal as ccodigop01, 'Pesos'  as moneda," +
                    " '1' as tipocambio, 0 as condpago, '(Ninguno)' as agente, '' as creferen01 " +
                   " FROM ((((tbl_operaciones AS o INNER JOIN tbl_puntosdeventa AS p ON o.id_punto_de_venta =p.id_pv02) INNER JOIN tbl_productos01 AS pr ON o.id_productos = pr.id_productos01) INNER JOIN tbl_clientes01 AS C ON p.empresa = C.nombrefiscal) INNER JOIN tbl_municipios AS m ON c.municipio = m.id_municipios) INNER JOIN tbl_estados AS e ON e.id_estados = c.estado " +
                   " WHERE o.id_punto_de_venta <> '' " +
                   " and o.id_punto_de_venta = p.id_pv02  " +
                   " and o.fechamov = #" + aFecha + "#" +
                   " and  " +
                   " p.Empresa ='" + aEmpresa + "'" +
                   " and p.nombre = '" + aNombre + "'" +
                   "ORDER BY o.fechamov DESC ";

            lregresa = "SELECT o.fechamov as cfecha, referencia_documento as cfolio, p.nombre as cReferen01, " +
                        " o.id_productos as ccodigop01, o.importe as cprecioc01, pr.nombreproducto as cnombrep01, o.cantidad as cunidades, o.observaciones01 as cobserva01 " +
                       "from  " + 
                        " (tbl_operaciones as o " + 
                        " INNER JOIN tbl_puntosdeventa AS p ON o.id_punto_de_venta =p.id_pv02) " + 
                        " INNER JOIN tbl_productos01 AS pr ON o.id_productos = pr.id_productos01 " + 
                        " where " + 
                        " o.id_punto_de_venta <> ''   " + 
                        " and  p.Empresa ='" + aEmpresa + "'" +
                        " and o.fechamov = #" + aFecha + "# " + 
                        " and referencia_documento <> '' " +
                        " ORDER BY referencia_documento, o.fechamov DESC ";
            
            return lregresa;

        }

 private string mProcesaItem(ref int aInicio,string sLine)
 {
     int lfin = sLine.IndexOf("|", aInicio) - aInicio;
     string lRegresa = sLine.Substring(aInicio, lfin);
     aInicio = sLine.IndexOf("|", aInicio) + 1;
     return lRegresa;
 }

 public string mBuscarDoctoArchivo(Boolean aRevisar)
 {
     string lrespuesta = ""; 
     string mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
     StringBuilder sb = new StringBuilder();
     List <RegDocto> misdoctos = new List<RegDocto>();
     _RegDoctos.Clear();
     string lrutacarpeta = @GetSettingValueFromAppConfigForDLL("RutaCarpeta");
     //lrutacarpeta = @lrutacarpeta;
     foreach (string txtName in Directory.GetFiles(lrutacarpeta , "*.txt"))
     {
         StreamReader objReader = new StreamReader(txtName);
         string sLine = "";
         
         ArrayList arrText = new ArrayList();
         RegDocto midocto = new RegDocto();
         int linicio = 0;
         while (sLine != null)
         {
             sLine = objReader.ReadLine();
             if (sLine != null)
             {
                 if (sLine != "")
                 {
                     if (sLine.Substring(0, 1) == "S")
                     {
                         linicio = 2;
                         string x  = mProcesaItem(ref linicio, sLine);
                         x = mProcesaItem(ref linicio, sLine);
                         x = mProcesaItem(ref linicio, sLine);
                         midocto.cNeto = double.Parse(x);
                         //midocto.cTextoExtra1 = "";
                         _RegDoctos.Add(midocto);
                         midocto = new RegDocto();
                     }
                     if (sLine.Substring(0, 2) == "H1")
                     {
                         //buscar contado
                         midocto.cContado = 0;
                         if (sLine.IndexOf("CONTADO") != -1)
                             midocto.cContado = 1;
                         linicio = 3;
                         midocto.cNombreArchivo = txtName;
                         midocto.cSerie = mProcesaItem(ref linicio, sLine);
                         midocto.cFolio = int.Parse(mProcesaItem(ref linicio, sLine));
                         midocto.cTextoExtra1 = mProcesaItem(ref linicio, sLine);
                         midocto.cMoneda = "Pesos";
                         midocto.cTipoCambio = 1;

                         // fecha
                         string y = mProcesaItem(ref linicio, sLine);
                         // 03/10/12
                         y = y.Substring(0, 6) + "20" + y.Substring(8, 2);
                         DateTime dt2 = DateTime.ParseExact(y, "dd/MM/yyyy", null);
                         midocto.cFecha = dt2;

                         //midocto.cTextoExtra1 = mProcesaItem(ref linicio, sLine);
                         string tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //cliente
                         midocto.cCodigoCliente = tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //moneda
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         midocto.cTextoExtra1 = tempo; // para movimiento

                         tempo = mProcesaItem(ref linicio, sLine);//carga y placas 
                         midocto.cReferencia = tempo.Substring(12, tempo.IndexOf("PLACAS") - 12);

                         tempo = tempo.Substring(tempo.IndexOf("PLACAS:") + 8);


                         midocto.cTextoExtra2 = tempo.Substring(0, tempo.IndexOf("RUTA"));

                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); 
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         midocto.cAgente = mProcesaItem(ref linicio, sLine); //agente
                         midocto.cCodigoConcepto = mProcesaItem(ref linicio, sLine); //concepto

                         midocto.cCodigoConcepto = midocto.cCodigoConcepto.Trim();



                     }
                     if (sLine.Substring(0, 1) == "D")
                     {
                         RegMovto movto = new RegMovto();
                         linicio = 2;
                         movto.cNombreProducto = mProcesaItem(ref linicio, sLine);

                         movto.cUnidades = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         //movto.cUnidades = decimal.Parse(mProcesaItem(ref linicio, sLine));
                         //MessageBox.Show(movto.cUnidades.ToString());
                         movto.cUnidad = mProcesaItem(ref linicio, sLine);

                         movto.cPrecio = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cneto = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         //movto.cPrecio = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cPorcent01 = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cImpuesto = decimal.Parse(mProcesaItem(ref linicio, sLine));

                         movto.cCodigoProducto = mProcesaItem(ref linicio, sLine);
                         movto.cCodigoAlmacen = "1";
                         movto.ctextoextra3 = midocto.cTextoExtra1;
                         
                         //MessageBox.Show(movto.cImpuesto.ToString());


                         midocto._RegMovtos.Add(movto);
                     }
                     if (sLine.Substring(0, 2) == "H2")
                     {
                         linicio = 3;
                         string tempo;
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //cliente
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine); //moneda
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         tempo = mProcesaItem(ref linicio, sLine);
                         midocto.cTextoExtra3 = mProcesaItem(ref linicio, sLine);
                     }
                     if (sLine.Substring(0, 2) == "H4")
                     {
                         linicio = 3;
                         midocto.cRazonSocial = mProcesaItem(ref linicio, sLine);
                         midocto.cRFC = mProcesaItem(ref linicio, sLine);
                         //midocto.cCodigoCliente = midocto.cRFC;
                         // linicio = lfin + 1;
                         //lfin = sLine.IndexOf("|", linicio) - linicio;
                         midocto._RegDireccion.cNombreCalle = mProcesaItem(ref linicio, sLine);
                         midocto._RegDireccion.cNumeroExterior = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cNumeroInterior = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cColonia = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cCiudad = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cEstado = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cPais = mProcesaItem(ref linicio, sLine); ;
                         midocto._RegDireccion.cCodigoPostal = mProcesaItem(ref linicio, sLine); ;
                         midocto.sMensaje = "";
                     }
                 }
             }
         }

         objReader.Close();    
     }
     
     return lrespuesta;
 }

        public  string mBuscarDoctoAccess(Boolean aRevisar)
        {
            OleDbCommand lcmd = new OleDbCommand();
            OleDbDataReader lreader;
            
             
            string lRespuesta = "";
            if (_con.State != 0)
                _con.Close();

            _con.Open();
            lcmd.Connection = _con;

            lcmd.CommandText = mConsultaEncabezado(0, "1");

            //OleDbDataAdapter lda = new OleDbDataAdapter(lcmd );
            //System.Data.DataSet xxx = new System.Data.DataSet ();
            //lda.Fill(xxx);

            

            

            try
            {
                lreader = null ;
                //lreader.Close();
                lreader = lcmd.ExecuteReader();
            }
            catch (Exception e)
            {
                lRespuesta = e.Message;
                _con.Close();
                return lRespuesta;
            }
            if (lreader.HasRows)
            {
                if (aRevisar == true)
                {

                    //if (mBuscarADM(aFolio, aTipo) == true)
                    //{
                    //    _con.Close();
                    //    return "Documento Ya existe en Adminpaq"; // documento ya existe
                    //}
                }
                //lreader.Read();
                lRespuesta = mLlenarDoctos(lreader);
            }

            else
            {
                lRespuesta = "Documento no Encontrado"; // documento no encontrado
            }

            _con.Close();
            return lRespuesta;
        }

        private Boolean mBuscarGeneradoADM(string aFolio, int aTipo )
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje="";
            lconexion = miconexion.mAbrirRutaGlobal(out amensaje);
            bool lrespuesta = false;

            //lconexion = miconexion.mAbrirConexionDestino();

            OleDbCommand lsql = new OleDbCommand("select * from interfaz where folioi = '" + aFolio + "' and tipodoc = " + aTipo, lconexion);
            OleDbDataReader lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lrespuesta = true;
            }
            lreader.Close();

            return lrespuesta ;
        }


        protected Boolean mBuscarADM(string aFolio, int aTipo)
        {
            bool lrespuesta = false;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();

            
            miconexion.mAbrirConexionDestino();
            string lcadena = "select cfolio from mgw10008 m8 join mgw10006 m6 on m6.cidconce01 = m8.cidconce01 where m8.cfolio = " + aFolio + " and m6.ccodigoc01 = '" + lCodigoConcepto + "'";
            OleDbCommand lsql = new OleDbCommand(lcadena, miconexion._conexion);
            OleDbDataReader  lreader;
            //long lIdDocumento = 0;
            lreader = lsql.ExecuteReader();
            if (lreader.HasRows)
            {
                lrespuesta = true;
            }
            lreader.Close();
            miconexion.mCerrarConexionDestino ();
            return lrespuesta;
        }
        public string mGrabarAdm1()
        {

            
                miconexion.mAbrirConexionDestino(1);
                bool lentre = true;
            
            //miconexion.mCerrarConexionDestino();
            miconexion.mCerrarConexionOrigen(1);
            _controlfp(0x9001F, 0xFFFFF); 
            // barra.Asignar(100);
            return "";
        }

        public List<string> mGrabarAdms(int opcion)
        {
            string lrespuesta = "";
            string lcadena =  "";

            List<string> lvar = new List<string>();

            int lcuantos = _RegDoctos.Count;
            int lindice = 1;

            if (_RegDoctos.Count == 0 )
            {
                lvar.Add("No existe documentos con los filtros seleccionados");
                return lvar; 
            }


            foreach (RegDocto _reg in _RegDoctos)
            {
                _RegDoctoOrigen = null;
                _RegDoctoOrigen = new RegDocto();
                _RegDoctoOrigen = _reg;
                string lCodigoConcepto ;
                if (opcion != 5)
                    lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                else
                    lCodigoConcepto = _reg.cCodigoConcepto ;

                //lrespuesta = _RegDoctoOrigen.sMensaje;
                //if (_RegDoctoOrigen.sMensaje == string.Empty)
                //{
                  lrespuesta = mGrabarAdm(_reg.cFolio.ToString(), _RegDoctoOrigen.cFolio, opcion);
                
                //}
                //mActualizarBarra((double)lindice / lcuantos);
                //lporcentaje = 0.0D;
                //lporcentaje = (double)lindice / lcuantos;
                //Notificar();
                Notificar((double)(lindice*100) / lcuantos);
                


                lindice++;
                if (lrespuesta != "")
                {
                    switch (opcion)
                    {
                        case 1:
                            lcadena = "La factura";
                            break;
                        case 2:
                            lcadena = "El pedido";
                            break;
                        case 3:
                            lcadena = "La nota de credito";
                            break;
                        case 4:
                            lcadena = "La nota de cargo";
                            break;
                    }
                    lcadena += " con folio " + _reg.cFolio.ToString() + " presento el siguiente problema " + lrespuesta + Convert.ToChar(13);

                    lvar.Add(lcadena);
                }
                else
                { 
                    //copiar el archivo
                    string lrutaorigen = GetSettingValueFromAppConfigForDLL("RutaCarpeta");
                    lrutaorigen += "\\"  + _RegDoctoOrigen.cNombreArchivo;
                    string lrutadestino = GetSettingValueFromAppConfigForDLL("RutaCarpetaBackup") ;
                    string larchivo = System.IO.Path.GetFileName(_RegDoctoOrigen.cNombreArchivo);
                    lrutadestino = lrutadestino + "\\" + larchivo;


                    File.Move(_RegDoctoOrigen.cNombreArchivo, lrutadestino);
                    
                }

                    

            }

            return lvar;
        }
          
        protected virtual void mActualizarBarra(double valor)
        {
            return;
            //lporcentaje = 0.0D;
                //lporcentaje = (double)lindice / lcuantos;
                //Notificar(lporcentaje);
        }

        private void mRegresarPrincipales(string lCodigoConcepto,ref long lidconce, ref long tipocfd, ref string cserie)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            lidconce = 0;
            tipocfd = 0;
            lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                cserie = lreader["cseriepo01"].ToString();
                lidconce = long.Parse(lreader["cidconce01"].ToString());
                tipocfd = long.Parse(lreader["cverfacele"].ToString());
            }
            else
                cserie = "";
            lreader.Close();
        }
        private string mValidarExisteDoc(long lidconce, string cserie, double afolionuevo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            lsql.CommandText = "select count(*) as cuantos from mgw10008 where cidconce01 = " + lidconce + " and cseriedo01 = '" + cserie + "' and cfolio = " + afolionuevo.ToString().Trim();
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                long cuantos = 0;
                cuantos = long.Parse(lreader["cuantos"].ToString());
                lreader.Close();
                if (cuantos > 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);

                    return "Documento ya existe en ADMINPAQ";
                }
            }
            lreader.Close();
            return "";
        }

        private string mGrabarEncabezado(double aFolio, string lCodigoConcepto)
        {
            long lret,lidconce=0,tipocfd=0;
            string cserie="";

            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            
            mRegresarPrincipales(lCodigoConcepto,ref lidconce, ref tipocfd, ref cserie);
            string lresp = mValidarExisteDoc(lidconce,cserie,aFolio);
            if (lresp != "")
                return lresp;

            fInsertarDocumento();
            lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);
            if (_RegDoctoOrigen.cSerie != "")
                lret = fSetDatoDocumento("cSerieDocumento", _RegDoctoOrigen.cSerie);
            else
                lret = fSetDatoDocumento("cSerieDocumento", "");
//            lret = fSetDatoDocumento("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            lresp = "";
            if (lret!=0)
            {
                lresp=mGrabarCliente();
            }
            if (lresp != "")
                return lresp;

            lret = fSetDatoDocumento("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
            lret = fSetDatoDocumento("cRFC", _RegDoctoOrigen.cRFC);
            if (_RegDoctoOrigen.cMoneda != "Pesos")
                lret = fSetDatoDocumento("cIdMoneda", "2");
            else
                lret = fSetDatoDocumento("cIdMoneda", "1");
            lret = fSetDatoDocumento("cTipoCambio", _RegDoctoOrigen.cTipoCambio.ToString());
            lret = fSetDatoDocumento("cReferencia", "Por Programa");
            //lret = fSetDatoDocumento("cObservaciones", _RegDoctoOrigen.cTextoExtra1 );
            lret = fSetDatoDocumento("cFolio", _RegDoctoOrigen.cFolio.ToString().Trim());


            try
            {
                //lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cFolio.ToString());
                //lret = fSetDatoDocumento("cTextoExtra1", _RegDoctoOrigen.cReferencia);
                lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cReferencia);
                lret = fSetDatoDocumento("cTextoExtra2", _RegDoctoOrigen.cTextoExtra2);
                lret = fSetDatoDocumento("cTextoExtra3", _RegDoctoOrigen.cTextoExtra3);
            }
            catch (Exception ee)
            {
            }

         
            DateTime lFechaVencimiento;
            lFechaVencimiento = _RegDoctoOrigen.cFecha.AddDays(int.Parse("0"));
            //lFechaVencimiento = DateTime.Today.AddDays(int.Parse(_RegDoctoOrigen.cCond) );

            string lfechavenc = "";
            lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            lret = fSetDatoDocumento("cCodigoAgente", _RegDoctoOrigen.cAgente );
            /*
            if (lret != 0)
            {
                miconexion.mCerrarConexionOrigen(1);
                _controlfp(0x9001F, 0xFFFFF);
                // barra.Asignar(100);
                return "Fecha Incorrecta";
            }
             * 
             */


            

            string lfechadocto = "";
            lfechadocto = _RegDoctoOrigen.cFecha.ToString();
            DateTime lFechaDocto;
            lFechaDocto = _RegDoctoOrigen.cFecha;

            lfechadocto = "";


            lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year
            lret = fSetDatoDocumento("cFecha", lfechadocto);

            lret = fSetDatoDocumento("cFechaVencimiento", lfechadocto);
            lret = fSetDatoDocumento("cTipoCambio", "1");
            lret = fGuardaDocumento();
            string serror= "";
            if (lret != 0)
            {
                //fError(lret, serror, 255);
                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return lret.ToString()+ " Documento ya Existe";
                

            }
            return "";



        }

        private string  mGrabarCliente()
        { 
            long lret = 0;
            fInsertaCteProv();
                lret = fSetDatoCteProv("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
                lret = fSetDatoCteProv("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRazonSocial;
                }
                lret = fSetDatoCteProv("cRFC", _RegDoctoOrigen.cRFC);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRFC;
                }
                lret = fSetDatoCteProv("CLISTAPRECIOCLIENTE", "1");
                lret = fSetDatoCteProv("CIDMONEDA", "1");

                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");
                lret = fSetDatoCteProv("CFECHAALTA", lfecha);
                if (lret != 0)
                {
                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cFecha.ToString();
                }
                lret = fSetDatoCteProv("CTIPOCLIENTE", "1");
                lret = fSetDatoCteProv("CESTATUS", "1");
                lret = fSetDatoCteProv("CIDADDENDA", "-1");

                lret = fSetDatoCteProv("CEMAIL1", _RegDoctoOrigen._RegDireccion.cEmail);
                lret = fSetDatoCteProv("CEMAIL2", _RegDoctoOrigen._RegDireccion.cEmail2);
                lret = fSetDatoCteProv("CBANCFD", "1");
                lret = fSetDatoCteProv("CTIPOENTRE", "6");





                lret = fGuardaCteProv();
                if (lret == 0)
                    return "";
                else
                    return "Error dar de alta Cliente";
        }

        public string mGrabarDireccion(long lIdDocumento)
        {
            long lret=0;
            mLeerDireccion();

            RegDireccion lRegDireccion = new RegDireccion();
            // la direccion del cliente pasarla a la direccion de la factura
            lRegDireccion = _RegDoctoOrigen._RegDireccion;
            if (lRegDireccion.cNombreCalle != null )
            {
                lret = fInsertaDireccion();
                lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString());
                lret = fSetDatoDireccion("cTipoCatalogo", "3");
                lret = fSetDatoDireccion("cTipoDireccion", "0");
                lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle);
                if (lRegDireccion.cNumeroExterior == string.Empty)
                    lret = fSetDatoDireccion("cNumeroExterior", "0");
                else
                    lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior);
                lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior);
                lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia);
                lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal);
                lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado);
                lret = fSetDatoDireccion("cPais", lRegDireccion.cPais);
                lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad);
                lret = fSetDatoDireccion("cEmail", lRegDireccion.cEmail);
                lret = fGuardaDireccion();
                if (lret != 0)
                {

                    _controlfp(0x9001F, 0xFFFFF);
                    miconexion.mCerrarConexionOrigen(1);
                    return "Se presento el error direccion" + lret.ToString();

                }
            }
            return "";
        }


        public string mGrabarAdm(string afolioant, double afolionuevo, int opcion)
        {
            miconexion.mAbrirConexionDestino(1);
            string lCodigoConcepto;
            //lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
            
            string lresp1 = mGrabarEncabezado(afolionuevo, lCodigoConcepto);
            if (lresp1 != "")
                return lresp1;

            //OleDbCommand lsql = new OleDbCommand();
            //OleDbDataReader lreader;
            
            string cserie;
            cserie = _RegDoctoOrigen.cSerie;
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));


            if (lIdDocumento == 0)
            {

                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return "no se encontro documento " + lCodigoConcepto + " " +
                    long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim());
            }

            string lresp = mGrabarDireccion(lIdDocumento);
            lresp = mGrabarMovimientos(lIdDocumento, opcion);

            string lrespuestas = mGrabarExtras(lIdDocumento,2,afolionuevo);

            int lret = fAfectaDocto_Param(lCodigoConcepto, cserie, afolionuevo, true);
            string lCodigoConceptoPago = "";
            if (_RegDoctoOrigen.cContado == 1)
            {
                lCodigoConceptoPago = "10";
            }
            if (_RegDoctoOrigen.cCodigoConcepto == "503")
            {
                lCodigoConceptoPago = "504";
            }
            if (lCodigoConceptoPago != "")
            {
                mGrabarEncabezado(afolionuevo, lCodigoConceptoPago);
                lIdDocumento = mBuscarIdDocumento(lCodigoConceptoPago, 0, cserie, long.Parse(afolionuevo.ToString().Trim()));
                _RegDoctoOrigen._RegMovtos.Clear();
                lresp = mGrabarMovimientos(lIdDocumento, 3);
                lrespuestas = mGrabarExtras(lIdDocumento, 3, afolionuevo);

                string lfechavenc = "";
                lfechavenc = String.Format("{0:MM/dd/yyyy}", _RegDoctoOrigen.cFecha); ;  // "8 08 008 2008"   year


                double importe = _RegDoctoOrigen.cNeto;
                string otroconcepto = "10";
                string sFolo = afolionuevo.ToString();

                lret = fAfectaDocto_Param("10", cserie, afolionuevo, true);
                long lret1 = fSaldarDocumento_Param (lCodigoConcepto, cserie, afolionuevo,
otroconcepto, cserie, afolionuevo, importe, 1, lfechavenc);
            }
            



            miconexion.mCerrarConexionOrigen(1);
            //miconexion.mCerrarConexionDestino(1);

            try
            {
                _controlfp(0x9001F, 0xFFFFF);
            }
            catch (Exception eee)
            {
                lrespuestas = eee.Message;
            }
            // barra.Asignar(100);
            return lrespuestas;
        }

        private string mGrabarExtras(long lIdDocumento,int opcion, double afolionuevo)
        {
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            
            string lresp= "";
            long lrespuesta = 0;
            string lrespuestas = "";
            if (lresp == "")
            {
                //double x = double.Parse(afolionuevo.ToString () );
                if (opcion == 1)
                    mActualizaDocumento(lIdDocumento, opcion, afolionuevo);

                lsql.CommandText = "select alltrim(cnombrec01) + ', '" +
                                    " + alltrim(cnumeroe01) + ', '" +
                                    " + alltrim(ccolonia) + ', '" +
                                    " + alltrim(cciudad) + ', '" +
                                    " + alltrim(cestado) + ', '" +
                                    " + alltrim(cpais) " +
                                   " from mgw10011 where ctipocat01 = 4";

                //miconexion.mAbrirConexionDestino();
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                string ldireccion = "";
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ldireccion = lreader[0].ToString().Trim();
                    lreader.Close();
                }



                string lcadena2 = "update mgw10008 set cobserva01 = '" + _RegDoctoOrigen.cTextoExtra1 + "', clugarexpe = '" + ldireccion.Trim() + "'  where ciddocum01 = " + lIdDocumento;

                OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                lsql4.ExecuteNonQuery();
                //lrespuesta = fAfectaDocto_Param(lCodigoConcepto, cserie , x, true);

                if (opcion == 1 && _RegDoctoOrigen.cTipoCambio != 1)
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    string lcadena1 = "update mgw10008 set ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + "  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    // lsql3.ExecuteNonQuery();



                }

                /* actualizar observaciones del movimiento */

                string lcadenaA = "update mgw10010 set cobserva01= '" + _RegDoctoOrigen.cTextoExtra1 + "' where ciddocum01 = " + lIdDocumento;
                    OleDbCommand lsqlA= new OleDbCommand(lcadenaA, miconexion._conexion);
                    lsqlA.ExecuteNonQuery();
                


                
                if (_RegDoctoOrigen._RegDireccion.cEmail != "")
                {

                    string lcadena21 = "update mgw10002 set cemail1 = '" + _RegDoctoOrigen._RegDireccion.cEmail + "', cemail2 = '" + _RegDoctoOrigen._RegDireccion.cEmail2 + "', cbancfd = 1, ctipoentre=6 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
                    OleDbCommand lsql212 = new OleDbCommand(lcadena21, miconexion._conexion);
                    lsql212.ExecuteNonQuery();
                }


                if (opcion == 3 || opcion == 4)
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    decimal limpuestos = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
                    limpuestos = decimal.Round(limpuestos, 4);
                    string lcadena1 = "update mgw10008 set cneto = " + _RegDoctoOrigen.cNeto.ToString() + ", cimpuesto1 = " + limpuestos + ",ctotal = " + ltotal.ToString() + ",cpendiente = " + ltotal.ToString() + ",ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + " where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    lsql3.ExecuteNonQuery();
                    //miconexion.mCerrarConexionDestino();
                }
                //miconexion.mCerrarConexionDestino();

                //mGrabarInterfaz(afolioant, opcion);
            }
            else
            {
                lrespuestas = "ocurrio error";
            }
            //miconexion.mCerrarConexionDestino ();
            return lrespuestas;
            // antes de cerrar grabar en la tabla de interfaz

        }


        public string mGrabarAdm5(string afolioant, double afolionuevo, int opcion)
        {

            //mGrabarEncabezado(afolionuevo);


            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            string cserie;

            long lret;
            string lCodigoConcepto;
                lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            
            miconexion.mAbrirConexionDestino(1);
            
            //long lidconce = 0;
            //long tipocfd = 0;
            //lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            //lsql.Connection = miconexion._conexion;
            //lreader = lsql.ExecuteReader();
            ////_RegDoctoOrigen._RegMovtos.Clear();
            //if (lreader.HasRows)
            //{
            //    lreader.Read();
            //    cserie = lreader["cseriepo01"].ToString();
            //    lidconce = long.Parse(lreader["cidconce01"].ToString());
            //    tipocfd = long.Parse(lreader["cverfacele"].ToString());
            //}
            //else
            //    cserie = "";
            //lreader.Close();


            //lsql.CommandText = "select count(*) as cuantos from mgw10008 where cidconce01 = " + lidconce + " and cseriedo01 = '" + cserie + "' and cfolio = " + afolionuevo.ToString().Trim();
            //lsql.Connection = miconexion._conexion;
            //lreader = lsql.ExecuteReader();
            ////_RegDoctoOrigen._RegMovtos.Clear();
            //if (lreader.HasRows)
            //{
            //    lreader.Read();
            //    long cuantos=0;
            //    cuantos = long.Parse(lreader["cuantos"].ToString());
            //    lreader.Close();
            //    if (cuantos > 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Documento ya existe en ADMINPAQ";
            //    }
            //}
            //lreader.Close();



            //fInsertarDocumento();
            //lret = fSetDatoDocumento("cCodigoConcepto", lCodigoConcepto);

            
            //if (_RegDoctoOrigen.cSerie != "")
            //    lret = fSetDatoDocumento("cSerieDocumento",  _RegDoctoOrigen.cSerie  );
            //else
            //    lret = fSetDatoDocumento("cSerieDocumento", "");
            //lret = fSetDatoDocumento("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            //lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            //if (lret != 0)
            //{
            //    fInsertaCteProv();
            //    lret = fSetDatoCteProv("CCODIGOCLIENTE", _RegDoctoOrigen.cCodigoCliente);
            //    lret = fSetDatoCteProv("cRazonSocial", _RegDoctoOrigen.cRazonSocial);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRazonSocial;
            //    }
            //    lret = fSetDatoCteProv("cRFC", _RegDoctoOrigen.cRFC);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cRFC;
            //    }
            //    lret = fSetDatoCteProv("CLISTAPRECIOCLIENTE", "1");
            //    lret = fSetDatoCteProv("CIDMONEDA", "1");

            //    string lfecha = _RegDoctoOrigen.cFecha.ToString();
            //    DateTime ldate = DateTime.Parse(lfecha);
            //    lfecha = ldate.ToString("MM/dd/yyyy");
            //    lret = fSetDatoCteProv("CFECHAALTA", lfecha);
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cFecha.ToString();
            //    }
            //    lret = fSetDatoCteProv("CTIPOCLIENTE", "1");
            //    lret = fSetDatoCteProv("CESTATUS", "1");
            //    lret = fSetDatoCteProv("CIDADDENDA", "-1");

            //    lret = fSetDatoCteProv("CEMAIL1", _RegDoctoOrigen._RegDireccion.cEmail);
            //    lret = fSetDatoCteProv("CEMAIL2", _RegDoctoOrigen._RegDireccion.cEmail2);
            //    lret = fSetDatoCteProv("CBANCFD", "1");
            //    lret = fSetDatoCteProv("CTIPOENTRE", "6");





            //    lret = fGuardaCteProv();
            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        bool sigue = false;
            //        sigue = mDarAltaCliente();
            //        if (sigue == true)
            //            lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);
            //        else
            //        {
            //            miconexion.mCerrarConexionOrigen(1);
            //            return "Se presento el error en clientes111 " + lret.ToString();
            //        }

            //    }
            //    else
            //        lret = fSetDatoDocumento("cCodigoCteProv", _RegDoctoOrigen.cCodigoCliente);

            //    if (lret != 0)
            //    {
            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error en clientes " + lret.ToString() + _RegDoctoOrigen.cCodigoCliente;
            //    }
            
            //}
            /*
            if (primerdocto.cCodigoCliente == null)
            {
                mModificaDatosClienteFlexo();
                primerdocto.cCodigoCliente = _RegDoctoOrigen.cCodigoCliente;
                primerdocto.cRazonSocial = _RegDoctoOrigen.cRazonSocial;
                primerdocto.cRFC = _RegDoctoOrigen.cRFC;
                primerdocto.cCond = _RegDoctoOrigen.cCond;
                primerdocto.cAgente = _RegDoctoOrigen.cAgente;
                primerdocto._RegDireccion = _RegDoctoOrigen._RegDireccion;
            }

            else
            {
                _RegDoctoOrigen.cCodigoCliente = primerdocto.cCodigoCliente;
                _RegDoctoOrigen.cRazonSocial = primerdocto.cRazonSocial;
                _RegDoctoOrigen.cRFC = primerdocto.cRFC;
                _RegDoctoOrigen.cCond = primerdocto.cCond;
                _RegDoctoOrigen.cAgente = primerdocto.cAgente;
                _RegDoctoOrigen._RegDireccion = primerdocto._RegDireccion;

            }
             */
            
            
            //lret = fSetDatoDocumento("cRazonSocial", _RegDoctoOrigen.cRazonSocial );
            //lret = fSetDatoDocumento("cRFC", _RegDoctoOrigen.cRFC );
            //if (_RegDoctoOrigen.cMoneda != "Pesos") 
            //    lret = fSetDatoDocumento("cIdMoneda", "2");
            //else
            //    lret = fSetDatoDocumento("cIdMoneda", "1");
            //lret = fSetDatoDocumento("cTipoCambio", _RegDoctoOrigen.cTipoCambio.ToString ());
            //lret = fSetDatoDocumento("cReferencia", "Por Programa");
            ////lret = fSetDatoDocumento("cObservaciones", _RegDoctoOrigen.cTextoExtra1 );
            //lret = fSetDatoDocumento("cFolio", _RegDoctoOrigen.cFolio.ToString().Trim());


            //try
            //{
            //    lret = fSetDatoDocumento("cReferencia", _RegDoctoOrigen.cFolio.ToString ());
            //    lret = fSetDatoDocumento("cTextoExtra1", _RegDoctoOrigen.cReferencia);
            //}
            //catch (Exception ee)
            //{ 
            //}

            ////lret = fSetDatoDocumento("cEsCFD", "1");
            ////lret = fSetDatoDocumento("cSerieDocumento", GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim ()   );
            //string lfechavenc = "";
            //if (opcion == 1)
            //{
            //    DateTime lFechaVencimiento;
            //    lFechaVencimiento = _RegDoctoOrigen.cFecha.AddDays(int.Parse("0"));

            //    //lFechaVencimiento = DateTime.Today.AddDays(int.Parse(_RegDoctoOrigen.cCond) );
                
            //    lfechavenc = "";
            //    lfechavenc = String.Format("{0:MM/dd/yyyy}", lFechaVencimiento); ;  // "8 08 008 2008"   year
            //    lret = fSetDatoDocumento("cFechaVencimiento", lfechavenc);
            //    //lret = fSetDatoDocumento("cCodigoAgente", _RegDoctoOrigen.cAgente );
            //    if (lret != 0)
            //    {
            //        miconexion.mCerrarConexionOrigen(1);
            //        _controlfp(0x9001F, 0xFFFFF);
            //        // barra.Asignar(100);
            //        return "Agente no existe";
            //    }


            //}
            ////lret = fSetDatoDocumento("cImpuesto1", _RegDoctoOrigen.cImpuestos.ToString ());
            

            //string lfechadocto = "";
            //lfechadocto = _RegDoctoOrigen.cFecha.ToString();
            //DateTime lFechaDocto;
            //lFechaDocto = _RegDoctoOrigen.cFecha;

            //lfechadocto = "";
            

            //lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  // "8 08 008 2008"   year
            ////if (opcion == 3 || opcion == 4)
            ////    lfechadocto = String.Format("{0:MM/dd/yyyy}", DateTime.Today); ;  // "8 08 008 2008"   year

            ////lfechadocto = String.Format("{0:MM/dd/yyyy}", lFechaDocto); ;  
            
            //lret = fSetDatoDocumento("cFecha", lfechadocto);

            //if (opcion == 1)
            //    lret = fSetDatoDocumento("cFechaVencimiento", lfechadocto);
            //lret = fSetDatoDocumento("cTipoCambio", "1");
            //lret = fGuardaDocumento();
            //if (lret != 0)
            //{

            //    _controlfp(0x9001F, 0xFFFFF); 
            //    miconexion.mCerrarConexionOrigen(1);
            //    return "Se presento el error " + lret.ToString () ;
                
            //}

            //lret = fSetDatoDocumento("cCodigoConcepto", "10");
            //lret = fGuardaDocumento();
            

            // buscar el id del documento generado
            //long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, GetSettingValueFromAppConfigForDLL("SerieDestino").ToString().Trim(), long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim()));
            cserie = _RegDoctoOrigen.cSerie;
            long lIdDocumento = mBuscarIdDocumento(lCodigoConcepto, 0, cserie , long.Parse (afolionuevo.ToString().Trim()));
                

            if (lIdDocumento == 0)
            {

                _controlfp(0x9001F, 0xFFFFF);
                miconexion.mCerrarConexionOrigen(1);
                return "no se encontro documento " + lCodigoConcepto + " " +
                    long.Parse(GetSettingValueFromAppConfigForDLL("FolioDestino").ToString().Trim()); 
//                +" " + 
  //                  lret.ToString();

            }

            string lresp = mGrabarDireccion(lIdDocumento);

            //mLeerDireccion();

            //RegDireccion lRegDireccion = new RegDireccion();
            //// la direccion del cliente pasarla a la direccion de la factura
            //lRegDireccion = _RegDoctoOrigen._RegDireccion;
            //if (lRegDireccion.cNombreCalle != null )
            //{
            //    lret = fInsertaDireccion();
            //    lret = fSetDatoDireccion("cIdCatalogo", lIdDocumento.ToString());
            //    lret = fSetDatoDireccion("cTipoCatalogo", "3");
            //    lret = fSetDatoDireccion("cTipoDireccion", "0");
            //    lret = fSetDatoDireccion("cNombreCalle", lRegDireccion.cNombreCalle);
            //    if (lRegDireccion.cNumeroExterior == string.Empty)
            //        lret = fSetDatoDireccion("cNumeroExterior", "0");
            //    else
            //        lret = fSetDatoDireccion("cNumeroExterior", lRegDireccion.cNumeroExterior);
            //    lret = fSetDatoDireccion("cNumeroInterior", lRegDireccion.cNumeroInterior);
            //    lret = fSetDatoDireccion("cColonia", lRegDireccion.cColonia);
            //    lret = fSetDatoDireccion("cCodigoPostal", lRegDireccion.cCodigoPostal);
            //    lret = fSetDatoDireccion("cEstado", lRegDireccion.cEstado);
            //    lret = fSetDatoDireccion("cPais", lRegDireccion.cPais);
            //    lret = fSetDatoDireccion("cCiudad", lRegDireccion.cCiudad);
            //    lret = fSetDatoDireccion("cEmail", lRegDireccion.cEmail);
            //    lret = fGuardaDireccion();
            //    if (lret != 0)
            //    {

            //        _controlfp(0x9001F, 0xFFFFF);
            //        miconexion.mCerrarConexionOrigen(1);
            //        return "Se presento el error direccion" + lret.ToString();

            //    }
            //}

            lresp = mGrabarMovimientos(lIdDocumento,opcion);


            //long lNumeroMov = 100;
            //if (_RegDoctoOrigen._RegMovtos.Count == 0  && (opcion ==3 || opcion == 4))
            //{

            //    RegMovto lRegmovto = new RegMovto();
            //    lRegmovto.cCodigoProducto = "(Ninguno)";
            //    lRegmovto.cNombreProducto = "(Ninguno)";
            //    lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
            //    lRegmovto.cSubtotal  = decimal.Parse(_RegDoctoOrigen.cNeto.ToString () ); 
            //    //lRegmovto.cTotal = decimal.Parse(lreader["cunidades"].ToString());
            //    lRegmovto.cImpuesto = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
            //    lRegmovto.cTotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
            //    lRegmovto.cCodigoAlmacen = "(Ninguno)";
            //    lRegmovto.cNombreAlmacen = "(Ninguno)";
            //    lRegmovto.cUnidad = "";
            //    lRegmovto.cUnidades = 1;
            //    lRegmovto.cReferencia = "";
            //    lRegmovto.ctextoextra1 = "";
            //    lRegmovto.ctextoextra2 = "";
            //    lRegmovto.ctextoextra3 = "";
            //    _RegDoctoOrigen._RegMovtos.Add(lRegmovto); 
            //}
            //foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            //{
            //    lret = fInsertarMovimiento();
            //    lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
            //    lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());
            //    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
            //    if (lret != 0)
            //    {
            //        fInsertaProducto();
            //        lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
            //        lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto );
            //        lret = fSetDatoProducto("CTIPOPRODUCTO", "1");
            //        lret = fSetDatoProducto("CMETODOCOSTEO", "1");
            //        lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
            //        lret = fSetDatoProducto("CIMPUESTO1", x.cPorcent01.ToString () );
            //        OleDbCommand cmdunidad = new OleDbCommand();
            //        cmdunidad.CommandText = "select * from mgw10026 where cnombreu01 = '" + x.cUnidad.ToUpper() + "'";
            //        miconexion.mAbrirConexionDestino();
            //        cmdunidad.Connection = miconexion._conexion  ;
            //        OleDbDataReader ldr = cmdunidad.ExecuteReader() ;
            //        int lidunidad ;
            //        if (ldr.HasRows == false)
            //        {
            //            ldr.Read();
            //            ldr.Close();
            //            lret = fSetDatoProducto("CCODIGOUNIDADBASE", x.cUnidad.ToUpper());
            //            if (lret != 0)
            //            {
            //                // dar de alta la unicad de medida y peso
            //                cmdunidad.CommandText = "select max(cidunidad) + 1 from mgw10026";
            //                ldr = cmdunidad.ExecuteReader();
            //                ldr.Read();

            //                 lidunidad = int.Parse(ldr[0].ToString());
            //                ldr.Close();
            //                cmdunidad.CommandText = "insert into mgw10026 values (" + lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','')";
            //                cmdunidad.ExecuteNonQuery();
            //                lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
            //            }
            //        }
            //        else
            //        {
            //            ldr.Read();
                        
            //            lidunidad = int.Parse(ldr[0].ToString());
            //            ldr.Close();
            //            lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
            //        }
            //        lret = fGuardaProducto();
            //        lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
            //    }
            //    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
            //    if (lret != 0)
            //    {
            //        fInsertaAlmacen();
            //        lret = fSetDatoAlmacen("CCODIGOALMACEN", x.cCodigoAlmacen);
            //        lret = fSetDatoAlmacen("CNOMBREALMACEN", x.cNombreAlmacen);
            //        lret = fGuardaAlmacen();
            //        lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);

            //    }
            //    decimal total ;
            //    if (opcion == 3 || opcion == 4)
            //    {
            //        lret = fGuardaMovimiento();
            //        if (opcion == 4)
            //        {
            //            //lret = fSetDatoMovimiento("cNETO", x.cSubtotal.ToString());
            //            //lret = fSetDatoMovimiento("cTotal", x.cSubtotal.ToString());
            //            //string lcadena = "update mgw10010 set cneto = " + x.cSubtotal + ", ctotal = " + x.cSubtotal + " where ciddocum01 = " + lIdDocumento;
            //            total = x.cSubtotal + x.cImpuesto;
            //            string lcadena55 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total +"  where ciddocum01 = " + lIdDocumento;
            //            OleDbCommand lsql22 = new OleDbCommand(lcadena55, miconexion._conexion);
            //            lsql22.ExecuteNonQuery();
            //        }
            //        else
            //        {
            //             total = x.cSubtotal + x.cImpuesto;
            //             string lcadena44 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + " where ciddocum01 = " + lIdDocumento;
            //            OleDbCommand lsql3 = new OleDbCommand(lcadena44, miconexion._conexion);
            //            lsql3.ExecuteNonQuery();
            //        }

            //        //lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());
            //    }
            //    else
            //    {
                    
            //        lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
            //        lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
            //        lret = fSetDatoMovimiento("cporcentajeimpuesto1", x.cPorcent01.ToString());

            //        try
            //        {
            //            lret = fSetDatoMovimiento("ctextoextra1", x.ctextoextra1);
            //            lret = fSetDatoMovimiento("cReferencia", x.cReferencia);
            //            lret = fSetDatoMovimiento("ctextoextra2", x.ctextoextra2);
            //            lret = fSetDatoMovimiento("ctextoextra3", x.ctextoextra3);
            //        }
            //        catch (Exception ee)
            //        { }

            //        lret = fGuardaMovimiento();
            //    }

                
            //    lNumeroMov += 100;

            //}
            long lrespuesta = 0;
            string lrespuestas = "";
            if (lresp == "")
            {
                //double x = double.Parse(afolionuevo.ToString () );
                    mActualizaDocumento(lIdDocumento, opcion, afolionuevo );

                    lsql.CommandText = "select alltrim(cnombrec01) + ', '" + 
                                        " + alltrim(cnumeroe01) + ', '" +  
                                        " + alltrim(ccolonia) + ', '" +  
                                        " + alltrim(cciudad) + ', '" + 
                                        " + alltrim(cestado) + ', '" + 
                                        " + alltrim(cpais) " +   
                                       " from mgw10011 where ctipocat01 = 4";

                    miconexion.mAbrirConexionDestino();
                    lsql.Connection = miconexion._conexion;
                    lreader = lsql.ExecuteReader();
                    string ldireccion= "" ;
                    if (lreader.HasRows)
                    {
                        lreader.Read();
                        ldireccion = lreader[0].ToString().Trim();
                        lreader.Close();
                    }


                    
                    string lcadena2 = "update mgw10008 set cobserva01 = '" + _RegDoctoOrigen.cTextoExtra1 + "', clugarexpe = '"+ ldireccion.Trim () + "'  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql4 = new OleDbCommand(lcadena2, miconexion._conexion);
                    lsql4.ExecuteNonQuery();
                //lrespuesta = fAfectaDocto_Param(lCodigoConcepto, cserie , x, true);
                    
                if (opcion == 1 && _RegDoctoOrigen.cTipoCambio != 1 )
                {

                    //miconexion.mAbrirConexionDestino(1);
                    //                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;
                    
                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    string lcadena1 = "update mgw10008 set ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + "  where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                   // lsql3.ExecuteNonQuery();

                    
                    
                }
                if (_RegDoctoOrigen._RegDireccion.cEmail != "")
                {
                
                    string lcadena21 = "update mgw10002 set cemail1 = '" + _RegDoctoOrigen._RegDireccion.cEmail + "', cemail2 = '" + _RegDoctoOrigen._RegDireccion.cEmail2 + "', cbancfd = 1, ctipoentre=6 where ccodigoc01 = '" + _RegDoctoOrigen.cCodigoCliente + "'";
                    OleDbCommand lsql212 = new OleDbCommand(lcadena21, miconexion._conexion);
                    lsql212.ExecuteNonQuery();
                }
                

                if (opcion == 3 || opcion == 4)
                {
                    
                    //miconexion.mAbrirConexionDestino(1);
//                    string lcadena1 = "update mgw10008 set cpendiente = ctotal where ciddocum01 = " + lIdDocumento;

                    double ltotal = _RegDoctoOrigen.cImpuestos + _RegDoctoOrigen.cNeto;
                    decimal limpuestos = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString ());
                    limpuestos = decimal.Round(limpuestos,4);
                    string lcadena1 = "update mgw10008 set cneto = "  +  _RegDoctoOrigen.cNeto.ToString () + ", cimpuesto1 = " + limpuestos  + ",ctotal = " + ltotal.ToString () + ",cpendiente = " + ltotal.ToString () + ",ctipocam01 = " + _RegDoctoOrigen.cTipoCambio + " where ciddocum01 = " + lIdDocumento;

                    OleDbCommand lsql3 = new OleDbCommand(lcadena1, miconexion._conexion);
                    lsql3.ExecuteNonQuery();
                    //miconexion.mCerrarConexionDestino();
                }
                miconexion.mCerrarConexionDestino();

                //mGrabarInterfaz(afolioant, opcion);
            }
            else
            {
                lrespuestas = "ocurrio error";
            }
            //miconexion.mCerrarConexionDestino ();

            // antes de cerrar grabar en la tabla de interfaz




            miconexion.mCerrarConexionOrigen(1);
            //miconexion.mCerrarConexionDestino(1);

            try
            {
                _controlfp(0x9001F, 0xFFFFF);
            }
            catch (Exception eee)
            {
                lrespuestas = eee.Message;
            }
           // barra.Asignar(100);
            return lrespuestas;
        }

        private string mGrabarMovimientos(long lIdDocumento, int opcion)
        {
            long lret = 0;
            long lNumeroMov = 100;
            if (_RegDoctoOrigen._RegMovtos.Count == 0 && (opcion == 3 || opcion == 4))
            {

                RegMovto lRegmovto = new RegMovto();
                lRegmovto.cCodigoProducto = "(Ninguno)";
                lRegmovto.cNombreProducto = "(Ninguno)";
                lRegmovto.cIdDocto = long.Parse(_RegDoctoOrigen.cIdDocto.ToString());
                lRegmovto.cSubtotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
                //lRegmovto.cTotal = decimal.Parse(lreader["cunidades"].ToString());
                lRegmovto.cImpuesto = decimal.Parse(_RegDoctoOrigen.cImpuestos.ToString());
                lRegmovto.cTotal = decimal.Parse(_RegDoctoOrigen.cNeto.ToString());
                lRegmovto.cCodigoAlmacen = "(Ninguno)";
                lRegmovto.cNombreAlmacen = "(Ninguno)";
                lRegmovto.cUnidad = "";
                lRegmovto.cUnidades = 1;
                lRegmovto.cReferencia = "";
                lRegmovto.ctextoextra1 = "";
                lRegmovto.ctextoextra2 = "";
                lRegmovto.ctextoextra3 = "";
                _RegDoctoOrigen._RegMovtos.Add(lRegmovto);
            }
            foreach (RegMovto x in _RegDoctoOrigen._RegMovtos)
            {
                lret = fInsertarMovimiento();
                lret = fSetDatoMovimiento("cIdDocumento", lIdDocumento.ToString());
                lret = fSetDatoMovimiento("cNumeroMovimiento", lNumeroMov.ToString());
                lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                if (lret != 0)
                {
                    fInsertaProducto();
                    lret = fSetDatoProducto("CCODIGOPRODUCTO", x.cCodigoProducto);
                    lret = fSetDatoProducto("CNOMBREPRODUCTO", x.cNombreProducto);
                    lret = fSetDatoProducto("CTIPOPRODUCTO", "3");
                    lret = fSetDatoProducto("CMETODOCOSTEO", "1");
                    lret = fSetDatoProducto("CCONTROLEXISTENCIA", "1");
                    x.ctextoextra1 = "";
                    lret = fSetDatoDocumento("COBSERVACIONES", x.ctextoextra1);
                    lret = fSetDatoProducto("CIMPUESTO1", x.cPorcent01.ToString());
                    OleDbCommand cmdunidad = new OleDbCommand();
                    cmdunidad.CommandText = "select * from mgw10026 where cnombreu01 = '" + x.cUnidad.ToUpper() + "'";
                    miconexion.mAbrirConexionDestino();
                    cmdunidad.Connection = miconexion._conexion;
                    OleDbDataReader ldr = cmdunidad.ExecuteReader();
                    int lidunidad;
                    if (ldr.HasRows == false)
                    {
                        ldr.Read();
                        ldr.Close();
                        lret = fSetDatoProducto("CCODIGOUNIDADBASE", x.cUnidad.ToUpper());
                        if (lret != 0)
                        {
                            // dar de alta la unicad de medida y peso
                            cmdunidad.CommandText = "select max(cidunidad) + 1 from mgw10026";
                            ldr = cmdunidad.ExecuteReader();
                            ldr.Read();

                            lidunidad = int.Parse(ldr[0].ToString());
                            ldr.Close();
                            cmdunidad.CommandText = "insert into mgw10026 values (" + lidunidad + ",'" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','" + x.cUnidad.ToUpper() + "','')";
                            cmdunidad.ExecuteNonQuery();
                            lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
                        }
                    }
                    else
                    {
                        ldr.Read();

                        lidunidad = int.Parse(ldr[0].ToString());
                        ldr.Close();
                        lret = fSetDatoProducto("CIDUNIDADBASE", lidunidad.ToString());
                    }
                    lret = fGuardaProducto();
                    lret = fSetDatoMovimiento("cCodigoProducto", x.cCodigoProducto);
                }
                lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);
                if (lret != 0)
                {
                    fInsertaAlmacen();
                    lret = fSetDatoAlmacen("CCODIGOALMACEN", x.cCodigoAlmacen);
                    lret = fSetDatoAlmacen("CNOMBREALMACEN", x.cNombreAlmacen);
                    lret = fGuardaAlmacen();
                    lret = fSetDatoMovimiento("cCodigoAlmacen", x.cCodigoAlmacen);

                }
                decimal total;
                if (opcion == 3 || opcion == 4)
                {
                    lret = fGuardaMovimiento();
                    if (opcion == 4)
                    {
                        //lret = fSetDatoMovimiento("cNETO", x.cSubtotal.ToString());
                        //lret = fSetDatoMovimiento("cTotal", x.cSubtotal.ToString());
                        //string lcadena = "update mgw10010 set cneto = " + x.cSubtotal + ", ctotal = " + x.cSubtotal + " where ciddocum01 = " + lIdDocumento;
                        total = x.cSubtotal + x.cImpuesto;
                        string lcadena55 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + "  where ciddocum01 = " + lIdDocumento;
                        OleDbCommand lsql22 = new OleDbCommand(lcadena55, miconexion._conexion);
                        lsql22.ExecuteNonQuery();
                    }
                    else
                    {
                        total = x.cSubtotal + x.cImpuesto;
                        string lcadena44 = "update mgw10010 set cneto = " + x.cSubtotal + ", cimpuesto1 = " + x.cImpuesto + ", ctotal = " + total + " where ciddocum01 = " + lIdDocumento;
                        OleDbCommand lsql3 = new OleDbCommand(lcadena44, miconexion._conexion);
                        lsql3.ExecuteNonQuery();
                    }

                    //lret = fSetDatoMovimiento("cImpuesto1", x.cImpuesto.ToString());
                }
                else
                {

                    lret = fSetDatoMovimiento("cUnidades", x.cUnidades.ToString());
                    lret = fSetDatoMovimiento("cPrecio", x.cPrecio.ToString());
                    lret = fSetDatoMovimiento("cporcentajeimpuesto1", x.cPorcent01.ToString());

                    try
                    {
                        lret = fSetDatoMovimiento("ctextoextra1", x.ctextoextra1);
                        lret = fSetDatoMovimiento("cReferencia", x.cReferencia);
                        lret = fSetDatoMovimiento("ctextoextra2", x.ctextoextra2);
                        lret = fSetDatoMovimiento("ctextoextra3", x.ctextoextra3);
                    }
                    catch (Exception ee)
                    { }

                    lret = fGuardaMovimiento();
                }


                lNumeroMov += 100;

            }
            return "";
        }

        protected virtual void  mLeerDireccion()
        {


            //_RegDoctoOrigen._RegDireccion; 
        }

        private Boolean mGrabarInterfaz(string aFolio, int aTipo)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string amensaje = "";
            lconexion = miconexion.mAbrirRutaGlobal(out amensaje );
            bool lrespuesta = false;

            //lconexion = miconexion.mAbrirConexionDestino();
            try
            {
                OleDbCommand lsql = new OleDbCommand("insert into interfaz values ('" + aFolio + "'," + aTipo + ")", lconexion);
                lsql.ExecuteNonQuery();
            }
            catch (Exception eeee)
            { }

            miconexion.mCerrarConexionGlobal();






            return lrespuesta;
        }

        private bool mDarAltaCliente()
        {

            OleDbCommand  lsql = new OleDbCommand();
            OleDbDataReader   lreader;
            long lidclien;
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            lsql.CommandText = "select max(cidclien01) + 1 as cidclien01 from mgw10002";
            lsql.Connection = miconexion._conexion  ;
            lreader = lsql.ExecuteReader();
            _RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                lidclien = long.Parse (lreader["cidclien01"].ToString ());
            }
            else
                lidclien = 1;
            lreader.Close();


            //OleDbConnection lconexion = new OleDbConnection();
           // lconexion = miconexion.mAbrirConexionDestino ();
            bool lrespuesta = false;
            string lfecha = _RegDoctoOrigen.cFecha.ToString();
            DateTime ldate = DateTime.Parse (lfecha);
            lfecha = ldate.ToString("dd/MM/yyyy");

            //lconexion = miconexion.mAbrirConexionDestino();
            string lcadena = "insert into mgw10002 (cidclien01, ccodigoc01,crazonso01,cfechaalta,crfc,cidmoneda, clistapr01, ctipocli01,cestatus) values (" +
                lidclien +
                ",'" + _RegDoctoOrigen.cCodigoCliente + "','" + _RegDoctoOrigen.cRazonSocial  + "'," +
                "ctod('" + lfecha + "'),'" +
                _RegDoctoOrigen.cRFC + "'" + 
                ",1,1,1,1)";
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion );
            try
            {
                lsql.CommandText = "SET NULL OFF";
                lsql.ExecuteNonQuery();

                lsql1.ExecuteNonQuery();
                lrespuesta = true;
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }

            //.mCerrarConexionDestino ();



            return lrespuesta;

        }



        protected virtual  bool mActualizaDocumento(long liddocum, int aopcion, double afolionuevo)
        {
            miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal=0;
            bool lrespuesta = false;
            
            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum ;
            
            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {

                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");


                long ctipocfd = 0;
                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                //"ConceptoDocumento"
                    string lCodigoConcepto = _RegDoctoOrigen.cCodigoConcepto;
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01, cverfacele from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'" ;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();
                    ctipocfd = long.Parse (lreader["cverfacele"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = "+ liddocum ;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();


                double x = double.Parse(afolionuevo.ToString().Trim());

                lsql.CommandText = "select max(cidfoldig) + 1 as cidclien01 from mgw10045";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidfoldig = long.Parse(lreader["cidclien01"].ToString());
                }
                else
                    cidfoldig = 1;
                lreader.Close();

                lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                                 " values (" + liddocum  + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim()  + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";
                //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado, cfechaemi,cestrad) " +
                //                 "values (8,4,3001,11,'B',444,1,ctod('" + lfecha + "'),3)";
                OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                lsql1.CommandText = "SET NULL OFF";
                lsql1.ExecuteNonQuery();

                lsql2.ExecuteNonQuery();
                lrespuesta = true;
            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                miconexion.mCerrarConexionDestino();
            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }
        private bool mActualizaDocumento2(long liddocum, long adestino, double afolio)
        {
            //if (adestino > 0)
            //    miconexion.mAbrirConexionOrigen(1);
            //else
                miconexion.mAbrirConexionDestino();
            OleDbCommand lsql = new OleDbCommand();
            OleDbDataReader lreader;
            long cidfoldig;
            long cidconce;
            long ciddocum01 = 0;
            string cserie = "";
            double ctotal = 0;
            bool lrespuesta = false;

            int cescfd = 0;
            string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
            lsql.CommandText = "select cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
            lsql.Connection = miconexion._conexion;
            lreader = lsql.ExecuteReader();
            //_RegDoctoOrigen._RegMovtos.Clear();
            if (lreader.HasRows)
            {
                lreader.Read();
                cescfd = int.Parse(lreader["cescfd"].ToString());
            }

            lreader.Close();
            if (cescfd == 0)
                return true;


            //OleDbParameter lparametroIdDocumento = new OleDbParameter("@p1", _RegDoctoOrigen.cIdDocto);
            string lcadena = "update mgw10008 set cescfd = 1 where ciddocum01 = " + liddocum;

            OleDbCommand lsql1 = new OleDbCommand(lcadena, miconexion._conexion);
            try
            {

                lsql1.ExecuteNonQuery();
                string lfecha = _RegDoctoOrigen.cFecha.ToString();
                DateTime ldate = DateTime.Parse(lfecha);
                lfecha = ldate.ToString("MM/dd/yyyy");




                //string lCodigoConcepto = GetSettingValueFromAppConfigForDLL("ConceptoDocumento").ToString().Trim();
                lsql.CommandText = "select cidconce01, ciddocum01,cseriepo01,cescfd from mgw10006 where ccodigoc01 = '" + lCodigoConcepto + "'";
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidconce = long.Parse(lreader["cidconce01"].ToString());
                    ciddocum01 = long.Parse(lreader["ciddocum01"].ToString());
                    cserie = lreader["cseriepo01"].ToString();

                }
                else
                    cidconce = 1;
                lreader.Close();
                cserie = GetSettingValueFromAppConfigForDLL("SerieFactura").ToString().Trim();

                lsql.CommandText = "select ctotal from mgw10008 where ciddocum01 = " + liddocum;
                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                if (lreader.HasRows)
                {
                    lreader.Read();
                    ctotal = double.Parse(lreader["ctotal"].ToString());

                }
                else
                    cidconce = 1;
                lreader.Close();


                //double x = double.Parse(GetSettingValueFromAppConfigForDLL("FolioFactura").ToString().Trim());
                double x = afolio;

                lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                lsql.CommandText = "select top 1 cidfoldig, cfolio, cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce+ " order by cidfoldig asc";

                lsql.Connection = miconexion._conexion;
                lreader = lsql.ExecuteReader();
                _RegDoctoOrigen._RegMovtos.Clear();
                //string cserie;
                if (lreader.HasRows)
                {
                    lreader.Read();
                    cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                    x = double.Parse(lreader["cfolio"].ToString());
                    cserie = lreader["cserie"].ToString();
                }
                else
                {
                    cidfoldig = 1;
                    x = 1;

                }

                //return false;
                lreader.Close();

                //lcadena = "insert into mgw10045 (cidfoldig,ciddoctode,cidcptodoc,ciddocto,cserie,cfolio,cestado,centregado, cfechaemi,cestrad,ctotal) " +
                //                 " values (" + liddocum + "," + ciddocum01 + "," + cidconce + "," + liddocum + ",'" + cserie.Trim() + "'," + x + ",1, 0, ctod('" + lfecha + "'),3," + ctotal + ")";

                try
                {
                    lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                    " where cidfoldig = " + cidfoldig + " and ciddocto = 0 ";

                    //lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                    //" where cidfoldig  in (select min(cidfoldig) from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc)";


                    OleDbCommand lsql2 = new OleDbCommand(lcadena, miconexion._conexion);
                    lsql1.CommandText = "SET NULL OFF";
                    lsql1.ExecuteNonQuery();

                    long lcuantos = lsql2.ExecuteNonQuery();
                    if (lcuantos == 0)
                    {

                        lsql.CommandText = "select min(cidfoldig) as cidfoldig, min(cfolio) as cfolio, min(cserie) as cserie from mgw10045 where ciddocto = 0 and cidcptodoc = " + cidconce + " group by cidcptodoc";
                        lsql.Connection = miconexion._conexion;
                        lreader = lsql.ExecuteReader();
                        _RegDoctoOrigen._RegMovtos.Clear();
                        //string cserie;
                        if (lreader.HasRows)
                        {
                            lreader.Read();
                            cidfoldig = long.Parse(lreader["cidfoldig"].ToString());
                            x = double.Parse(lreader["cfolio"].ToString());
                            cserie = lreader["cserie"].ToString();
                        }
                        else
                        {
                            cidfoldig = 1;
                            x = 1;

                        }


                        lcadena = "update mgw10045 set ciddocto=" + liddocum + ",cestado=1,cfechaemi=ctod('" + lfecha + "')" +
                        " where cidfoldig = " + cidfoldig;
                        lsql2.ExecuteNonQuery();
                    }


                    lcadena = "update mgw10008 set cfolio=" + x + ",cseriedo01='" + cserie + "'" +
                    " where ciddocum01 = " + liddocum;
                    lsql2.CommandText = lcadena;
                    lsql2.ExecuteNonQuery();
                    lrespuesta = true;
                }
                catch (Exception eeeee)
                {
                    OleDbCommand lsql3 = new OleDbCommand(lcadena, miconexion._conexion);
                    lcadena = "delete from  mgw10008 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();

                    lcadena = "delete from  mgw10010 " +
                    " where ciddocum01 = " + liddocum;
                    lsql3.CommandText = lcadena;
                    lsql3.ExecuteNonQuery();
                    lrespuesta = false;
                }









            }
            catch (Exception eee)
            {
                lrespuesta = true;
            }
            finally
            {
                if (adestino > 0)
                    miconexion.mCerrarConexionOrigen(1);
                else
                    miconexion.mCerrarConexionDestino();

            }
            //.mCerrarConexionDestino ();



            return lrespuesta;

        }

        public virtual string mBuscarDoctosArchivo(string aNombreArchivo)
        {
            
            return "";
        }




        #region ISujeto Members

        public void Registrar(IObservador obs)
        {
            //throw new NotImplementedException();
            lista.Clear();
            lista.Add(obs);
            Notificar(0);
        }

        public void Notificar()
        {
            //throw new NotImplementedException();
            foreach (IObservador x in lista)
                x.Actualizar(0);
        }

        public void Notificar(double lavance)
        {
            //throw new NotImplementedException();
            foreach (IObservador x in lista)
                x.Actualizar(lavance);
        }

        #endregion


        public void Registrar()
        {
            throw new NotImplementedException();
        }
    }
    
}


/*x.AppendLine(",cseriedo01");
            x.AppendLine(",cfolio");
            x.AppendLine(",cfecha");

            x.AppendLine(",cidclien01");
            x.AppendLine(",crazonso01");
            x.AppendLine(",crfc");
            x.AppendLine(",cidagente");
            x.AppendLine(",cfechave01");
            x.AppendLine(",cfechapr01");
            x.AppendLine(",cfechaen01");
            x.AppendLine(",cfechaul01");
            x.AppendLine(",cidmoneda");
            x.AppendLine(",ctipocam01");
            x.AppendLine(",creferen01");
            x.AppendLine(",cobserva01");
            x.AppendLine(",cnatural01");
            x.AppendLine(",ciddocum03");
            x.AppendLine(",cplantilla");
            x.AppendLine(",cusaclie01");
            x.AppendLine(",cusaprov01");
            x.AppendLine(",cafectado");
            x.AppendLine(",cimpreso");
            x.AppendLine(",ccancelado");
            x.AppendLine(",cdevuelto");
            x.AppendLine(",cidprepo01");
            x.AppendLine(",cidprepo02");
            x.AppendLine(",cestadoc01");
            x.AppendLine(",cneto");
            x.AppendLine(",cimpuesto1");
            x.AppendLine(",cimpuesto2");
            x.AppendLine(",cimpuesto3");
            x.AppendLine(",cretenci01");
            x.AppendLine(",cretenci02");
            x.AppendLine(",cdescuen01");
            x.AppendLine(",cdescuen02");
            x.AppendLine(",cdescuen03");
            x.AppendLine(",cgasto1");
            x.AppendLine(",cgasto2");
            x.AppendLine(",cgasto3");
            x.AppendLine(",ctotal");
            x.AppendLine(",cpendiente");
            x.AppendLine(",ctotalun01");
            x.AppendLine(",cdescuen04");
            x.AppendLine(",cporcent01");
            x.AppendLine(",cporcent02");
            x.AppendLine(",cporcent03");
            x.AppendLine(",cporcent04");
            x.AppendLine(",cporcent05");
            x.AppendLine(",cporcent06");
            x.AppendLine(",ctextoex01");
            x.AppendLine(",ctextoex02");
            x.AppendLine(",ctextoex03");
            x.AppendLine(",cfechaex01");
            x.AppendLine(",cimporte01");
            x.AppendLine(",cimporte02");
            x.AppendLine(",cimporte02");
            x.AppendLine(",cimporte03");
            x.AppendLine(",cimporte03");
            x.AppendLine(",cimporte04");
            x.AppendLine(",cdestina01");
            x.AppendLine(",cnumerog01");
            x.AppendLine(",cmensaje01");
            x.AppendLine(",ccuentam01");
            x.AppendLine(",cnumeroc01");
            x.AppendLine(",cpeso");
            x.AppendLine(",cbanobse01");
            x.AppendLine(",cbandato01");
            x.AppendLine(",cbancond01");
            x.AppendLine(",cbangastos");
            x.AppendLine(",cunidade01");
            x.AppendLine(",ctimestamp");
            x.AppendLine(",cimpcheq01");
            x.AppendLine(",csistorig");
            x.AppendLine(",cidmonedca");
            x.AppendLine(",ctipocamca");
            x.AppendLine(",cescfd");
            x.AppendLine(",ctienecfd");
            x.AppendLine(",clugarexpe");
            x.AppendLine(",cmetodopag");
            x.AppendLine(",cnumparcia");
            x.AppendLine(",ccantparci");
            x.AppendLine(",ccondipago");
            x.AppendLine(",cnumctapag) ");*/