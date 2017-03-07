using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Configuration;
using System.IO;


namespace LibreriaDoctos
{
    public class ClassConexion
    {
        //public string llaveregistry = "SOFTWARE\\Wow6432Node\\Computación en Acción, SA CV\\AdminPAQ";
        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public string sError = "";
        public string aRutaExe = "";
        [DllImport("MGW_SDK.DLL")] static extern int fInicializaSDK();
        [DllImport("MGW_SDK.DLL")] private static extern void fTerminaSDK();
        [DllImport("MGW_SDK.DLL")] private static extern void fCierraEmpresa();
        [DllImport("MGW_SDK.DLL")] static extern int fAbreEmpresa(String aRuta);
        [DllImport("KERNEL32.DLL")] static extern int SetCurrentDirectory(string pPtrDirActual);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscaProducto(String aCodigoProducto);
        [DllImport("MGW_SDK.DLL")] static extern int fInsertaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fGuardaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fSetDatoDireccion(string aCampo, string aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscaDireccionDocumento(long aIdDocumento, byte aValor);
        [DllImport("MGW_SDK.DLL")] static extern int fEditaDireccion();
        [DllImport("MGW_SDK.DLL")] static extern int fLeeDatoProducto(string aCampo, string aValor, long aLongitud);
        [DllImport("MGW_SDK.DLL")] static extern int fAfectaDocto_Param(string aConcepto, string aSerie, double aFolio, Boolean aAfecta);
        [DllImport("MGW_SDK.DLL")] static extern int fBuscarIdDocumento(int aIdDocumento);
        [DllImport("MGW_SDK.DLL")] static extern int fEditarDocumento();
        [DllImport("MGW_SDK.DLL")] static extern int fBuscarDocumento(string aConcepto, string aFolio, string aSerie);
        [DllImport("MGW_SDK.DLL")] static extern int fBorraDocumento();
        [DllImport("MGW_SDK.DLL")] static extern long fError(long aNumErrror, string aError, long aLen);
        [DllImport("MGW_SDK.DLL")] static extern long fLeeDatoDocumento(string aCampo, ref string aValor, long aLongitud);
        

        public string rutaorigen;
        public string rutadestino;
        public const string _NombreAplicacionCompleto = "SaldosIniciales.exe";
        public const string _NombreAplicacion = "SaldosIniciales";

        public OleDbConnection _conexion ;
        public void borrar()
        { 
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
        }
        public OleDbConnection  mAbrirConexionOrigen()
        {
            _conexion = null;
            rutaorigen = GetSettingValueFromAppConfigForDLL( "RutaEmpresaADM");
            if (rutaorigen != "c:\\" && rutaorigen != "LibreriaDoctos.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion.Open();
            }
            return _conexion;
         
        }
        public OleDbConnection mAbrirConexionOrigen(int a)
        {

            rutaorigen = GetSettingValueFromAppConfigForDLL( "RutaEmpresaSamira");
            //rutaorigen = "c:\\compacw\\empresas\\adtala";
            //rutaorigen = Properties.Settings.Default.RutaEmpresaSamira;
             _conexion =new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen ;
            _conexion.Open();
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            if (obc == null)
            {
                sError = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            SetCurrentDirectory(obc.ToString());
            
            fInicializaSDK();
            fAbreEmpresa(rutaorigen); 
            return _conexion;
        }
        
        public  void mCerrarConexionOrigen()
        {
            _conexion.Close();
        }

        public void mCerrarConexionOrigen(int a)
        {
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            SetCurrentDirectory(obc.ToString());
            
            _conexion.Close();
            fCierraEmpresa();
            fTerminaSDK();
        }

        
        public void mCerrarConexionDestino()
        {
            _conexion.Close();
        }

        public void mCerrarConexionGlobal()
        {
            _conexion.Close();
        }

        public OleDbConnection mAbrirRutaGlobal(out string amensaje)
        {
            amensaje = "";
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = null;
            try
            {
                 obc = hklp.GetValue("DIRECTORIODATOS");
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
            }
                //amensaje = obc.ToString ();
            if (obc == null)
            {
                amensaje = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + obc.ToString();
            //_conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + "\\toshiba-pc" + asc(92) +  "empresas";
            try
            {
                _conexion.Open();
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
            }
            return _conexion ;

        }
        public OleDbConnection mAbrirConexionAccess(out string msg)
        {
            msg = "";
                string rutaaccess = GetSettingValueFromAppConfigForDLL("RutaAccess");
                //msg = rutaaccess;
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + rutaaccess    + ";User Id=admin;Password=";
            //_conexion.Open();

            
            return _conexion;
 
        }
        public OleDbConnection  mAbrirConexionDestino(int a)
        {
//            rutadestino = "c:\\compacw\\empresas\\adtala2";
            rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaADM");
            //string lruta3 = obc.ToString();
            string lruta4 = @rutadestino;
             _conexion =new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutadestino ;
            _conexion.Open();
            
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry );
            Object obc = hklp.GetValue("DIRECTORIOBASE");
            string lruta1 = obc.ToString();
            string lruta2 = @lruta1;
            SetCurrentDirectory(obc.ToString());
            
            long lret;
            try
            {
                //fTerminaSDK();
                lret = fInicializaSDK();
            }
            catch (Exception eeeee)
            { fTerminaSDK();
            lret = fInicializaSDK();
            }
            lret = fAbreEmpresa(rutadestino);
            //fCierraEmpresa();
            //fTerminaSDK();
            return _conexion;
         
        }
        public OleDbConnection mAbrirConexionDestino()
        {
            _conexion = null;
            rutadestino = GetSettingValueFromAppConfigForDLL("RutaEmpresaDestino");
            if (rutadestino != "c:\\" && rutadestino != "LibreriaDoctos.RegEmpresa")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutadestino;
                _conexion.Open();
            }
            return _conexion;

        }
        private string GetSettingValueFromAppConfigForDLL(string aNombreSetting)
        {
            string lrutadminpaq = Directory.GetCurrentDirectory();
            if (Directory.GetCurrentDirectory() != aRutaExe)
                Directory.SetCurrentDirectory(aRutaExe);

            string value ="";
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
                Directory.SetCurrentDirectory(lrutadminpaq);
            
            return value;
        }
        


    }
}
