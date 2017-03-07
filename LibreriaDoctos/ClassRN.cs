using System;
using System.Collections.Generic;
using System.Text;
using Interfaces ;
//using BarradeProgreso;

namespace LibreriaDoctos
{
    public class ClassRN
    {
        public string productos = "";
        public string almacenes = "";
        public ClassBD lbd = new ClassBD();

        public int mValidaSQLConexion(string server, string bd, string user, string psw)
        {
            return lbd.mValidaSQLConexion(server, bd, user, psw);
        }

        public int mEjecutarComando(string comando, int aclientes, int lporcodigo)
        {
            return lbd.mEjecutarComando(comando,aclientes,lporcodigo);
        }

        public int mEjecutarComando2(string comando, int aclientes, int lporcodigo, string empresa)
        {
            return lbd.mEjecutarComando2(comando, aclientes, lporcodigo, empresa);
        }

        public int mEjecutarComando3(string comando, int aclientes, int lporcodigo, string empresaorigen, string empresadestino)
        {
            return lbd.mEjecutarComando3(comando, aclientes, lporcodigo, empresaorigen, empresadestino);
        }
        public void mInicializar(string aRutaOrigen, string aRutaDestino)
        {
            //Properties.Settings.Default.RutaEmpresaSamira = aRutaOrigen; 
        }
        //public ClassBD lbd = new ClassBD();
        //public ClassBD lbd;
        public Boolean mBuscar(long aFolio, string aConcepto, string aSerie, int aTipo)
        {
            return lbd.mBuscar(aFolio, aConcepto, aSerie, aTipo);
        }


        public string mBuscarDocto(string aFolio, int aTipo, bool aRevisar)
        {
           return lbd.mBuscarDocto(aFolio,  aTipo, aRevisar );
        }

        public   virtual  string mBuscarDoctoFlex(string aFolio, int aTipo, bool aRevisar)
        {
            return lbd.mBuscarDoctoAccess(aRevisar);
        }

        public virtual string mBuscarDoctos(long aFolioinicial, long afoliofinal , int aTipo, bool aRevisar)
        {
            return lbd.mBuscarDoctos(aFolioinicial, afoliofinal , aTipo, aRevisar);
        }

        public Boolean mValidarConexionIntell(string aRuta)
        {
            return lbd.mValidarConexionIntell(aRuta);
        }

        public Boolean mValidarConexionIntell(string aServidor, string aBd, string ausu, string apwd)
        {
            return lbd.mValidarConexionIntell(aServidor, aBd, ausu, apwd);
        }

        public string mGrabarAdm(string afolioant, double afolionuevo, int opcion)
        {
            return lbd.mGrabarAdm(afolioant, afolionuevo , opcion);
        }

        public List<string> mGrabarAdms(int opcion)
        {
            lbd.primerdocto = new RegDocto ();
            return lbd.mGrabarAdms(opcion);
        }


        public string mGrabarDestinos( )
        {
            string lregresa = "";
            lregresa =lbd.mGrabarDestinos();
            productos = lbd.productos;
            almacenes = lbd.almacenes ;
            return lregresa;
        }

        public List<RegConcepto> mCargarConceptosFactura()
        {
            return lbd.mCargarConceptos(4,0);
        }

        public List<RegProveedor> mCargarClientes()
        {
            return lbd.mCargarClientes ();
        }

        public List<RegConcepto> mCargarConceptosPedido()
        {
            return lbd.mCargarConceptos(2, 0);
        }

        public List<RegConcepto> mCargarConceptosDevolucion()
        {
            return lbd.mCargarConceptos(5,0);
        }
        public List<RegConcepto> mCargarConceptosNotaCredito()
        {
            return lbd.mCargarConceptos(7, 0);
        }
        public List<RegConcepto> mCargarConceptosNotaCargo()
        {
            return lbd.mCargarConceptos(13, 0);
        }

        public List<RegConcepto> mCargarConceptosCompraOrigen()
        {
            return lbd.mCargarConceptos(19, 0);
        }
        public RegProveedor mBuscarCliente(string aCliente)
        {
            return lbd.mBuscarCliente(aCliente,0,0);
        }
        public RegProveedor mBuscarProveedor(string aProveedor)
        {
            return lbd.mBuscarCliente(aProveedor, 1, 1);
        }

        public  void mSeteaDirectorio(string aRuta)
        {
            lbd.mAsignaRuta( aRuta);
        }
        public List<RegEmpresa> mCargarEmpresas(out string mensaje)
        {
            return lbd.mCargarEmpresas(out mensaje);
        }

        public List<RegPuntodeVenta> mCargarPuntoVenta(string aEmpresa, out string mensaje)
        {
            return lbd.mCargarPuntoVenta( aEmpresa, out mensaje);
        }

        public List<RegEmpresas> mCargarEmpresasAccess(out string mensaje)
        {
            return lbd.mCargarEmpresasAccess(out mensaje);
        }

        public virtual string mBuscarDoctosArchivo(string aArchivo)
        {
            return "";
        }
    }
}
