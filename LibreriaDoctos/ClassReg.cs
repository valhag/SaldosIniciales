using System;
using System.Collections.Generic;
using System.Text;

namespace LibreriaDoctos
{

    public class RegEmpresas
    {
        private string _Empresa;
        public string cEmpresa
        {
            get { return _Empresa; }
            set { _Empresa = value; }
        }
    }

    public class RegPuntodeVenta
    {
        private string _Empresa;
        private string _Nombre;

        public string cEmpresa
        {
            get { return _Empresa; }
            set { _Empresa = value; }
        }
        public string cNombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
    }


    public class RegDocto
    {
        public List<RegMovto> _RegMovtos = new List<RegMovto>();
        public RegDireccion _RegDireccion = new RegDireccion ();
        private long _cIdDocto;
        private string _cCodigoCliente;
        private string _cCodigoConcepto;
        private long _cIdConcepto;
        private string _cRFC;
        private string _cRazonSocial;
        private string _cMoneda;
        private string _cCond;
        private string _cTextoExtra1="";
        private string _cTextoExtra2 = "";
        private string _cTextoExtra3 = "";
        private string _cNombreArchivo = "";


        private int _cContado = 0;
        public int cContado
        {
            get { return _cContado; }
            set { _cContado = value; }
        }


        public string cNombreArchivo
        {
            get { return _cNombreArchivo; }
            set { _cNombreArchivo = value; }
        }
        

        public string cTextoExtra1
        {
            get { return _cTextoExtra1; }
            set { _cTextoExtra1 = value; }
        }

        public string cTextoExtra2
        {
            get { return _cTextoExtra2; }
            set { _cTextoExtra2 = value; }
        }

        public string cTextoExtra3
        {
            get { return _cTextoExtra3; }
            set { _cTextoExtra3 = value; }
        }
        private string _sMensaje;

        


        public string cReferencia { get; set; }



        public string sMensaje
        {
            get { return _sMensaje; }
            set { _sMensaje = value; }
        }

        public string cCond
        {
            get { return _cCond; }
            set { _cCond = value; }
        }
        private string _cAgente;

        public string cAgente
        {
            get { return _cAgente; }
            set { _cAgente = value; }
        }



        private double _cNeto;

        public double cNeto
        {
            get { return _cNeto; }
            set { _cNeto = value; }
        }
        private double _cImpuestos;

        public double cImpuestos
        {
            get { return _cImpuestos; }
            set { _cImpuestos = value; }
        }

        public string cMoneda
        {
            get { return _cMoneda; }
            set { _cMoneda = value; }
        }
        private decimal _cTipoCambio;

        public decimal cTipoCambio
        {
            get { return _cTipoCambio; }
            set { _cTipoCambio = value; }
        }

        public string cRazonSocial
        {
            get { return _cRazonSocial; }
            set { _cRazonSocial = value; }
        }

        public string cRFC
        {
            get { return _cRFC; }
            set { _cRFC = value; }
        }


        public long cIdConcepto
        {
            get { return _cIdConcepto; }
            set { _cIdConcepto = value; }
        }


        public string cCodigoConcepto
        {
            get { return _cCodigoConcepto; }
            set { _cCodigoConcepto = value; }
        }
        private DateTime _cFecha;

        public DateTime cFecha
        {
            get { return _cFecha; }
            set { _cFecha = value; }
        }
        private long _cFolio;

        public long cFolio
        {
            get { return _cFolio; }
            set { _cFolio = value; }
        }


        public long cIdDocto
        {
            get { return _cIdDocto; }
            set { _cIdDocto = value; }
        }
        public string cCodigoCliente
        {
            get { return _cCodigoCliente; }
            set { _cCodigoCliente = value; }
        }

        private string _cSerie;

        public string cSerie
        {
            get { return _cSerie; }
            set { _cSerie = value; }
        }




    }
    public class RegMovto
    {


        
        public string cReferencia { get; set; }
        public string ctextoextra1 { get; set; }
        public string ctextoextra2 { get; set; }
        public string ctextoextra3 { get; set; }
        private string _cUnidad;

        public string cUnidad
        {
            get { return _cUnidad; }
            set { _cUnidad = value; }
        }

        
        private decimal _cMargenUtilidad;

        public decimal cMargenUtilidad
        {
            get { return _cMargenUtilidad; }
            set { _cMargenUtilidad = value; }
        }

        private string _cCodigoAlmacen;

        public string cCodigoAlmacen
        {
            get { return _cCodigoAlmacen; }
            set { _cCodigoAlmacen = value; }
        }

        private string _cNombreAlmacen;

        public string cNombreAlmacen
        {
            get { return _cNombreAlmacen; }
            set { _cNombreAlmacen = value; }
        }

        private long _cIdMovto;

        public long cIdMovto
        {
            get { return _cIdMovto; }
            set { _cIdMovto = value; }
        }
        private long _cIdDocto;

        public long cIdDocto
        {
            get { return _cIdDocto; }
            set { _cIdDocto = value; }
        }
        private string _cNombreProducto;

        public string cNombreProducto
        {
            get { return _cNombreProducto; }
            set { _cNombreProducto = value; }
        }

        private string _cCodigoProducto;

        public string cCodigoProducto
        {
            get { return _cCodigoProducto; }
            set { _cCodigoProducto = value; }
        }
        private decimal _cUnidades;

        public decimal cUnidades
        {
            get { return _cUnidades; }
            set { _cUnidades = value; }
        }
        
        private decimal _cPrecio;

        public decimal cPrecio
        {
            get { return _cPrecio; }
            set { _cPrecio = value; }
        }
        
        private decimal _cSubtotal;

        public decimal cSubtotal
        {
            get { return _cSubtotal; }
            set { _cSubtotal = value; }
        }
        private decimal _cTotal;

        public decimal cTotal
        {
            get { return _cTotal; }
            set { _cTotal = value; }
        }
        private decimal _cImpuesto;

        public decimal cImpuesto
        {
            get { return _cImpuesto; }
            set { _cImpuesto = value; }
        }

        private decimal _cPorcent01;
        public decimal cPorcent01
        {
            get { return _cPorcent01; }
            set { _cPorcent01 = value; }
        }

        private decimal _cneto;
        public decimal cneto
        {
            get { return _cneto; }
            set { _cneto = value; }
        }



    }

    public class RegProveedor
    {
        private long _Id;

        public long Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _RazonSocial;

        public string RazonSocial
        {
            get { return _RazonSocial; }
            set { _RazonSocial = value; }
        }
        private string _RFC;

        public string RFC
        {
            get { return _RFC; }
            set { _RFC = value; }
        }
        private int _DiasCredito;

        public int DiasCredito
        {
            get { return _DiasCredito; }
            set { _DiasCredito = value; }
        }
        


    }

    public class RegConcepto
    {
        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _sTipocfd;

        public string Tipocfd
        {
            get { return _sTipocfd; }
            set { _sTipocfd = value; }
        }
    }

    public class RegEmpresa
    {
        public string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _Ruta;

        public string Ruta
        {
            get { return _Ruta; }
            set { _Ruta = value; }
        }
    }

    public class RegDireccion
    {
        private string _cEmail = "";
        private string _cEmail2 = "";

        public string cEmail
        {
            get { return _cEmail; }
            set { _cEmail = value; }
        }
        public string cEmail2
        {
            get { return _cEmail2; }
            set { _cEmail2 = value; }
        }
        private string _cNombreCalle;

        public string cNombreCalle
        {
            get { return _cNombreCalle; }
            set { _cNombreCalle = value; }
        }
        private string _cNumeroExterior;

        public string cNumeroExterior
        {
            get { return _cNumeroExterior; }
            set { _cNumeroExterior = value; }
        }
        private string _cNumeroInterior;

        public string cNumeroInterior
        {
            get { return _cNumeroInterior; }
            set { _cNumeroInterior = value; }
        }
        private string _cColonia;

        public string cColonia
        {
            get { return _cColonia; }
            set { _cColonia = value; }
        }
        private string _cCodigoPostal;

        public string cCodigoPostal
        {
            get { return _cCodigoPostal; }
            set { _cCodigoPostal = value; }
        }
        private string _cEstado;

        public string cEstado
        {
            get { return _cEstado; }
            set { _cEstado = value; }
        }
        private string _cPais;

        public string cPais
        {
            get { return _cPais; }
            set { _cPais = value; }
        }
        private string _cCiudad;

        public string cCiudad
        {
            get { return _cCiudad; }
            set { _cCiudad = value; }
        }
    }
    public class RegOrigen
    {
        private string _Folio;

        public string Folio
        {
            get { return _Folio; }
            set { _Folio = value; }
        }
        private string _Fecha;

        public string Fecha
        {
            get { return _Fecha; }
            set { _Fecha = value; }
        }
        private string _CodigoProducto;

        public string CodigoProducto
        {
            get { return _CodigoProducto; }
            set { _CodigoProducto = value; }
        }

        private string _NombreProducto;

        public string NombreProducto
        {
            get { return _NombreProducto; }
            set { _NombreProducto = value; }
        }
        private decimal _Cantidad;

        public decimal Cantidad
        {
            get { return _Cantidad; }
            set { _Cantidad = value; }
        }
        private decimal _Precio;

        public decimal Precio
        {
            get { return _Precio; }
            set { _Precio = value; }
        }

    }


}
