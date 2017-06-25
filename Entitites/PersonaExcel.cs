using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel.Attributes;

namespace EntititesExcel
{
    public class PersonaExcel
    {
        [ExcelColumn("Id")]
        public int id { get; set; }

        [ExcelColumn("FECHA de Ingreso")]
        public DateTime? fechaDeIngreso { get; set; }

        [ExcelColumn("Actualizado")]
        public DateTime actualizado { get; set; }

        [ExcelColumn("Titulo")]
        public string titulo { get; set; }

        [ExcelColumn("Apellido")]
        public string apellido { get; set; }

        [ExcelColumn("Nombre de Pila")]
        public string nombreDePila { get; set; }

        [ExcelColumn("Nombre Completo")]
        public string nombreCompleto { get; set; }

        [ExcelColumn("Domicilio")]
        public string domicilio { get; set; }

        [ExcelColumn("Codigo Postal")]
        public int codigoPostal { get; set; }

        [ExcelColumn("Localidad")]
        public string localidad { get; set; }
        
        [ExcelColumn("Cliente")]
        public string cliente { get; set; }

        [ExcelColumn("Imprime carta")]
        public bool imprimeCarta { get; set; }

        [ExcelColumn("ESTADO ACTUAL")]
        public string estadoActual { get; set; }

        [ExcelColumn("CELULAR")]
        public string celular { get; set; }

        [ExcelColumn("TEL COMERCIAL")]
        public string telComercial { get; set; }

        [ExcelColumn("Tel Particular")]
        public string telParticular { get; set; }
        
        [ExcelColumn("E-MAIL")]
        public string email { get; set; }

        [ExcelColumn("Nacimiento")]
        public DateTime nacimiento { get; set; }

        [ExcelColumn("Dia de Cumpleaños")]
        public int diaCumpleanios { get; set; }

        [ExcelColumn("Nacionalidad")]
        public string nacionalidad { get; set; }

        [ExcelColumn("DNI Titular")]
        public string dniTitular { get; set; }

        [ExcelColumn("CUIT CUIL TITULAR")]
        public string cuitCuilTitular { get; set; }
        
        [ExcelColumn("ESTADO CIVIL")]
        public string estadoCivil { get; set; }

        [ExcelColumn("Apellido Conyuge")]
        public string apellidoConyuge { get; set; }

        [ExcelColumn("Nombre de pila conyuge")]
        public string nombrePilaConyuge { get; set; }

        [ExcelColumn("Nombre Completo Conyuge")]
        public string nombreCompletoConyuge { get; set; }

        [ExcelColumn("DNI Conyuge")]
        public string dniConyuge { get; set; }

        [ExcelColumn("CUIT CUIL CONYUGE")]
        public string cuitCuilConyuge { get; set; }

        [ExcelColumn("Nacimiento Conyuge")]
        public DateTime nacimientoConyuge { get; set; }

        [ExcelColumn("NACIONALIDAD CONYUGE")]
        public string nacionalidadConyuge { get; set; }

        [ExcelColumn("CELULAR DEL CONYUGE")]
        public string celularConyuge { get; set; }

        [ExcelColumn("MAIL DEL CONYUGE")]
        public string mailConyuge { get; set; }

        [ExcelColumn("CATEGORIA")]
        public string categoria { get; set; }

        [ExcelColumn("OBSERVACIONES")]
        public string observaciones { get; set; }
        
        public PersonaExcel()
        { }

        public PersonaExcel(int _id, DateTime? _fechaDeIngreso, string _titulo, string _apellido, string _nombreDePila, string _nombreCompleto, string _domicilio,
        int _codigoPostal, string _localidad, string _cliente, bool _imprimeCarta, string _estadoActual, string _celular, string _telComercial, string _telParticular,
        string _email, int _diaCumpleanios, string _nacionalidad, string _dniTitular, string _cuitCuilTitular, string _estadoCivil, string _apellidoConyuge,
        string _nombrePilaConyuge, string _nombreCompletoConyuge, string _dniConyuge, string _cuitCuilConyuge, string _nacionalidadConyuge, string _celularConyuge, string _mailConyuge, string _categoria, string _observaciones)
        {
            this.id = _id;
            this.fechaDeIngreso = _fechaDeIngreso;
            this.titulo = _titulo;
            this.apellido = _apellido;
            this.nombreDePila = _nombreDePila;
            this.nombreCompleto = _nombreCompleto;
            this.domicilio = _domicilio;
            this.codigoPostal = _codigoPostal;
            this.localidad = _localidad;
            this.cliente = _cliente;
            this.imprimeCarta = _imprimeCarta;
            this.estadoActual = _estadoActual;
            this.celular = _celular;
            this.telComercial = _telComercial;
            this.telParticular = _telParticular;
            this.email = _email;
            this.diaCumpleanios = _diaCumpleanios;
            this.nacionalidad = _nacionalidad;
            this.dniTitular = _dniTitular;
            this.cuitCuilTitular = _cuitCuilTitular;
            this.estadoCivil = _estadoCivil;
            this.apellidoConyuge = _apellidoConyuge;
            this.nombrePilaConyuge = _nombrePilaConyuge;
            this.nombreCompletoConyuge = _nombreCompletoConyuge;
            this.dniConyuge = _dniConyuge;
            this.cuitCuilConyuge = _cuitCuilConyuge;
            this.nacionalidadConyuge = _nacionalidadConyuge;
            this.celularConyuge = _celularConyuge;
            this.mailConyuge = _mailConyuge;
            this.categoria = _categoria;
            this.observaciones = _observaciones;
        }

        //COMPLETO
        //public PersonaExcel(int _id, DateTime _fechaDeIngreso, DateTime _actualizado, string _titulo, string _apellido, string _nombreDePila, string _nombreCompleto, string _domicilio,
        //int _codigoPostal, string _localidad, string _cliente, bool _imprimeCarta, string _estadoActual, string _celular, string _telComercial, string _telParticular,
        //string _email, DateTime _nacimiento, int _diaCumpleanios, string _nacionalidad, string _dniTitular, string _cuitCuilTitular, string _estadoCivil, string _apellidoConyuge,
        //string _nombrePilaConyuge, string _nombreCompletoConyuge, string _dniConyuge, string _cuitCuilConyuge, DateTime _nacimientoConyuge, string _nacionalidadConyuge, string _celularConyuge, string _mailConyuge, string _categoria, string _observaciones)
        //{
        //    this.id = _id;
        //    this.fechaDeIngreso = _fechaDeIngreso;
        //    this.actualizado = _actualizado;
        //    this.titulo = _titulo;
        //    this.apellido = _apellido;
        //    this.nombreDePila = _nombreDePila;
        //    this.nombreCompleto = _nombreCompleto;
        //    this.domicilio = _domicilio;
        //    this.codigoPostal = _codigoPostal;
        //    this.localidad = _localidad;
        //    this.cliente = _cliente;
        //    this.imprimeCarta = _imprimeCarta;
        //    this.estadoActual = _estadoActual;
        //    this.celular = _celular;
        //    this.telComercial = _telComercial;
        //    this.telParticular = _telParticular;
        //    this.email = _email;
        //    this.nacimiento = _nacimiento;
        //    this.diaCumpleanios = _diaCumpleanios;
        //    this.nacionalidad = _nacionalidad;
        //    this.dniTitular = _dniTitular;
        //    this.cuitCuilTitular = _cuitCuilTitular;
        //    this.estadoCivil = _estadoCivil;
        //    this.apellidoConyuge = _apellidoConyuge;
        //    this.nombrePilaConyuge = _nombrePilaConyuge;
        //    this.nombreCompletoConyuge = _nombreCompletoConyuge;
        //    this.dniConyuge = _dniConyuge;
        //    this.cuitCuilConyuge = _cuitCuilConyuge;
        //    this.nacimientoConyuge = _nacimientoConyuge;
        //    this.nacionalidadConyuge = _nacionalidadConyuge;
        //    this.celularConyuge = _celularConyuge;
        //    this.mailConyuge = _mailConyuge;
        //    this.categoria = _categoria;
        //    this.observaciones = _observaciones;
        //}

        //public PersonaExcel(string _nombre, string _apellido, string _direccion, string _email)
        //{
        //    this.nombreCompleto = _nombre;
        //    this.apellido = _apellido;
        //    this.domicilio = _direccion;
        //    this.email = _email;
        //}
    }
}
