﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EntititesExcel;
using BusinessExcel;

namespace ConsoleLinqToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<PersonaExcel> listaPersonas = new List<PersonaExcel>();

            listaPersonas = CargarLista();
            foreach (var p in listaPersonas)
            {
                Console.WriteLine("Nombre: " + p.nombreCompleto + ", " + p.apellido + " Cumple dia: " + p.diaCumpleanios);
            }

            Console.ReadKey();
        }

        public static List<PersonaExcel> CargarLista()
        {
            try
            {

                ExcelBusiness excelBusiness = new ExcelBusiness("C:\\Leo\\Dropbox\\Desarrollos\\Librerias\\ImportarExcel\\Pruebas\\BASE DE DATOS - CUMPLEAÑOS (ACTUAL).xlsx");
                List<PersonaExcel> lista = (List<PersonaExcel>)excelBusiness.RetornarRowExcel("BASE TOTAL DE CLIENTES CUMPLE");

                foreach (PersonaExcel persona in lista)
                {
                    string nombre = persona.nombreCompleto;
                }

                return lista;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
