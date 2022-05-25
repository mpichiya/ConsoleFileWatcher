/// <summary>
///   Herramienta para monitorear archivos Excel (*.xls*)
///   
///   Author: Miguel Leonardo Pichiyá Catú
///   Contacto: miguelpichiya@hotmail.com
/// </summary>
/// 

using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;



namespace ConsoleFileWatcher
{




   class Program
   {



      /// <summary>
      ///   Programa principal
      ///   Se ingresa el directorio a analizar y se invoca el método para monitorear y procesar los libros de excel.
      /// </summary>
      static void Main(string[] args)
      {
         string sDestino = "Libro99.xlsx";

         Console.WriteLine();
         Console.WriteLine("------------------");
         Console.WriteLine("File Watcher for Excel files (.xls*).");
         Console.WriteLine("-------------------------------------------");
         Console.WriteLine();

         string sPath = null;

         Console.Write("Please, enter the path of the folder to watch it (e.g. c:\\temp): ");
         sPath = Console.ReadLine();
         Console.WriteLine();
         Console.WriteLine();

         Console.WriteLine("Watching files in {0}",sPath);
         Console.WriteLine("---");
         Console.WriteLine("Watching... Please copy/move files to this directory, press any key to exit.", sPath);
         Console.WriteLine();
         Console.Write("...");


         ClsWatching clsWathFiles = new ClsWatching(sPath, sPath + "\\" + sDestino);
         clsWathFiles.ClsMonitorFiles();

         Console.ReadKey();
         Console.WriteLine("Finishing.");
         Console.WriteLine();
         Console.WriteLine("All worksheets are consildated in Libro99.xlsx. Please review.");
         Console.WriteLine();
         Console.WriteLine("Workbooks procesed are in directory PROCESSED");
         Console.WriteLine("Workbooks not applicables are in directory NA");
         Console.WriteLine();
         Console.WriteLine("Bye.");
         Console.WriteLine();

      }



   }




}














