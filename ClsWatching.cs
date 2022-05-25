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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace ConsoleFileWatcher
{

   /// <summary>
   ///   Clase para monitorear los archivos Excel
   /// </summary>
   public class ClsWatching
   {

      private string _sPath;
      private string _sLibroDestino;

      public ClsWatching(string sPath, string sLibroDestino)
      {

         this._sPath = sPath;
         this._sLibroDestino = sLibroDestino;

         // crear directorios
         try
         {
            Directory.CreateDirectory(sPath + "\\PROCESSED");
            Directory.CreateDirectory(sPath + "\\NA");
         }
         catch (Exception ex)
         {

         }


         // crear libro destino
         try
         {
            if (File.Exists(sLibroDestino))
               File.Delete(sLibroDestino);

            _Application nexcelapp = new Application();
            Workbook nLibroDestino = nexcelapp.Workbooks.Add(Type.Missing);

            nLibroDestino.SaveCopyAs(sLibroDestino);
            nLibroDestino.Close();

            nexcelapp.Quit();
         }
         catch (Exception ex)
         {

         }

      }


      /// <summary>
      ///   Clase para monitorear los archivos Excel
      /// </summary>
      public void ClsMonitorFiles()
      {

         FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();

         fileSystemWatcher.Path = _sPath;

         fileSystemWatcher.Created += FileSystemWatcher_Created;
         fileSystemWatcher.EnableRaisingEvents = true;

      }



      /// <summary>
      ///   Evento que monitorea los archivos a copiar en el directorio a analizar
      /// </summary>
      private void FileSystemWatcher_Created(object sender, FileSystemEventArgs e)
      {

         string sExt = Path.GetExtension(e.FullPath);
         
         try
         {
            bool isHidden = ((File.GetAttributes(e.FullPath) & FileAttributes.Hidden) == FileAttributes.Hidden);
            Console.Write("Analyzing file: {0}{1}", e.FullPath, "...");


            if ((sExt.Equals(".xls") || sExt.Equals(".xlsx")) && (!isHidden))
            {
               ClsExcel clsExcelFile = new ClsExcel();
               clsExcelFile.pu_fn_CopiarHojas(e.FullPath, _sLibroDestino);

               Console.WriteLine("...Processed.", e.FullPath);

               // Mover archivos a carpeta de procesados
               try
               {
                  File.Move(e.FullPath, _sPath + "\\PROCESSED\\" + e.Name, true);
               }
               catch (Exception ex)
               {

               }

            }
            else
            {
               Console.WriteLine("......Not applicable.", e.FullPath);

               // Mover archivos a carpeta de No Aplicables
               try
               {
                  File.Move(e.FullPath, _sPath + "\\NA\\" + e.Name, true);
               }
               catch (Exception ex)
               {
               }
            }

            Console.WriteLine();
            Console.Write("...");

         }
         catch (Exception ex)
         {

         }

      }








   }





}
