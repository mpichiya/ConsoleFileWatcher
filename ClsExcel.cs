/// <summary>
///   Herramienta para monitorear archivos Excel (*.xls*)
///   
///   Author: Miguel Leonardo Pichiyá Catú
///   Contacto: miguelpichiya@hotmail.com
/// </summary>
/// 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace ConsoleFileWatcher
{


   /// <summary>
   ///   Clase para procesar las hojas del archvio origen hacia el archivo destino
   /// </summary>
   public class ClsExcel
   {



      public ClsExcel()
      {
      }



    /// <summary>
    ///     Función para copiar hojas de un libro de excel a otro libro destino.
    ///     Las hojas se consolidan en Libro99.xlsx
    /// </summary>
    /// <param name="sExcelFile">
    ///     El libro origen de donde copiar las hojas
    /// </param>
    /// <param name="sExcelDestino">
    ///     El libro destino a donde consolidar las hojas
    /// </param>
    /// <returns>0:Error, 1:OK</returns>
    public int pu_fn_CopiarHojas(string sExcelFile, string sExcelDestino)
      {

         _Application excelapp = new Application();


         Workbook workbookOrigen = null;
         Workbook workbookDestino = null;
         Worksheet nuevaHojaDestino = null;
         Object defaultArg = Type.Missing;


         try
         {
            // Abrir libro .xls, .xlsx
            workbookOrigen = excelapp.Workbooks.Open(sExcelFile,Type.Missing,true);

            // Recorrer todas las hojas
            foreach (Worksheet wsh in workbookOrigen.Sheets)
            {
               // Copiar contenido de hoja actual
               wsh.UsedRange.Copy(defaultArg);

               // Pegar contenido de hoja actual en Libro Destino, agregando una hoja hasta la derecha
               workbookDestino = excelapp.Workbooks.Open(sExcelDestino);
               nuevaHojaDestino = (Worksheet)workbookDestino.Worksheets.Add(After: workbookDestino.Sheets[workbookDestino.Sheets.Count]);
               try
               {
                  nuevaHojaDestino.Name = wsh.Name;
               }
               catch (Exception ee)
               {
               }
               nuevaHojaDestino.UsedRange.PasteSpecial(XlPasteType.xlPasteAll);

            }

            //Guardar libro destino y cerrar libro origen
            workbookDestino.Save();
            workbookDestino.Close();

            workbookOrigen.Close();

            //Cerrar aplicación de Excel
            excelapp.Quit();

            return 1;  // Todas las hojas del libro procesadas correctamente

         }
         catch (Exception ex)
         {

         }

         return 0;  // Error en procesar las hojas
      }






   }






}






