using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace WordToPDFConversorWS.Tools
{
    public class FileUtils
    {

        public static void initLog()
        {
            log4net.Config.XmlConfigurator.Configure();
        }
        /// <summary>
        /// This method converts a File to a Base64 String of this File
        /// </summary>
        /// <param name="path">Path from file</param>
        /// <returns>Base64 String representation of file</returns>
        public static String FileToBase64(String path)
        {
            FileStream fis = new FileStream(path, FileMode.Open, FileAccess.Read);
            Byte[] barray = new Byte[fis.Length];
            fis.Read(barray, 0, (int)fis.Length);
            fis.Close();
            return Convert.ToBase64String(barray);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="base64File"></param>
        /// <returns></returns>
        public static String Base64ToTempFile(String base64File, String extension)
        {
            String result = GenerateTempPath(extension);
            FileStream fos = new FileStream(result, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            Byte[] barray = Convert.FromBase64String(base64File);
            fos.Write(barray, 0, barray.Length);
            fos.Flush();
            fos.Close();
            return result;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        private static string GenerateTempPath(string extension)
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Path.DirectorySeparatorChar + DateTime.Now.Ticks + extension;            
        }

        public static void DeleteTempFilesIgnoreExtension(String filename)
        {

            //try
            //{
            //    foreach (String s in Directory.GetFiles(Environment.SpecialFolder.ApplicationData))
            //    {

            //    }
            //}
            //catch (Exception)
            //{
                
                
            //}
        }
    }
}