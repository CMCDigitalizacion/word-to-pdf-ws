using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WordToPDFConversorWS.Managers;
using WordToPDFConversorWS.Tools;

namespace WordToPDFConversorWS
{
    public class WordToPdfController : ApiController
    {
        /// <summary>
        /// Convert the given word document into pdf
        /// </summary>
        /// <param name="fileBase64">Word document in Base64 Format</param>
        /// <returns>PDF document in Base64 Format</returns>
        public String Get(String fileBase64)
        {
            String path = FileUtils.Base64ToTempFile(fileBase64, Consts.Consts.WORD_EXT);
            path = WordManager.Instance().SaveDocAsPdf(path);
            return FileUtils.FileToBase64(path);
        }

        public String Post([FromBody]String fileBase64)
        {
            ILog Log = LogManager.GetLogger("Word2PdfLog.log");
            try
            {
                String path = FileUtils.Base64ToTempFile(fileBase64, Consts.Consts.WORD_EXT);
                path = WordManager.Instance().SaveDocAsPdf(path);
                return FileUtils.FileToBase64(path);
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
                Log.Error(ex.Source);
                throw;
            }
        }

        public String Get()
        {
            return "Alive";
        }
    }
}
