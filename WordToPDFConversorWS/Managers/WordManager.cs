using log4net;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using WordToPDFConversorWS.Consts;

namespace WordToPDFConversorWS.Managers
{
    public class WordManager
    {
        private static WordManager instance;

        private WordManager()
        {

        }

        public static WordManager Instance(){
            if(instance==null){
                instance = new WordManager();
            }
            return instance;
        }


		public String SaveDocAsPdf(String path)
		{
            ILog Log = LogManager.GetLogger("Word2PdfLog.log");
            // Create a new Microsoft Word application object
            Log.Debug("Initializing Word Interop");
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Log.Debug("Application Word working");
            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

			FileInfo wordFile = new FileInfo(path);
            Object filename = (Object)wordFile.FullName;
            Log.Debug("Opening Word file");
	
            // Use the dummy value as a placeholder for optional arguments
            Document doc = word.Documents.Open(ref filename, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,		
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);	
            doc.Activate();
            Log.Debug("Word file opened");
            object outputFileName = wordFile.FullName.Replace(Consts.Consts.WORD_EXT, Consts.Consts.PDF_EXT);	
            object fileFormat = WdSaveFormat.wdFormatPDF;
            Log.Debug("PDF file created");
	
            // Save document into PDF Format	
            doc.SaveAs(ref outputFileName,ref fileFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,		
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            Log.Debug("PDF File saved");
            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the	
            // correct Close method.                	
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;	
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);	
            doc = null;
            Log.Debug("Word doc closed");
            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            Log.Debug("Word Exit");
            word = null;
            return outputFileName.ToString();
		}
 

    }
}