using System;
using P = DocumentFormat.OpenXml.Presentation;
using ODD = DocumentFormat.OpenXml.Wordprocessing;
using ODP = DocumentFormat.OpenXml.Drawing;

namespace DocToPptC
{
    internal class Program
    {
        private static void Main(string[] args)
        {
           /*Вводная параграф*/
            OfficeEx.Word = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 1, 0, 0);

            
            /*заголовок TODO*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
//            OfficeEx.Word = OfficeEx.ReadWordIp(5);
//            OfficeEx.Txt = "";
//            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 1, 0, 0);
         
           /*температура*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
//            OfficeEx.Word = OfficeEx.ReadWordIp(2);
//            OfficeEx.Txt = OfficeEx.Word;
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 1, 1, 0);

            /*осадки*/
/*            string osadB = OfficeEx.ReadWordIp(9);
              OfficeEx.Word = osadA + " " + osadB;*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
//            string osadA = OfficeEx.ReadWordIp(8);
//            OfficeEx.Word = osadA;
//            OfficeEx.Txt = OfficeEx.Word;
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 2, 1, 0); 

            /*Направление и скорость ветра*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
//            string veterA = OfficeEx.ReadWordIp(5).ToUpper().Remove(1, 5).Remove(3, 7);
//            string veterB = OfficeEx.ReadWordIp(6);
//            string veterC = OfficeEx.ReadWordIp(7);
//            OfficeEx.Word = veterA + " " + veterB + " " + veterC;
//            OfficeEx.Txt = OfficeEx.Word.Remove(9, 2);
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 3, 1, 0);

            /*Пострадало*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
//            string gervA = OfficeEx.ReadWordIp(53);
//            OfficeEx.Word = gervA;
//            OfficeEx.Txt = OfficeEx.Word;
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 1, 1, 0);

            /*Погибло*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
//            string gibelA = OfficeEx.ReadWordIp(57);
//            OfficeEx.Word = gibelA;
//            OfficeEx.Txt = OfficeEx.Word;
//            Console.WriteLine(Offic
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 2, 1, 0);

            /*Госпитализированно*/
//            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
//            string gospitA = OfficeEx.ReadWordIp(60);
//            OfficeEx.Word = gospitA;
//            OfficeEx.Txt = OfficeEx.Word;
//            Console.WriteLine
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 3, 1, 0);



            /*От МЧС л/с, чел.*/
//            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 8, 2);
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 3, 1, 0);

            /*От МЧС тех.ед*/
//            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 8, 3);
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 3, 2, 0);

            /*Всего л/с, чел.*/
//            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 20, 2);
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 4, 1, 0);

            /*Всего тех.ед*/
//            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 20, 3);
//            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 4, 2, 0);



        }

     
    }
}
