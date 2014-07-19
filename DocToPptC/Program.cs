using System;
using System.Collections.Generic;
using P = DocumentFormat.OpenXml.Presentation;
using ODD = DocumentFormat.OpenXml.Wordprocessing;
using ODP = DocumentFormat.OpenXml.Drawing;


namespace DocToPptC
{
    internal class Program
    {
        private static void Main(string[] args)
        {
           List<String> str = new List<string>();
           /*Карта района ЧС*/
           /*Вводная параграф*/
            OfficeEx.Word = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 1, 0, 0);
            
            /*заголовок*/
            /*по параграфам 2-3 в заголовке*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            /*первый параграф*/
            str.Add(OfficeEx.ReadWordIp(17));/*1*/
            str.Add(OfficeEx.ReadWordIp(18));/*2*/
            str.Add(OfficeEx.ReadWordIp(20));/*3*/
            str.Add(OfficeEx.ReadWordIp(21));/*4*/
            str.Add(OfficeEx.ReadWordIp(22));/*5*/
            str.Add(OfficeEx.ReadWordIp(23));/*6*/
            
            /*второй параграф параграф*/
            str.Add(OfficeEx.ReadWordIp(27));/*7*/
            str.Add(OfficeEx.ReadWordIp(28));/*8*/
            str.Add(OfficeEx.ReadWordIp(29));/*9*/
            
            /*первый параграф вывод*/
            OfficeEx.Word = str[1] + " " + str[2] + " " + str[3] + " " + str[4] + " " + str[5] + " " + str[6];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 1, 0);

            /*второй параграф параграф вывод*/
            OfficeEx.Word = str[7] + "." + str[8] + " " + str[9];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 2, 0);

            /*температура*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            OfficeEx.Word = OfficeEx.ReadWordIp(2);
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 1, 1, 0);

            /*осадки*/
/*            string osadB = OfficeEx.ReadWordIp(9);
              OfficeEx.Word = osadA + " " + osadB;*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            str.Add(OfficeEx.ReadWordIp(8));/*10*/
            OfficeEx.Word = str[10];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 2, 1, 0); 

            /*Направление и скорость ветра*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            str.Add(OfficeEx.ReadWordIp(5).ToUpper().Remove(1, 5).Remove(3, 7));/*11*/
            str.Add(OfficeEx.ReadWordIp(6));/*12*/
            str.Add(OfficeEx.ReadWordIp(7));/*13*/
            OfficeEx.Word = str[11] + " " + str[12] + " " + str[13];
            OfficeEx.Txt = OfficeEx.Word.Remove(9, 2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 3, 1, 0);

            /*Пострадало*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(53));/*14*/
            OfficeEx.Word = str[14];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 1, 1, 0);

            /*Погибло*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(57));/*15*/
            OfficeEx.Word = str[15];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 2, 1, 0);

            /*Госпитализированно*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(57));/*16*/
            OfficeEx.Word = str[16];
            OfficeEx.Txt = OfficeEx.Word;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 3, 1, 0);

            /*От МЧС л/с, чел.*/
            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 8, 2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 3, 1, 0);

            /*От МЧС тех.ед*/
            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 8, 3);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 3, 2, 0);

            /*Всего л/с, чел.*/
            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 20, 2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 4, 1, 0);

            /*Всего тех.ед*/
            OfficeEx.Txt = OfficeEx.ExcelGetVal(OfficeEx.ExcelFile1, 0, 0, 20, 3);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 4, 4, 2, 0);



        }

     
    }
}
