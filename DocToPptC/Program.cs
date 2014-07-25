﻿using System;
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
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 1, 0, 0);
            
            /*заголовок*/
            /*по параграфам 2-3 в заголовке*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            
//            Console.WriteLine(OfficeEx.Txt.IndexOf("погибло", System.StringComparison.Ordinal));
//            Console.ReadKey();
            /*первый параграф*/
            str.Add(OfficeEx.ReadWordIp(17));/*0*/
            str.Add(OfficeEx.ReadWordIp(18));/*1*/
            str.Add(OfficeEx.ReadWordIp(20));/*2*/
            str.Add(OfficeEx.ReadWordIp(21));/*3*/
            str.Add(OfficeEx.ReadWordIp(22));/*4*/
            str.Add(OfficeEx.ReadWordIp(23));/*5*/
            
            /*второй параграф параграф*/
            str.Add(OfficeEx.ReadWordIp(27));/*6*/
            str.Add(OfficeEx.ReadWordIp(28));/*7*/
            str.Add(OfficeEx.ReadWordIp(29));/*8*/

            /*первый параграф вывод*/
            OfficeEx.Txt = str[0] + " " + str[1] + " " + str[2] + " " + str[3] + " " + str[4] + " " + str[5];
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 1, 0);

            /*второй параграф параграф вывод*/
            OfficeEx.Txt = str[6] + "." + str[7] + " " + str[8];
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 2, 0);

            /*температура*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            OfficeEx.Txt = OfficeEx.ReadWordIp(2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 1, 1, 0);

            /*осадки*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            str.Add(OfficeEx.ReadWordIp(8));/*9*/
            OfficeEx.Txt = str[9];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 2, 1, 0); 

            /*Направление и скорость ветра*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            str.Add(OfficeEx.ReadWordIp(5).ToUpper().Remove(1, 5).Remove(3, 7));/*10*/
            str.Add(OfficeEx.ReadWordIp(6));/*11*/
            str.Add(OfficeEx.ReadWordIp(7));/*12*/
            OfficeEx.Txt = (str[10] + " " + str[11] + " " + str[12]).Remove(9, 2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 3, 1, 0);

            /*Пострадало*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(53));/*13*/
            OfficeEx.Txt = str[13];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 1, 1, 0);

            /*Погибло*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(57));/*14*/
            OfficeEx.Txt = str[14];
            
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 2, 1, 0);

            /*Госпитализированно*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            str.Add(OfficeEx.ReadWordIp(57));/*15*/
            OfficeEx.Txt = str[15];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 3, 1, 0);
            Console.WriteLine(OfficeEx.Txt);
            Console.ReadKey();

              /*От МЧС л/с, чел. */
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
