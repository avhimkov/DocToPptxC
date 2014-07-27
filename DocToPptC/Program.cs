using System;
using System.Collections.Generic;

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
            
            /*первый параграф*/
            int parag1 = OfficeEx.SearchString("пожаре");
            str.Add(OfficeEx.ReadWordIp(parag1));/*7*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 1));/*8*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 3));/*10*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 4));/*21*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 5));/*22*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 6));/*23*/
           
            /*второй параграф параграф*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 10));/*6 27*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 11));/*7 28*/
            str.Add(OfficeEx.ReadWordIp(parag1 + 12));/*8 29*/

            /*первый параграф вывод*/
            OfficeEx.Txt = str[0].Remove(5) + " " + str[1] + " " + str[2] + " " + str[3]  + " " + str[4] + " " + str[5];
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 1, 0);

            /*второй параграф параграф вывод*/
            OfficeEx.Txt = str[6] + "." + str[7] + "." + " " + str[8];
            OfficeEx.PptxGetPar(OfficeEx.PptxFile, 0, 0, 2, 0);

            /*температура*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            int parag2 = OfficeEx.SearchString("температура");
            OfficeEx.Txt = OfficeEx.ReadWordIp(parag2 + 1);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 1, 1, 0);

            /*осадки TODO*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            str.Add(OfficeEx.ReadWordIp(8));/*9*/
//            str.Add(OfficeEx.ReadWordIp(9));/*10*/
            OfficeEx.Txt = str[9]/*+ " " + str[10]*/;
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 2, 1, 0); 

            /*Направление и скорость ветра*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 6, 0, 0);
            int skorvetr = OfficeEx.SearchString("ветер");
            str.Add(OfficeEx.ReadWordIp(skorvetr + 1));/*10*/
            str.Add(OfficeEx.ReadWordIp(skorvetr + 2));/*11*/
            str.Add(OfficeEx.ReadWordIp(skorvetr + 3));/*12*/
            OfficeEx.Txt = (str[10].ToUpper().Remove(1, 5).Remove(3, 7) + " " + str[11] + " " + str[12]).Remove(9, 2);
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 1, 3, 1, 0);

            /*Пострадало*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            int postrd = OfficeEx.SearchString("погибло");
            str.Add(OfficeEx.ReadWordIp(postrd - 6));/*13*/
            OfficeEx.Txt = str[13];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 1, 1, 0);

            /*Погибло*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            int pogib = OfficeEx.SearchString("погибло");
            str.Add(OfficeEx.ReadWordIp(pogib - 2));/*14*/
            OfficeEx.Txt = str[14];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 2, 1, 0);

            /*Госпитализированно*/
            OfficeEx.Txt = OfficeEx.DocGetPar(OfficeEx.DocxFile, 5, 0, 0);
            int gospit = OfficeEx.SearchString("погибло");
            str.Add(OfficeEx.ReadWordIp(gospit + 1));/*15*/
            OfficeEx.Txt = str[15];
            OfficeEx.PptxGetTab(OfficeEx.PptxFile, 0, 3, 3, 1, 0);
            

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
