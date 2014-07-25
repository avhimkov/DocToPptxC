using System;
using P = DocumentFormat.OpenXml.Presentation;
using Slide = DocumentFormat.OpenXml.Presentation.Slide;
using ODD = DocumentFormat.OpenXml.Wordprocessing;
using ODP = DocumentFormat.OpenXml.Drawing;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;


namespace DocToPptC
{
    public class OfficeEx
    {
        public static string DocxFile = @"D:\DocToPPTС\Doc.docx";
        public static string PptxFile = @"D:\DocToPPTС\Presentation1.pptx";
        public static string ExcelFile1 = @"D:\DocToPPTС\Presentation1.xlsx";
//        public static string ExcelFile2 = @"Справка  СиС.xlsx";
        
        private static string _txt = "";

        public static string Txt
        {
            get { return _txt; }
            set { _txt = value; }
        }

        public static string ReadWordIp(int index)
        {
            /*Разбиваем параграф из DocGetPar на слова с индексом*/
            string[] words = Txt.Split(new[] { ' ', ',', ':', '?', '!', '.' }, StringSplitOptions.RemoveEmptyEntries);
            string word = words[index];
            return word;
        }

        public static string ExcelGetVal(string filepatch, int indexWsp, int indexSheetD, int indexRow, int indexCell)
        {
            /*ставим указатель на место вставки в таблице*/
            using (SpreadsheetDocument docExcel = SpreadsheetDocument.Open(filepatch, true))
            {
                WorkbookPart workbookPart = docExcel.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(indexWsp);
                SheetData sheetData = worksheetPart.Worksheet.ChildElements.OfType<SheetData>().ElementAt(indexSheetD);
                Row row1 = sheetData.ChildElements.OfType<Row>().ElementAt(indexRow);
                Cell cell1 = row1.ChildElements.OfType<Cell>().ElementAt(indexCell);
                CellValue cellValue1 = cell1.GetFirstChild<CellValue>();
                if (cellValue1.InnerText != "")
                {
                    Txt = cellValue1.InnerText;
                }
            }
            return Txt;
        }

        public static void PptxGetTab(string filepatch, int indexSlide, int indexGraphicFrame, int indexRow, int indexCell, int indexRun)
        /*ставим указатель на место вставки текста в таблицу*/
        {
            using (PresentationDocument prstDoc = PresentationDocument.Open(filepatch, true))
            {
                Slide firstSlide = prstDoc.PresentationPart.SlideParts.ElementAt(indexSlide).Slide;
                P.CommonSlideData commonSlideData1 = firstSlide.GetFirstChild<P.CommonSlideData>();
                P.ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<P.ShapeTree>();
                P.GraphicFrame graphicFrame2 = shapeTree1.Elements<P.GraphicFrame>().ElementAt(indexGraphicFrame);
                ODP.Graphic graphic2 = graphicFrame2.GetFirstChild<ODP.Graphic>();
                ODP.GraphicData graphicData2 = graphic2.GetFirstChild<ODP.GraphicData>();
                ODP.Table table1 = graphicData2.GetFirstChild<ODP.Table>();
                ODP.TableRow tableRow1 = table1.Elements<ODP.TableRow>().ElementAt(indexRow);
                ODP.TableCell tableCell1 = tableRow1.Elements<ODP.TableCell>().ElementAt(indexCell);
                ODP.TextBody textBody1 = tableCell1.GetFirstChild<ODP.TextBody>();
                ODP.Paragraph paragraph1 = textBody1.GetFirstChild<ODP.Paragraph>();
                ODP.Run run1 = paragraph1.ChildElements.OfType<ODP.Run>().ElementAt(indexRun);
                ODP.Text text1 = run1.GetFirstChild<ODP.Text>();
                text1.Text = Txt;
              }
        }

        public static void PptxGetPar(string filepatch, int indexslide, int indexshape, int indexpara, int indexrun)
        /*ставим указатель на место вставки текста*/
        {
            using (PresentationDocument prstDoc = PresentationDocument.Open(filepatch, true))
            {
                Slide firstSlide = prstDoc.PresentationPart.SlideParts.ElementAt(indexslide).Slide;
                Shape firstShape = firstSlide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().ElementAt(indexshape);
                ODP.Paragraph paraP = firstShape.TextBody.ChildElements.OfType<ODP.Paragraph>().ElementAt(indexpara);
                ODP.Text t = paraP.ChildElements.OfType<ODP.Run>().ElementAt(indexrun).Text;
                t.Text = Txt;
                prstDoc.PresentationPart.Presentation.Save();
            }
        }

        public static string DocGetPar(string filepatch, int indexparagraf, int indexrun, int indextext)
        /*Берем параграфы и слова из DOCX и возвращаем текст*/
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepatch, true))
            {
                ODD.Body body = wordDoc.MainDocumentPart.Document.Body;
                ODD.Paragraph para = body.ChildElements.OfType<ODD.Paragraph>().ElementAt(indexparagraf);
                ODD.Run run = para.ChildElements.OfType<ODD.Run>().ElementAt(indexrun);
                ODD.Text t = run.ChildElements.OfType<ODD.Text>().ElementAt(indextext);
                if (para.InnerText != "")
                {
                    Txt = para.InnerText;
                }
            }
            return Txt;
        } 
    }
}