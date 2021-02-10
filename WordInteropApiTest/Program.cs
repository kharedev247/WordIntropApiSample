using Microsoft.Office.Interop.Word;
using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace WordInteropApiTest
{
    class Program
    {
        object MissingType = Type.Missing;

        static void Main(string[] args)
        {
            try
            {
                var rootPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
                var prgObj = new Program();
                var wordDocPath = rootPath + "\\resource\\TestDoc.docx";
                var wordApp = prgObj.OpenWordDocument(wordDocPath);

                Thread.Sleep(2000);
                var newFilePath = rootPath + "\\resource\\newSnap.png";

                prgObj.ReplaceShape(wordApp, newFilePath);

            } catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            Console.ReadLine();
        }

        private Application OpenWordDocument(string docPath) {
            var app = new Application();
            app.Documents.Open(docPath);
            app.Visible = true;
            return app;
        }

        private void ReplaceShape(Application app, string newFilePath)
        {
            Console.WriteLine("Starting At:" + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));

            var activeDocument = app.ActiveDocument;
            activeDocument.Activate();

            InlineShapes shapes = activeDocument.InlineShapes;

            Console.WriteLine("InlineShapesCount::" + shapes.Count);

            var counter = 100;
            foreach (InlineShape oWordShape in shapes)
            {
                counter++;
                if (oWordShape.AlternativeText != null && oWordShape.AlternativeText.Contains(counter.ToString()))
                {
                    Console.WriteLine("Index:" + counter + "||Start::ReplaceImageinInlineShapeAt:" + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
                    object rangeAddress = oWordShape.Range;
                    var newShape = app.ActiveDocument.InlineShapes.AddPicture(newFilePath, ref MissingType, ref MissingType, ref rangeAddress);
                    newShape.AlternativeText = oWordShape.AlternativeText;
                    oWordShape.Delete();
                    Console.WriteLine("Index:" + counter + "End::ReplaceImageinInlineShapeAt:" + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
                }
            }
        }
    }
}
