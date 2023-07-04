using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Presentation;
using System.Drawing.Drawing2D;
using Rectangle = System.Drawing.Rectangle;

namespace GenerateTPSReport
{

    class ImageUtility
    {

        /// <summary>
        /// Resize the image to the specified width and height.
        /// </summary>
        /// <param name="image">The image to resize.</param>
        /// <param name="width">The width to resize to.</param>
        /// <param name="height">The height to resize to.</param>
        /// <returns>The resized image.</returns>
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
    }
    class ExcelUtility
    {

        /// <summary>
        /// Add an image to a specific cell of an existing excel spreadsheet/sheet
        /// Create a new spreadsheet and add an image to a specific cell
        /// Add an image to a specific cell of a WorksheetPart(so other operations can be performed before it's saved)
        /// Add the image from a stream
        /// Add the image from a file
        /// Use different image formats
        /// </summary>
            public static ImagePartType GetImagePartTypeByBitmap(Bitmap image)
            {
                if (ImageFormat.Bmp.Equals(image.RawFormat))
                    return ImagePartType.Bmp;
                else if (ImageFormat.Gif.Equals(image.RawFormat))
                    return ImagePartType.Gif;
                else if (ImageFormat.Png.Equals(image.RawFormat))
                    return ImagePartType.Png;
                else if (ImageFormat.Tiff.Equals(image.RawFormat))
                    return ImagePartType.Tiff;
                else if (ImageFormat.Icon.Equals(image.RawFormat))
                    return ImagePartType.Icon;
                else if (ImageFormat.Jpeg.Equals(image.RawFormat))
                    return ImagePartType.Jpeg;
                else if (ImageFormat.Emf.Equals(image.RawFormat))
                    return ImagePartType.Emf;
                else if (ImageFormat.Wmf.Equals(image.RawFormat))
                    return ImagePartType.Wmf;
                else
                    throw new Exception("Image type could not be determined.");
            }

            public static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
            {
                IEnumerable<Sheet> sheets =
                   document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                   Elements<Sheet>().Where(s => s.Name.ToString().Contains(sheetName));

                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist
                    return null;
                }

                string relationshipId = sheets.First().Id.Value;
                return (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            }

            public static void AddImage(bool createFile, string excelFile, string sheetName,
                                        string imageFileName, string imgDesc,
                                        int colNumber, int rowNumber)
            {
                using (var imageStream = new FileStream(imageFileName, FileMode.Open))
                {
                    AddImage(createFile, excelFile, sheetName, imageStream, imgDesc, colNumber, rowNumber);
                }
            }

            public static void AddImage(WorksheetPart worksheetPart,
                                        string imageFileName, string imgDesc,
                                        int colNumber, int rowNumber)
            {
                using (var imageStream = new FileStream(imageFileName, FileMode.Open))
                {
                AddImage(worksheetPart, imageStream, imgDesc, colNumber, rowNumber);
                }
            }

            public static void AddImage(bool createFile, string excelFile, string sheetName,
                                        Stream imageStream, string imgDesc,
                                        int colNumber, int rowNumber)
            {
                SpreadsheetDocument spreadsheetDocument = null;
                WorksheetPart worksheetPart = null;
                if (createFile)
                {
                    // Create a spreadsheet document by supplying the filepath
                    spreadsheetDocument = SpreadsheetDocument.Create(excelFile, SpreadsheetDocumentType.Workbook);

                    // Add a WorkbookPart to the document
                    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart
                    worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    // Add Sheets to the Workbook
                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                        AppendChild<Sheets>(new Sheets());

                    // Append a new worksheet and associate it with the workbook
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = sheetName
                    };
                    sheets.Append(sheet);
                }
                else
                {
                    // Open spreadsheet
                    spreadsheetDocument = SpreadsheetDocument.Open(excelFile, true);

                    // Get WorksheetPart
                    worksheetPart = GetWorksheetPartByName(spreadsheetDocument, sheetName);
                }

                AddImage(worksheetPart, imageStream, imgDesc, colNumber, rowNumber);

                worksheetPart.Worksheet.Save();

                spreadsheetDocument.Close();
            }

            public static void AddImage(WorksheetPart worksheetPart,
                                        Stream imageStream, string imgDesc,
                                        int colNumber, int rowNumber)
            {
                // We need the image stream more than once, thus we create a memory copy
                MemoryStream imageMemStream = new MemoryStream();
                imageStream.Position = 0;
                imageStream.CopyTo(imageMemStream);
                imageStream.Position = 0;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                    drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
                {
                    worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                }

                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;

                Bitmap bm = new Bitmap(imageMemStream);
                var imagePart = drawingsPart.AddImagePart(GetImagePartTypeByBitmap(bm));
                imagePart.FeedData(imageStream);

                A.Extents extents = new A.Extents();
                var extentsCx = bm.Width * (long)(914400 / bm.HorizontalResolution);
                var extentsCy = bm.Height * (long)(914400 / bm.VerticalResolution);
                bm.Dispose();

                var colOffset = 0;
                var rowOffset = 0;

                var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
                var nvpId = nvps.Count() > 0
                    ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                    : 1U;

                var oneCellAnchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker
                    {
                        ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                        RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                        ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                        RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                    },
                    new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDesc },
                            new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                        ),
                        new Xdr.BlipFill(
                            new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                worksheetDrawing.Append(oneCellAnchor);
            }
        /*
        public static void AddImage(WorksheetPart worksheetPart, string imageFileName, int rowNumber, int colNumber, int rowOffset, int colOffset)
        {
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            var imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(imageFileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            Bitmap bm = new Bitmap(imageFileName);
            DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            var extentsCx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
            var extentsCy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
            bm.Dispose();

            // var colOffset = 0;
            //  var rowOffset = 0;
            // int colNumber = 5;
            // int rowNumber = 10;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0 ?
                (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imageFileName },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);

           
        }
        public static void CreatePackage(string sFile, string imageFileName, int colNumber, int rowNumber, string SheetName, int  colOffset = 0, int rowOffset = 0)
        {
            try
            {
                // Create a spreadsheet document by supplying the filepath. 
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                    Create(sFile, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document. 
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart. 
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook. 
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook. 
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = SheetName
                };
                sheets.Append(sheet);


                //AddImage(worksheetPart, imageFileName, rowNumber, colNumber, rowOffset, colOffset);
                AddImage(worksheetPart, imageFileName, rowNumber, colNumber + 8, rowOffset, colOffset);
               // AddImage(worksheetPart, imageFileName, 6, rowNumber, rowOffset, colOffset);

                workbookpart.Workbook.Save();

                // Close the document. 
                spreadsheetDocument.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }*/
    }
}
