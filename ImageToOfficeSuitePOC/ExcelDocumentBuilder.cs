using System;

namespace ImageToOfficeSuitePOC
{
    using System.Drawing;
    using System.IO;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class ExcelDocumentBuilder
    {
        public static void BuildDocumentWithImage(string filename, string sImagePath)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    Workbook workbook = new Workbook();

                    FileVersion fileVersion = new FileVersion();
                    fileVersion.ApplicationName = "Microsoft Office Excel";

                    Worksheet worksheet = new Worksheet();
                    SheetData sheetData = new SheetData();

                    DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                    ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png, worksheetPart.GetIdOfPart(drawingsPart));

                    using (FileStream fileStream = new FileStream(sImagePath, FileMode.Open))
                    {
                        imagePart.FeedData(fileStream);
                    }

                    NonVisualDrawingProperties nvdp = new NonVisualDrawingProperties();
                    nvdp.Id = 1025;
                    nvdp.Name = "Picture 1";
                    DocumentFormat.OpenXml.Drawing.PictureLocks picLocks =
                        new DocumentFormat.OpenXml.Drawing.PictureLocks
                        {
                            NoChangeAspect = true,
                            NoChangeArrowheads = true
                        };

                    NonVisualPictureDrawingProperties drawingProperties = new NonVisualPictureDrawingProperties
                    {
                        PictureLocks = picLocks
                    };

                    NonVisualPictureProperties pictureProperties = new NonVisualPictureProperties
                    {
                        NonVisualDrawingProperties = nvdp,
                        NonVisualPictureDrawingProperties = drawingProperties
                    };

                    DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch
                    {
                        FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle()
                    };

                    BlipFill blipFill = new BlipFill();
                    DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip
                    {
                        Embed = drawingsPart.GetIdOfPart(imagePart),
                        CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                    };

                    blipFill.Blip = blip;
                    blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
                    blipFill.Append(stretch);

                    DocumentFormat.OpenXml.Drawing.Transform2D transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
                    DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset
                    {
                        X = 0,
                        Y = 0
                    };
                    transform2D.Offset = offset;
                    Bitmap bitmap = new Bitmap(sImagePath);
                    //http://en.wikipedia.org/wiki/English_Metric_Unit#DrawingML
                    //http://stackoverflow.com/questions/1341930/pixel-to-centimeter
                    //http://stackoverflow.com/questions/139655/how-to-convert-pixels-to-points-px-to-pt-in-net-c
                    DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents
                    {
                        Cx = bitmap.Width * (long)(914400 / bitmap.HorizontalResolution),
                        Cy = bitmap.Height * (long)(914400 / bitmap.VerticalResolution)
                    };
                    bitmap.Dispose();
                    transform2D.Extents = extents;
                    ShapeProperties shapeProperties = new ShapeProperties
                    {
                        BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto,
                        Transform2D = transform2D
                    };

                    DocumentFormat.OpenXml.Drawing.PresetGeometry presetGeometry =
                        new DocumentFormat.OpenXml.Drawing.PresetGeometry
                        {
                            Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle,
                            AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                        };

                    shapeProperties.Append(presetGeometry);
                    shapeProperties.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

                    DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture
                        {
                            NonVisualPictureProperties = pictureProperties,
                            BlipFill = blipFill,
                            ShapeProperties = shapeProperties
                        };

                    Position pos = new Position { X = 0, Y = 0 };
                    Extent ext = new Extent { Cx = extents.Cx, Cy = extents.Cy };
                    AbsoluteAnchor anchor = new AbsoluteAnchor { Position = pos, Extent = ext };

                    anchor.Append(picture);
                    anchor.Append(new ClientData());

                    WorksheetDrawing worksheetDrawing = new WorksheetDrawing();
                    worksheetDrawing.Append(anchor);
                    Drawing drawing = new Drawing { Id = drawingsPart.GetIdOfPart(imagePart) };

                    worksheetDrawing.Save(drawingsPart);

                    worksheet.Append(sheetData);
                    worksheet.Append(drawing);

                    worksheetPart.Worksheet = worksheet;
                    worksheetPart.Worksheet.Save();

                    Sheets sheets = new Sheets();
                    Sheet sheet = new Sheet();

                    sheet.Name = "Sheet1";
                    sheet.SheetId = 1;
                    sheet.Id = workbookPart.GetIdOfPart(worksheetPart);

                    sheets.Append(sheet);
                    workbook.Append(fileVersion);
                    workbook.Append(sheets);

                    spreadsheetDocument.WorkbookPart.Workbook = workbook;
                    spreadsheetDocument.WorkbookPart.Workbook.Save();
                    spreadsheetDocument.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadLine();
            }
        }
    }
}
