using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.IO;
using System.Windows;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Drawing.Imaging;
using System.Drawing;
using Color = System.Drawing.Color;
using Pen = System.Drawing.Pen;
using Brush = System.Drawing.Brush;

namespace MyExport
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string imagePath = "barchart.png";
            string docxPath = "test.docx";  // 指定已有的DOCX文件路径

            // 生成柱状图
            CreateBarChart(imagePath);

            // 打开已有的 DOCX 文件并附加内容
            AppendToExistingDocx(docxPath, imagePath);
        }

        public void CreateBarChart(string imagePath)
        {
            int width = 500;
            int height = 300;
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                Pen pen = new Pen(Color.Black);
                Brush brush = new SolidBrush(Color.Blue);

                // 示例数据
                int[] data = { 10, 20, 30, 40, 50 };
                int barWidth = 40;
                int spacing = 20;
                int maxValue = 50;

                for (int i = 0; i < data.Length; i++)
                {
                    int barHeight = (data[i] * (height - 50)) / maxValue;
                    g.FillRectangle(brush, spacing + i * (barWidth + spacing), height - barHeight - 30, barWidth, barHeight);
                    g.DrawRectangle(pen, spacing + i * (barWidth + spacing), height - barHeight - 30, barWidth, barHeight);
                }
            }

            bmp.Save(imagePath, ImageFormat.Png);
        }

        public void AppendToExistingDocx(string docxPath, string imagePath)
        {
            // 打开已有的 Word 文档
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docxPath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;

                // 添加段落
                Paragraph para = new Paragraph(new Run(new Text("下面是生成的柱状图：")));
                body.AppendChild(para);

                // 添加柱状图图片
                ImagePart imagePart = wordDocument.MainDocumentPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                // 创建图片元素并插入到文档中
                AddImageToBody(wordDocument, wordDocument.MainDocumentPart.GetIdOfPart(imagePart));

                // 保存文档
                wordDocument.MainDocumentPart.Document.Save();
            }
        }

        private void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            var element =
                 new Drawing(
                     new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DocumentFormat.OpenXml.Drawing.Graphic(
                             new DocumentFormat.OpenXml.Drawing.GraphicData(
                                 new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                     new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.png"
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                         new DocumentFormat.OpenXml.Drawing.Blip()
                                         {
                                             Embed = relationshipId
                                         },
                                         new DocumentFormat.OpenXml.Drawing.Stretch(
                                             new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                     new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                         new DocumentFormat.OpenXml.Drawing.Transform2D(
                                             new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                             new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                             new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                         )
                                         { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                     );

            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}