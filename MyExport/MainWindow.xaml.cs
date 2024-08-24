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
using System;
using Font = System.Drawing.Font;
using System.Drawing.Drawing2D;
using LiveCharts.Wpf;
using LiveCharts;
using System.Windows.Media.Imaging;
using Size = System.Windows.Size;
using System.Windows.Media;
using LiveCharts.Definitions.Charts;
using System.Windows.Controls;
using Point = System.Windows.Point;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using Separator = LiveCharts.Wpf.Separator;
using Brushes = System.Windows.Media.Brushes;

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

            BuildPngOnClick(sender, e);
            AppendToExistingDocx(docxPath, imagePath);
        }

        private void BuildPngOnClick(object sender, RoutedEventArgs e)
        {
            var myChart = new CartesianChart
            {
                DisableAnimations = true,
                Width = 600,
                Height = 200,
                Series = new SeriesCollection
                {
                    new ColumnSeries
                    {
                        DataLabels =  true,
                        Values = new ChartValues<double> {1, 6, 7, 2, 9, 3, 6, 5}
                    }
                },
                AxisX = new AxesCollection
                {
                    new Axis
                    {
                        Title = "X Axis",
                        Labels = new[] {"A", "B", "C", "D", "E", "F", "G", "H"},
                        Position = AxisPosition.LeftBottom, // 这表明这个轴是X轴，并且位于底部
                        Separator = new Separator
                        {
                            IsEnabled = true, // 隐藏网格线
                            Stroke = new SolidColorBrush(Colors.Gray), // X轴线条颜色
                            StrokeThickness = 0.5 // X轴线条粗细
                        },
                        
                    }
                },
                AxisY = new AxesCollection
                {
                    new Axis
                    {
                        Title = "Y Axis",
                        LabelFormatter = value => value.ToString("N"),
                        Separator = new Separator
                        {
                            IsEnabled = true, // 隐藏网格线
                            Stroke = new SolidColorBrush(Colors.Gray), // X轴线条颜色
                            StrokeThickness = 0.5 // X轴线条粗细
                        },
                    }
                }
            };

            var viewbox = new Viewbox();
            viewbox.Child = myChart;
            viewbox.Measure(new Size(myChart.Width, myChart.Height));
            viewbox.Arrange(new Rect(new Point(0, 0), new Size(myChart.Width, myChart.Height)));
            myChart.Update(true, true); // 强制重绘
            viewbox.UpdateLayout();

            var encoder = new PngBitmapEncoder();
            var bitmap = new RenderTargetBitmap((int)myChart.ActualWidth, (int)myChart.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(viewbox);
            var frame = BitmapFrame.Create(bitmap);
            encoder.Frames.Add(frame);
            using (var stream = File.Create("barchart.png")) encoder.Save(stream);
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