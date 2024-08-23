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
            // 创建一个柱状图
            var cartesianChart = new CartesianChart
            {
                Width = 600,
                Height = 400
            };

            // 创建数据系列
            var seriesCollection = new SeriesCollection
            {
                new ColumnSeries
                {
                    Title = "2024",
                    Values = new ChartValues<double> { 10, 50, 39, 50 }
                }
            };

            // 添加数据到柱状图
            cartesianChart.Series = seriesCollection;

            // 添加X轴标签
            cartesianChart.AxisX.Add(new Axis
            {
                Title = "Categories",
                Labels = new[] { "Category 1", "Category 2", "Category 3", "Category 4" }
            });

            // 添加Y轴
            cartesianChart.AxisY.Add(new Axis
            {
                Title = "Values",
                LabelFormatter = value => value.ToString("N")
            });

            // 将柱状图添加到窗口内容中（如果需要显示）
            this.Content = cartesianChart;

            // 确保图表已经正确布局并渲染
            cartesianChart.Measure(new Size(cartesianChart.Width, cartesianChart.Height));
            cartesianChart.Arrange(new Rect(new Size(cartesianChart.Width, cartesianChart.Height)));
            cartesianChart.UpdateLayout(); // 强制更新布局

            // 等待布局更新完成
            cartesianChart.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
            {
                cartesianChart.UpdateLayout();
                cartesianChart.InvalidateVisual();
            }));

            // 延迟以确保渲染完成
            System.Threading.Thread.Sleep(100);

            // 渲染并保存为图片
            var renderBitmap = new RenderTargetBitmap((int)cartesianChart.Width, (int)cartesianChart.Height, 96d, 96d, PixelFormats.Pbgra32);
            renderBitmap.Render(cartesianChart);

            using (var fileStream = new FileStream(imagePath, FileMode.Create))
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderBitmap));
                encoder.Save(fileStream);
            }

            MessageBox.Show($"柱状图已保存到 {imagePath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
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