using FelisOpenXml.FelisShape;
using FelisOpenXml.FelisShape.Text;
using System.Drawing;

namespace TestFelisShape
{
    internal class Program
    {
        static readonly string TestPresentationFilePath = @"assets\test.pptx";
        static bool waitKey = true;
        static bool outputSaveFile = true;

        static void Main(string[] _args)
        {
            Console.WriteLine($"Current Directory is \"{Environment.CurrentDirectory}\"");
            if (_args.Length <= 0)
            {
                TestAll(_args);
            }
            else
            {
                var methodInfo = typeof(Program).GetMethod($"Test{_args[0]}", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public, new[] { typeof(string[]) });
                if (null != methodInfo)
                {
                    methodInfo.Invoke(null, new[] { _args.Skip(1).ToArray() });
                }
                else
                {
                    Console.WriteLine($"Cannot found the test entry named \"Test{_args[0]}\"");
                    var methods = typeof(Program).GetMethods(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public).Where(e =>
                    {
                        var paramList = e.GetParameters();
                        return e.Name.StartsWith("Test") && (paramList?.Length == 1) && (paramList?[0].ParameterType == typeof(string[]));
                    });
                    Console.WriteLine($"Available argument can be one of follow:");
                    foreach (var method in methods) 
                    {
                        Console.WriteLine($"    {method.Name.Substring(4)}");
                    }
                }
            }
        }

        static void TipAndWait(string _tip)
        {
            Console.WriteLine(_tip);
            if (waitKey)
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey(true);
            }
        }

        public static void TestAll(string[] _args)
        {
            TestList(_args);
            TestCopyAndChange(_args);
        }

        public static void TestList(string[] _args)
        {
            TipAndWait("List the shapes and some properties of the shapes in each slide");

            Action<IEnumerable<FelisShape>, string?>? fnList = null;

            fnList = (IEnumerable<FelisShape> list, string? _prefix) =>
            {
                bool? isEditable = null;

                foreach (var shape in list)
                {
                    if (null == isEditable)
                    {
                        isEditable = FelisSlide.RetrospectToSlide(shape.Element)?.Writable;
                    }

                    Console.WriteLine($"{_prefix}Shape ({shape.Id} : {shape.Name}) {shape.GetType().Name}");
                    var r = shape.Rect;
                    var rr = shape.RelativeRect;
                    Console.WriteLine($"{_prefix}> [{r.x}, {r.y}, {r.cx}, {r.cy}] - [{rr.x}, {rr.y}, {rr.cx}, {rr.cy}] & ({shape.HasTextBody}, {shape.HasTextContent})");
                    if (shape.IsPlaceHolder)
                    {
                        Console.WriteLine($"{_prefix}> Placeholder = {shape.PlaceHolderTypeText}");
                    }
                    var textBody = shape.TextBody;
                    if (null != textBody)
                    {
                        var families = textBody.Paragraphs.SelectMany(e => e.TextRuns).SelectMany(e => e.Properties?.FontFamilies.Select(e => $"{e.Key}:{e.Value}") ?? Array.Empty<string>());
                        if (families.Any()) Console.WriteLine($"{_prefix}> FontFamily: {string.Join(",", families)}");
                        var sizes = textBody.Paragraphs.SelectMany(e => e.TextRuns).Select(e => Convert.ToString(e.Properties?.FontSize)).Where(e => !string.IsNullOrWhiteSpace(e));
                        if (sizes.Any()) Console.WriteLine($"{_prefix}> FontSizes: {string.Join(", ", sizes)}");
                        var colors = textBody.Paragraphs.SelectMany(e => e.TextRuns).Where(e => e.Properties?.Color?.HasDefined ?? false).Select(e => e.Properties!.Color!.Value).Select(e => e.ToArgb().ToString("X8")).Where(e => !string.IsNullOrWhiteSpace(e));
                        if (colors.Any()) Console.WriteLine($"{_prefix}> FontColors: {string.Join(", ", colors)}");
                        Console.WriteLine($"{_prefix}> TextContent:");
                        foreach (FelisTextParagraph paraph in textBody.Paragraphs)
                        {
                            Console.WriteLine($"{_prefix}> {paraph.Text}");
                        }
                    }

                    if (shape is IFelisShapeTree shapeTree)
                    {
                        fnList?.Invoke(shapeTree.Shapes, $"{_prefix} |  ");
                    }
                }
            };

            using (var pres = new FelisPresentation(File.OpenRead(TestPresentationFilePath)))
            {
                int idxSlide = 0;
                foreach (var slide in pres.Slides)
                {
                    Console.WriteLine($"== [Slide {++idxSlide}] ===============================");
                    fnList(slide.Shapes, null);
                }

                Console.WriteLine();
                TipAndWait("<<< Listing complete. Test searching special shape follow... >>>");
                Console.WriteLine();

                idxSlide = 0;
                foreach (var slide in pres.Slides)
                {
                    idxSlide++;
                    var shape = slide.GetShapeById(17, true);
                    if (null != shape)
                    {
                        Console.WriteLine($"Slide {idxSlide} has shape with id 17 named \"{shape.Name}\"");
                    }
                    var pictures = slide.GetShapes<FelisPicture>(true);
                    if (pictures.Any())
                    {
                        Console.WriteLine($"Slide {idxSlide} has picture shapes");
                        foreach (var item in pictures)
                        {
                            Console.WriteLine($"    Picture (id: {item.Id}, name: \"{item.Name})\"");
                        }
                    }
                }
            }
        }

        public static void TestCopyAndChange(string[] _args)
        {
            TipAndWait("Copy the slide and change the content of some shapes");
            using (var sourcePres = new FelisPresentation(File.OpenRead(TestPresentationFilePath)))
            {
                using (var targetPres = FelisPresentation.From(sourcePres))
                {
                    var slide = targetPres.InsertSlide(sourcePres.Slides.ElementAt(2));
                    var shape = slide?.GetShapeById(4);
                    if (null != shape)
                    {
                        var paragraph = shape.TextBody?.Paragraphs.Add(0);
                        var run = paragraph?.TextRuns.Add();
                        if (null != run)
                        {
                            run.Text = "Hello ";
                            var props = run.Properties;
                            props!.Color!.Value = Color.Red;
                            props!.Bold = true;
                            props!.Submit();
                        }
                        run = paragraph?.TextRuns.Add();
                        if (null != run)
                        {
                            run.Text = "World";
                            var props = run.Properties;
                            props!.Color!.Value = Color.Blue;
                            props!.FontSize = 2000;
                            props!.Submit();
                        }
                        paragraph = shape.TextBody?.Paragraphs.Add();
                        run = paragraph?.TextRuns.Add();
                        if (null != run)
                        {
                            run.Text = "Continue";
                            var props = run.Properties;
                            props!.Color!.Value = Color.Green;
                            props!.Italic = true;
                            props!.Underline = "heavy";
                            props!.Submit();
                        }
                    }

                    slide = targetPres.InsertSlide(sourcePres.Slides.ElementAt(7));
                    var pic = slide?.GetFirstShape<FelisPicture>();
                    if (null != pic)
                    {
                        var newImgStream = typeof(Program).Assembly.GetManifestResourceStream("TestFelisShape.assets.cat.png");
                        if (null != newImgStream)
                        {
                            using (newImgStream)
                            {
                                pic.Set(newImgStream, "png");
                            }
                        }
                    }

                    slide = targetPres.InsertSlide(sourcePres.Slides.ElementAt(4));
                    var chart = slide?.GetFirstShape<FelisChart>();
                    if (null != chart)
                    {
                        chart.SeriesCollection.Clear();
                        var series = chart.SeriesCollection.Add();
                        if (null != series)
                        {
                            series.Title = "Object-1";
                            series.Categories.ReWrite("Width", "Height", "Weight");
                            series.Values.ReWrite(100, 200, 120);
                        }
                        series = chart.SeriesCollection.Add();
                        if (null != series)
                        {
                            series.Title = "Object-2";
                            series.Categories.ReWrite("Width", "Height", "Weight");
                            series.Values.ReWrite(90, 170, 97);
                        }
                        series = chart.SeriesCollection.Add();
                        if (null != series)
                        {
                            series.Title = "Object-3";
                            series.Categories.ReWrite("Width", "Height", "Weight");
                            series.Values.ReWrite(260, 130, 150);
                        }
                    }

                    slide = targetPres.InsertSlide(sourcePres.Slides.ElementAt(5));
                    var table = slide?.GetFirstShape<FelisTable>();
                    if (null != table)
                    {
                        if (null != table.Cell(0, 0)?.TextBody)
                        {
                            table.Cell(0, 0)!.TextBody!.Text = "Hello World";
                            var props = table.Cell(0, 0)!.TextBody!.Paragraphs[0]!.TextRuns[0]!.Properties;
                            if (null != props)
                            {
                                props.Color!.Value = Color.Brown;
                                props.Submit();
                            }
                        }
                    }

                    slide = targetPres.InsertSlide(targetPres.Slides.ElementAt(0));
                    shape = slide?.GetShapeById(4);
                    if (null != shape)
                    {
                        var paragraph = shape.TextBody?.Paragraphs.Add();
                        var run = paragraph?.TextRuns.Add();
                        if (null != run)
                        {
                            run.Text = "The END";
                            var props = run.Properties;
                            props!.Bold = true;
                            props!.Submit();
                        }
                    }

                    if (outputSaveFile)
                    {
                        var saveFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "testModified.pptx");
                        targetPres.Save(saveFilePath);
                    }
                }
            }
        }

        public static void TestPressure(string[] _args)
        {
            const uint DefaultRepeatCount = 1000;
            uint repeatCount;
            if (_args.Length > 0)
            {
                if (!uint.TryParse(_args[0], out repeatCount))
                {
                    repeatCount = DefaultRepeatCount;
                }
            }
            else
            {
                repeatCount = DefaultRepeatCount;
            }
            var backupWaitKey = waitKey;
            var backupOutputSaveFile = outputSaveFile;
            try
            {
                waitKey = false;
                outputSaveFile = false;
                Console.WriteLine($"Pressure test (repeat for {repeatCount} times) ...");
                for (; repeatCount > 0; repeatCount--)
                {
                    TestList(Array.Empty<string>());
                    TestCopyAndChange(Array.Empty<string>());
                }
                Console.WriteLine("Done");
            }
            catch (Exception err)
            {
                Console.WriteLine($"[Error]{Environment.NewLine}{err.ToString()}");
            }
            finally
            {
                waitKey = backupWaitKey;
                outputSaveFile = backupOutputSaveFile;
            }
        }
    }
}