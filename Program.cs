using System;
using System.Linq;
using System.IO;
using System.Net;
using Syncfusion.Presentation;
using System.Collections.Generic;

namespace Hueta
{    
    class Program
    {
        public class InputData
        {
            public string FolderPath { get; set; }
            public string ImgPath { get; set; }
        }

        static void Main(string[] args)
        {
            var data = GetInputData();
            var files = SearchPresentationInSpeceficFolder(data.FolderPath);
            MakeProgram(files, data.ImgPath);
        }
        static void MakeProgram(string[] paths, string img)
        {
            for (int i = 0; i < paths.Length; i++)
            {
                ChangePresentation(paths[i], img);
            }
        }
        
        static InputData GetInputData()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("Путь папки от куда будет пакасть:");
            Console.ForegroundColor = ConsoleColor.Magenta;
            string path = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write("Путь к фотографии:");
            Console.ForegroundColor = ConsoleColor.Magenta;
            string img = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Gray;
            InputData res = new InputData();
            res.FolderPath = path;
            res.ImgPath = img;
            return res;
        }

        static string[] SearchPresentationInSpeceficFolder(string folder)
        {
            if (Directory.Exists(folder))
            {
                Stack<string> paths = new Stack<string>();
                var ds = Directory.GetDirectories(folder);
                if (ds.Length > 0)
                {
                    foreach (var d in ds)
                    {
                        var files = SearchPresentationInSpeceficFolder(d);
                        foreach (var f in files)
                        {
                            paths.Push(f);
                        }
                    }
                    //return paths.ToArray();

                }
                string[] extensions = { ".pptx", ".pptm", ".potx", ".potm" };
                var filePaths = Directory.GetFiles(folder)
                    .Where(f => extensions.Contains(new FileInfo(f).Extension.ToLower())).ToArray();
                foreach (var f in filePaths)
                {
                    paths.Push(f);
                }
                return paths.ToArray();
            }
            else
            {
                Console.WriteLine("Папки не существует");
                return new string[] { };
            }
        }
        static void ChangePresentation(string path, string img = "")
        {
            IPresentation pptxDoc;
            try
            {
                pptxDoc = Presentation.Open(path);
            }
            catch
            {
                return;
            }
            ILayoutSlide layout = pptxDoc.Masters.First().LayoutSlides.First(l => l.Name == "Title and Content");
            var slide = pptxDoc.Slides.Add(layout);
            slide.Background.Fill.FillType = FillType.Picture;
            slide.Background.Fill.PictureFill.TileMode = TileMode.Stretch;
            slide.Background.Fill.PictureFill.ImageBytes = GetImgFromFolder(img);
            var title = slide.Shapes.First(s => s.ShapeName.ToLower().Contains("title")) as IShape;
            var placeholder = slide.Shapes.First(s => s.ShapeName.ToLower().Contains("placeholder")) as IShape;
            title.TextBody.AddParagraph("Презентация была захвачена");
            title.TextBody.AnchorCenter = true;
            title.Description = "short message";
            placeholder.TextBody.AddParagraph("Презентация была захвачена подразделением \"почемучки\" манную кашу, конфеты и игрушки на стол, БЫСТРО!");
            placeholder.TextBody.WrapText = true;
            try
            {
                pptxDoc.Save(path);
            }
            catch
            {
                Console.WriteLine($"На {path} произошда ошибка");
            }
            pptxDoc.Close();
        }
        static byte[] GetImg(string uri)
        {
            WebRequest request = WebRequest.CreateHttp(uri);
            var res = request.GetResponse();
            var stream = res.GetResponseStream();
            MemoryStream memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            var bs = memoryStream.ToArray();
            memoryStream.Close();
            stream.Close();
            res.Close();
            return bs;
        }
        static byte[] GetImgFromFolder(string path)
        {
            var exts = new string[] {".gif",".jpg",".png",".tif",".bmp" };

            var res = new byte[] { };
            if (File.Exists(path) && exts.Contains(new FileInfo(path).Extension))
            {
                res = File.ReadAllBytes(path);
            }

            return res;
        }
    }
}
