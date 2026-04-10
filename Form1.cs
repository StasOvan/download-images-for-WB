using System.Diagnostics;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;


namespace Выкачиваем_картинки_для_WB__700х900____3
{
    public partial class Form1 : Form
    {
        // после формирования rrr.csv преобразовать в 1251 (Windows) для Excel

        const string FOLDER_FOR_DOWNLOAD_IMAGES = "D:\\Temp\\images\\"; // В Temp должна быть папка images
        const string FILE_EXCEL = "артикулы+фото.xlsx"; // Путь к файлу Excel


        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e) // скачивание файлов в  FOLDER_FOR_DOWNLOAD_IMAGES
        {
            List<string> articles = [];
            List<string> imagesUrl = [];

            // Считываем данные из Ехель
            using (FileStream file = new FileStream(Application.StartupPath + "\\" + FILE_EXCEL, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);

                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    IRow currentRow = sheet.GetRow(row);
                    if (currentRow.GetCell(0) == null || currentRow.GetCell(1) == null) continue;

                    articles.Add(currentRow.GetCell(0).CellType == CellType.Numeric ?
                        currentRow.GetCell(0).NumericCellValue.ToString() : currentRow.GetCell(0).StringCellValue);
                    imagesUrl.Add(currentRow.GetCell(1).StringCellValue);
                }

            }

            // Выкачиваем все картинки по refers
            HttpClient client = new HttpClient();
            string[] refers;
            string article;
            string temp;
            string fileName;

            for (int i = 0; i < articles.Count; i++)
            {
                // в article заменяются /\*? на _

                article = articles[i].Replace("\\", "_").Replace("/", "_").Replace("*", "_").Replace("?", "_");
                refers = imagesUrl[i].Split(' ');
                foreach (string refer in refers)
                {
                    temp = refer.Replace("https://static.insales-cdn.com/images/", "");
                    temp = temp.Replace("\\", "_").Replace("/", "_").Replace("*", "_").Replace("?", "_").Replace("!", "_");
                    fileName = FOLDER_FOR_DOWNLOAD_IMAGES + article + "---" + temp;
                    if (article == "Артикул" || article == "") continue;

                    if (Path.GetExtension(fileName) == ".dll") fileName += ".jpg";
                    if (Path.GetExtension(fileName) == ".ashx") fileName += ".png";
                    if (Path.GetExtension(fileName) == "") fileName += ".png";

                    fileName = fileName.Replace("#", "%23");

                    if (File.Exists(fileName))
                        textBox2.Text = i.ToString();
                    else
                    {
                        try
                        {
                            byte[] data;
                            HttpResponseMessage response = await client.GetAsync(refer);
                            using (HttpContent content = response.Content)
                            {
                                data = await content.ReadAsByteArrayAsync();
                                using (FileStream file = File.Create(fileName))
                                    file.Write(data, 0, data.Length);
                            }
                            textBox1.Text = $"{fileName}";
                            textBox2.Text = i.ToString() + " cool";
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine("bad! - " + i.ToString() + " --- " + article);
                            Debug.WriteLine(ex.Message);
                        }
                    }

                    Application.DoEvents();

                } // end for refers
            } // end for articles
        }



        private void button2_Click(object sender, EventArgs e) // обработка изображений (новое фото будет X_product)
        {
            // Получаем список файлов из указанной папки
            string[] imageFiles = System.IO.Directory.GetFiles(FOLDER_FOR_DOWNLOAD_IMAGES, "*.*");

            // Цикл для обработки каждого изображения
            int width = 700;
            int height = 900;
            string originalFilePath;
            bool FLAG;

            for (var q = 0; q < imageFiles.Length; q++)
            {
                string imagePath = imageFiles[q];

                originalFilePath = imagePath;
                FLAG = false;
                textBox1.Text = originalFilePath;
                textBox2.Text = (q + 1).ToString();

                if (Path.GetExtension(imagePath) == ".webp") continue;

                GC.Collect();

                try
                {
                    Bitmap originalImage = new Bitmap(imagePath);

                    if (originalImage.Width >= originalImage.Height)
                    {
                        width = originalImage.Width;
                        height = width / originalImage.Height * originalImage.Height;
                    }
                    if (originalImage.Width <= originalImage.Height)
                    {
                        height = originalImage.Height;
                        width = height / originalImage.Width * originalImage.Width;
                    }

                    if (originalImage.Width <= 700 && originalImage.Height <= 900)
                    {
                        width = 700;
                        height = 900;
                    }

                    if (width < 700) width = 700;
                    if (height < 900) height = 900;


                    if (originalImage.Width < 700 || originalImage.Height < 900)
                    {

                        using (Bitmap newImage = new Bitmap(width, height))
                        {
                            using (Graphics g = Graphics.FromImage(newImage))
                            {
                                g.Clear(Color.White); // Заливаем фон белым цветом

                                // Тут считаем  координаты для размещения исходного изображения по центру
                                int x = (width - originalImage.Width) / 2;
                                int y = (height - originalImage.Height) / 2;

                                // Налагаем исходное изображение на новый белый фон
                                g.DrawImage(originalImage, x, y, originalImage.Width, originalImage.Height);
                            }

                            // Сохраняем новое изображение в той же папке с указанием нового имени файла Х_
                            string newImagePath = imagePath.Replace("product", "X_product");
                            newImage.Save(newImagePath); // Сохраняем новое изображение
                            FLAG = true; // Выполнили =))
                            newImage.Dispose();
                        }

                    }

                    originalImage.Dispose();

                    if (FLAG) File.Delete(originalFilePath);

                    Application.DoEvents();
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    continue;

                }
            } // end for ImageFiles
        }

        private void button3_Click(object sender, EventArgs e) // формирование rrr.csv
        {
            File.Delete("rrr.csv");

            // Получаем список файлов из указанной папки
            string[] imageFiles = System.IO.Directory.GetFiles(FOLDER_FOR_DOWNLOAD_IMAGES, "*.*");

            var q = 0;

            for (var i = 0; i < imageFiles.Length; i++)
            {
                var imageFile = imageFiles[i];
                var article = imageFile.Split("---")[0];
                string pathFileName = "";

                if (article != "")
                {
                    for (var c = i; c < imageFiles.Length; c++)
                    {
                        if (article == imageFiles[c].Split("---")[0])
                        {
                            pathFileName += "https://myqu.ru/images/" + imageFiles[c] + " ";
                            imageFiles[c] = "---";
                        }
                    }

                    string row = ("\"" + article.Replace("%23", "#") + "\"" + ";" + "\"" + pathFileName + "\"" + "\r\n").Replace(FOLDER_FOR_DOWNLOAD_IMAGES, "");
                    File.AppendAllText("rrr.csv", row);
                    q++;
                }

                textBox2.Text = (q + 1).ToString();
                Application.DoEvents();

            } // end for ImageFiles
        }
    }
}



