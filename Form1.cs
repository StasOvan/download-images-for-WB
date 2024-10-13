using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Diagnostics.Eventing.Reader;
using System.Diagnostics;
using NPOI.SS.Formula.Functions;
using System.Security.Policy;
using System.Net;

namespace Выкачиваем_картинки_для_WB__700х900____2
{
    public partial class Form1 : Form
    {

        string[] refers;
        string[] arr;
        string imagesFolder = "D:\\Temp\\images\\";

        public Form1()
        {
            InitializeComponent();
        }

        private  void button1_Click(object sender, EventArgs e) // скачивание файлов в  D:\temp\images
        {

            // В Temp должна быть папка images
            // при формировании rrr.csv преобразовать в 1251 для Excel

            // Путь к файлу Excel
            string filePath = Application.StartupPath + "\\Товары на улучшение фото.xlsx";
            

            string article;
            string imageUrl;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);
                

                for (int row = 0; row <= 9000; row++)
                {
                    IRow currentRow = sheet.GetRow(row);
                    if (currentRow.GetCell(0) == null || currentRow.GetCell(1) == null) continue;
                 
                    article = currentRow.GetCell(0).CellType == CellType.Numeric ? 
                        currentRow.GetCell(0).NumericCellValue.ToString() : currentRow.GetCell(0).StringCellValue;

                    imageUrl = currentRow.GetCell(1).StringCellValue;
                    refers = imageUrl.Split(' ');

                    for (int i = 0; i < refers.Length; i++)
                    {
                        using (WebClient webClient = new WebClient())
                        {
                            string localFileName = refers[i].Replace("https://static.insales-cdn.com/images/", "");
                            localFileName = localFileName.Replace("/", "_");
                            localFileName = localFileName.Replace("\\", "_");
                            string fileName = imagesFolder + article.Replace("/", "_").Replace("\\", "_").Replace("*", "x") + "---" + localFileName;
                            if (Path.GetExtension(fileName) == ".dll") fileName = fileName + ".jpg";
                            if (Path.GetExtension(fileName) == ".ashx") fileName = fileName + ".png";
                            if (File.Exists(fileName))
                            {
                                textBox2.Text = row.ToString();
                            }
                            else
                            {                               
                                try
                                {
                                    webClient.DownloadFile(refers[i], fileName);
                                    textBox1.Text = $"{fileName}";
                                    textBox2.Text = row.ToString() + " cool";
                                }
                                catch (Exception ex) 
                                {
                                    Debug.WriteLine("bad! - " + row.ToString() + " --- " + article);
                                    Debug.WriteLine(ex.Message);
                                }
                            }
                        }

                        Application.DoEvents();

                    } // END for refers

                } // END for rows in Excel
            }
        }

        private void button2_Click(object sender, EventArgs e) // обработка изображений (новое фото будет X_product)
        {
            // Получаем список файлов из указанной папки imagesFolder
            string[] imageFiles = System.IO.Directory.GetFiles(imagesFolder, "*.*");

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
                textBox2.Text = q.ToString();

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

                                // Рассчитываем координаты для размещения исходного изображения по центру
                                int x = (width - originalImage.Width) / 2;
                                int y = (height - originalImage.Height) / 2;

                                // Налагаем исходное изображение на новый белый фон
                                g.DrawImage(originalImage, x, y, originalImage.Width, originalImage.Height);
                            }

                            // Сохраняем новое изображение в той же папке с указанием нового имени файла
                            string newImagePath = imagePath.Replace("product", "X_product");
                            newImage.Save(newImagePath); // Сохраняем новое изображение
                            FLAG = true;
                            newImage.Dispose();
                        }

                    }

                    originalImage.Dispose();

                    if (FLAG) File.Delete(originalFilePath);

                    Application.DoEvents();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message); continue;
                }
            } // END for
        }

        private void button3_Click(object sender, EventArgs e) // формирование rrr.csv
        {
            File.Delete("rrr.csv");

            string[] imageFiles = System.IO.Directory.GetFiles(imagesFolder, "*.*");

            var q = 0;
            
            for (var i = 0; i < imageFiles.Length; i++)
            {
                var imageFile = imageFiles[i];
                var article = imageFile.Split("---")[0];
                string pathFileName = "";

                if (article != "") 
                { 
                    for (var c = i; c < imageFiles.Length; c++)
                        if (article == imageFiles[c].Split("---")[0])
                        {
                            pathFileName += "https://kostyak.site/images/" + imageFiles[c] + " ";
                            imageFiles[c] = "---";
                        }
                        
                    string row = ("\"" + article + "\"" + ";" + "\"" + pathFileName + "\"" + "\r\n").Replace(imagesFolder,"");
                    File.AppendAllText("rrr.csv", row);
                    q++;
                }

                textBox2.Text = q.ToString();
                Application.DoEvents();

            } // END for
        }
    }
}



