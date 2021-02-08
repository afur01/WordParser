using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
//using Aspose.Words;
using DocumentFormat.OpenXml.Packaging;
//using GemBox.Document;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
//deneme

namespace WordParser.Controllers
{
    public class ParserController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public ParserController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Parser(string Message)
        {
            ViewBag.Message = Message;
            return View();
        }
        [HttpPost]
        public IActionResult Parser(IFormFile File, string DocumentType)
        {
            string Message = "";
            string[] Words;
            if (File != null)
            {
                if (DocumentType == "Word")
                {

                    //var fileName = System.IO.Path.GetFileName(File.FileName);
                    //var fileNameWithOutExtension = System.IO.Path.GetFileNameWithoutExtension(File.FileName);
                    //var myUniqueFileName = Convert.ToString(Guid.NewGuid());
                    //var fileExtension = System.IO.Path.GetExtension(fileName);
                    //var newFileName = String.Concat(myUniqueFileName, fileExtension);

                    //string Upload = System.IO.Path.Combine(_hostingEnvironment.WebRootPath, @"WordFile\", newFileName);
                    //File.CopyTo(new FileStream(Upload, FileMode.Create));

                    string[] RomeNumber = new string[2] { "I", "II" };                   
                    #region OpenXml
                    string ResultText = "";
                    using (var ms = new MemoryStream())
                    {                      
                        File.CopyTo(ms);
                        var fileBytes = ms.ToArray();
                        string s = Convert.ToBase64String(fileBytes);
                        // act on the Base64 data
                        using (WordDocument document = new WordDocument(ms, ""))
                        {

                            ResultText = document.GetText();
                            Words = ResultText.Split("\n");

                        }
                    }
                    string letter5;
                    letter5 = "Giriş bölümü bulunamadı!  ya da başlık yanlış yazılmıştır başlık formatı ='GİRİŞ'";
                    string letter;                                                                      //hataları daha sonra view kısmında görüntülemek için değişkenler oluşturuyoruz
                    string letter2;
                    string letter3;
                    string letter4;
                    string letter6;
                    letter6 = "PROJE PLANI başlığı bulunamadı! ya da başlık yanlış yazılmıştır başlık formatı ='PROJE PLANI'";
                    string letter7;
                    string letter8;
                    string letter9;
                    string letter10;
                    string letter11;
                    letter11 = "SİSTEM ÇÖZÜMLEME başlığı bulunamadı! ya da başlık yanlış yazılmıştır başlık formatı= 'sistem çözümleme' ";
                    string letter12;
                    string letter13;
                    string letter14;
                    letter14 = "SİSTEM TASARIMI başlığı bulunamadı! ya da başlık yanlış yazılmıştır başlık formatı= 'SİSTEM TASARIMI'  ";
                    string letter15;
                    string letter16;
                    string letter17;
                    letter17 = "GERÇEKLEŞTİRİM başlığı bulunamadı! ya da başlık yanlış yazılmıştır başlık formatı= 'GERÇEKLEŞTİRİM'";
                    string letter18;
                    string letter19;
                    string letter20;
                    foreach (var item in Words)                                                         //burada okuduğumuz değişkenin herbir satırını tek tek dolaşıyoruz
                    {
                        string Controitem = item.Replace("\r", string.Empty);
                        if (Controitem.Trim() == "GİRİŞ".Trim())                                        //koşul ile giriş bölümüne girip girmediğimizi kontrol ediyoruz
                        {
                            letter5 = "  ";
                            int order;                                                                  // bize aradığpımız değerin hangi satırda olacağını verecek bir değişken tanımlıyoruz
                            order = 0;
                            int line4 = 0;
                            int line3 = 0;
                            for (int j = 0; j < Words.Length; j++)
                            {
                                if (Words[j] == "GİRİŞ\r")
                                {
                                    order = j;                                                          //// burada giriş bölünün hangi satırda tutulduğunu bulup ona göre işlkem yapacağız
                                    for (int q = 25; q < 60; q++)                                       // giriş bölümümüz eğer ikinci sayfada ise
                                    {
                                        if (order == q)
                                        {
                                            letter5 = "  ";                                              // hata almıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            letter5 = "GİRİŞ başlığı 2. sayfadan başlamalıdır";            // fakat değilse hata alıyoruz
                                        }
                                    }
                                    for (int kk = 60; kk < Words.Length; kk++)                              //buraradada aynı şekilde giriş bölümü başka bir sayfadada tanımlamışmı onu kontrol ediyoruz
                                    {
                                        if (order == kk)
                                        {
                                            letter5 = "GİRİŞ başlığı 2. sayfada değildir";
                                            break;
                                        }

                                    }
                                    break;

                                }


                            }
                            letter2 = "GİRİŞ bölümünde  ' 1.1 Projenin Amacı ' ve ' 1.2 Projenin Kapsamı ' başlığı eklenmemiş, bu başlıklar zorunludur. ";
                            letter3 = "  ";
                            letter4 = "  ";
                            for (int i = order; i < order + 31; i++)
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "1.1 Projenin Amacı \r")                                //alt başlıklarında giriş bölümünde olup olmadığını kontrol ediyoruz
                                {


                                    asrs = i;
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "1.1 Projenin Amacı \r")                        //eğer alt başlığımız giriş bölümünde değil ise 
                                        {
                                            line3 = t;                                                                                                 //hatanım bulunduğu satırı bu değikene atıyoruz
                                            letter3 = " ' 1.1 Projenin Amacı ' başlığı sadece 1 kez ve GİRİŞ bölümünde kullanılmalıdır";           //ve hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)                       //eğer alt başlığımız giriş bölümünde değil ise
                                    {
                                        if (Words[k] == "1.1 Projenin Amacı \r")
                                        {
                                            line3 = k;                                                                                         //hatanım bulunduğu satırı bu değikene atıyoruz
                                            letter3 = " ' 1.1 Projenin Amacı ' başlığı sadece 1 kez ve GİRİŞ bölümünde kullanılmalıdır";    //ve hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                        }
                                    }

                                    for (int s = order; s < order + 31; s++)
                                    {

                                        if (Words[s] == "1.2 Projenin Kapsamı\r")   // burada ise alt başlığımızın giriş kısmına eklenip eklenmediğini kontrol ediyoruz
                                        {
                                            letter2 = "  ";
                                            break;
                                        }
                                        letter2 = "GİRİŞ bölümünde ' 1.2 Projenin Kapsamı ' başlığı eklenmemiş ya da başlık yanlış yazılmıştır başlık formatı ='1.2 Projenin Kapsamı '";            // hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                    }



                                }
                                int arsrs;
                                arsrs = 0;

                                if (Words[i] == "1.2 Projenin Kapsamı\r")
                                {

                                    arsrs = i;                                              //alt başlığımızın bulunduğu satırı buluyoruz tabi giriş bölümünde ise
                                    for (int t = 0; t < arsrs; t++)                         //ilk satırdan alt başlığın olduğu satıra kadar kontrol ediyoruz
                                    {
                                        if (Words[t] == "1.2 Projenin Kapsamı\r")           //eğer bi altbaşlık daha var ise hata alıyoruz
                                        {
                                            line4 = t;
                                            letter4 = " ' 1.2 Projenin Kapsamı ' başlığı sadece 1 kez ve GİRİŞ bölümünde kullanılmalıdır";      // hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                        }
                                    }

                                    for (int k = arsrs + 1; k < Words.Length; k++)          //alt başlığın olduğu satırdan son satıra  kadar kontrol ediyoruz
                                    {
                                        if (Words[k] == "1.2 Projenin Kapsamı\r")           //eğer bi altbaşlık daha var ise hata alıyoruz     
                                        {
                                            line4 = k;
                                            letter4 = " ' 1.2 Projenin Kapsamı ' başlığı sadece 1 kez ve GİRİŞ bölümünde kullanılmalıdır";              // hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                        }
                                    }
                                    for (int ş = arsrs; ş < order + 31; ş++)
                                    {
                                        if (Words[ş] == "1.1 Projenin Amacı \r")                // burada ise altbaşlıklarımızn sırası doğrumu diye kontrol ediyoruz
                                        {
                                            letter2 = "Giriş bölümünde ' 1.1 Projenin Amacı ' kısmı  ' 1.2 Projenin Kapsamı ' kısmından önce olmalıdır. ";              // hata ile birlikte view kısmına göndemek üzere letter değişkenimize alıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            for (int ş2 = order; ş2 < order + 31; ş2++)
                                            {
                                                if (Words[ş2] == "1.1 Projenin Amacı \r")
                                                {
                                                    letter2 = "  ";
                                                    break;
                                                }
                                                letter2 = "Giriş bölümünde ' 1.1 Projenin Amacı ' kısmı  ' Eklenmemiş  ya da başlık yanlış yazılmıştır başlık formatı ='1.1 Projenin Amacı '";

                                            }
                                        }

                                    }

                                }
                            }
                            if (letter2 != "  ")                        // tüm bu koşullara girilmemiş ise 
                            {
                                for (int t = 0; t < order; t++)         // başlıklardan eğer biri bile giriş bölümüne yazılmamış ise
                                {
                                    if (Words[t] == "1.1 Projenin Amacı \r" || Words[t] == "1.2 Projenin Kapsamı")
                                    {
                                        letter2 = "1.1 Projenin Amacı veya 1.2 Projenin Kapsamı başlığı GİRİŞ bölümünde bulunmak zorundadır.";      // bu hatayı alıyoruz
                                    }
                                }
                                for (int k = order + 32; k < Words.Length; k++)
                                {
                                    if (Words[k] == "1.1 Projenin Amacı \r" || Words[k] == "1.2 Projenin Kapsamı")
                                    {
                                        letter2 = "1.1 Projenin Amacı veya 1.2 Projenin Kapsamı başlığı GİRİŞ bölümünde bulunmak zorundadır.";
                                    }
                                }
                            }
                            ViewData["hata2"] = letter2; ViewData["order"] = order;               //burada ise hatalarımızı view kısmına yollamış olduk
                            ViewData["hata3"] = letter3; ViewData["line3"] = line3;
                            ViewData["hata4"] = letter4; ViewData["line4"] = line4;


                        }
                        ViewData["hata5"] = letter5;

                        string Controitem2 = item.Replace("\r", string.Empty);
                        if (Controitem2.Trim() == Words[1].Trim())                      //İlk satırı kontrol ediyoruz
                        {
                            int line = 1;

                            if (Words[1] != "\r")                                       // eğer boş değil ise
                            {
                                letter = "İlk satır daima boş olmalıdır.";              // hata alıyoruz

                            }
                            else
                            {
                                letter = "  ";                                           // boş ise hata almıyoruz

                            }
                            ViewData["hata1"] = letter;                                 // ve hatayı boş yada dolu view kısmına gönderiyoruz
                            ViewData["a"] = line;
                        }

                        string Controitem3 = item.Replace("\r", string.Empty);
                        if (Controitem3.Trim() == "PROJE PLANI".Trim())                 // 2. başlığımız olan proje planı başlığını kontrol ediyoruz
                        {
                            letter6 = "  ";
                            int order2;
                            order2 = 0;
                            for (int j = 0; j < Words.Length; j++)                      // tüm satırları tek tek dolaşarak
                            {
                                if (Words[j] == "PROJE PLANI \r")                       // proje planının bulunduğu satırı buluyoruz
                                {
                                    order2 = j;
                                    for (int q = 30; q < 66; q++)                       // eğer 3. sayfada ise
                                    {
                                        if (order2 == q)
                                        {
                                            letter6 = "  ";                             // hata almıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            letter6 = "PROJE PLANI başlığı 3. sayfadan başlamalıdır";    // değil ise hata alıyoruz 
                                        }
                                    }
                                    for (int kk = 66; kk < Words.Length; kk++)
                                    {
                                        if (order2 == kk)
                                        {
                                            letter6 = "PROJE PLANI başlığı 3. sayfadan başlamalıdır";
                                            break;
                                        }

                                    }
                                    break;

                                }

                            }
                            int line7 = 0;
                            letter7 = "2.1 Proje Plan Kapsamı başlığı bulunamadı";
                            for (int i = order2; i < order2 + 90; i++)                              //proje planuı bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "2.1 Proje Plan Kapsamı\r")
                                {
                                    letter7 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "2.1 Proje Plan Kapsamı\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line7 = t;
                                            letter7 = " ' 2.1 Proje Planı Kapsamı ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "2.1 Proje Plan Kapsamı\r")
                                        {
                                            line7 = k;
                                            letter7 = " ' 2.1 Proje Planı Kapsamı ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line7 = order2;
                                    letter7 = "'2.1 Proje Planı Kapsamı ' başlığı  PROJE PLANI bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='2.1 Proje Plan Kapsamı' ";

                                }
                            }
                            ViewData["order2"] = order2;
                            ViewData["hata7"] = letter7; ViewData["line7"] = line7;

                            int line8 = 0;
                            letter8 = "2.2 Proje Ekibi başlığı bulunamadı";

                            for (int i = order2; i < order2 + 90; i++)                              //proje ekibi bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "2.2 Proje Ekibi\r")
                                {
                                    letter8 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == " 2.2 Proje Ekibi\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line8 = t;
                                            letter8 = " ' 2.2 Proje Ekibi ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "2.2 Proje Ekibi\r")
                                        {
                                            line8 = k;
                                            letter8 = " '2.2 Proje Ekibi ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line8 = order2;
                                    letter8 = "'2.2 Proje Ekibi ' başlığı  PROJE PLANI bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='2.2 Proje Ekibi' ";

                                }
                            }
                            ViewData["hata8"] = letter8; ViewData["line8"] = line8;

                            int line9 = 0;
                            letter9 = "2.2 Yöntem ve Metodolojiler başlığı bulunamadı";

                            for (int i = order2; i < order2 + 90; i++)                              //   Yöntem ve Metodolojiler bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "2.3 Yöntem ve Metodolojiler\r")
                                {
                                    letter9 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "2.3 Yöntem ve Metodolojiler\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line9 = t;
                                            letter9 = " ' 2.3 Yöntem ve Metodolojiler ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "2.3 Yöntem ve Metodolojiler\r")
                                        {
                                            line9 = k;
                                            letter9 = " '2.3 Yöntem ve Metodolojiler ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line9 = order2;
                                    letter9 = "'2.3 Yöntem ve Metodolojiler ' başlığı  PROJE PLANI bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='2.3 Yöntem ve Metodolojiler' ";

                                }
                            }
                            ViewData["hata9"] = letter9; ViewData["line9"] = line9;

                            int line10 = 0;
                            letter10 = "2.4 Planlar başlığı bulunamadı";

                            for (int i = order2; i < order2 + 90; i++)                              //  2.4 Planlar bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "2.4 Planlar\r")
                                {
                                    letter10 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "2.4 Planlar\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line10 = t;
                                            letter10 = " ' 2.4 Planlar ' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "2.4 Planlar\r")
                                        {
                                            line10 = k;
                                            letter10 = " '2.4 Planlar' başlığı sadece 1 kez ve PROJE PLANI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line10 = order2;
                                    letter10 = "'2.4 Planlar ' başlığı  PROJE PLANI bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='2.4 Planlar' ";

                                }
                            }
                            ViewData["hata10"] = letter10; ViewData["line10"] = line10;

                        }
                        ViewData["hata6"] = letter6;    

                        string Controitem4 = item.Replace("\r", string.Empty);
                        if (Controitem4.Trim() == "SİSTEM ÇÖZÜMLEME".Trim())                  // 3. başlığımız olanSİSTEM ÇÖZÜMLEME başlığını kontrol ediyoruz
                        {
                            letter11 = "  ";
                            int order3;
                            order3 = 0;
                            for (int j = 0; j < Words.Length; j++)                      // tüm satırları tek tek dolaşarak
                            {
                                if (Words[j] == "\t\t\t\tSİSTEM ÇÖZÜMLEME\r")                       // SİSTEM ÇÖZÜMLEME bulunduğu satırı buluyoruz
                                {
                                    order3 = j;
                                    for (int q = 115; q < 200; q++)                       // eğer 6. sayfada ise
                                    {
                                        if (order3 == q)
                                        {
                                            letter11 = "  ";                             // hata almıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            letter11 = "SİSTEM ÇÖZÜMLEME başlığı 6. sayfadan başlamalıdır";    // değil ise hata alıyoruz 
                                        }
                                    }
                                    for (int kk = 200; kk < Words.Length; kk++)
                                    {
                                        if (order3 == kk)
                                        {
                                            letter11 = "SİSTEM ÇÖZÜMLEME başlığı 6. sayfadan başlamalıdır";
                                            break;
                                        }

                                    }
                                    break;

                                }

                            }
                            int line12 = 0;
                            letter12 = "Mevcut Sİstem İncelemesi başlığı bulunamadı";
                            int number1 = 0;
                            for (int i = order3; i < order3 + 90; i++)                              //  Mevcut Sİstem İncelemesi bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "3.1 Mevcut Sİstem İncelemesi\r")
                                {
                                    letter12 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "3.1 Mevcut Sİstem İncelemesi\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line12 = t;
                                            letter12 = " '3.1 Mevcut Sİstem İncelemesi ' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "3.1 Mevcut Sİstem İncelemesi\r")
                                        {
                                            line12 = k;
                                            letter12 = " '3.1 Mevcut Sİstem İncelemesi' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    number1 = order3 + 90;
                                    line12 = order3;
                                    letter12 = "'3.1 Mevcut Sİstem İncelemesi ' başlığı  SİSTEM ÇÖZÜMLEME bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='3.1 Mevcut Sİstem İncelemesi' ";

                                }
                            }
                            ViewData["hata12"] = letter12; ViewData["line12"] = line12; ViewData["number1"] = number1;

                            int line13 = 0;
                            letter13 = "3.2 Arayüz başlığı bulunamadı";

                            for (int i = order3; i < order3 + 90; i++)                              //  3.2 Arayüz bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "3.2 Arayüz\r")
                                {
                                    letter13 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "3.2 Arayüz\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line13 = t;
                                            letter13 = " '3.2 Arayüz' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "3.2 Arayüz\r")
                                        {
                                            line13 = k;
                                            letter13 = " '3.2 Arayüz' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line13 = order3;
                                    letter13 = "'3.2 Arayüz' başlığı  SİSTEM ÇÖZÜMLEME bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='3.2 Arayüz' ";

                                }
                            }
                            ViewData["hata13"] = letter13; ViewData["line13"] = line13;

                        }
                        ViewData["hata11"] = letter11;

                        string Controitem5 = item.Replace("\r", string.Empty);
                        if (Controitem5.Trim() == "SİSTEM TASARIMI".Trim())                  // 3. başlığımız olan SİSTEM TASARIMI başlığını kontrol ediyoruz
                        {
                            letter14 = "  ";
                            int order4;
                            order4 = 0;
                            for (int j = 0; j < Words.Length; j++)                      // tüm satırları tek tek dolaşarak
                            {
                                if (Words[j] == "\t\t\t\tSİSTEM TASARIMI\r")                       // SİSTEM TASARIMI bulunduğu satırı buluyoruz
                                {
                                    order4 = j;
                                    for (int q = 150; q < 250; q++)                       // eğer 6. sayfada ise
                                    {
                                        if (order4 == q)
                                        {
                                            letter14 = "  ";                             // hata almıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            letter14 = "SİSTEM TASARIMI başlığı 10. sayfadan başlamalıdır";    // değil ise hata alıyoruz 
                                        }
                                    }
                                    for (int kk = 250; kk < Words.Length; kk++)
                                    {
                                        if (order4 == kk)
                                        {
                                            letter14 = "SİSTEM TASARIMI başlığı 10. sayfadan başlamalıdır";
                                            break;
                                        }

                                    }
                                    break;

                                }

                            }
                            
                            int line15 = 0;
                            letter15 = "4.1 Genel Tasarım İncelemesi başlığı bulunamadı";
                            
                            for (int i = order4; i < order4 + 90; i++)                              //  Sistem tasarımı bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "4.1 Genel Tasarım Bilgileri\r")
                                {
                                    letter15 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "4.1 Genel Tasarım Bilgileri\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line15 = t;
                                            letter15 = " '4.1 Genel Tasarım ' başlığı sadece 1 kez ve SİSTEM TASARIMI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "4.1 Genel Tasarım Bilgileri\r")
                                        {
                                            line15 = k;
                                            letter15 = " '4.1 Genel Tasarım' başlığı sadece 1 kez ve SİSTEM TASARIMI bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    
                                    line15 = order4;
                                    letter15 = "'4.1 Genel Tasarım' başlığı  SİSTEM TASARIMI bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='4.1 Genel Tasarım' ";

                                }
                            }
                            ViewData["hata15"] = letter15; ViewData["line15"] = line15; 

                            int line16 = 0;
                            letter16 = "4.2 ORTAK ALT SİSTEM TASARIMI başlığı bulunamadı";

                            for (int i = order4; i < order4 + 90; i++)                              //  ORTAK ALT SİSTEM TASARIMI bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "4.2 ORTAK ALT SİSTEM TASARIMI \r")
                                {
                                    letter16 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "4.2 ORTAK ALT SİSTEM TASARIMI \r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line16 = t;
                                            letter16 = " '4.2 ORTAK ALT SİSTEM TASARIMI' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "4.2 ORTAK ALT SİSTEM TASARIMI \r")
                                        {
                                            line16 = k;
                                            letter16 = " '4.2 ORTAK ALT SİSTEM TASARIMI' başlığı sadece 1 kez ve SİSTEM ÇÖZÜMLEME bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {
                                    line16 = order4;
                                    letter16 = "'4.2 ORTAK ALT SİSTEM TASARIMI' başlığı  SİSTEM ÇÖZÜMLEME bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='4.2 ORTAK ALT SİSTEM TASARIMI' ";

                                }
                            }
                            ViewData["hata16"] = letter16; ViewData["line16"] = line16;

                        }
                        ViewData["hata14"] = letter14;

                        string Controitem6 = item.Replace("\r", string.Empty);
                        if (Controitem6.Trim() == "GERÇEKLEŞTİRİM".Trim())                  // 5. başlığımız olan GERÇEKLEŞTİRİM başlığını kontrol ediyoruz
                        {
                            letter17 = "  ";
                            int order5;
                            order5 = 0;
                            for (int j = 0; j < Words.Length; j++)                      // tüm satırları tek tek dolaşarak
                            {
                                if (Words[j] == "GERÇEKLEŞTİRİM\r")                       // GERÇEKLEŞTİRİM bulunduğu satırı buluyoruz
                                {
                                    order5 = j;
                                    for (int q = 200; q < 300; q++)                       // eğer 6. sayfada ise
                                    {
                                        if (order5 == q)
                                        {
                                            letter17 = "  ";                             // hata almıyoruz
                                            break;
                                        }
                                        else
                                        {
                                            letter17 = "GERÇEKLEŞTİRİM başlığı 12. sayfadan başlamalıdır";    // değil ise hata alıyoruz 
                                        }
                                    }
                                    for (int kk = 300; kk < Words.Length; kk++)
                                    {
                                        if (order5 == kk)
                                        {
                                            letter17 = "GERÇEKLEŞTİRİM başlığı 12. sayfadan başlamalıdır";
                                            break;
                                        }

                                    }
                                    break;

                                }

                            }

                            int line18 = 0;
                            letter18 = "5.1 GİRİŞ başlığı bulunamadı";

                            for (int i = order5; i < order5 + 90; i++)                              //  GERÇEKLEŞTİRİM bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "5.1 GİRİŞ\r")
                                {
                                    letter18 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "5.1 GİRİŞ\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line18 = t;
                                            letter18 = " '5.1 GİRİŞ ' başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] =="5.1 GİRİŞ\r")
                                        {
                                            line18 = k;
                                            letter18 = " '5.1 GİRİŞ başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {

                                    line18 = order5;
                                    letter18 = "'5.1 GİRİŞ  GERÇEKLEŞTİRİM bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='5.1 GİRİŞ' ";

                                }
                            }
                            ViewData["hata18"] = letter18; ViewData["line18"] = line18;

                            int line19 = 0;
                            letter19 = "5.2 YAZILIM GELİŞTİRME ORTAMLARI başlığı bulunamadı";

                            for (int i = order5; i < order5 + 90; i++)                              //  GERÇEKLEŞTİRİM bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "5.2 YAZILIM GELİŞTİRME ORTAMLARI\r")
                                {
                                    letter19 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "5.2 YAZILIM GELİŞTİRME ORTAMLARI\r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line19 = t;
                                            letter19 = " '5.2 YAZILIM GELİŞTİRME ORTAMLARI' başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "5.2 YAZILIM GELİŞTİRME ORTAMLARI\r")
                                        {
                                            line19 = k;
                                            letter19 = " '5.2 YAZILIM GELİŞTİRME ORTAMLARI' başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {

                                    line19 = order5;
                                    letter19 = "'5.2 YAZILIM GELİŞTİRME ORTAMLARI' GERÇEKLEŞTİRİM bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='5.2 YAZILIM GELİŞTİRME ORTAMLARI' ";

                                }
                            }
                            ViewData["hata19"] = letter19; ViewData["line19"] = line19;

                            int line20 = 0;
                            letter20 = "5.3 Olağan Dışı Durum Çözümleme başlığı bulunamadı";

                            for (int i = order5; i < order5 + 90; i++)                              //  GERÇEKLEŞTİRİM bölümünde arama yapıyoruz 
                            {
                                int asrs;
                                asrs = 0;
                                if (Words[i] == "5.3 Olağan Dışı Durum Çözümleme \r")
                                {
                                    letter20 = "  ";

                                    asrs = i;                                           //eğer alt başlığımız doğru bölümde ise onun satırını bu değişkene atıyoruz
                                    for (int t = 0; t < asrs; t++)
                                    {
                                        if (Words[t] == "5.3 Olağan Dışı Durum Çözümleme \r")             // bu alt başlık eğer 1 den fazla yazılmış ise hata alıyoruz
                                        {
                                            line20 = t;
                                            letter20 = " '5.3 Olağan Dışı Durum Çözümleme' başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }

                                    for (int k = asrs + 1; k < Words.Length; k++)
                                    {
                                        if (Words[k] == "5.3 Olağan Dışı Durum Çözümleme \r")
                                        {
                                            line20 = k;
                                            letter20 = " '5.3 Olağan Dışı Durum Çözümleme' başlığı sadece 1 kez ve GERÇEKLEŞTİRİM bölümünde kullanılmalıdır";
                                            break;
                                        }
                                    }
                                    break;

                                }
                                if (asrs == 0)
                                {

                                    line20 = order5;
                                    letter20 = "'5.3 Olağan Dışı Durum Çözümleme' GERÇEKLEŞTİRİM bölümünde bulunmamaktadır ya da başlık yanlış yazılmıştır başlık formatı ='5.3 Olağan Dışı Durum Çözümleme' ";

                                }
                            }
                            ViewData["hata20"] = letter20; ViewData["line20"] = line20;




                        }
                        ViewData["hata17"] = letter17;
                    }

                    
                    #endregion

                    #region Aspos
                    //Document doc = new Document(@"E:\PROJELER\asp.net proje çalışmalrı\WordParser\WordParser\wwwroot\WordFile\b317db3f-3fbc-48f0-baba-e26eb11803e9.docx");
                    //var builder = new DocumentBuilder(doc);
                    //var mBuilder = new StringBuilder();
                    //var paragraphs = builder.Document.GetChildNodes(NodeType.Paragraph, true).ToArray().ToList();
                    //paragraphs.ForEach(
                    //    x =>
                    //    {
                    //        ((Aspose.Words.Paragraph)x).Runs.ToArray().ToList().ForEach(y => mBuilder.Append(y.Text));
                    //        mBuilder.Append(Environment.NewLine);
                    //    }
                    //);
                    //string test = mBuilder.ToString();
                    #endregion

                    #region Gembox
                    //// If using Professional version, put your serial key below.
                    //ComponentInfo.SetLicense("FREE-LIMITED-KEY");
                    //// Load Word document from file's path.
                    //var document = DocumentModel.Load(Upload);
                    //// Get Word document's plain text.
                    //string text = document.Content.ToString();
                    //// Get Word document's count statistics.
                    //int charactersCount = text.Replace(Environment.NewLine, string.Empty).Length;
                    //int wordsCount = Regex.Matches(text, @"[\S]+").Count;
                    //int paragraphsCount = document.GetChildElements(true, ElementType.Paragraph).Count();
                    //int pageCount = document.GetPaginator().Pages.Count;

                    //// Display file's count statistics.
                    //ViewData["karakterbulma"] = charactersCount;
                    //ViewData["kelimesayısı"] = wordsCount;
                    //ViewData["paragrafsayısı"] = paragraphsCount;
                    //ViewData["sayfasayısı"] = pageCount;                                    

                    // Display file's text content.
                    //Console.WriteLine(text);
                    #endregion


                    return View();
                }
                else if (DocumentType == "Pdf")
                {

                    using (var fileStram = File.OpenReadStream())
                    {
                        Words = pdfText(fileStram).Split("\n");

                    }
                }

                Message = "Döküman Kontrolü Başarılı";
            }
            else
            {
                Message = "Lütfen Döküman Seçiniz";
            }


            return Parser(Message);
        }



        public static string pdfText(Stream data)
        {
            PdfReader reader = new PdfReader(data);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }
    }
}
