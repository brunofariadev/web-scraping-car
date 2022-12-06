using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Net.Mail;
using System.Globalization;

namespace WebScrapingCar
{
    public class Car
    {
        public string Titulo { get; set; }

        public string Descricao { get; set; }

        public string Valor { get; set; }

        public string DiaPublicacao { get; set; }

        public string HoraPublicacao { get; set; }

        public string Local { get; set; }

        public string Link { get; set; }

        public DateTime ObtenhaDataFormatada()
        {
            var mes = 0;
            var dia = 0;
            if (DiaPublicacao.Trim() == "Hoje")
            {
                mes = DateTime.Now.Month;
                dia = DateTime.Now.Day;
            }
            else if (DiaPublicacao.Trim() == "Ontem")
            {
                mes = DateTime.Now.AddDays(-1).Month;
                dia = DateTime.Now.AddDays(-1).Day;
            }
            else
            {
                var splitDiaMes = DiaPublicacao.Split(" ");
                mes = ObtenhaMes(splitDiaMes[1].Trim());
                dia = int.Parse(splitDiaMes[0].Trim());
            }
            return new DateTime(DateTime.Now.Year, mes, dia);
        }

        private int ObtenhaMes(string mes)
        {
            switch (mes)
            {
                case "jan":
                    return 1;
                case "fev":
                    return 2;
                case "mar":
                    return 3;
                case "abr":
                    return 4;
                case "mai":
                    return 5;
                case "jun":
                    return 6;
                case "jul":
                    return 7;
                case "ago":
                    return 8;
                case "set":
                    return 9;
                case "out":
                    return 10;
                case "nov":
                    return 11;
                case "dez":
                    return 12;
                default:
                    return 0;
            }
        }
    }

    public class EmailRequest
    {
        public string From { get; set; } = "brunogyn1@gmail.com";
        public string FromName { get; set; } = "Web Scraping - Corolla XEI";
        public string SmtpUsername { get; set; } = "brunogyn1@gmail.com";
        public string SmtpPassword { get; set; } = "psycho3gyn";
        public string Host { get; set; } = "smtp.gmail.com";
        public int Port { get; set; } = 587;
        public MemoryStream File { get; set; }
        public string FileName { get; set; } = "cars.xlsx";
        public IEnumerable<string> Emails { get; set; } = new List<string> { "brunogyn1@gmail.com" };
        public string Subject { get; set; } = "Relatório Corolla XEI";
        public string Body { get; set; } = "Relatório Corolla XEI de hoje";
    }

    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        //string url = "https://www.webmotors.com.br/carros/go/volkswagen/virtus";

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            //while (!stoppingToken.IsCancellationRequested)
            //{
            //    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
            //    try
            //    {
            //        var cars = ObtenhaCars();
            //        var arquivoStream = CriarArquivoExcel(cars.OrderByDescending(c => c.ObtenhaDataFormatada()).ToList());
            //        var emailRequest = new EmailRequest();
            //        emailRequest.File = arquivoStream;
            //        EnviarEmail(emailRequest);
            //        await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            //    }
            //    catch (Exception ex)
            //    {
            //        _logger.LogError("Erro geral do webscraping: {mensagemError}", ex.Message);
            //    }
            //}
        }

        private List<Car> ObtenhaCars()
        {
            List<Car> cars = new List<Car>();
            string lastPage = GetLastPageOfList();
            for (int i = 0; i < int.Parse(lastPage); i++)
            {
                var htmlDocument = GetHtmlDocument(i + 1);
                foreach (var node in htmlDocument.GetElementbyId("ad-list").ChildNodes)
                {
                    if (node.Attributes.Count <= 0)
                    {
                        continue;
                    }
                    try
                    {
                        var nodeHref = node.Descendants();
                        string link = nodeHref.FirstOrDefault()?.Attributes["href"]?.Value;
                        if (string.IsNullOrWhiteSpace(link))
                        {
                            continue;
                        }
                        var valor = WebUtility.HtmlDecode(nodeHref.FirstOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Equals("sc-ifAKCX eoKYee"))?.InnerText);
                        var numberValor = string.IsNullOrWhiteSpace(valor) ? 0 : Convert.ToDouble(valor.Split("R$").Last().Trim(), CultureInfo.GetCultureInfo("pt-BR"));
                        if (numberValor < 74000 || numberValor > 90000)
                        {
                            continue;
                        }
                        var car = CrieCar(nodeHref, link, valor);
                        cars.Add(car);
                        _logger.LogInformation("Carro '{titulo}' importado", car.Titulo);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Erro ao percorrer elementos do html: {mensagemError}", ex.Message);
                    }
                }
            }

            return cars;
        }

        private static Car CrieCar(IEnumerable<HtmlAgilityPack.HtmlNode> nodeHref, string link, string valor)
        {
            var car = new Car();
            car.Valor = valor;
            car.Link = link;
            car.Titulo = WebUtility.HtmlDecode(nodeHref.FirstOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Contains("sc-1mbetcw-0 fKteoJ sc-ifAKCX"))?.InnerText);
            car.Descricao = WebUtility.HtmlDecode(nodeHref.FirstOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Equals("sc-1j5op1p-0 lnqdIU sc-ifAKCX eLPYJb"))?.InnerText);
            car.DiaPublicacao = WebUtility.HtmlDecode(nodeHref.FirstOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Equals("wlwg1t-1 fsgKJO sc-ifAKCX eLPYJb"))?.InnerText);
            car.HoraPublicacao = WebUtility.HtmlDecode(nodeHref.LastOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Equals("wlwg1t-1 fsgKJO sc-ifAKCX eLPYJb"))?.InnerText);
            car.Local = WebUtility.HtmlDecode(nodeHref.FirstOrDefault(c => c.Attributes["class"] != null && c.Attributes["class"].Value.Equals("sc-7l84qu-1 ciykCV sc-ifAKCX dpURtf"))?.InnerText);
            return car;
        }

        private string GetLastPageOfList()
        {
            var htmlDocumentPage = GetHtmlDocument(1);
            var divParentPage = htmlDocumentPage.GetElementbyId("listing-main-content-slot").ChildNodes.FirstOrDefault(n => n.Attributes["class"] != null && n.Attributes["class"].Value.Contains("h3us20-6 eCFxPX"));
            var hrefLastPage = divParentPage.Descendants().FirstOrDefault(t => t.Attributes["data-lurker-detail"] != null && t.Attributes["data-lurker-detail"].Value.Equals("last_page"))?.Attributes["href"]?.Value;
            var lastPage = hrefLastPage.Substring(0, hrefLastPage.IndexOf("&")).Split("?o=").Last();
            return lastPage;
        }

        private HtmlAgilityPack.HtmlDocument GetHtmlDocument(int page)
        {
            string url = page == 1 ? "https://go.olx.com.br/?q=corolla%20xei" : $"https://go.olx.com.br/?o={page}&q=corolla%20xei";
            string siteContent = WebClient(url);
            var htmlDocument = new HtmlAgilityPack.HtmlDocument();
            htmlDocument.LoadHtml(siteContent);
            return htmlDocument;
        }

        private string WebClient(string url)
        {
            var wc = new WebClient();
            wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36");
            var siteContent = wc.DownloadString(url);
            return siteContent;
        }

        private MemoryStream CriarArquivoExcel(List<Car> cars)
        {
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Cells.LoadFromCollection(cars, true);
                package.Save();
            }
            stream.Position = 0;
            return stream;
        }

        public bool EnviarEmail(EmailRequest emailRequest)
        {
            var message = new MailMessage { IsBodyHtml = true, From = new MailAddress(emailRequest.From, emailRequest.FromName) };

            foreach (var email in emailRequest.Emails)
            {
                message.To.Add(new MailAddress(email));
            }

            message.Subject = emailRequest.Subject;
            message.SubjectEncoding = System.Text.Encoding.Unicode;
            message.Body = emailRequest.Body;

            message.Attachments.Add(new Attachment(emailRequest.File, emailRequest.FileName));

            using var client = new SmtpClient(emailRequest.Host, emailRequest.Port)
            {
                UseDefaultCredentials = true,
                Credentials = new NetworkCredential(emailRequest.SmtpUsername, emailRequest.SmtpPassword),
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network
            };

            try
            {
                _logger.LogInformation($"Send Email!");
                client.Send(message);
                _logger.LogInformation($"Email sent!");
                emailRequest.File.Close();
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error send email message. Detalhes: {ex.Message}");
                emailRequest.File.Close();
                return false;
            }
        }

        private string HttpRequest(string url)
        {
            string siteContent;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36");
            request.AutomaticDecompression = DecompressionMethods.GZip;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())  // Go query google
            using (Stream responseStream = response.GetResponseStream())               // Load the response stream
            using (StreamReader streamReader = new StreamReader(responseStream))       // Load the stream reader to read the response
            {
                siteContent = streamReader.ReadToEnd(); // Read the entire response and store it in the siteContent variable
            }

            return siteContent;
        }
    }
}
