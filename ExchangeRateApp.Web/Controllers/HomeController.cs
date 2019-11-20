using ExchangeRateApp.Web.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Net;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

namespace ExchangeRateApp.Web.Controllers
{
    public class HomeController : Controller
    {
        private Random rand = new Random();
        private XmlDocument docStartDate, docEndDate;
        private List<ExchangeRate> startDateExchangeRates, endDateExchangeRates;

        public ActionResult Index()
        {
            return View();
        }

        // Başlangıç ve bitiş tarihi girilip Gönder butonuna basıldığında bu metoda istek yapıyoruz. 
        // Başlangıç ve bitiş tarihini almak için, oluşturduğumuz RequestData modelini kullanıyoruz.
        // Bu metoddan geriye bir Json çıktı döndürerek Chart üzerinde gerekli verileri gösteriyoruz.
        [HttpPost]
        public JsonResult Index(RequestData requestData)
        {            
            docStartDate = new XmlDocument();
            docEndDate = new XmlDocument();
            startDateExchangeRates = new List<ExchangeRate>();
            endDateExchangeRates = new List<ExchangeRate>();

            string start_date_str = requestData.start_date;
            string end_date_str = requestData.end_date;
            if (start_date_str == "" || end_date_str == "") {
                ViewBag.Message = "Başlangıç ve bitiş tarihi boş olamaz.";
                return null;
            } else {
                // Bitiş tarihinin başlangıç tarihinden küçük olup olmadığını kontrol etmek için 
                // gelen ifadeleri DateTime'a çeviriyoruz.
                DateTime start_date = DateTime.ParseExact(start_date_str, "yyyy-MM-dd", CultureInfo.InstalledUICulture);
                DateTime end_date = DateTime.ParseExact(end_date_str, "yyyy-MM-dd", CultureInfo.InstalledUICulture);

                if (start_date > end_date) {
                    ViewBag.Message = "Bitiş tarihi başlangıç tarihinden küçük olamaz.";
                    return null;
                } else {
                    // Başlangıç ve bitiş tarihini DateTime'a çevirdikten sonra gün ve ay kısmının başındaki 0'ları atıyor.
                    // Tcmb Url'i oluşturulurken sorun olmaması adına gün ve ay eğer tek haneli ise başlarına 0 ekliyoruz.
                    // Ör: 1/1/2019 => 01/01/2019
                    string startDateDay = start_date.Day.ToString();
                    if (startDateDay.Length == 1) startDateDay = "0" + startDateDay;
                    string startDateMonth = start_date.Month.ToString();
                    if (startDateMonth.Length == 1) startDateMonth = "0" + startDateMonth;
                    string startDateYear = start_date.Year.ToString();
                    string startDate = startDateDay + '.' + startDateMonth + '.' + startDateYear;                    

                    string endDateDay = end_date.Day.ToString();
                    if (endDateDay.Length == 1) endDateDay = "0" + endDateDay;
                    string endDateMonth = end_date.Month.ToString();
                    if (endDateMonth.Length == 1) endDateMonth = "0" + endDateMonth;
                    string endDateYear = end_date.Year.ToString();
                    string endDate = endDateDay + '.' + endDateMonth + '.' + endDateYear;

                    // Excel çıktısı için son girilen başlangıç ve bitiş tarihlerini Session'a atıyoruz.
                    Session["start_date_excel"] = startDate;
                    Session["end_date_excel"] = endDate;

                    // Başlangıç ve bitiş tarihi eğer daha önceden de girilmişse Tcmb linkine gitmeden
                    // veriyi Session üzerinden getiriyoruz.
                    if (Session[startDate + "-" + endDate] != null) {
                        return Json(Session[startDate + "-" + endDate], JsonRequestBehavior.AllowGet);
                    }

                    // Chart üzerinde verileri göstermek için ChartData modeli oluşturduk.
                    var chartData = new ChartData();                    

                    // Chart üzerindeki label'lara başlangıç ve bitiş tarihlerini yazdırıyoruz.
                    List<string> dates = new List<string>();
                    dates.Add(startDate);
                    dates.Add(endDate);
                    chartData.labels = dates;

                    // Girilen başlangıç ve bitiş tarihlerine ilişkin Tcmb üzerinde veri olup olmadığını kontrol ederek
                    // veri yoksa hata döndürüyoruz.
                    string startDateLink = "https://www.tcmb.gov.tr/kurlar/" + startDateYear + startDateMonth + "/" + startDateDay + startDateMonth + startDateYear + ".xml";                    
                    try {
                        docStartDate.Load(startDateLink);
                    } catch {
                        Response.StatusCode = (int)HttpStatusCode.BadRequest;
                        string error = "Resmi tatil, hafta sonu ve yarım iş günü çalışılan günlerde gösterge niteliğinde kur bilgisi yayımlanmamaktadır.";
                        return Json(error);
                    }

                    string endDateLink = "https://www.tcmb.gov.tr/kurlar/" + endDateYear + endDateMonth + "/" + endDateDay + endDateMonth + endDateYear + ".xml";
                    try {
                        docEndDate.Load(endDateLink);
                    } catch {
                        Response.StatusCode = (int)HttpStatusCode.BadRequest;
                        string error = "Resmi tatil, hafta sonu ve yarım iş günü çalışılan günlerde gösterge niteliğinde kur bilgisi yayımlanmamaktadır.";
                        return Json(error);
                    }

                    // Tcmb üzerinden Döviz Cinsi ve Döviz Satış değerlerine ulaşıyoruz.
                    foreach (XmlNode node in docStartDate.DocumentElement)
                    {
                        string name = node["Isim"].InnerText;
                        string selling = node["ForexSelling"].InnerText;

                        var exchangeRate = new ExchangeRate();
                        exchangeRate.name = name;
                        exchangeRate.selling = selling;
                        startDateExchangeRates.Add(exchangeRate);
                    }                    
                    
                    foreach (XmlNode node in docEndDate.DocumentElement)
                    {
                        string name = node["Isim"].InnerText;
                        string selling = node["ForexSelling"].InnerText;

                        var exchangeRate = new ExchangeRate();
                        exchangeRate.name = name;
                        exchangeRate.selling = selling;
                        endDateExchangeRates.Add(exchangeRate);
                    }
                    
                    List<ChartDataset> list = new List<ChartDataset>();

                    for (int i = 0; i < startDateExchangeRates.Count; i++)
                    {
                        // Chart üzerinde her döviz cinsi için ayrı bir renk oluşturmak adına random renk değerleri atıyoruz.
                        int firstColor = rand.Next(256);
                        int secondColor = rand.Next(256);
                        int thirdColor = rand.Next(256);
                        string color = "rgb(" + firstColor + ", " + secondColor + ", " + thirdColor + ")";

                        // Chart üzerindeki verilere ilişkin ChartDataset modelimizi oluşturarak
                        // bu verileri listeye dolduruyoruz.
                        var chartDataSet = new ChartDataset();
                        chartDataSet.label = startDateExchangeRates[i].name;
                        chartDataSet.backgroundColor = color;
                        chartDataSet.borderColor = color;
                        List<string> data = new List<string>();
                        data.Add(startDateExchangeRates[i].selling);
                        data.Add(endDateExchangeRates[i].selling);
                        chartDataSet.data = data;
                        chartDataSet.fill = false;

                        list.Add(chartDataSet);
                    }              

                    chartData.datasets = list;

                    // Girilen başlangıç ve bitiş tarihine ilişkin veriyi Session'a atıyoruz.
                    Session[startDate + "-" + endDate] = chartData;

                    // Excel çıktısı için son elde edilen verileri Session'a atıyoruz.
                    Session["list_excel"] = list;

                    return Json(chartData, JsonRequestBehavior.AllowGet);                    
                }
            }                
        }
        
        public void ExportToExcel()
        {
            // Excel'deki ilk satırın sütunlarına DÖVİZ CİNSİ, başlangıç ve bitiş tarihini yazdırıyoruz.
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[3] {
                        new DataColumn("DÖVİZ CİNSİ"),
                        new DataColumn(Session["start_date_excel"].ToString()),
                        new DataColumn(Session["end_date_excel"].ToString())});
            // Excel'deki diğer satırlara Session'da bulunan verileri yazdırıyoruz.
            List<ChartDataset> list = (List<ChartDataset>)Session["list_excel"];
            foreach (var item in list) {
                // Excel'e nokta(.) olarak giden veriyle ilgili bir sıkıntı oluyor.
                // Ör: 5.1234 verisi 51.234 olarak kaydediliyor. 
                // Bunu önlemek adına nokta olan verileri virgüle(,) çeviriyoruz.
                string data1 = item.data[0].Replace('.', ',');
                string data2 = item.data[1].Replace('.', ',');
                dt.Rows.Add(item.label, data1, data2);
            }

            string dosyaAdi = "exchange_rate_app";
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            var grid = new GridView();
            grid.DataSource = dt;
            grid.DataBind();            
            grid.AllowPaging = false;
            grid.RenderControl(hw);

            Response.Clear();
            Response.ClearHeaders();
            Response.ClearContent();
            Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1254");
            Response.Charset = "windows-1254";
            Response.AddHeader("content-disposition", "attachment;filename=" + dosyaAdi + ".xls");
            Response.ContentType = "application/vnd.ms-excel ";            
            Response.Output.Write(" <meta http-equiv='Content-Type' content='text/html; charset=windows-1254' />" + sw.ToString());            
            Response.Flush();
            Response.End();
        }
    }
}