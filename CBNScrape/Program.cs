using HtmlAgilityPack;
using System;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using RandomSolutions;
using System.IO;

namespace CBNScrape
{
    public class SecuritiesModel
    {
        public string AuctionDate { get; set; }
        public string SecurityType { get; set; }
        public string Tenor { get; set; }
        public string AuctionNo { get; set; }
        public string Auction { get; set; }
        public string MaturityDate { get; set; }
        public string TotalSubscription { get; set; }
        public string TotalSuccessful { get; set; }
        public string RangeBid { get; set; }
        public string SuccessfulBidRate { get; set; }
        public string Description { get; set; }
        public string Rate { get; set; }
        public string TrueYield { get; set; }
        public string AmountOffered { get; set; }
    }
    class Program
    {
        private static readonly string BaseUrl = "https://www.cbn.gov.ng";

       private static readonly HtmlWeb scraper = new HtmlWeb();
        static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");


            var page = await scraper.LoadFromWebAsync(BaseUrl);

            var govtSecurityElement = page.DocumentNode.SelectSingleNode("//*[@id=\"header\"]/div[2]/div/ul/li[8]/div/div[1]/div[4]/ol/li[1]/a");

            var link = govtSecurityElement.Attributes.FirstOrDefault(x => x.Name == "href")?.Value;

            if (link == null)
            {
                Console.WriteLine("No link found");
                return;
            }

            var govtSecurityPage = await scraper.LoadFromWebAsync(BaseUrl + link);

            var pageLinks = govtSecurityPage.DocumentNode.SelectNodes("//*[@id=\"ContentTextinner\"]/a")?
                .Select(x =>
                x.Attributes.FirstOrDefault(x => x.Name == "href")?.Value);


            var listExtract = new List<SecuritiesModel>();
            foreach (var lk in pageLinks)
            {
              var extractedInfo = await ExtractPageInfo(BaseUrl + "/rates/" + lk);
                listExtract.AddRange(extractedInfo);
            }


            var bytes = listExtract.ToExcel();
           await File.WriteAllBytesAsync("sec.xlsx", bytes);
        }

        private static async Task<List<SecuritiesModel>> ExtractPageInfo(string url)
        {
            var pageData = await scraper.LoadFromWebAsync(url);

            var pageTableRows = pageData.DocumentNode.SelectNodes("//*[@id=\"mytables\"]/tr");

            var listSecurityData = new List<SecuritiesModel>();

            for (int i = 0; i < pageTableRows.Count / 15; i++)
            {
                var data = pageTableRows.Skip(i * 15).Take(15).ToList();


                var model = new SecuritiesModel
                {
                    AmountOffered = data[13].ChildNodes[1].InnerText,
                    Auction = data[4].ChildNodes[1].InnerText,
                    AuctionDate = data[0].ChildNodes[1].InnerText,
                    AuctionNo = data[3].ChildNodes[1].InnerText,
                    Description = data[10].ChildNodes[1].InnerText,
                    MaturityDate = data[5].ChildNodes[1].InnerText,
                    RangeBid = data[8].ChildNodes[1].InnerText,
                    Rate = data[11].ChildNodes[1].InnerText,
                    SecurityType = data[1].ChildNodes[1].InnerText,
                    SuccessfulBidRate = data[9].ChildNodes[1].InnerText,
                    Tenor = data[2].ChildNodes[1].InnerText,
                    TotalSubscription = data[6].ChildNodes[1].InnerText,
                    TotalSuccessful = data[7].ChildNodes[1].InnerText,
                    TrueYield = data[12].ChildNodes[1].InnerText,
                };

                listSecurityData.Add(model);

            }

            return listSecurityData;
        }
    }
}
