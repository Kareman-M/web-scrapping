

using ClosedXML.Excel;
using HtmlAgilityPack;

namespace WebScrapper
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var factories = new List<Factories>();
            var httpClient = new HttpClient();
            List<string> urls = new List<string>
            ([
                "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=2",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=3",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=4",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=5",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=6",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=7",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=8",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/40/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%84%D8%A7%D9%85?page=9",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=8",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/32/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%88%D8%B1%D8%B3%D8%B9%D9%8A%D8%AF?page=9",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=8",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/33/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D8%A5%D8%B3%D9%85%D8%A7%D8%B9%D9%8A%D9%84%D9%8A%D8%A9?page=9",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/51/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=6",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/41/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A8%D9%86%D9%8A-%D8%B3%D9%88%D9%8A%D9%81?page=8",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/42/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D9%85%D9%86%D9%8A%D8%A7?page=6",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/43/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%8A%D9%88%D8%B7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/43/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%8A%D9%88%D8%B7?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/43/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%8A%D9%88%D8%B7?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/43/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%8A%D9%88%D8%B7?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/43/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%8A%D9%88%D8%B7?page=5",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1?page=2",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1?page=3",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1?page=4",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1?page=5",
"https://egyptianindustry.com/SearchR/1/Page/1/Zone/44/Industry/0/%D8%B3%D9%88%D9%87%D8%A7%D8%AC%D9%88%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%83%D9%88%D8%AB%D8%B1?page=6",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/46/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D9%82%D9%86%D8%A7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/46/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D9%82%D9%86%D8%A7?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/46/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D9%82%D9%86%D8%A7?page=3",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/47/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A7%D9%84%D8%A3%D9%82%D8%B5%D8%B1",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/48/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%A3%D8%B3%D9%88%D8%A7%D9%86",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/53/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D9%85%D8%B7%D8%B1%D9%88%D8%AD%D9%88%D8%A7%D9%84%D8%B3%D8%A7%D8%AD%D9%84",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/53/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D9%85%D8%B7%D8%B1%D9%88%D8%AD%D9%88%D8%A7%D9%84%D8%B3%D8%A7%D8%AD%D9%84?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/54/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%88%D8%A7%D8%AF%D9%8A%D8%A7%D9%84%D8%AC%D8%AF%D9%8A%D8%AF",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/54/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9%D8%A7%D9%84%D9%88%D8%A7%D8%AF%D9%8A%D8%A7%D9%84%D8%AC%D8%AF%D9%8A%D8%AF?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/55/Industry/0/%D8%A7%D9%84%D9%85%D9%86%D8%A7%D8%B7%D9%82-%D8%A7%D9%84%D8%AD%D8%B1%D8%A9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/55/Industry/0/%D8%A7%D9%84%D9%85%D9%86%D8%A7%D8%B7%D9%82-%D8%A7%D9%84%D8%AD%D8%B1%D8%A9?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/55/Industry/0/%D8%A7%D9%84%D9%85%D9%86%D8%A7%D8%B7%D9%82-%D8%A7%D9%84%D8%AD%D8%B1%D8%A9?page=3",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%B4%D9%82-%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=8",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/34/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%B4%D9%82%D8%A7%D9%84%D8%AB%D8%B9%D8%A8%D8%A7%D9%86?page=10",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/35/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%AA%D8%A8%D9%8A%D9%86",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/35/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%AA%D8%A8%D9%8A%D9%86?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/36/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%A8%D8%B3%D8%A7%D8%AA%D9%8A%D9%86?page=6",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%AC%D8%B3%D8%B1-%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=8",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=10",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=11",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=12",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=13",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=14",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/37/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%AC%D8%B3%D8%B1%D8%A7%D9%84%D8%B3%D9%88%D9%8A%D8%B3?page=15",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/38/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%AD%D8%B1%D9%81%D9%8A%D9%8A%D9%86",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/38/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%AD%D8%B1%D9%81%D9%8A%D9%8A%D9%86?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/38/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9-%D8%A7%D9%84%D8%AD%D8%B1%D9%81%D9%8A%D9%8A%D9%86?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/56/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%B9%D8%A8%D8%A7%D8%B3%D9%8A%D8%A9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/56/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%B9%D8%A8%D8%A7%D8%B3%D9%8A%D8%A9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/56/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%B9%D8%A8%D8%A7%D8%B3%D9%8A%D8%A9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=4",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=5",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=6",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=7",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=8",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=10",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/58/Industry/0/%D9%85%D9%86%D8%B7%D9%82%D8%A9%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D8%B9%D8%A7%D9%84%D8%B5%D9%86%D8%A7%D8%B9%D9%8A%D8%A9?page=11",
            "https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/59/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%AC%D9%85%D8%B5%D8%A9",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/59/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%AC%D9%85%D8%B5%D8%A9?page=2",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/59/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%AC%D9%85%D8%B5%D8%A9?page=3",
"https://www.egyptianindustry.com/SearchR/1/Page/1/Zone/59/Industry/0/%D9%85%D8%AF%D9%8A%D9%86%D8%A9-%D8%AC%D9%85%D8%B5%D8%A9?page=4",
            ]);


            foreach (var link in urls)
            {

                var linkhtml = await httpClient.GetStringAsync(link);
                var linkhtmlDocument = new HtmlDocument();
                linkhtmlDocument.LoadHtml(linkhtml);

                var divs = linkhtmlDocument.DocumentNode.Descendants("div").Where(node => node.HasClass("f-listings-item")).ToList();
                foreach (var div in divs)
                {
                    var insideDives = div.Descendants("div");
                    var name = div.Descendants("h1").Where(x => x.HasClass("f-listings-item__title")).First().ChildNodes[0].InnerHtml.Trim();

                    var adresses = div.Descendants("address").Where(x => x.HasClass("f-listings-item__address")).First().ChildNodes;

                    var phone = adresses.Count > 2 ? adresses[2].InnerHtml.Trim() : null;
                    var desc = adresses.Count > 6 ? adresses[6].InnerHtml.Trim() : null;

                    var otherdata2 = div.Descendants("address").First().Descendants("a").ToList();
                    var mail = otherdata2.Count >= 1 ? otherdata2[0].GetAttributes("href").Last().Value.Trim() : null;
                    var website = otherdata2.Count >= 2 ? otherdata2[1].GetAttributes("href").Last().Value.Trim() : null;


                    var contents = insideDives.Where(x => x.HasClass("listing-single__content")).First().Descendants("h2").First().InnerHtml.Split("/");

                    var city = contents[0].Trim();
                    var sector = contents[1].Trim();

                    factories.Add(new Factories
                    {
                        CityName = city,
                        Description = desc,
                        FactoryName = name.ToString(),
                        Link = link,
                        Mail = mail,
                        Website = website,
                        Sectort = sector,
                        Phone = phone,
                    });
                }

            }

            SaveToExcel(factories);
            Console.ReadLine();




        }

        private static int getMaxNumer(List<string> pagingNumbers)
        {
            var maxNumer = 0;
            foreach (var number in pagingNumbers)
            {
                if (int.Parse(number) > maxNumer) maxNumer = int.Parse(number);
            }
            return maxNumer;
        }

        private static void SaveToExcel(List<Factories> itemList)
        {
            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");
            var ws = workbook.Worksheet("sheetName");
            int row = 1;
            ws.Cell("A" + row.ToString()).Value = "FactoryName";
            ws.Cell("B" + row.ToString()).Value = "Phone";
            ws.Cell("C" + row.ToString()).Value = "Description";
            ws.Cell("D" + row.ToString()).Value = "Mail";
            ws.Cell("E" + row.ToString()).Value = "Website";
            ws.Cell("F" + row.ToString()).Value = "Sectort";
            ws.Cell("G" + row.ToString()).Value = "CityName";
            ws.Cell("H" + row.ToString()).Value = "Link";
            row++;
            foreach (var item in itemList)
            {
                ws.Cell("A" + row.ToString()).Value = item.FactoryName;
                ws.Cell("B" + row.ToString()).Value = item.Phone;
                ws.Cell("C" + row.ToString()).Value = item.Description;
                ws.Cell("D" + row.ToString()).Value = item.Mail;
                ws.Cell("E" + row.ToString()).Value = item.Website;
                ws.Cell("F" + row.ToString()).Value = item.Sectort;
                ws.Cell("G" + row.ToString()).Value = item.CityName;
                ws.Cell("H" + row.ToString()).Value = item.Link;
                row++;
            }

            workbook.SaveAs("factories.xlsx");
        }
    }

}