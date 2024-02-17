
using HtmlAgilityPack;
using System.Net;
using OfficeOpenXml;


namespace Program
{

    public class Program {
        public async static Task Main(){

           
            
            var (productName, Category, productsNbr) = GetInfoFromUser();


        
        
            List<ProductInfo> result = await LoadProductsInfo(productName, Category, productsNbr);
            // foreach(ProductInfo productInfo in result){
            //     System.Console.WriteLine(productInfo);
            // }

            LoadDataOnExcelFile(result);

           
        }

        public static string ConstructUrl(string productName, string Category, int pageNbr = 1){
            string url = $"https://www.amazon.com/s?k={productName}&i={Category}-intl-ship&page={pageNbr}&ref=sr_pg_2";

            return url;
        }
       

        public async static Task<List<ProductInfo>> LoadProductsInfo(string productName, string Category, int ProductsNbr = 16){
            int page = 1;
            List<ProductInfo> productInfos = new();

            do{

                string url = ConstructUrl(productName, Category, page);
                
                string result = await Program.LoadHtmlPage(url);

                if(result.Length == 0)
                    throw new Exception("problem with loading page");

                HtmlDocument htmlDoc = new();
                htmlDoc.LoadHtml(result);

                // Searched Result container
                HtmlNode SearchedResultContainer = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"search\"]/div[1]/div[1]/div/span[1]/div[1]");

                // System.Console.WriteLine(SearchedResultContainer.Name);
                
                // list of target product
                IEnumerable<HtmlNode> SearchedResult = SearchedResultContainer.ChildNodes.Where(c => c.Name == "div" && c.Attributes.Contains("data-component-type") );

                // foreach(var h in SearchedResult)
                //     System.Console.WriteLine(h.Attributes["data-asin"].Value);

                // product Infos list
                

                foreach(HtmlNode productNode in SearchedResult){
                    // LoadProductInfo(productNode);
                    productInfos.Add(LoadProductInfo(productNode));
                    // System.Console.WriteLine(productNode.Attributes["data-asin"].Value);
                }

                page++;

            }while(productInfos.Count() < ProductsNbr);


            

            return productInfos.Take(ProductsNbr).ToList();

            

        }

        public static ProductInfo LoadProductInfo(HtmlNode productNode){
            
           

            string? Title = productNode.SelectSingleNode($".//div/div/span/div/div/div/div[2]/div/div/div[1]/h2/a/span")?.InnerText;
            string? Price = productNode.SelectSingleNode(".//div/div/span/div/div/div/div[2]/div/div/div[3]/div[1]/div/div[1]/div[2]/div[1]/a/span/span[1]")?.InnerHtml;
            string? ImageUrl = productNode.SelectSingleNode(".//div/div/span/div/div/div/div[1]/div/div[2]/div/span/a/div/img")?.Attributes["src"]?.Value;

          

            return new ProductInfo(Title, Price, ImageUrl);


        }

     
        public static async Task<String> LoadHtmlPage(string url){
            string result = "";
            HttpClientHandler handler = new HttpClientHandler();

            // Set user agent header to mimic a real browser
            handler.DefaultProxyCredentials = CredentialCache.DefaultCredentials;
            handler.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            // handler.AutomaticDecompression = DecompressionMethods.GZip;
            handler.AllowAutoRedirect = true;
            handler.UseCookies = true;
            handler.CookieContainer = new CookieContainer();
            handler.UseDefaultCredentials = false;

            using HttpClient httpClient = new(handler);

            httpClient.DefaultRequestHeaders.Add(
                "User-Agent", 
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
            );

        
            try{

                    var response = await httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();
                    
                    // string result
                    result = response.Content.ReadAsStringAsync().Result;
                    
                    
                    
            }catch(HttpRequestException e){
                System.Console.WriteLine($"Exception : {e.Message}");
            }

            return result; 
        }
   
        public static void LoadDataOnExcelFile(List<ProductInfo> productInfos, string ProductName = "productX"){
            
            try{
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a worksheet to the package
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Populate the worksheet with data
                    worksheet.Cells["A1"].Value = "ProductTitle";
                    worksheet.Cells["B1"].Value = "ProductPrice";
                    worksheet.Cells["C1"].Value = "ImageUrl";

                    for(int i = 2; i < productInfos.Count() + 2; i++){
                        var (Title, Price, ImageUrl) = productInfos[i - 2];

                        worksheet.Cells[$"A{i}"].Value = Title;
                        worksheet.Cells[$"B{i}"].Value = Price;
                        worksheet.Cells[$"C{i}"].Value = ImageUrl;
                    }

                    string date = DateTime.UtcNow.ToLongDateString();

                    // Save the Excel package to a file
                    string filePath = $"example-{ProductName}-{date}.xlsx";
                    excelPackage.SaveAs(new System.IO.FileInfo(filePath));
                    
                    Console.WriteLine($"Excel file created: {filePath}");
                }
            }catch{
                System.Console.WriteLine("Exception : Something went wrong while constructing excel file");
            }
        

        }
        

        public static (string ProductName, string Category, int ProductsNbr) GetInfoFromUser(){
            // Define the categories array
            string[] categories = new string[]
            {
                "Electronics",
                "Clothing-Shoes-Jewelry",
                "Home-Kitchen",
                "Books",
                "Health-Household",
                "Toys-Games",
                "Beauty-Personal-Care",
                "Tools-Home-Improvement",
                "Sports-Outdoors",
                "Automotive",
                "Grocery-Gourmet-Food",
                "Pet-Supplies",
                "Office-Products",
                "Musical-Instruments",
                "Industrial-Scientific",
                "Baby",
                "Arts-Crafts-Sewing",
                "Patio-Lawn-Garden",
                "Software",
                "Movies-TV",
                "Video-Games",
                "Home-Audio-Theater",
                "Computers-Accessories",
                "Cell-Phones-Accessories",
                "Kindle-Store",
                "Digital-Music",
                "Appliances",
                "Electronics-Accessories-Supplies",
                "Luggage-Travel-Gear",
                "Watches"
            };

            // Prompt user for input
            Console.WriteLine("Welcome to the Product Information App!");
            Console.WriteLine("Please enter the following information:");

            // Get ProductName from user
            Console.Write("Product Name: ");
            string productName = Console.ReadLine() ?? "book";

            // Get ProductNumber from user
            int productNumber;
            while(true){

                Console.Write("Product Number: ");
                

                if(int.TryParse(Console.ReadLine(), out productNumber)){
                    break;
                }else{
                    System.Console.WriteLine("Product Number must be Number !");
                }

            }

            // Get Category from user and validate
            string category;
            while (true)
            {
                Console.WriteLine("Available categories:");
                foreach (string cat in categories)
                {
                    Console.WriteLine("- " + cat);
                }
                Console.Write("Category: ");
                category = Console.ReadLine() ?? "";

                if (Array.IndexOf(categories, category) != -1)
                {
                    break; // Category is valid
                }
                else
                {
                    Console.WriteLine("Invalid category. Please choose from the provided list.");
                }
            }

            // Display the entered information
            Console.WriteLine("\nProduct Information:");
            Console.WriteLine($"Product Name: {productName}");
            Console.WriteLine($"Product Number: {productNumber}");
            Console.WriteLine($"Category: {category}");


            return (productName, category, productNumber);
        }
 }

   public record ProductInfo(String? Title, String? Price, String? ImageUrl);
}





