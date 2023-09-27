using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class Program
{
    static List<Product> products = new List<Product>();
    static List<Sale> sales = new List<Sale>();

    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        InitializeProducts();
        InitializeExcelFile();

        while (true)
        {
            Console.Clear();
            Console.WriteLine("1- Petrol Sahibi Girişi");
            Console.WriteLine("2- Kasiyer Girişi");
            Console.WriteLine("3- Çıkış");
            Console.Write("Seçiminizi yapın: ");
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    PetrolSahibiMenu();
                    break;
                case "2":
                    KasiyerMenu();
                    break;
                case "3":
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("Geçersiz seçenek! Lütfen tekrar deneyin.");
                    break;
            }
        }
    }

    static void PetrolSahibiMenu()
    {
        while (true)
        {
            Console.Clear();
            Console.WriteLine("Petrol Sahibi Menüsü");
            Console.WriteLine("1- Ürünleri Görüntüle");
            Console.WriteLine("2- Toplam Satışı Görüntüle");
            Console.WriteLine("3- Ürün Stoklarına Ekle");
            Console.WriteLine("4- Çıkış");
            Console.Write("Seçiminizi yapın: ");
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    ListProducts();
                    break;
                case "2":
                    ShowTotalSales();
                    break;
                case "3":
                    AddToStock();
                    break;
                case "4":
                    return;
                default:
                    Console.WriteLine("Geçersiz seçenek! Lütfen tekrar deneyin.");
                    break;
            }
        }
    }

    static void KasiyerMenu()
    {
        while (true)
        {
            Console.Clear();
            Console.WriteLine("Kasiyer Menüsü");
            Console.WriteLine("1- Ürün Satışı");
            Console.WriteLine("2- Çıkış");
            Console.Write("Seçiminizi yapın: ");
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    SellProduct();
                    break;
                case "2":
                    return;
                default:
                    Console.WriteLine("Geçersiz seçenek! Lütfen tekrar deneyin.");
                    break;
            }
        }
    }

    static void InitializeProducts()
    {
        products.Add(new Product("Su\t", 10, 100));
        products.Add(new Product("Çikolata", 15, 100));
        products.Add(new Product("Soğuk Kahve", 25, 100));
        products.Add(new Product("Hediyelik Eşya", 100, 50));
    }

    static void InitializeExcelFile()
    {
        using (var package = new ExcelPackage(new FileInfo("MarketKayitlari.xlsx")))
        {
            var workbook = package.Workbook;

            if (workbook.Worksheets.All(ws => ws.Name != "Ürünler"))
            {
                var productsSheet = workbook.Worksheets.Add("Ürünler");
                productsSheet.Cells.LoadFromCollection(products, true);
            }

            if (workbook.Worksheets.All(ws => ws.Name != "Satışlar"))
            {
                var salesSheet = workbook.Worksheets.Add("Satışlar");
                salesSheet.Cells.LoadFromCollection(sales, true);
            }

            package.Save();
        }
    }

    static void ListProducts()
    {
        Console.Clear();
        Console.WriteLine("Ürünler:");
        Console.WriteLine("ID\tÜrün Adı\tFiyat\tStok");
        foreach (var product in products)
        {
            Console.WriteLine($"{product.ID}\t{product.Name}\t{product.Price}\t{product.Stock}");
        }
        Console.WriteLine("Devam etmek için bir tuşa basın...");
        Console.ReadKey();
    }

    static void ShowTotalSales()
    {
        Console.Clear();
        double totalSales = 0;
        foreach (var sale in sales)
        {
            totalSales += sale.TotalPrice;
        }
        Console.WriteLine($"Toplam Satış: {totalSales} TL");
        Console.WriteLine("Devam etmek için bir tuşa basın...");
        Console.ReadKey();
    }

    static void AddToStock()
    {
        Console.Clear();
        ListProducts();
        Console.Write("Stok eklemek istediğiniz ürünün ID'sini girin ya da çıkış yapmak için q yazın: ");
        string cevap = Console.ReadLine();
        if (int.TryParse(cevap, out int productId))
        {
            Console.Write("Eklemek istediğiniz stok miktarını girin: ");
            if (int.TryParse(Console.ReadLine(), out int quantity))
            {
                Product product = products.Find(p => p.ID == productId);
                if (product != null)
                {
                    product.Stock += quantity;
                    Console.WriteLine($"{product.Name} ürününün stoku güncellendi. Yeni stok: {product.Stock}");

                    // Excel sayfasını güncelle
                    UpdateExcelSheet("Ürünler", products);
                }
                else
                {
                    Console.WriteLine("Geçersiz ürün ID'si!");
                }
            }
            else
            {
                Console.WriteLine("Geçersiz miktar!");
            }
        }
        else if (cevap.Equals("q"))
        {
            return;
        }
        else
        {
            Console.WriteLine("Geçersiz ID!");
        }

        Console.WriteLine("Devam etmek için bir tuşa basın...");
        Console.ReadKey();
    }

    static void SellProduct()
    {
        Console.Clear();
        ListProducts();
        Console.Write("Satış yapmak istediğiniz ürünün ID'sini girin ya da çıkmak için q yazın: ");
        string cevap = Console.ReadLine();
        if (int.TryParse(cevap, out int productId))
        {
            Product product = products.Find(p => p.ID == productId);
            if (product != null && product.Stock > 0)
            {
                Console.Write("Satış miktarını girin: ");
                if (int.TryParse(Console.ReadLine(), out int quantity) && quantity <= product.Stock)
                {
                    double totalPrice = quantity * product.Price;
                    sales.Add(new Sale(product.Name, quantity, totalPrice, DateTime.Now));
                    product.Stock -= quantity;
                    Console.WriteLine($"Satış başarılı! Toplam fiyat: {totalPrice} TL");

                    // Excel sayfalarını güncelle
                    UpdateExcelSheet("Ürünler", products);
                    UpdateExcelSheet("Satışlar", sales);
                }
                else
                {
                    Console.WriteLine("Geçersiz miktar veya yetersiz stok!");
                }
            }
            else
            {
                Console.WriteLine("Geçersiz ürün ID'si veya stokta ürün yok!");
            }
        }
        else if (cevap.Equals("q"))
        {
            return;
        }
        else
        {
            Console.WriteLine("Geçersiz ID!");
        }

        Console.WriteLine("Devam etmek için bir tuşa basın...");
        Console.ReadKey();
    }

    static void UpdateExcelSheet<T>(string sheetName, List<T> data)
    {
        using (var package = new ExcelPackage(new FileInfo("MarketKayitlari.xlsx")))
        {
            var workbook = package.Workbook;

            var sheet = workbook.Worksheets.FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(sheetName);
            }

            sheet.Cells.Clear();
            sheet.Cells.LoadFromCollection(data, true);

            package.Save();
        }
    }
}

class Product
{
    public int ID { get; set; }
    public string Name { get; set; }
    public double Price { get; set; }
    public int Stock { get; set; }

    private static int nextID = 1;

    public Product(string name, double price, int stock)
    {
        ID = nextID++;
        Name = name;
        Price = price;
        Stock = stock;
    }
}

class Sale
{
    public string ProductName { get; set; }
    public int Quantity { get; set; }
    public double TotalPrice { get; set; }
    public DateTime SaleDate { get; set; }

    public Sale(string productName, int quantity, double totalPrice, DateTime saleDate)
    {
        ProductName = productName;
        Quantity = quantity;
        TotalPrice = totalPrice;
        SaleDate = saleDate;
    }
}
