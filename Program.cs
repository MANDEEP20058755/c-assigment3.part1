using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

public class User
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Username { get; set; }
    public string Email { get; set; }
    public Address Address { get; set; }
    public string Phone { get; set; }
    public string Website { get; set; }
    public Company Company { get; set; }
}

public class Address
{
    public string Street { get; set; }
    public string Suite { get; set; }
    public string City { get; set; }
    public string Zipcode { get; set; }
    public Geo Geo { get; set; }
}

public class Geo
{
    public string Lat { get; set; }
    public string Lng { get; set; }
}

public class Company
{
    public string Name { get; set; }
    public string CatchPhrase { get; set; }
    public string Bs { get; set; }
}

public class Program
{
    public static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string apiUrl = "https://jsonplaceholder.typicode.com/users";

        HttpClient httpClient = new HttpClient();
        var userData = await FetchUserDataAsync(httpClient, apiUrl);
        DisplayUserData(userData);
        SaveToExcel(userData, "C:\\path\\to\\your\\directory\\user_data.xlsx");
    }

    private static async Task<User[]> FetchUserDataAsync(HttpClient client, string url)
    {
        HttpResponseMessage response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        string responseBody = await response.Content.ReadAsStringAsync();
        return JsonConvert.DeserializeObject<User[]>(responseBody);
    }

    private static void DisplayUserData(User[] users)
    {
        Console.WriteLine("{0,-30} {1,-30} {2,-15} {3,-50}", "Name", "Email", "Phone", "Address");
        Console.WriteLine(new string('-', 120));

        foreach (var user in users)
        {
            Console.WriteLine("{0,-30} {1,-30} {2,-15} {3,-50}", user.Name, user.Email, user.Phone, $"{user.Address.Street}, {user.Address.Suite}, {user.Address.City}, {user.Address.Zipcode}");
        }
    }

    private static void SaveToExcel(User[] users, string filePath)
    {
        using (var excelPackage = new ExcelPackage())
        {
            var worksheet = excelPackage.Workbook.Worksheets.Add("UserDetails");

            worksheet.Cells["A1"].Value = "Name";
            worksheet.Cells["B1"].Value = "Email";
            worksheet.Cells["C1"].Value = "Phone";
            worksheet.Cells["D1"].Value = "Address";

            for (int i = 0; i < users.Length; i++)
            {
                var user = users[i];
                worksheet.Cells[$"A{i + 2}"].Value = user.Name;
                worksheet.Cells[$"B{i + 2}"].Value = user.Email;
                worksheet.Cells[$"C{i + 2}"].Value = user.Phone;
                worksheet.Cells[$"D{i + 2}"].Value = $"{user.Address.Street}, {user.Address.Suite}, {user.Address.City}, {user.Address.Zipcode}";
            }

            FileInfo file = new FileInfo(filePath);
            excelPackage.SaveAs("/Users/mssaini/C#3/ConsoleApp1/user.xlsx");
        }
    }
}