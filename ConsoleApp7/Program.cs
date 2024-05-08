using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;

public class DoctorScraper
{

	public static async Task<string> SearchInstagramAsync(IWebDriver searchDriver, string name, Dictionary<string, string> instagramProfileUrls)
	{
		string profileUrl = null;
		IWebDriver instagramDriver = null;
		try
		{
			string instagramSearchUrl = "https://www.instagram.com/";

			instagramDriver = new EdgeDriver(); // Создаем новый экземпляр драйвера Instagram
			instagramDriver.Navigate().GoToUrl(instagramSearchUrl);
			await Task.Delay(1000);
			var searchEMail = instagramDriver.FindElement(By.XPath("//*[@id=\"loginForm\"]/div/div[1]/div/label/input"));
			searchEMail.Click();
			searchEMail.SendKeys("kukakarakuzov@gmail.com");
			var searchpass = instagramDriver.FindElement(By.XPath("//*[@id=\"loginForm\"]/div/div[2]/div/label/input"));
			searchpass.Click();
			searchpass.SendKeys("Kuanis2006");
			var searchin = instagramDriver.FindElement(By.XPath("//*[@id=\"loginForm\"]/div/div[3]/button"));
			searchin.Click();
			await Task.Delay(8000);
			var searchprop = instagramDriver.FindElement(By.XPath("/html/body/div[3]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]"));
			searchin.Click();
			var searchInput = instagramDriver.FindElement(By.XPath("//*[@id=\"mount_0_0_pQ\"]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div[2]/span/div/a/div/div[1]/div/div"));
			searchInput.Click();
			await Task.Delay(1000);
			var search1 = instagramDriver.FindElement(By.XPath("//*[@id=\"mount_0_0_XI\"]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/input"));
			search1.SendKeys(name);
			var search2 = instagramDriver.FindElement(By.XPath("//*[@id=\"mount_0_0_fc\"]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/a[1]"));
			profileUrl = search2.GetAttribute("href");

			instagramProfileUrls[name] = profileUrl;
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Произошла ошибка при поиске на Instagram для {name}: {ex.Message}");
		}
		finally
		{
			instagramDriver?.Quit(); // Закрыть окно драйвера после завершения работы
		}
		return profileUrl;
	}




	public static async Task<ExcelPackage> ScrapeDoctorsAsync(string[] cityUrls, string excelFilePath, Dictionary<string, string> vkProfileUrls, Dictionary<string, string> instagramProfileUrls)
	{
		try
		{
			ExcelPackage excelPackage;

			if (File.Exists(excelFilePath))
			{
				var existingFile = new FileInfo(excelFilePath);
				using (var stream = existingFile.Open(FileMode.Open, FileAccess.ReadWrite))
				{
					excelPackage = new ExcelPackage(stream);
				}
			}
			else
			{
				excelPackage = new ExcelPackage();
			}

			var searchDriverPath = "C:\\Users\\Askha\\OneDrive\\Рабочий стол\\ConsoleApp7\\ConsoleApp7\\bin\\Debug\\net8.0\\msedgedriver.exe";
			var searchEdgeOptions = new EdgeOptions();
			searchEdgeOptions.PageLoadTimeout = TimeSpan.FromSeconds(60);
			searchEdgeOptions.AddArgument("--start-maximized");

			foreach (var cityUrl in cityUrls)
			{
				using (var searchDriver = new EdgeDriver(searchDriverPath, searchEdgeOptions))
				{
					var driverPath = "C:\\Users\\Askha\\OneDrive\\Рабочий стол\\ConsoleApp7\\ConsoleApp7\\bin\\Debug\\net8.0\\msedgedriver.exe";
					var edgeOptions = new EdgeOptions();
					edgeOptions.PageLoadTimeout = TimeSpan.FromSeconds(600);
					edgeOptions.AddArgument("--start-maximized");

					using (var driver = new EdgeDriver(driverPath, edgeOptions))
					{
						driver.Navigate().GoToUrl(cityUrl);

						IReadOnlyList<IWebElement> profiles = driver.FindElements(By.XPath("//div[@class='profile' and @data-id]"));

						foreach (var profile in profiles)
						{
							string dataId = profile.GetAttribute("data-id");
							IWebElement nameElement = profile.FindElement(By.XPath(".//a[contains(@class, 'profile--basic__title')]"));
							string fullName = nameElement.Text;

							string[] nameParts = fullName.Split(' ');
							string name = nameParts.Length > 1 ? $"{nameParts[0]} {nameParts[1]}" : fullName;

							// Вызываем метод SearchVKontakteAsync для поиска профиля VKontakte
							await SearchVKontakteAsync(searchDriver, name, vkProfileUrls);

							// Вызываем метод SearchInstagramAsync для поиска профиля Instagram
							await SearchInstagramAsync(searchDriver, name, instagramProfileUrls);

							string cityName = GetCityName(cityUrl);
							var worksheet = excelPackage.Workbook.Worksheets[cityName] ?? excelPackage.Workbook.Worksheets.Add(cityName);
							if (worksheet.Dimension == null)
							{
								worksheet.Cells["A1"].Value = "Name";
								worksheet.Cells["B1"].Value = "Data ID";
								worksheet.Cells["C1"].Value = "VKontakte";
								worksheet.Cells["D1"].Value = "INSTAGRAMM";
							}

							int rowCount = worksheet.Dimension?.End.Row + 1 ?? 1;

							// Проверяем уникальность данных по ID
							if (!IsDuplicate(worksheet, dataId))
							{
								worksheet.Cells[rowCount, 1].Value = name;
								worksheet.Cells[rowCount, 2].Value = dataId;
								worksheet.Cells[rowCount, 3].Value = vkProfileUrls.ContainsKey(name) ? vkProfileUrls[name] : "";
								worksheet.Cells[rowCount, 4].Value = instagramProfileUrls.ContainsKey(name) ? instagramProfileUrls[name] : "";
							}
						}
					}
				}
			}

			await SaveExcelFileAsync(excelPackage, excelFilePath);

			return excelPackage;
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Произошла ошибка при сборе данных о докторах: {ex.Message}");
			return null;
		}
	}



	public static bool IsDuplicate(ExcelWorksheet worksheet, string dataId)
	{
		for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
		{
			if (worksheet.Cells[row, 2].Value?.ToString() == dataId)
			{
				return true;
			}
		}
		return false;
	}


	public static async Task<string> SearchVKontakteAsync(IWebDriver searchDriver, string name, Dictionary<string, string> vkProfileUrls)
	{
		string profileUrl = null;
		try
		{
			string vkSearchUrl = "https://vk.com/search";

			searchDriver.Navigate().GoToUrl(vkSearchUrl);

			var searchInput = searchDriver.FindElement(By.XPath("//*[@id=\"search_query\"]"));
			searchInput.SendKeys(name);

			var searchButton = searchDriver.FindElement(By.XPath("//*[@id=\"search_query_wrap\"]/div/div[1]/button"));
			searchButton.Click();

			await WaitForPageLoad(searchDriver);

			profileUrl = searchDriver.Url;

			string[] nameParts = name.Split(' ');
			string lastName = nameParts[0];
			string firstName = nameParts.Length > 1 ? nameParts[1] : "";
			string key = $"{lastName} {firstName}";
			vkProfileUrls[key] = profileUrl;
		}
		catch (Exception ex)
		{
			Console.WriteLine($"Произошла ошибка при поиске на VKontakte для {name}: {ex.Message}");
		}
		return profileUrl;
	}

	public static async Task WaitForPageLoad(IWebDriver driver)
	{
		await Task.Delay(1000);
		while (true)
		{
			var isLoaded = (bool)((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState == 'complete'");
			if (isLoaded)
				break;
			await Task.Delay(500);
		}
	}

	public static async Task SaveExcelFileAsync(ExcelPackage excelPackage, string filePath)
	{
		await Task.Run(() => excelPackage.SaveAs(new FileInfo(filePath)));
	}

	public static string GetCityName(string cityUrl)
	{
		var parts = cityUrl.Split('/');
		var cityName = parts[parts.Length - 2];
		cityName = cityName.Replace("-", " ");
		Console.WriteLine($"City Name: {cityName}"); // Добавляем вывод названия города для отладки

		return cityName;
	}

	public static async Task Main(string[] args)
	{
/*соедини все в один юрл и исправь по мелочам там добавь сохранение старых данных и тд*/
		string[] cityUrls1 = {
			"https://idoctor.kz/astana/doctors",
			"https://idoctor.kz/almaty/doctors",
			"https://idoctor.kz/atirau/doctors",
			"https://idoctor.kz/aktau/doctors",
			"https://idoctor.kz/shymkent/doctors",
			"https://idoctor.kz/karaganda/doctors",
			"https://idoctor.kz/aktobe/doctors",
			"https://idoctor.kz/taldykorgan/doctors",
			"https://idoctor.kz/semey/doctors",
			"https://idoctor.kz/kostanay/doctors",
			"https://idoctor.kz/temirtau/doctors",
			"https://idoctor.kz/petropavlovsk/doctors",
			"https://idoctor.kz/uralsk/doctors",
			"https://idoctor.kz/kokshetau/doctors" ,
			"https://idoctor.kz/taraz/doctors",
			"https://idoctor.kz/pavlodar/doctors",
			"https://idoctor.kz/turkestan/doctors",
			"https://idoctor.kz/ust-kamenogorsk/doctors",
			"https://idoctor.kz/ekibastuz/doctors",
			"https://idoctor.kz/jezkazgan/doctors"
		};

		/*string[] cityUrls2 = {
			"https://idoctor.kz/taldykorgan/doctors",
			"https://idoctor.kz/semey/doctors",
			"https://idoctor.kz/kostanay/doctors",
			"https://idoctor.kz/temirtau/doctors",
			"https://idoctor.kz/petropavlovsk/doctors",
			"https://idoctor.kz/uralsk/doctors",
			"https://idoctor.kz/kokshetau/doctors"
		};*/

		/*string[] cityUrls3 = {
			"https://idoctor.kz/taraz/doctors",
			"https://idoctor.kz/pavlodar/doctors",
			"https://idoctor.kz/turkestan/doctors",
			"https://idoctor.kz/ust-kamenogorsk/doctors",
			"https://idoctor.kz/ekibastuz/doctors",
			"https://idoctor.kz/jezkazgan/doctors"
		};*/

		var vkProfileUrls1 = new Dictionary<string, string>();
		var vkProfileUrls2 = new Dictionary<string, string>();
		var vkProfileUrls3 = new Dictionary<string, string>();
		var instagramProfileUrls1 = new Dictionary<string, string>();
		var instagramProfileUrls2 = new Dictionary<string, string>();
		var instagramProfileUrls3 = new Dictionary<string, string>();

		ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

		var excelFilePath1 = "DoctorsData1.xlsx";
		var excelFilePath2 = "DoctorsData2.xlsx";
		var excelFilePath3 = "DoctorsData3.xlsx";

		var task1 = ScrapeDoctorsAsync(cityUrls1, excelFilePath1, vkProfileUrls1 , instagramProfileUrls1);
		var task2 = ScrapeDoctorsAsync(cityUrls1, excelFilePath2, vkProfileUrls2 , instagramProfileUrls2);
		var task3 = ScrapeDoctorsAsync(cityUrls1, excelFilePath3, vkProfileUrls3 , instagramProfileUrls3);


		// Ожидаем завершения всех задач
		await Task.WhenAll(task1, task2, task3);
	}
}
/*пробуй в бесплатном онлайн вивере а не в офиц от гугла там не работает почему то??*/