using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
namespace selenium16
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click1(object sender, RoutedEventArgs e)
        {

            string path = @"C:\Users\Admin\Desktop\SQA\DataTest.xlsx";
            int sheet = 1;

            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(path);
            Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[sheet];
            int startRow = 2, endRow = 23;
            List<string> nameList = getInputs(ws, startRow, endRow, 2); 
            List<string> dobList = getInputs(ws, startRow, endRow, 3);
            List<string> sexList = getInputs(ws, startRow, endRow, 4);
            List<string> addressList = getInputs(ws, startRow, endRow, 5);
            List<string> phoneList = getInputs(ws, startRow, endRow, 6);
            List<string> roleIDList = getInputs(ws, startRow, endRow, 7);
            List<string> falcultyList = getInputs(ws, startRow, endRow, 8);
            List<string> classList = getInputs(ws, startRow, endRow, 9);
            List<string> codeList = getInputs(ws, startRow, endRow, 10);
            List<string> passList = getInputs(ws, startRow, endRow, 11);

            ChromeDriver chromeDriver = new ChromeDriver();

            chromeDriver.Url = "http://localhost:8080/qldiem/dang-nhap?action=login";
            chromeDriver.Navigate();

            chromeDriver.Manage().Window.Maximize();

            //var username = chromeDriver.FindElements(By.XPath("//input[@class='fadeIn second']"))[0];
            var username = chromeDriver.FindElement(By.XPath("//input[@class='fadeIn second']"));
            username.SendKeys("admin");

            var password = chromeDriver.FindElement(By.XPath("//input[@id='password']"));
            password.SendKeys("123456");

            var loginBtn = chromeDriver.FindElement(By.XPath("//input[@class='fadeIn fourth']"));
            loginBtn.Click();

            chromeDriver.FindElement(By.LinkText("Quản lý người dùng")).Click();

           

            string pathImage = @"C:\Users\Admin\Desktop\SQA\TC";

            for (int i = 0; i < nameList.Count; i++)
            {
                chromeDriver.Navigate().Refresh();
                chromeDriver.FindElement(By.XPath("//input[@class='btn btn-success']")).Click();

                var name = chromeDriver.FindElement(By.XPath("//input[@name='fullName']"));
                name.SendKeys(nameList[i]);

                var dob = chromeDriver.FindElement(By.XPath("//input[@name='dateOfBirth']"));
                dob.SendKeys(dobList[i]);

                var sex = chromeDriver.FindElement(By.XPath("(//input[@name='gender'])[1]"));
                //if (sexList[i] == "Nam") sex = chromeDriver.FindElement(By.XPath("(//input[@name='gender'])[0]"));
                sex.Click();

                var address = chromeDriver.FindElement(By.XPath("(//input[@name='address'])"));
                address.SendKeys(addressList[i]);

                var phone = chromeDriver.FindElement(By.XPath("(//input[@name='phone'])"));
                phone.SendKeys(phoneList[i]);

                SelectElement roleID = new SelectElement(chromeDriver.FindElement(By.XPath("(//select[@name='roleId'])")));
                roleID.SelectByText("Sinh viên");

                if (falcultyList[i] != "")
                {
                    SelectElement falculty = new SelectElement(chromeDriver.FindElement(By.XPath("(//select[@name='faculty'])")));
                    falculty.SelectByText(falcultyList[i]);
                }

                var classes = chromeDriver.FindElement(By.XPath("//input[@name='classroom']"));
                classes.SendKeys(classList[i]);

                var code = chromeDriver.FindElement(By.XPath("(//input[@name='userName'])"));
                code.SendKeys(codeList[i]);

                var pass = chromeDriver.FindElement(By.XPath("(//input[@name='password'])"));
                pass.SendKeys(passList[i]);

                // chụp màn hình
                Screenshot saveScreenShot = ((ITakesScreenshot)chromeDriver).GetScreenshot();
                string caseID = (i + 1).ToString("#00");
                saveScreenShot.SaveAsFile(pathImage + caseID + ".png", ScreenshotImageFormat.Png);

                var saveBtn = chromeDriver.FindElement(By.XPath("(//input[@value='Lưu'])"));
                saveBtn.Click();
            }

            wb.Close();
            excel.Quit();
        }

        private List<string> getInputs(Worksheet ws, int startRow, int endRow, int col)
        {
            List<string> list = new List<string>();

            for (int i = startRow; i <= endRow; i++)
            {
                string tmp = "";

                if ((ws.Cells[i, col] as _Excel.Range).Value != null)
                {
                    Microsoft.Office.Interop.Excel.Range cell = ws.Cells[i, col] as _Excel.Range;
                    tmp = cell.Value.ToString();
                }

                list.Add(tmp);
            }

            return list;
        }
    }
}
