import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class Selenium {

        private static String  strID = "ID";
        private static String strHoten = "Hoten";
        private static String strNgaysinh = "Ngaysinh";
        private static String strGt = "GT";
        private static String strDc = "DiaChi";
        private static String strSdt = "Sdt";
        private static String strCv = "Chucvu";
        private static String strKhoa = "Khoa";
        private static String strLop ="Lop";
        private static String strMa ="Ma";
        private static String strMatkhau ="Matkhau";

        private static int intID;
        private static int intHoten;
        private static int intNgaysinh;
        private static int intGT;
        private static int intDc;
        private static int intSdt;
        private static int intChuvu;
        private static int intkhoa;
        private static int intLop;
        private static int intma;
        private static int intMatkhau;

        public static XSSFWorkbook  excelWBook;
        public  static XSSFSheet excelWSheet;
        public static XSSFCell cell;

        //This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
        public static void setExcelFile(String path, String sheetName) throws Exception {
            try {
                File file = new File(path);
                FileInputStream excelFile = new FileInputStream(file);
                // Access the test data sheet
                excelWBook = new XSSFWorkbook(excelFile);
                excelWSheet = excelWBook.getSheet(sheetName);
            } catch (Exception e) {
                throw (e);
            }
        }

        //This method is to read the test data from the Excel cell, in this we are passing parameters as Row num and Col num
        public static String getCellData(int rowNum, int colNum) throws Exception {
            try {
                cell = excelWSheet.getRow(rowNum).getCell(colNum);
                DataFormatter formatter = new DataFormatter();
                String cellData = formatter.formatCellValue(cell);
                return cellData;
            } catch (Exception e) {
                System.out.println(e);
                return "";
            }
        }

        public static int getRowBaseOnTCID(String TestCaseID) throws Exception {
            //Find number of rows in excel file
            int rowCount = excelWSheet.getLastRowNum();
            int i = 0;
            while (i <= rowCount) {
                if (getCellData(i, 0).equals(TestCaseID))
                    return i;
                else
                    i++;
            }
            return i;
        }

        public static int getColBaseOnFieldName(String fieldName) throws Exception {
//        get ca cot , add bat ki cai cot nao vao code
            int firstRow = excelWSheet.getFirstRowNum();
            int lastCol = excelWSheet.getRow(firstRow).getLastCellNum();
            int j = 0;
            while (j <= lastCol) {
                if (getCellData(firstRow, j).equals(fieldName))
                    return j;
                else
                    j++;
            }
            return j;
        }

    public static void main(String args[]) throws Exception {
        WebDriver driver;
        WebDriverWait wait;
        String baseURL = "http://localhost:8080/qldiem/trang-chu";
        String pathProject = System.getProperty("user.dir");
        String browser = " Chrome ";
        System.setProperty("webdriver.chrome.driver", pathProject + "/libs/chromedriver.exe");
        driver = new ChromeDriver();
        driver.get(baseURL);
        driver.manage().window().maximize();

            String strTestID = "TC02";
            int rowTC = getRowBaseOnTCID(strTestID);
            intID = getColBaseOnFieldName(strID);
            intHoten = getColBaseOnFieldName(strHoten);
            intNgaysinh = getColBaseOnFieldName(strNgaysinh);
            intGT = getColBaseOnFieldName(strGt);
            intDc = getColBaseOnFieldName(strDc);
            intSdt = getColBaseOnFieldName(strSdt);
            intChuvu = getColBaseOnFieldName(strCv);
            intkhoa = getColBaseOnFieldName(strKhoa);
            intLop = getColBaseOnFieldName(strLop);
            intma = getColBaseOnFieldName(strMa);
            intMatkhau = getColBaseOnFieldName(strMatkhau);

        // Đăng nhập
        driver.findElement(By.linkText("Đăng nhập")).click();
        wait= new WebDriverWait(driver, 90);
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='wrapper fadeInDown']")));
        driver.findElement(By.xpath("//input[@class='fadeIn second']")).sendKeys("admin");
        driver.findElement(By.xpath("//input[@id='password']")).sendKeys("123456");
        driver.findElement(By.xpath("//input[@class='fadeIn fourth']")).click();
        setExcelFile(pathProject + "/testdata/DataTest.xlsx", "Data01");
            //Thêm mới
            driver.findElement(By.linkText("Quản lý người dùng")).click();
            driver.findElement(By.xpath("//input[@class='btn btn-success']")).click();
            driver.findElement(By.xpath("//input[@name='fullName']")).sendKeys(getCellData(rowTC, intHoten));
            driver.findElement(By.xpath("//input[@name='dateOfBirth']")).sendKeys(getCellData(rowTC, intNgaysinh));
            driver.findElement(By.xpath("(//input[@name='gender'])[1]")).click();
            driver.findElement(By.xpath("(//input[@name='address'])")).sendKeys(getCellData(rowTC, intDc));
            driver.findElement(By.xpath("(//input[@name='phone'])")).sendKeys(getCellData(rowTC, intSdt));

            Select monday = new Select(driver.findElement(By.xpath("(//select[@name='roleId'])")));
            monday.selectByVisibleText("Sinh viên");
            monday.selectByIndex(1);
            Select khoa = new Select(driver.findElement(By.xpath("(//select[@name='faculty'])")));
            khoa.selectByVisibleText("CNTT");
            khoa.selectByIndex(1);

            driver.findElement(By.xpath("//input[@name='classroom']")).sendKeys(getCellData(rowTC, intLop));
            driver.findElement(By.xpath("(//input[@name='userName'])")).sendKeys(getCellData(rowTC, intma));
            driver.findElement(By.xpath("(//input[@name='password'])")).sendKeys(getCellData(rowTC, intMatkhau));
            driver.findElement(By.xpath("(//input[@value='Lưu'])")).click();
    }

}



