package MBR;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.JavascriptExecutor;

import java.util.Map;
import java.util.NoSuchElementException;
import java.util.HashMap;
import javax.swing.*;

import java.time.format.DateTimeFormatter;
import java.time.LocalDateTime;

//{@literal @RunWith(JUnit4.class)}
public class Mbr extends App {

        private static ChromeDriverService service;
        private WebDriver driver;

        // {@literal @BeforeClass}
        public static void createAndStartService() throws IOException {
                service = new ChromeDriverService.Builder()
                                .usingDriverExecutable(new File("Resources/chromedriver.exe")).usingAnyFreePort()
                                .build();
                service.start();
        }

        // {@literal @AfterClass}
        public static void createAndStopService() {
                service.stop();
        }

        // {@literal @Before}
        public void createDriver() {
                driver = new RemoteWebDriver(service.getUrl(), DesiredCapabilities.chrome());
        }

        // {@literal @After}
        public void quitDriver() {
                driver.quit();
        }

        Sheet sheet;
        Cell cell;
        String mp, datran, gl, typvnd, mancode, sl;

        public int ketData() throws IOException {

                FileInputStream finput = null;
                int k;

                finput = new FileInputStream(new File("MBRGenerator.xlsm"));
                Workbook workbook = WorkbookFactory.create(finput);

                sheet = workbook.getSheetAt(0);

                k = sheet.getLastRowNum();

                return k;
        }

        public void getValues(int j) {

                cell = sheet.getRow(j).getCell(0);
                mp = cell.getStringCellValue();
                cell = sheet.getRow(j).getCell(1);
                datran = cell.getStringCellValue();
                cell = sheet.getRow(j).getCell(2);
                gl = cell.getStringCellValue();
                cell = sheet.getRow(j).getCell(3);
                sl = cell.getStringCellValue();
                cell = sheet.getRow(j).getCell(4);
                typvnd = cell.getStringCellValue();
                cell = sheet.getRow(j).getCell(5);
                mancode = cell.getStringCellValue();
        }

        // {@literal @Test}
        public void mbrgen(int k) throws InterruptedException, IOException {
                WebDriverWait wait = new WebDriverWait(driver, 100);
                driver.get("https://aves-beet-eu.aka.amazon.com/businessReporting?tenantId=WinstonEU");
                Thread.sleep(3000);
                Map<String, String> map = new HashMap<>();
                map.put("ALL", "//li//div[@title='ALL']");
                map.put("PAN-EU", "//li//div[@title='Pan-EU']");
                map.put("NL", "//li//div[@title='NL']");
                map.put("IT", "//li//div[@title='IT']");
                map.put("UK", "//li//div[@title='UK']");
                map.put("FR", "//li//div[@title='FR']");
                map.put("ES", "//li//div[@title='ES']");
                map.put("DE", "//li//div[@title='DE']");
                map.put("TR", "//li//div[@title='TR']");
                map.put("Vendor Child", "//li/div[@title='Vendor Child']");
                map.put("Vendor Parent", "//li/div[@title='Vendor Parent']");
                map.put("Vendor Company", "//li/div[@title='Vendor Company']");
                map.put("Brand", "//li/div[@title='Brand']");
                map.put("Manufacturer", "//li/div[@title='Manufacturer']");
                map.put("Manufacturer Parent", "//li/div[@title='Manufacturer Parent']");
                map.put("Manufacturer Company", "//li/div[@title='Manufacturer Company']");

                // wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//label[@for='password_field']")));

                try {
                        // wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'Clear')]")));
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                                        "//label[contains(text(),'Marketplace(s)')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']")));
                } catch (NoSuchElementException | TimeoutException e) {
                        // e.notify();
                        Thread.sleep(60000);
                }

                driver.findElement(By.xpath(
                                "//label[contains(text(),'Marketplace(s)')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                .click();
                driver.findElement(By.xpath(map.get(mp))).click();
                try {
                        // wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'Clear')]")));
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                                        "//label[contains(text(),'Date Range')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']")));
                } catch (NoSuchElementException | TimeoutException e) {
                        // e.notify();
                        Thread.sleep(6000);
                }
                driver.findElement(By.xpath(
                                "//label[contains(text(),'Date Range')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                .click();
                driver.findElement(By.xpath("//li//div[@data-value='PREVIOUS_MONTH']")).click();
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                                "//div[@class='awsui-select-trigger-wrapper']//span//span[contains(text(),'Add/remove GLs')]")));
                driver.findElement(By.xpath(
                                "//div[@class='awsui-select-trigger-wrapper']//span//span[contains(text(),'Add/remove GLs')]"))
                                .click();
                driver.findElement(By.xpath("//input[@id='awsui-input-3']|//input[@placeholder='Search']"))
                                .sendKeys(gl);
                driver.findElement(By.xpath("//div[contains(text(),'" + gl
                                + "')]/ancestor::label[@class='awsui-checkbox']//input[@class='awsui-checkbox-native-input']"))
                                .click();

                if (typvnd.equals("Vendor Code")) {
                        if (k != 2) {
                                driver.findElement(By.xpath(
                                                "//label[contains(text(),'Type of vendor')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                                .click();
                                driver.findElement(By.xpath("//li/div[@title='Vendor Child']")).click();
                        } else {
                                // driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]")).click();
                                driver.findElement(By.xpath(
                                                "//label[contains(text(),'Type of vendor')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                                .click();
                                driver.findElement(By.xpath("//li/div[@title='Manufacturer']")).click();
                        }
                        // driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]")).click();
                } else if (typvnd.equals("Company Code")) {
                        if (k != 2) {
                                driver.findElement(By.xpath(
                                                "//label[contains(text(),'Type of vendor')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                                .click();
                                driver.findElement(By.xpath("//li/div[@title='Vendor Company']")).click();
                        } else {
                                // driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]")).click();
                                driver.findElement(By.xpath(
                                                "//label[contains(text(),'Type of vendor')]/ancestor::div[@class='awsui-form-field awsui-form-field-stretch']//span[@class='awsui-select-trigger-textbox']"))
                                                .click();
                                driver.findElement(By.xpath("//li/div[@title='Manufacturer Company']")).click();
                        }
                        // driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]")).click();
                } else {
                        JOptionPane.showMessageDialog(null, "No Type of Vendor value");

                }

                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[1]/div[2]/div[2]/awsui-form-field[1]/div[1]/div[1]/div[1]/div[1]/span[1]/div[1]/div[1]/input[1]"))
                                .sendKeys(mancode);
                // driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[1]/div[2]/div[2]/awsui-form-field[1]/div[1]/div[1]/div[1]/div[1]/span[1]/div[1]/div[1]/input[1]")).click();
                // click on dashboard
                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/ul[1]/li[1]/a[1]/span[1]"))
                                .click();
                Thread.sleep(3000);
                try {
                        wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
                                        "//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]/button[1]")));
                } catch (NoSuchElementException | TimeoutException e) {
                        e.printStackTrace();
                }
                // click on submit
                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[2]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/div[2]/div[1]/div[2]/div[1]/div[3]/awsui-button[1]/button[1]"))
                                .click();
                try {
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                                        "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/awsui-cards[1]/div[1]/ol[1]/li[1]/div[1]/div[2]/div[1]/span[1]/awsui-spinner[1]/div[1]/div[2]")));
                } catch (NoSuchElementException | TimeoutException e) {
                        e.printStackTrace();
                }
                JavascriptExecutor jse = (JavascriptExecutor) driver;
                jse.executeScript("window.scrollBy(0,900)", "");
                try {
                        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
                                        "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/awsui-cards[1]/div[1]/ol[1]/li[1]/div[1]/div[2]/div[1]/span[1]/div[1]/div[3]/div[1]/section[1]/div[1]/div[1]/div[3]/div[1]/awsui-spinner[1]/div[1]/div[2]")));
                } catch (NoSuchElementException | TimeoutException e) {
                        Thread.sleep(70000);
                }
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
                                "//span[.='Export']/ancestor::awsui-button[@class='awsui-util-mb-n']//button[@type='submit']")));
                // click export
                driver.findElement(By.xpath(
                                "//span[.='Export']/ancestor::awsui-button[@class='awsui-util-mb-n']//button[@type='submit']"))
                                .click();
                try {
                        wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
                                        "//span[contains(text(),'Both - Excel and PowerPoint')]/ancestor::label[@class='awsui-radio-button-label']//input[@class='awsui-radio-native-input']")));
                } catch (NoSuchElementException | TimeoutException e) {
                        Thread.sleep(6000);
                }
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Both - Excel and PowerPoint')]/ancestor::label[@class='awsui-radio-button-label']//input[@class='awsui-radio-native-input']"))
                                .click();

                if (sl.equals("softlines")) {
                        driver.findElement(By.xpath(
                                        "//div[contains(text(),'Teen')]/ancestor::label[@class='awsui-checkbox']//input[@class='awsui-checkbox-native-input']"))
                                        .click();

                } else {

                }
                driver.findElement(By.xpath(
                                "//body[1]/div[1]/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-field[2]/div[1]/div[1]/div[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[3]/awsui-column-layout[1]/div[1]/span[1]/span[1]/div[2]/awsui-checkbox[1]/label[1]/input[1]"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//body[1]/div[1]/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-field[2]/div[1]/div[1]/div[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[3]/awsui-column-layout[1]/div[1]/span[1]/span[1]/div[3]/awsui-checkbox[1]/label[1]/input[1]"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//div[@awsui-form-section-region='header' and .='Metrics']")).click();
                driver.findElement(By.xpath(
                                "//div[contains(text(),'Public')]/ancestor::label[@class='awsui-checkbox']//input[@class='awsui-checkbox-native-input']"))
                                .click();
                Thread.sleep(1000);
                /*
                 * driver.findElement(By.xpath(
                 * "//span[contains(text(),'SnS Revenue Penetration')]/ancestor::label//input[@class='awsui-toggle-native-input']"
                 * )) .click();
                 */
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'ASP')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'CCOGS as a % of PCOGS')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Net PPM%')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Net Receipts')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                try {
                        Thread.sleep(1000);
                        driver.findElement(By.xpath(
                                        "//span[contains(text(),'OPS')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                        .click();

                        Thread.sleep(1000);
                        driver.findElement(By.xpath(
                                        "//span[contains(text(),'Ordered Units')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                        .click();
                } catch (org.openqa.selenium.NoSuchElementException e) {
                        e.printStackTrace();
                }
                Thread.sleep(1000);
                List<WebElement> wb = driver.findElements(By.xpath(
                                "//span[contains(text(),'PPM')]/ancestor::label//input[@class='awsui-toggle-native-input']"));

                for (WebElement w : wb) {
                        Thread.sleep(1000);
                        w.click();
                }
                List<WebElement> wbb = driver.findElements(By.xpath(
                                "//span[contains(text(),'Revenue')]/ancestor::label//input[@class='awsui-toggle-native-input']"));

                for (WebElement w : wbb) {
                        Thread.sleep(1000);
                        w.click();
                }
                // PPM
                /*
                 * driver.findElement(By.xpath(
                 * "//body[1]/div[1]/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[2]/awsui-tabs[1]/div[1]/div[1]/div[1]/span[1]/div[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[1]/awsui-column-layout[1]/div[1]/span[1]/span[1]/div[10]/awsui-toggle[1]/label[1]/input[1]"
                 * )) .click(); Thread.sleep(1000); //Revenue driver.findElement(By.xpath(
                 * "//body[1]/div[1]/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[2]/awsui-tabs[1]/div[1]/div[1]/div[1]/span[1]/div[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[1]/awsui-column-layout[1]/div[1]/span[1]/span[1]/div[11]/awsui-toggle[1]/label[1]/input[1]"
                 * )) .click();
                 */
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Net PPM%')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                /*
                 * driver.findElement(By.xpath(
                 * "//span[contains(text(),'SnS Revenue Penetration')]/ancestor::label//input[@class='awsui-toggle-native-input']"
                 * )) .click();
                 */
                // Thread.sleep(1000);
                // Availability & Operational efficiency
                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[2]/awsui-tabs[1]/div[1]/ul[1]/li[3]/a[1]"))
                                .click();
                Thread.sleep(1000);
                try {
                        driver.findElement(By.xpath(
                                        "//span[contains(text(),'Rep OOS % [DEPRECATED]')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                        .click();
                        Thread.sleep(1000);
                        driver.findElement(By.xpath(
                                        "//span[contains(text(),'Retail Fast Track %')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                        .click();
                } catch (org.openqa.selenium.NoSuchElementException e) {
                        e.printStackTrace();
                }
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'TheyPay Vendor Time To Deliver (VTTD)')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Total Inventory Cost')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                driver.findElement(By.xpath(
                                "//span[contains(text(),'Total Inventory Units')]/ancestor::label//input[@class='awsui-toggle-native-input']"))
                                .click();
                Thread.sleep(1000);
                // Conversion
                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[2]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/awsui-form-section[1]/awsui-expandable-section[1]/div[1]/div[2]/span[1]/div[1]/span[1]/awsui-column-layout[1]/div[1]/span[1]/div[1]/div[2]/awsui-tabs[1]/div[1]/ul[1]/li[5]/a[1]"))
                                .click();
                Thread.sleep(1000);
                // Click on Export
                driver.findElement(By.xpath(
                                "//body/div[@id='root']/div[1]/div[3]/awsui-tabs[1]/div[1]/div[1]/span[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/awsui-modal[1]/div[2]/div[1]/div[1]/div[3]/span[1]/span[1]/awsui-button[2]/button[1]/span[1]"))
                                .click();
                Thread.sleep(5000);
                wait.until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath("//div[@class='awsui-spinner-circle awsui-spinner-circle-right']")));

                try {
                        wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
                                        "//span[.='Export']/ancestor::awsui-button[@class='awsui-util-mb-n']//button[@type='submit']")));
                } catch (NoSuchElementException | TimeoutException e) {
                        Thread.sleep(240000);
                }
                Thread.sleep(1000);
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy");
                LocalDateTime now = LocalDateTime.now();
                System.out.println(dtf.format(now));
                String username = System.getProperty("user.name");
                if (typvnd.equals("Vendor Code")) {
                        if (k != 2) {
                                File file1 = new File("C:\\Users\\" + username + "\\Downloads\\AVSBusinessData_"
                                                + dtf.format(now) + ".xlsx");
                                File file2 = new File("C:\\Users\\" + username + "\\Downloads\\AVS Report " + mancode
                                                + ".pptx");
                                File file3 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_vend.xlsx");
                                File file4 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_vend.pptx");
                                file1.renameTo(file3);
                                String f2 = file2.getAbsolutePath();
                                File f22 = new File(f2);
                                String f1 = file1.getAbsolutePath();
                                File f11 = new File(f1);
                                f22.renameTo(file4);
                                f11.renameTo(file3);
                        } else {
                                File file1 = new File("C:\\Users\\" + username + "\\Downloads\\AVSBusinessData_"
                                                + dtf.format(now) + ".xlsx");
                                File file2 = new File("C:\\Users\\" + username + "\\Downloads\\AVS Report " + mancode
                                                + ".pptx");
                                File file3 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_manu.xlsx");
                                File file4 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_manu.pptx");
                                file1.renameTo(file3);
                                String f2 = file2.getAbsolutePath();
                                File f22 = new File(f2);
                                String f1 = file1.getAbsolutePath();
                                File f11 = new File(f1);
                                f22.renameTo(file4);
                                f11.renameTo(file3);
                        }
                } else if (typvnd.equals("Company Code")) {
                        if (k != 2) {
                                File file1 = new File("C:\\Users\\" + username + "\\Downloads\\AVSBusinessData_"
                                                + dtf.format(now) + ".xlsx");
                                File file2 = new File("C:\\Users\\" + username + "\\Downloads\\AVS Report " + mancode
                                                + ".pptx");
                                File file3 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_vendc.xlsx");
                                File file4 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_vendc.pptx");

                                file1.renameTo(file3);
                                String f2 = file2.getAbsolutePath();
                                File f22 = new File(f2);
                                String f1 = file1.getAbsolutePath();
                                File f11 = new File(f1);
                                f22.renameTo(file4);
                                f11.renameTo(file3);
                        } else {
                                File file1 = new File("C:\\Users\\" + username + "\\Downloads\\AVSBusinessData_"
                                                + dtf.format(now) + ".xlsx");
                                File file2 = new File("C:\\Users\\" + username + "\\Downloads\\AVS Report " + mancode
                                                + ".pptx");
                                File file3 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_manuc.xlsx");
                                File file4 = new File("C:\\Users\\" + username + "\\Downloads\\" + mancode + "_" + mp
                                                + "_manuc.pptx");
                                file1.renameTo(file3);
                                String f2 = file2.getAbsolutePath();
                                File f22 = new File(f2);
                                String f1 = file1.getAbsolutePath();
                                File f11 = new File(f1);
                                f22.renameTo(file4);
                                f11.renameTo(file3);
                        }
                } else {
                }
                /*
                 * try { Runtime.getRuntime().exec( "tskill /A powerpnt"); } catch (IOException
                 * e) { System.out.println(e); System.exit(0); }
                 */
                String manu, vend, merged;
                if (k != 2) {
                } else {
                        if (typvnd.equals("Vendor Code")) {
                                manu = mancode + "_" + mp + "_manu";
                                vend = mancode + "_" + mp + "_vend";
                                merged = mancode + "_" + mp;
                                try {
                                        Runtime.getRuntime().exec(
                                                        "\"C:\\Program Files\\Microsoft Office\\Office16\\powerpnt\" /m MBR_Gen.pptm Workbook_Open /"
                                                                        + manu + "/" + vend + "/" + merged);
                                } catch (IOException e) {
                                        System.out.println(e);
                                        System.exit(0);
                                }
                        } else if (typvnd.equals("Company Code")) {
                                manu = mancode + "_" + mp + "_manuc";
                                vend = mancode + "_" + mp + "_vendc";
                                merged = mancode + "_" + mp;
                                try {
                                        Runtime.getRuntime().exec(
                                                        "\"C:\\Program Files\\Microsoft Office\\Office16\\powerpnt\" /m MBR_Gen.pptm Workbook_Open /"
                                                                        + manu + "/" + vend + "/" + merged);
                                } catch (IOException e) {
                                        System.out.println(e);
                                        System.exit(0);
                                }
                        } else {
                        }
                }
        }
}