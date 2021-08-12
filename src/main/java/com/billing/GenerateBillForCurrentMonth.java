package com.billing;

import com.spire.xls.Workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.threeten.extra.Temporals;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;

import lombok.SneakyThrows;

import static com.billing.BigDecimalUtils.format;
import static com.google.common.base.Strings.isNullOrEmpty;
import static com.spire.xls.FileFormat.PDF;
import static java.lang.Integer.parseInt;
import static java.math.BigDecimal.ONE;
import static java.math.BigDecimal.valueOf;
import static java.nio.file.StandardOpenOption.WRITE;
import static java.time.LocalDate.now;
import static java.time.format.DateTimeFormatter.ofPattern;
import static java.util.Map.entry;
import static java.util.concurrent.TimeUnit.SECONDS;
import static org.apache.poi.ss.usermodel.BorderStyle.THIN;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;
import static org.apache.poi.ss.usermodel.IndexedColors.LIGHT_CORNFLOWER_BLUE;

public class GenerateBillForCurrentMonth {

    public static final String RAPORT_DE_ACTIVITATE = "Raport_de_activitate.xls";

    public static String bilXlsxPath = "/Users/ionut/Documents/%s/Facturi/Arnia/Factura.xlsx";

    public static String generatedBillsPath = "/Users/ionut/Documents/%s/Facturi/Arnia/Facturi_generate";

    public static String generatedReportsPath = "/Users/ionut/Documents/%s/Facturi/Arnia/Rapoarte_generate";

    public static String GLOBAL_NUMBER_FILE = "bills_current_number_%s.txt";

    public static final BigDecimal DAILY_RATE = BigDecimal.valueOf(300);

    private static WebDriver driver;

    private static ChromeOptions options = new ChromeOptions();

    private static Scanner scanner = new Scanner(System.in);


    public static final Map<Integer, String> months = Map.ofEntries(
            entry(1, "Ianuarie"),
            entry(2, "Februarie"),
            entry(3, "Martie"),
            entry(4, "Aprilie"),
            entry(5, "Mai"),
            entry(6, "Iunie"),
            entry(7, "Iulie"),
            entry(8, "August"),
            entry(9, "Septembrie"),
            entry(10, "Octombrie"),
            entry(11, "Noiembrie"),
            entry(12, "Decembrie")
    );

    public static void main(String... args) throws Exception {
        System.out.println("Welcome to the most advanced billing generator!\n" +
                "Please enter the billing option or enter for default PFA: (PFA/SRL)");
        var rawOption = scanner.nextLine().trim();
        var billingOption = isNullOrEmpty(rawOption) ? "PFA" : rawOption;

        System.out.println("Please enter the number of working days:");
        var workingDays = scanner.nextInt();

        bilXlsxPath = String.format(bilXlsxPath, billingOption);
        generatedBillsPath = String.format(generatedBillsPath, billingOption);
        generatedReportsPath = String.format(generatedReportsPath, billingOption);
        GLOBAL_NUMBER_FILE = String.format(GLOBAL_NUMBER_FILE, billingOption);

        int billNumber = generateBill(workingDays);

        saveBillAsPdf(billNumber);
        generateReport(workingDays, billNumber);
        saveReportAsPdf();
    }

    private static void saveReportAsPdf() throws IOException {
        var path = Paths.get(generatedReportsPath).resolve(now().getYear() + "");
        if (!Files.exists(path)) {
            Files.createDirectory(path);
        }
        Workbook workbook = new Workbook();
        workbook.loadFromFile(path.getParent().getParent().resolve(RAPORT_DE_ACTIVITATE).toString());

        //Fit to page
        workbook.getConverterSetting().setSheetFitToPage(true);

        //Save as PDF document
        workbook.saveToFile(path.resolve("Raport-" + months.get(now().getMonthValue()) + "-" + now().getYear() + ".pdf").toString(), PDF);
        var desktop = Desktop.getDesktop();
        desktop.open(new File(path.toString()));
        desktop.open(path.toFile());
    }

    private static int generateBill(int workingDays) throws IOException {
        var inputStream = new FileInputStream(bilXlsxPath);
        var workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        var billNumber = updateBillNumber(sheet);
        updateCurrentDate(sheet);
        var currentExchangeRate = updateCurrentExchangeRate(sheet);
        updateAmounts(sheet, workingDays, currentExchangeRate);
        inputStream.close();

        FileOutputStream output = new FileOutputStream(bilXlsxPath);
        workbook.write(output);
        output.close();

        return billNumber;
    }

    private static void generateReport(int workingDays, int billNumber) throws Exception {
        shiftRows(workingDays);
        var raportPath = Paths.get(bilXlsxPath).getParent().resolve("Raport_de_activitate.xls").toFile();
        var inputStream = new FileInputStream(raportPath);
        var workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        updateBillNumber(billNumber, sheet);

        for (int i = findFooterRowIndex(sheet) - 1; i >= 8; i--) {
            Row row = sheet.createRow(i);
            for (int j = 0; j <= 2; j++) {
                var cell = row.createCell(j, CellType.BLANK);
                cell.setCellStyle(getCellStyle(workbook));
            }
            row.getCell(1).setCellValue("Servicii de consultanta IT");
            row.getCell(2).setCellValue(format(ONE, 1));
        }
        updateDates(workingDays, sheet);
        updateTotalDays(findFooterRowIndex(sheet), sheet, workingDays);
        inputStream.close();

        FileOutputStream output = new FileOutputStream(raportPath);
        workbook.write(output);
        output.close();
    }

    private static void updateTotalDays(int footerRowIndex, Sheet sheet, int workingDays) {
        var totalCell = sheet.getRow(footerRowIndex).getCell(2);
        totalCell.setCellValue(format(valueOf(workingDays), 1));
    }

    private static void updateDates(int workingDays, Sheet sheet) {
        var firstOfMonth = now().with(TemporalAdjusters.firstDayOfMonth());
        var nextWorkingDay = firstOfMonth.with(Temporals.nextWorkingDayOrSame());
        var workingDates = new ArrayList<LocalDate>();

        while (nextWorkingDay.getMonthValue() == now().getMonthValue()) {
            workingDates.add(nextWorkingDay);
            nextWorkingDay = nextWorkingDay.with(Temporals.nextWorkingDay());
        }

        if (workingDays == workingDates.size()) {
            for (int i = 8; i <= workingDays + 7; i++) {
                sheet.getRow(i).getCell(0).setCellValue(workingDates.get(i - 8).format(ofPattern("M.d.yyyy")));
            }
        } else {
            throw new RuntimeException("WrongNumber for working days");
        }
    }

    private static void updateBillNumber(int billNumber, Sheet sheet) {
        Cell cell = sheet.getRow(4).getCell(0);
        cell.setCellValue("Referitor la factura numarul: " + billNumber + "/" + now().format(ofPattern("dd.MM.yyyy")));
    }

    private static void shiftRows(int workingDays) throws IOException {
        var raportPath = Paths.get(bilXlsxPath).getParent().resolve("Raport_de_activitate.xls").toFile();

        var inputStream = new FileInputStream(raportPath);
        var workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int headerRowIndex = 7;
        int footerRowIndex = findFooterRowIndex(sheet);
        if (headerRowIndex + workingDays + 1 != footerRowIndex) {
            sheet.shiftRows(footerRowIndex, footerRowIndex + 7,
                    (headerRowIndex + workingDays + 1) - footerRowIndex, true, false);
        }
        correctBug(sheet);
        inputStream.close();

        FileOutputStream output = new FileOutputStream(raportPath);
        workbook.write(output);
        output.close();

    }

    private static CellStyle getCellStyle(org.apache.poi.ss.usermodel.Workbook workbook) {
        CellStyle backgroundStyle = workbook.createCellStyle();

        backgroundStyle.setFillForegroundColor(LIGHT_CORNFLOWER_BLUE.getIndex());
        backgroundStyle.setFillPattern(SOLID_FOREGROUND);

        backgroundStyle.setBorderBottom(THIN);
        backgroundStyle.setBottomBorderColor(BLACK.getIndex());
        backgroundStyle.setBorderLeft(THIN);
        backgroundStyle.setLeftBorderColor(BLACK.getIndex());
        backgroundStyle.setBorderRight(THIN);
        backgroundStyle.setRightBorderColor(BLACK.getIndex());
        backgroundStyle.setBorderTop(THIN);
        backgroundStyle.setTopBorderColor(BLACK.getIndex());
        backgroundStyle.setAlignment(HorizontalAlignment.CENTER);
        backgroundStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return backgroundStyle;
    }

    private static int findFooterRowIndex(Sheet sheet) {
        for (int i = 8; i < 50; i++) {
            var total = false;
            try {
                total = sheet.getRow(i).getCell(0).getStringCellValue().trim().equalsIgnoreCase("total");
            } catch (Exception e) {
            }

            if (total) {
                return i;
            }
        }
        throw new RuntimeException("I couldn't find the footer row");
    }

    private static void saveBillAsPdf(int billNumber) throws Exception {
        var today = now().format(ofPattern("dd-MM-yyyy"));
        var path = Paths.get(generatedBillsPath).resolve("Factura_Arnia_nr" + billNumber + "_din_" + today + ".pdf");

        Workbook workbook = new Workbook();
        workbook.loadFromFile(bilXlsxPath);

        //Fit to page
        workbook.getConverterSetting().setSheetFitToPage(true);

        //Save as PDF document
        workbook.saveToFile(path.toString(), PDF);
        var desktop = Desktop.getDesktop();
        desktop.open(new File(generatedBillsPath));
        desktop.open(path.toFile());
    }

    private static int updateBillNumber(XSSFSheet sheet) {
        Cell cell = sheet.getRow(1).getCell(4); //E2
        var currentBillNumber = getNextAvailableNumberFromGlobaBillingFile();
        cell.setCellValue(currentBillNumber);
        return currentBillNumber;
    }

    private static void updateCurrentDate(XSSFSheet sheet) {
        Cell dateCell = sheet.getRow(8).getCell(4);//E9
        Cell detailActivityCell = sheet.getRow(14).getCell(1);//B15
        var current = detailActivityCell.getStringCellValue();
        detailActivityCell.setCellValue(current.substring(0, current.indexOf("lunii") + 6) + months.get(now().getMonthValue()));
        dateCell.setCellValue(now().format(ofPattern("dd-MMM-yyyy")));
    }

    private static BigDecimal updateCurrentExchangeRate(XSSFSheet sheet) {
        Cell exchangeCell = sheet.getRow(15).getCell(1);//B16
        var currentExchange = getCurrentExchange();
        exchangeCell.setCellValue(format(new BigDecimal(currentExchange), 4));
        driver.close();
        return new BigDecimal(currentExchange);

    }

    private static void updateAmounts(XSSFSheet sheet, int workingDays, BigDecimal currentExchangeRate) {
        Cell quantityCell = sheet.getRow(14).getCell(4);//E15
        Cell unitaryPriceCell = sheet.getRow(14).getCell(5);//F15
        Cell monthPriceCell = sheet.getRow(14).getCell(6);//G15
        Cell totalPriceCell = sheet.getRow(17).getCell(6);//G18

        var unitaryPrice = currentExchangeRate.multiply(DAILY_RATE);
        var monthPrice = unitaryPrice.multiply(BigDecimal.valueOf(workingDays));

        quantityCell.setCellValue(workingDays);
        unitaryPriceCell.setCellValue(format(unitaryPrice, 2));
        monthPriceCell.setCellValue(format(monthPrice, 2));
        totalPriceCell.setCellValue(format(monthPrice, 2));
    }

    private static String getCurrentExchange() {
        navigate("https://www.cursbnr.ro");

        return driver.findElement(By.xpath("//*[@id=\"table-currencies\"]/tbody"))
                     .findElements(By.tagName("tr"))
                     .stream()
                     .filter(element -> element.getText().contains("EUR"))
                     //ultimul update de curs
                     .map(element -> element.findElement(By.xpath(".//td[3]")).getText())
                     .peek(exchange -> System.out.println("Cursul de schimb de azi este: " + exchange))
                     .findFirst()
                     .orElseThrow(() -> new NoSuchElementException("I couldn't find the current exchange rate"));
    }

    private static int getNextAvailableNumberFromGlobaBillingFile() {
        try {
            var parent = Paths.get(bilXlsxPath).getParent().getParent().resolve(GLOBAL_NUMBER_FILE);
            var currentNumberStr = Files.readString(parent);
            var updatedNumber = parseInt(currentNumberStr.trim()) + 1;
            Files.writeString(parent, String.valueOf(updatedNumber), WRITE);
            return updatedNumber;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @SneakyThrows
    private static void navigate(String url) {
        System.setProperty("webdriver.chrome.silentOutput", "true");
        Logger.getLogger("org.openqa.selenium").setLevel(Level.OFF);
//        options.addArguments("--headless");
        options.addArguments("--log-level=4");
        options.addArguments("--silent");
        driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(3, SECONDS);

        JavascriptExecutor js = (JavascriptExecutor) driver;
        driver.get(url);
        driver.manage().window().maximize();
        js.executeScript("window.scrollBy(0,1000)");

        Thread.sleep(4000);
    }


    private static void correctBug(Sheet sheet) {
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xSSFSheet = (XSSFSheet) sheet;
            // correcting bug that shiftRows does not adjusting references of the cells
            // if row 3 is shifted down, then reference in the cells remain r="A3", r="B3", ...
            // they must be adjusted to the new row thoug: r="A4", r="B4", ...
            // apache poi 3.17 has done this properly but had have other bugs in shiftRows.
            for (int r = xSSFSheet.getFirstRowNum(); r < sheet.getLastRowNum() + 1; r++) {
                XSSFRow row = xSSFSheet.getRow(r);
                if (row != null) {
                    long rRef = row.getCTRow().getR();
                    for (Cell cell : row) {
                        String cRef = ((XSSFCell) cell).getCTCell().getR();
                        ((XSSFCell) cell).getCTCell().setR(cRef.replaceAll("[0-9]", "") + rRef);
                    }
                }
            }
            // end correcting bug
        }
    }
}
