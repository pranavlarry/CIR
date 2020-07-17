import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ScrapingCheck {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("D:\\Auki\\Java\\CIRMigration\\cir2-internal_html1.xlsx"));
        XSSFSheet sheet = workbook.getSheetAt(1);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell link = row.getCell(0);
            Cell scrapedData = row.getCell(8);
            if(scrapedData.getStringCellValue().equalsIgnoreCase("") || scrapedData.getStringCellValue() == null) {
                System.out.println(link);
            }
        }
    }
}
