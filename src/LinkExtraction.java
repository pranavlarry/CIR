import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LinkExtraction {
    static int pdfCount = 0,nonPdfCount = 0;
    public static void main(String[] args) throws Exception {

        //input excel sheet
        String sheetLink = "D:\\Auki\\Java\\CIRMigration\\Weekly Announcements_Output_V1_latest.xlsx";
        XSSFWorkbook excel = new XSSFWorkbook(new FileInputStream(sheetLink));

        //Output
        XSSFWorkbook output = new XSSFWorkbook();


        Iterator<XSSFSheet> sheetIterator = excel.iterator();

        while (sheetIterator.hasNext()) {
            XSSFSheet sheet = sheetIterator.next();
            XSSFSheet outputSheet = output.createSheet(sheet.getSheetName());

            System.out.println(sheet.getSheetName());
            Iterator<Row> row = sheet.rowIterator();
            while (row.hasNext()) {
                Row currentRow = row.next();
                Row outputRow = outputSheet.createRow(currentRow.getRowNum());
                Iterator<Cell> cellIterator = currentRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //writing every cell to output sheet
                    switch (cell.getCellType()) {
                        case HSSFCell.CELL_TYPE_STRING:
                            outputRow.createCell(cell.getColumnIndex()).setCellValue(cell.getStringCellValue());

                            //if the cell contains scrapped data process the links
                            if(cell.getColumnIndex() == 4) {
                                try {
                                    String val = cell.getStringCellValue();
                                    Pattern pattern = Pattern.compile("href=\"(.*?)\"");
                                    Matcher matcher = pattern.matcher(cell.getStringCellValue());
                                    String pdfLinks = "", nonPdfLinks= "";
                                    while (matcher.find()) {
                                        String link = matcher.group(1);
                                        Pattern mail = Pattern.compile("mailto:(.*?).com");
                                        Matcher mailMatcher = mail.matcher(link);
                                        if(!mailMatcher.find()) {
                                            Pattern pdf = Pattern.compile(("(.*?)pdf|PDF"));
                                            Matcher pdfMatcher = pdf.matcher(link);

                                            Pattern nonPdf = Pattern.compile(("https://www.cir2.com|file://(.*?)"));
                                            Matcher nonPdfMatcher = nonPdf.matcher(link);

                                            if(pdfMatcher.find()) {
                                                pdfLinks = checkLink(nonPdfLinks,link,true);

                                            }

                                            else if(nonPdfMatcher.find()) {
                                                nonPdfLinks = checkLink(nonPdfLinks,link,false);
                                            }
                                        }
                                    }

                                    if(!pdfLinks.equals("")) {
                                        outputRow.createCell(7).setCellValue(pdfLinks);
                                    }
                                    if(!nonPdfLinks.equals("")) {
                                        outputRow.createCell(8).setCellValue(nonPdfLinks);
                                    }

                                }
                                catch (NullPointerException | IllegalStateException e) {

                                }
                            }
                            break;

                        case HSSFCell.CELL_TYPE_NUMERIC:
                            outputRow.createCell(cell.getColumnIndex()).setCellValue(cell.getNumericCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_BLANK:
                            outputRow.createCell(cell.getColumnIndex()).setCellValue("");
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN:
                            outputRow.createCell(cell.getColumnIndex()).setCellValue(cell.getBooleanCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_ERROR:
                            break;
                        default:
                            System.out.println("Shouldn't reach here!!!");
                            break;
                    }


                }
            }
        }


        //Output sheet, create an empty excel file on this path otherwise it will through error
        FileOutputStream out = new FileOutputStream(new File("D:\\Auki\\Java\\CIRMigration\\sample.xlsx"));
        output.write(out);
        out.close();
        System.out.println(pdfCount +"\n" + nonPdfCount);

    }

    //checking links for duplicates within a cell (single page)
    public static String checkLink(String proceesedLinks,String link,boolean pdf) {
        boolean flag = true;
        if(proceesedLinks.equals("")) {
            proceesedLinks = link;
            if (pdf) {
                pdfCount++;
            } else {
                nonPdfCount++;
            }
            System.out.println("PDF---->" + link);

        }
        else {
            String[] links = proceesedLinks.split("\n");
            for(int i=0;i<links.length;i++) {
                if(links[i].equals(link)) {
                    flag=false;
                }
            }
            if(flag) {
                proceesedLinks = proceesedLinks + "\n" + link;
                if (pdf) {
                    pdfCount++;
                } else {
                    nonPdfCount++;
                }
            }
        }

        return proceesedLinks;
    }
}