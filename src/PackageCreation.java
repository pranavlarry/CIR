import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.*;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPathExpressionException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class PackageCreation {

    public static void main(String[] args) throws ParserConfigurationException, IOException, SAXException, TransformerException {
        String rootFld = "D:\\Auki\\Java\\CIRMigration\\testing\\jcr_root\\content\\we-retail\\us\\en\\";
        String filter = "D:\\Auki\\Java\\CIRMigration\\testing\\META-INF\\vault\\filter.xml";

        FileInputStream excelSheet = new FileInputStream("D:\\Auki\\Java\\CIRMigration\\testingExcel.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(excelSheet);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        File xmlTemplate = new File("D:\\Auki\\Java\\CIRMigration\\template.xml");
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder db = dbf.newDocumentBuilder();
        Document doc = db.parse(xmlTemplate);
        doc.getDocumentElement().normalize();

        Document filterDoc = db.parse(filter);
        filterDoc.getDocumentElement().normalize();

        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);

            Element root = filterDoc.getDocumentElement();
            Element e1 = filterDoc.createElement("filter");
            e1.setAttribute("root","/content/we-retail/us/en/"+cell.getStringCellValue().toLowerCase());
//            filterDoc.appendChild(e1);
            root.appendChild(e1);

            NodeList list = doc.getElementsByTagName("jcr:content");
            Element e = (Element) list.item(0);


            e.setAttribute("jcr:title",cell.getStringCellValue());
            String finalFld = rootFld+cell.getStringCellValue().toLowerCase();
            File file = new File(finalFld);
            file.mkdir();
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.METHOD, "xml");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "5");
            DOMSource source = new DOMSource(doc);
            StreamResult result = new StreamResult(new File(finalFld+"\\.content.xml"));

            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.METHOD, "xml");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "5");
            DOMSource sourceFilter = new DOMSource(filterDoc);
            StreamResult resultFilter = new StreamResult(new File("D:\\Auki\\Java\\CIRMigration\\testing\\META-INF\\vault\\filter.xml"));
            transformer.transform(sourceFilter, resultFilter);

        }


//        System.out.println("Root element: " + doc.getDocumentElement().getNodeName());

//        System.out.println(e.getAttribute("jcr:primaryType"));


    }
}
