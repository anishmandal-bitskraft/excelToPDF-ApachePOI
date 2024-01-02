import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class MainApp {
    public static void main(String[] args) throws IOException, DocumentException {
//        for excel file data
        FileInputStream fileInputStream =  new FileInputStream(new File("C:\\Users\\Anish\\Documents\\test.xlsx"));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet =  workbook.getSheetAt(0);
//        getRowData(sheet);
//        header data
        List<String> headerList = setHeader(sheet);

//        row data
        for(int i = 1; i< sheet.getPhysicalNumberOfRows(); i++) {
            //            pdf creation
//            open document
            Document document = new Document();
//            provide path
            String fileName = "C:\\Users\\Anish\\Documents\\pdfs\\pdf_"+ i + ".pdf";
            PdfWriter.getInstance(document, new FileOutputStream(fileName));
//            open document
            document.open();

//            create table
            PdfPTable table  = new PdfPTable(sheet.getRow(0).getPhysicalNumberOfCells());

//            add header data to PDF
            addPDFData(true, headerList, table);
            List<String> rowData = getRow(i, sheet);
            if(rowData.isEmpty()){
                document.add(table);
                document.close();
                Files.delete(Paths.get(fileName));
                continue;
            }
            addPDFData(false, rowData, table);
            document.add(table);
            document.close();
        }
    }

    public static List<String> setHeader(Sheet sheet){
    return getRow(0, sheet);
    }
    public static List<String> getRow(int index, Sheet sheet){
        List<String> list = new ArrayList<>();
            for(Cell cell : sheet.getRow(index)){
//                Iterating each cell
                switch(cell.getCellType()){
                    case STRING:
                        list.add(cell.getStringCellValue());
//                        System.out.println(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        list.add(String.valueOf(cell.getNumericCellValue()));
//                        System.out.println(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        list.add(String.valueOf(cell.getBooleanCellValue()));
//                        System.out.println(cell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        list.add(cell.getCellFormula().toString());
//                        System.out.println(cell.getCellFormula());
                        break;
                }
        }
            return list;
    }

    public static void addPDFData(boolean isHeader, List<String> list, PdfPTable table){
        list.stream().forEach(column -> {
            PdfPCell row = new PdfPCell();
            if(isHeader){
                row.setBackgroundColor(BaseColor.LIGHT_GRAY);
                row.setBorderWidth(2);
            }
            row.setPhrase(new Phrase(column));
            table.addCell(row);
        });
    }
}
