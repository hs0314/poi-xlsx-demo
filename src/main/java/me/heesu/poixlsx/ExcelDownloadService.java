package me.heesu.poixlsx;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import javax.servlet.http.HttpServletResponse;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.InputStream;

@Service
public class ExcelDownloadService {

    // reader 설정
    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    // 엑셀 read
    private void processOneSheet(File file) throws Exception {
        //OPCPackage 파일 read, write 할 수 있는 컨테이너 생성
        OPCPackage pkg = OPCPackage.open(file);

       // XSSFWorkbook xssfWorkbook = new XSSFWorkbook(pkg);

        //reader를 통해서 적은 메모리로 sax parsing
        XSSFReader xssfReader = new XSSFReader( pkg );
        //StylesTable styles = xssfReader.getStylesTable();

        SharedStringsTable sst = xssfReader.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);
        // To look up the Sheet Name / Sheet Order / rID,
        //  you need to process the core Workbook stream.
        // Normally it's of the form rId# or rSheet
        try(InputStream sheet = xssfReader.getSheetsData().next()){
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            // sheet.close();    try-with-resources문 -> 자동으로 close시켜줌
        }
    }

    //엑셀 다운로드
    public void processExcelDownload(){ //HttpServletResponse response){
        String filename = "static/test.xlsx";
        try {
            //File file = new File(filename);
            File file = new ClassPathResource(filename).getFile();
            processOneSheet(file);

            SXSSFWorkbook wb = new SXSSFWorkbook(100);
            Sheet sheet = wb.createSheet("sheet1");
            Row
            response.setHeader("Set-Cookie", "fileDownload=true; path=/");
            response.setHeader("Content-Disposition", String.format("attachment; filename=\"test.xlsx\""));
            wb.write(response.getOutputStream());


        }catch(Exception e) {
            e.printStackTrace();
        }
        //wb.dispose
        //wb.close();
    }

}
