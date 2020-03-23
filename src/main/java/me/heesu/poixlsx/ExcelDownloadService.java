package me.heesu.poixlsx;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.stereotype.Service;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import javax.servlet.http.HttpServletResponse;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

@Service
public class ExcelDownloadService {
    private static String FILE_PATH = "/Users/heesu/tmpDir/";

    // 엑셀 read
    public void createXlsxExcelFileBySax(HttpServletResponse response) throws Exception {
        String targetFileName = "test_xlsx_small.xlsx";
        File file = new File(FILE_PATH + targetFileName);

        if (!file.exists()) {
            throw new Exception("파일이 존재하지 않습니다.");
        }

        OPCPackage pkg = OPCPackage.open(file);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);

        //첫번째 시트에 대한 처리
        InputStream sheet = r.getSheetsData().next();
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();

        List<String[]> res = new LinkedList<String[]>();
        res = SheetHandler.getRowCache();

        createXlsxFile(res, response);

    }

    // reader 설정
    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    // 엑셀 생성 로직
    private void createXlsxFile (List<String[]> resultList, HttpServletResponse response) {
        try {
            String fileName = "OOM_TEST";
            SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(100); // 메모리에 들고있는 최대 row수 설정
            Sheet sheet = sxssfWorkbook.createSheet(); // Sheet
            Row row = null; // Row
            Cell cell = null; // Cell

            if (resultList != null && !resultList.isEmpty()) {
                row = sheet.createRow(0);
                String[] headerResult = resultList.get(0);
                int totalRowNum = (resultList != null) ? resultList.size() : 0;
                int totalColNum = (headerResult != null) ? headerResult.length : 0 ;
                // 첫 row - 컬럼명, 따로 style추가 가능
                for (int i = 0;i < totalColNum ; i++) {
                    cell = row.createCell(i);
                    cell.setCellValue(headerResult[i]);
                }

                // 그 이후 row
                for (int i = 1; i < totalRowNum; i++) {
                    row = sheet.createRow(i);
                    String[] result = resultList.get(i);
                    for (int j = 0; j< totalColNum; j++) {
                        String cellVal = result[j];
                        cell = row.createCell(j);
                        if (cellVal != null) {
                            cell.setCellValue(cellVal);
                        } else {
                            cell.setCellValue("");
                        }
                    }
                    //System.out.println(i);
                }
            }

            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            sxssfWorkbook.write(response.getOutputStream());
            ((SXSSFWorkbook)sxssfWorkbook).dispose();

        }catch(Exception e) {
            //: TODO
            e.printStackTrace();
        }
    }

    private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean inlineStr;

        private static final String ROW_EVENT = "row";
        private static final String CELL_EVENT = "c";

        private static List<String> cellCache = new LinkedList<String>();
        private static List<String[]> rowCache = new LinkedList<String[]>();

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            if(CELL_EVENT.equals(name)) {
                String cellType = attributes.getValue("t");

                nextIsString = (cellType != null && cellType.equals("s"));
                inlineStr = (cellType != null && cellType.equals("inlineStr"));

            }else if(ROW_EVENT.equals(name)) {
                if(!cellCache.isEmpty()) {
                    rowCache.add(cellCache.toArray(new String[cellCache.size()]));
                }
                cellCache.clear();
            }

            // Clear contents cache
            lastContents = "";
        }
        @Override
        public void endElement(String uri, String localName, String name)
                throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();  // sst.getItemAt(idx).getString();
                nextIsString = false;
            }
            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v") || (inlineStr && name.equals("c"))) {
                //System.out.println(lastContents);
                cellCache.add(lastContents);
            }
        }
        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }

        @Override
        public void endDocument() throws SAXException {
            //위에 rowCache에 마지막 row event는 호출되지 않으므로 파싱이 종료될때 현재 cellCache에 있는 값을 rowCache에 넣어준다
            if(!cellCache.isEmpty()) {
                rowCache.add(cellCache.toArray(new String[cellCache.size()]));
            }

            System.out.println("######################END TO READ EXCEL DOCUMENT");
        }

        public static List<String[]> getRowCache() {
            return rowCache;
        }
    }
}
