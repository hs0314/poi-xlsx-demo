package me.heesu.poixlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;

@Service
public class ExcelCreateService {

    private static String FILE_PATH = "/Users/heesu/tmpDir";
    private static String XLSX_EXTENSION = ".xlsx";
    private static int COLUMN_SIZE = 20;

    /* 테스트용 대용량 엑셀 */
    public void createXlsxExcel(int rowSize){
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
        SXSSFSheet sheet = workbook.createSheet();
        String filename = "test_xlsx";

        for (int r=0; r <= rowSize; r++) {
            Row row = sheet.createRow(r);
            for (int c=0 ; c <= COLUMN_SIZE; c++) {
                Cell cell = row.createCell(c);
                cell.setCellValue("Cell " + r + "-" + c);
            }
        }

        //엑셀파일 세팅 후 파일 생성
        try {
            File file = new File(FILE_PATH);

            FileOutputStream fileOutputStream = new FileOutputStream(file + File.separator + filename + XLSX_EXTENSION);
            //생성한 엑셀파일을 outputStream 해줍니다.
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
