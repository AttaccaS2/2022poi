package poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class Excel2 {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook workBook = new HSSFWorkbook(); 
		
        CellStyle defaultStyle = workBook.createCellStyle();
		
        // 테두리 설정
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);
		
        // 줄 바꿈 및 중앙 정렬
        defaultStyle.setWrapText(true);
        defaultStyle.setAlignment(HorizontalAlignment.CENTER);
        defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
        // 시트 생성 및 셀 높이 설정
        HSSFSheet sheet = workBook.createSheet();
        sheet.setDefaultRowHeightInPoints(30);
		
        for (int i=0; i<10; i++) {
            Row row = sheet.createRow(i);
			
            for (int j=0; j<5; j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(defaultStyle);
                cell.setCellValue("셀 생성 (" + i + ", " + j + ")");
				
                sheet.setColumnWidth(j, 3000);
            }
        }
		
        try {
            File xlsFile = new File("test.xls"); //현재 내 프로젝트 안에서 만들어짐
            FileOutputStream fileOut = new FileOutputStream(xlsFile);
            workBook.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            workBook.close();
        }
    }
}