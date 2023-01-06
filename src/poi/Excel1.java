package poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel1 {

	public static void main(String[] args) {
		//.xlsx 확장자 지원
		try {
			int rowNo = 0 ; //행의 갯수
			XSSFWorkbook xssfWb = new XSSFWorkbook(); //객체 생성

			XSSFSheet xssfSheet = xssfWb.createSheet("워크 시트1"); // 워크시트 이름 설정

			// 폰트 스타일
			XSSFFont font = xssfWb.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL); // 폰트 스타일
			font.setFontHeightInPoints((short)20); // 폰트 크기
			font.setBold(true); // Bold 설정
			//font.setColor(new XSSFColor(Color.decode("#457ba2"))); //비추천 줄그어짐
			font.setColor(IndexedColors.BLUE.getIndex());// 폰트 색 지정
			//테이블 셀 스타일
			CellStyle cellStyle = xssfWb.createCellStyle();
			xssfSheet.setColumnWidth(0, (xssfSheet.getColumnWidth(0))+(short)2048); // 0번째 컬럼 넓이 조절

			cellStyle.setFont(font); // cellStyle에 font를 적용
			cellStyle.setAlignment(HorizontalAlignment.CENTER); // 정렬

			//셀병합
			xssfSheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 4)); //첫행, 마지막행, 첫열, 마지막열 병합

			// 타이틀 생성
			XSSFRow xssfRow= xssfSheet.createRow(rowNo++); // 행 객체 추가
			XSSFCell xssfCell = xssfRow.createCell((short) 0); // 추가한 행에 셀 객체 추가
			xssfCell.setCellStyle(cellStyle); // 셀에 스타일 지정
			xssfCell.setCellValue("타이틀 입니다"); // 데이터 입력

			xssfSheet.createRow(rowNo++);
			xssfRow = xssfSheet.createRow(rowNo++);  // 빈행 추가

			//테이블 스타일 설정
			CellStyle tableCellStyle = xssfWb.createCellStyle();
			tableCellStyle.setBorderTop(BorderStyle.MEDIUM);    // 테두리 위쪽
			tableCellStyle.setBorderBottom(BorderStyle.THIN); // 테두리 아래쪽
			tableCellStyle.setBorderLeft(BorderStyle.THIN);   // 테두리 왼쪽
			tableCellStyle.setBorderRight(BorderStyle.THIN);  // 테두리 오른쪽

			xssfRow = xssfSheet.createRow(rowNo++);
			xssfCell = xssfRow.createCell((short) 0);
			xssfCell.setCellStyle(tableCellStyle);
			xssfCell.setCellValue("셀1");
			xssfCell = xssfRow.createCell((short) 1);
			xssfCell.setCellStyle(tableCellStyle);
			xssfCell.setCellValue("셀2");
			xssfCell = xssfRow.createCell((short) 2);
			xssfCell.setCellStyle(tableCellStyle);
			xssfCell.setCellValue("셀3");
			xssfCell = xssfRow.createCell((short) 3);
			xssfCell.setCellStyle(tableCellStyle);
			xssfCell.setCellValue("셀4");
			xssfCell = xssfRow.createCell((short) 4);
			xssfCell.setCellStyle(tableCellStyle);

			String localFile = "excelTest.xlsx";

			File file = new File(localFile);
			FileOutputStream fos = null;

			fos = new FileOutputStream(file);

			xssfWb.write(fos);

			if (fos != null) fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}


	}
}
