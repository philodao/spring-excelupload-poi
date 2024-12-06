package com.philodao.excelupload.controller;

import com.philodao.excelupload.dto.ExcelData;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItemFactory;
import org.json.JSONObject;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class ExcelUploadController {

    @GetMapping("/excel/upload")
    public String excelUpload(Model model) {

        return "excelUpload";
    }

    @ResponseBody
    @PostMapping("/excel/uploadExcel")
    public void uploadExcel(@RequestParam(value="excelFile",required=false) MultipartFile file, Model model, HttpServletRequest request, HttpServletResponse response) throws Exception {
        System.out.println("success1");

        response.setContentType("application/json");
        response.setCharacterEncoding("utf-8");

        int rowindex = 0;
        int columnindex = 0;
        // JSON 응답 객체 생성
        JSONObject jsonResponse = new JSONObject();

        List<Map<String,String>> dataList = new ArrayList<>();

        Sheet worksheet = getRows(file);

        Sheet sheet = worksheet.getWorkbook().getSheetAt(0);

        //마지막 행
        int rows = (sheet.getLastRowNum()+1);
        //마지막 셀
        int maxCells = 0;
        for (rowindex = 0; rowindex < rows; rowindex++) { // 세로
            System.out.println("rows"+rows);
            Row row = sheet.getRow(rowindex);
            if(row!=null) {
                int cells = (row.getLastCellNum());
                if (cells > maxCells)
                    maxCells = cells;
            }

        }
        jsonResponse.put("maxRow",rows);
        jsonResponse.put("maxCol",maxCells);

        //병합된 셀 기록
        String[][] merge = new String[rows][maxCells];
        for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
            CellRangeAddress range = sheet.getMergedRegion(i);

            int mergeRow = range.getFirstRow(); 	//병합 셀의 시작Row
            int mergeCol = range.getFirstColumn();	//병합 셀의 시작Col
            int rowLength = range.getLastRow() - range.getFirstRow() + 1; 		//병합 셀의 Row 길이 계산
            int colLength = range.getLastColumn() - range.getFirstColumn() + 1;	//병합 셀의 Col 길이 계산

            //merge[][] 에 병합된 셀의 정보 기록
            for (int r = 0; r < rowLength; r++) {
                for (int c = 0; c < colLength; c++) {

                    if (r == 0 && c == 0) {//병합된 셀의 시작부분은 [Row, Col] 형태로 길이 제정
                        merge[mergeRow][mergeCol] = rowLength + "," + colLength;
                    } else { //이외의 부분은 mergeCell로 표시
                        merge[mergeRow + r][mergeCol + c] = "mergeCell";
                    }

                }
            }
        }

        //셀의 내용 저장
        String[][] text = new String[rows][maxCells];
        for (rowindex = 0; rowindex < rows; rowindex++) { // Row

            Row row = sheet.getRow(rowindex);
            if (row != null) {
                int cells = row.getLastCellNum();
                for (columnindex = 0; columnindex <= cells; columnindex++) { // Col

                    Cell cell = row.getCell(columnindex);

                    String value = "";
                    // 셀이 빈값일경우를 위한 널체크
                    if (cell == null) {
                        continue;
                    } else {
                        // 타입별로 내용 조회
                        switch (cell.getCellType()) {
                            case FORMULA:
                                value = cell.getCellFormula();
                                break;
                            case NUMERIC:
                                value = cell.getNumericCellValue() + "";
                                break;
                            case STRING:
                                value = cell.getStringCellValue() + "";
                                break;
                            case BLANK:
                                value = cell.getBooleanCellValue() + "";
                                break;
                            case ERROR:
                                value = cell.getErrorCellValue() + "";
                                break;
                        }
                    }
                    //내용 저장
                    text[rowindex][columnindex] = value;
                }

            }
        }

        jsonResponse.put("merge", merge); 	//병합정보
        jsonResponse.put("text", text);	//내용정보

        /*


        for (int i = 0; i < worksheet.getPhysicalNumberOfRows(); i++) {

            Row row = worksheet.getRow(i);

            System.out.println("success3"+row.getCell(0).getNumericCellValue());
            System.out.println("success3"+row.getCell(1).getNumericCellValue());
            System.out.println("success3"+row.getCell(2).getNumericCellValue());

            Map<String,String> data = new HashMap<>();

            data.put("data",String.valueOf(row.getCell(0).getNumericCellValue()));
            data.put("data1",String.valueOf(row.getCell(1).getNumericCellValue()));
            data.put("data2",String.valueOf(row.getCell(2).getNumericCellValue()));

            dataList.add(data);
        }

        jsonResponse.put("status", "success");
        jsonResponse.put("datas",dataList);
        */

        jsonResponse.put("status", "success");
        response.getWriter().write(jsonResponse.toString());
    }

    private static Sheet getRows(MultipartFile file) throws IOException {
        String extension = FilenameUtils.getExtension(file.getOriginalFilename()); // 3

        if (!extension.equals("xlsx") && !extension.equals("xls")) {
            throw new IOException("엑셀파일만 업로드 해주세요.");
        }

        Workbook workbook = null;

        if (extension.equals("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (extension.equals("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }

        Sheet worksheet = workbook.getSheetAt(0);
        return worksheet;
    }

    // ExcelDataMissingException은 엑셀 파일 처리 중 필요한 데이터가 누락되었을 때 발생하는 예외를 나타냄
    // 이 클래스는 RuntimeException을 확장하여, 런타임 중 발생할 수 있는 예외 상황을 표현합니다.
    public static class ExcelDataMissingException extends RuntimeException {

        // 생성자는 예외 발생 시 전달될 메시지를 인자로 받습니다.
        // @param message 예외 발생 시 표시될 메시지
        public ExcelDataMissingException(String message) {
            super(message); // 슈퍼클래스인 RuntimeException의 생성자를 호출하여 메시지를 전달합니다.
        }
    }






}
