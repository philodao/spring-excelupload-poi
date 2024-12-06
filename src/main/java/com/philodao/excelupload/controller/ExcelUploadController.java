package com.philodao.excelupload.controller;

import com.philodao.excelupload.dto.ExcelData;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

        // JSON 응답 객체 생성
        JSONObject jsonResponse = new JSONObject();

        List<Map<String,String>> dataList = new ArrayList<>();

        Sheet worksheet = getRows(file);

        //행의수
        System.out.println("success2"+worksheet.getPhysicalNumberOfRows());

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
