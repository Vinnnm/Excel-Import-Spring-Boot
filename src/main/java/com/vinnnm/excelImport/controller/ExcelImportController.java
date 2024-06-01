package com.vinnnm.excelImport.controller;

import com.vinnnm.excelImport.dto.ExcelImportDto;
import com.vinnnm.excelImport.service.ExcelService;
import lombok.AllArgsConstructor;
import org.apache.poi.EmptyFileException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;

@Controller
@AllArgsConstructor
public class ExcelImportController {

    private final ExcelService excelService;

    @GetMapping("/")
    public String importExcel(ModelMap model){
        model.addAttribute("excelImportDto", new ExcelImportDto());
        return "importExcel";
    }
    @PostMapping("/importExcel")
    public String uploadFile(@RequestParam("file") MultipartFile file, @RequestParam("sheetName") String sheetName,
                             ModelMap model) {
        try {
            InputStream inputStream = file.getInputStream();

            if (file.isEmpty()) {
                model.addAttribute("message", "Uploaded file is empty.");
                model.addAttribute("dto", new ExcelImportDto());
                return "importExcel";
            }

            if (sheetName == null || sheetName.trim().isEmpty()) {
                model.addAttribute("message", "Sheet name cannot be null or empty.");
                model.addAttribute("dto", new ExcelImportDto());
                return "importExcel";
            }

            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                model.addAttribute("message", "No sheet found with the name: " + sheetName);
                model.addAttribute("dto", new ExcelImportDto());
                return "hr/importExcel";
            }

            excelService.readExcelAndInsertIntoDatabase(inputStream, sheetName, workbook);
        } catch (IOException | SQLException | EmptyFileException e) {
            if (e instanceof EmptyFileException) {
                model.addAttribute("message", "Uploaded file is empty.");
            } else {
                e.printStackTrace();
                model.addAttribute("message", "Error uploading file: " + e.getMessage());
            }
        }
        return "redirect:/dashboard";
    }
}
