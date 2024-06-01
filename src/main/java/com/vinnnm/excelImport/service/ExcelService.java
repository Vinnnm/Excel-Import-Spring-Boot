package com.vinnnm.excelImport.service;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.SQLException;

@Service
public interface ExcelService {
    void readExcelAndInsertIntoDatabase(InputStream inputStream, String sheetName, Workbook workbook) throws SQLException;
}
