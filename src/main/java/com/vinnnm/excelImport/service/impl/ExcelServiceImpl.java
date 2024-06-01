package com.vinnnm.excelImport.service.impl;

import com.vinnnm.excelImport.models.*;
import com.vinnnm.excelImport.repo.*;
import com.vinnnm.excelImport.service.ExcelService;
import jakarta.persistence.EntityNotFoundException;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.sql.*;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {
    @Value("${spring.datasource.url}")
    private String dbUrl;

    @Value("${spring.datasource.username}")
    private String dbUsername;

    @Value("${spring.datasource.password}")
    private String dbPassword;

    private final DivisionRepository divisionRepository;
    private final DepartmentRepository departmentRepository;
    private final TeamRepository teamRepository;
    private final RoleRepository roleRepository;
    private final UserRepository userRepository;
    @Override
    public void readExcelAndInsertIntoDatabase(InputStream inputStream, String sheetName, Workbook workbook) throws SQLException {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet != null) {
            createTableFromSheet(sheet);
            insertDataIntoTable(sheet);
        }
        insertDataIntoUser(sheetName, workbook);
    }

    private void createTableFromSheet(Sheet sheet) throws SQLException {
        Row headerRow = sheet.getRow(3);
        List<String> columnNames = new ArrayList<>();
        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING) {
                String columnName = cell.getStringCellValue().replaceAll("[^a-zA-Z0-9]", "_");
                columnNames.add(columnName);
            } else if (cell.getCellType() == CellType.NUMERIC) {
                columnNames.add("COLUMN_" + cell.getColumnIndex());
            }
        }

        String tableName = sheet.getSheetName();
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE IF NOT EXISTS ")
                .append(tableName)
                .append(" (");
        for (String columnName : columnNames) {
            createTableQuery.append(columnName).append(" VARCHAR(255), ");
        }
        createTableQuery.setLength(createTableQuery.length() - 2);
        createTableQuery.append(")");

        try (Connection connection = DriverManager.getConnection(dbUrl, dbUsername, dbPassword);
             Statement statement = connection.createStatement()) {
            statement.executeUpdate(createTableQuery.toString());
        }
    }

    private void insertDataIntoTable(Sheet sheet) throws SQLException {
        String tableName = sheet.getSheetName();
        try (Connection connection = DriverManager.getConnection(dbUrl, dbUsername, dbPassword);
             Statement statement = connection.createStatement()) {
            for (Row row : sheet) {
                if (row.getRowNum() < 4) {
                    continue; // Skip header row
                }
                StringBuilder insertQuery = new StringBuilder("INSERT INTO ")
                        .append(tableName)
                        .append(" VALUES (");
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.NUMERIC) {
                        insertQuery.append(cell.getNumericCellValue()).append(", ");
                    } else {
                        insertQuery.append("'").append(cell.getStringCellValue()).append("', ");
                    }
                }
                insertQuery.setLength(insertQuery.length() - 2);
                insertQuery.append(")");

                statement.executeUpdate(insertQuery.toString());
            }
        }
    }

    public List<List<String>> getTableRows(String currentSheetName) {
        List<List<String>> rows = new ArrayList<>();
        String url = dbUrl;
        String username = dbUsername;
        String password = dbPassword;

        try (Connection connection = DriverManager.getConnection(url, username, password);
             Statement statement = connection.createStatement()) {
            String selectQuery = "SELECT * FROM `" + currentSheetName + "`";
            ResultSet resultSet = statement.executeQuery(selectQuery);

            ResultSetMetaData metaData = resultSet.getMetaData();
            int columnCount = metaData.getColumnCount();
            while (resultSet.next()) {
                List<String> row = new ArrayList<>();
                for (int i = 1; i <= columnCount; i++) {
                    String cellValue = resultSet.getString(i);
                    row.add(cellValue);
                }
                rows.add(row);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return rows;
    }
    private List<Integer> getColumnIndicesContainingKeyword(Map<String, Integer> columnIndices, String keyword) {
        List<Integer> indices = new ArrayList<>();
        for (Map.Entry<String, Integer> entry : columnIndices.entrySet()) {
            if (entry.getKey().contains(keyword)) {
                indices.add(entry.getValue() - 1);
            }
        }
        return indices;
    }
    public Map<String, Integer> getColumnIndices(String currentSheetName) throws SQLException {
        Map<String, Integer> columnIndices = new HashMap<>();
        String url = dbUrl;
        String username = dbUsername;
        String password = dbPassword;

        try (Connection connection = DriverManager.getConnection(url, username, password);
             Statement statement = connection.createStatement()) {
            String selectQuery = "SELECT * FROM `" + currentSheetName + "` LIMIT 1";
            ResultSet resultSet = statement.executeQuery(selectQuery);

            ResultSetMetaData metaData = resultSet.getMetaData();
            int columnCount = metaData.getColumnCount();
            for (int i = 1; i <= columnCount; i++) {
                String columnName = metaData.getColumnName(i);
                columnIndices.put(columnName, i);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

        return columnIndices;
    }

    private void insertDataIntoUser(String sheetName, Workbook workbook) throws SQLException {
        Map<String, Integer> columnIndices = getColumnIndices(sheetName);
        List<List<String>> rows = getTableRows(sheetName);
        Set<String> uniqueDivisionInfo = new HashSet<>();
        Set<String> uniqueDepartmentInfo = new HashSet<>();
        Set<String> uniqueTeamInfo = new HashSet<>();

        List<Integer> divisionIndices = getColumnIndicesContainingKeyword(columnIndices, "Division");
        List<Integer> departmentIndices = getColumnIndicesContainingKeyword(columnIndices, "Dept");
        List<Integer> teamIndices = getColumnIndicesContainingKeyword(columnIndices, "Team");
        List<Integer> staffIDIndices = getColumnIndicesContainingKeyword(columnIndices, "Staff_ID");
        List<Integer> nameIndices = getColumnIndicesContainingKeyword(columnIndices, "Name");
        List<Integer> doorLogNoIndices = getColumnIndicesContainingKeyword(columnIndices, "DoorLog");
        List<Integer> statusIndices = getColumnIndicesContainingKeyword(columnIndices, "Status");
        List<Integer> roleIndices = getColumnIndicesContainingKeyword(columnIndices, "Role");

        for (List<String> row : rows) {
            if (!row.isEmpty()) {

                User user = new User();

                for (Integer divisionIndex : divisionIndices) {
                    String divisionName = row.get(divisionIndex);
                    Division division = divisionRepository.findByName(divisionName)
                            .orElseGet(() -> {
                                Division newDivision = new Division();
                                newDivision.setName(divisionName);
                                divisionRepository.save(newDivision);
                                return newDivision;
                            });
                    uniqueDivisionInfo.add(divisionName);
                    user.setDivision(division);
                }

                for (Integer departmentIndex : departmentIndices) {
                    String departmentName = row.get(departmentIndex);
                    Department department = departmentRepository.findByName(departmentName)
                            .orElseGet(() -> {
                                Department newDepartment = new Department();
                                newDepartment.setName(departmentName);
                                for (Integer divisionIndex : divisionIndices) {
                                    String divisionName = row.get(divisionIndex);
                                    Division division = divisionRepository.findByName(divisionName)
                                            .orElseThrow(() -> new EntityNotFoundException("Division not found"));
                                    newDepartment.setDivision(division);
                                }
                                departmentRepository.save(newDepartment);
                                return newDepartment;
                            });
                    uniqueDepartmentInfo.add(departmentName);
                    user.setDepartment(department);
                }

                for (Integer teamIndex : teamIndices) {
                    String teamName = row.get(teamIndex);
                    Team team = teamRepository.findByName(teamName)
                            .orElseGet(() -> {
                                Team newTeam = new Team();
                                newTeam.setName(teamName);
                                for (Integer departmentIndex : departmentIndices) {
                                    String departmentName = row.get(departmentIndex);
                                    Department department = departmentRepository.findByName(departmentName)
                                            .orElseThrow(() -> new EntityNotFoundException("Department not found"));
                                    newTeam.setDepartment(department);
                                }
                                teamRepository.save(newTeam);
                                return newTeam;
                            });
                    uniqueTeamInfo.add(teamName);
                    user.setTeam(team);
                }

                for (Integer staffIDIndex : staffIDIndices) {
                    user.setStaffId(row.get(staffIDIndex));
                }

                for (Integer nameIndex : nameIndices) {
                    user.setName(row.get(nameIndex));
                }

                for (Integer doorLogNoIndex : doorLogNoIndices) {
                    user.setDoorLogNo(row.get(doorLogNoIndex));
                }

                for (Integer statusIndex : statusIndices) {
                    user.setStatus(row.get(statusIndex));
                }

                for (Integer roleIndex : roleIndices) {
                    String roleName = row.get(roleIndex);
                    Role role = roleRepository.findByName(roleName)
                            .orElseGet(() -> {
                                Role newRole = new Role();
                                newRole.setName(roleName);
                                roleRepository.save(newRole);
                                return newRole;
                            });
                    user.setRoles(Set.of(role));
                }

                user.setPassword("123@dirace");

                userRepository.save(user);
            }
        }
    }


}
