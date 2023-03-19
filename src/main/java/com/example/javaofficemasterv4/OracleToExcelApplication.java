package com.example.javaofficemasterv4;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;

@SpringBootApplication
public class JavaOfficeMasterV4Application implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(JavaOfficeMasterV4Application.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        List<Employee> employees = employeeRepository.findAllWithHint();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employees");

        int rowIndex = 0;
        Row headerRow = sheet.createRow(rowIndex++);
        createHeader(headerRow);

        for (Employee employee : employees) {
            Row row = sheet.createRow(rowIndex++);
            createEmployeeRow(employee, row);
        }

        try (FileOutputStream fileOut = new FileOutputStream("employees.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();
        System.out.println("엑셀 파일이 생성되었습니다.");
    }

    private void createHeader(Row headerRow) {
        Cell cell = headerRow.createCell(0);
        cell.setCellValue("ID");

        cell = headerRow.createCell(1);
        cell.setCellValue("Name");

        // 다른 필드를 추가하려면 여기에 추가하세요.
    }

    private void createEmployeeRow(Employee employee, Row row) {
        Cell cell = row.createCell(0);
        cell.setCellValue(employee.getId());
        cell = row.createCell(1);
        cell.setCellValue(employee.getName());

        // 다른 필드를 추가하려면 여기에 추가하세요.
    }



}
