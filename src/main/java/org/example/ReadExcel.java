package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcel {
    public List<Employee> readExcel(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
        Iterator<Row> itr = sheet.iterator();    //iterating over excel file
        List<Employee> employees = new ArrayList<>();
        itr.next();
        while (itr.hasNext()) {
            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();//iterating over each column
            Employee employee = new Employee();
            while (cellIterator.hasNext()) {
                if (employee.getId() == 0) {
                    employee.setId((int) cellIterator.next().getNumericCellValue());
                }
                if (employee.getName() == null) {
                    employee.setName(cellIterator.next().getStringCellValue());
                }
                if (employee.getSalary() == null) {
                    employee.setSalary(cellIterator.next().getNumericCellValue());
                }
                break;
            }
            employees.add(employee);
        }
        return employees;
    }
}
