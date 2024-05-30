package org.example;

import com.itextpdf.text.DocumentException;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, DocumentException {
        ReadExcel re = new ReadExcel();
        WritePdf wp = new WritePdf();
        List<Employee> employees = re.readExcel(new File("src/main/resources/myExcel.xlsx"));
        wp.writePdf(wp.createText(employees.get(0)));
    }
}
