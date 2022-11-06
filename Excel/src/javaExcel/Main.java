package javaExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class Main {
public static void main(String[] args) {
String path = "C:\\Реестр_домов_v10.xlsx"; //путь до Excel файла
String square, floor, entrance, apartment;

FileInputStream fis = new FileInputStream(path);
Workbook wb = new HSSFWorkbook(fis);

//считываем площади домов из файла
for (int i = 1; i<33399; i++){
    square = wb.getSheetAt(0).getRow(1).getCell(i).getStringCellValue();
    System.out.println(square);
}

//считываем этажи из файла
for (int i = 1; i<33399; i++){
        floor = wb.getSheetAt(0).getRow(3).getCell(i).getStringCellValue;
        System.out.println(floor);
    }

//считываем подъезды из файла
for (int i = 1; i<33399; i++){
        entrance = wb.getSheetAt(0).getRow(4).getCell(i).getStringCellValue;
        System.out.println(entrance);
    }

//считываем квартиры из файла
for (int i = 1; i<33399; i++){
        apartment = wb.getSheetAt(0).getRow(5).getCell(i).getStringCellValue;
        System.out.println(apartment);
    }


}
}

