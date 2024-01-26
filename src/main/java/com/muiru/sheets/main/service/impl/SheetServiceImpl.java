package com.muiru.sheets.main.service.impl;

import com.muiru.sheets.main.service.SheetService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
@Service
@RequiredArgsConstructor
public class SheetServiceImpl implements SheetService {
    private final Environment environment;


    @Override
    public Path exportExcelFile() throws IOException {
        // blank worksheet
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("test sheet");

        Map<String,Object[]> data = new TreeMap<>();
        // writing data to sheet
        data.put("1",new Object[]{
                "ID","NAME","LASTNAME"
        });
        data.put("2",new Object[] { 1, "Pankaj", "Kumar" });

        // iterating over data and writing it to sheet
        Set<String> keyset = data.keySet();

        int rowNum =0;
        for(String key:keyset){
            Row row = sheet.createRow(rowNum++);
            Object[] objectArr = data.get(key);
            int cellNum=0;

            for(Object object: objectArr){
                Cell cell = row.createCell(cellNum++);
                if(object instanceof String){
                    cell.setCellValue((String) object);
                }
                else if(object instanceof Integer){
                    cell.setCellValue((Integer) object);
                }
            }
        }

        try{
            Path location = Paths.get(environment.getProperty("sheets.location")+"/test.xlsx");
            FileOutputStream out = new FileOutputStream(
                    location.toFile()
            );
            workbook.write(out);
            out.close();

            return location;

        }
        catch (Exception e){
            throw e;
        }


    }
}
