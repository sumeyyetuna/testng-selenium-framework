package excelautomation;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelReadDemo {

    @Test
    public void readXLFile() throws Exception{
       String path="./src/test/resources/Countries.xlsx";

        FileInputStream inputStream= new FileInputStream(path);
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet worksheet= workbook.getSheetAt(0);
        Row row=worksheet.getRow(0);
        Cell cell1=row.getCell(0);
        Cell cell2=row.getCell(1);
        System.out.println(cell1.toString());

        String country1=worksheet.getRow(1).getCell(0).toString();
        String capital1=workbook.getSheetAt(0).getRow(1).getCell(1).toString();
        System.out.println(country1);
        System.out.println(capital1);

        int rowCount=worksheet.getLastRowNum();
        System.out.println(rowCount);

//        for (int i = 1; i <=rowCount ; i++) {
//            System.out.println("Country #"+i+"i"+worksheet.getRow(i).getCell(0)+
//                    " ==> "+worksheet.getRow(i).getCell(1));
//
//
//        }
        Map<String,String> maplist= new HashMap<>();
        for (int i = 1; i <=rowCount ; i++) {
            maplist.put(worksheet.getRow(i).getCell(0).toString(),worksheet.getRow(i).getCell(1).toString());

        }
        System.out.println(maplist);
        Cell column=worksheet.getRow(0).getCell(2);
        if(column==null){
            column=worksheet.getRow(0).createCell(2);
        }
        column.setCellValue("Continent");

        Cell cont1=worksheet.getRow(1).getCell(2);
        if(cont1==null){
            cont1=worksheet.getRow(1).createCell(2);
        }
        cont1.setCellValue("North America");
        FileOutputStream out= new FileOutputStream(path);
        workbook.write(out);


        out.close();
        workbook.close();
        inputStream.close();






    }

}
