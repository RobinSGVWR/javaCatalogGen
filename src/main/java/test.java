import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class test {
    public static void main(String [ ] args) {
        String sheetName;
        Iterator<Row> rowIt;
        Row row;
        Iterator cellIt;
        int cpt=0;


//        try {
//
//            Workbook wb=WorkbookFactory.create(new File("C:/Users/robin.saint-georges/Desktop/newTete.xlsx")) ;
//
//            Sheet sheet = wb.getSheetAt(0);
//
//            Row r = sheet.getRow(1);
//            int maxCell=  r.getLastCellNum();
//
//            Map<String, ArrayList<String>> map = new HashMap<>();
//
//    //        pourchaques column jusqu'a la fin
//            for (int i=1; i<maxCell;i++){
//                String columnLetter = CellReference.convertNumToColString(i);
//                System.out.println(i);
//
//                Cell cell =  wb.getSheetAt(0).
//                        getRow(1).
//                        getCell(i);
//
//                String tmpHeaderName = cell.getStringCellValue();
//
//                ArrayList<String> values = new ArrayList<>();
//
//                for(int j=2; j<30;j++){
//                    String cellValue = wb.getSheetAt(0).
//                            getRow(j).
//                            getCell(i).getStringCellValue();
//
//                    if(!cellValue.equals("") &&
//                            !cellValue.equals("OK") &&
//                            !cellValue.equals("Not found")){
//
//                        values.add(cellValue);
//
//                    }
//                }
//
////                ajout au nom de column des paramèters
//                map.put(tmpHeaderName, values);
//
//
//
//            }
//            Set<String> keys = map.keySet();
//            String[] array = keys.toArray(new String[0]);
//
//            System.out.println("LES ITEMS");
//
//            for (String item : array) {
//                System.out.println(item);
//            }
//
//            System.out.println("LES ITEMS");
//
//        } catch (IOException | InvalidFormatException e) {
//            e.printStackTrace();
//        }

        try {

            FileInputStream excelFile = new FileInputStream(new File("C:/Users/robin.saint-georges/Desktop/transMapped.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            Map<String, ArrayList<String>> map = new HashMap<>();


            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                int acc=0;

                ArrayList<String> values = new ArrayList<>();
                String tmpheader = null;

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String value = currentCell.getStringCellValue();
//                    System.out.print(currentCell.getStringCellValue() + "--");

                    if (acc==0){
                         tmpheader = value;
                    }
                    else {
                        values.add(value);
                    }
                    acc++;
                }
                map.put(tmpheader, values);

                System.out.println();



                }
            for (String name: map.keySet()){

                String key =name.toString();
                String value = map.get(name).toString();
                System.out.println(key + " " + value);


            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }



    }

}
