import com.mongodb.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.UnknownHostException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;



public class MongoWriter {

    private static final DateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm");
    private String txt;
    private HSSFWorkbook wb;
    private List<String> rowL;
    private Reader reader;


    MongoWriter(String txt, List<String> rowL, Reader reader) {
        this.txt = txt;
        this.rowL = rowL;
        wb= new HSSFWorkbook();
        this.reader = reader;
    }

    void generateMongo() throws UnknownHostException {

/**le but est de parcourir chaque mel number, et pour chacune de ces reférences, aller
 * dans le tableau associé (map)
 * regardé les attributs dispobible, faire une requete et, si il y a un resultat, l'ajouter dans le excel dans
 * le header correspondant
 */

        reader.addToConsole("generating mongo");

        MongoClient mongoClient = new MongoClient();
        DB db = mongoClient.getDB("catalog");
        DBCollection collection = db.getCollection("products");

        String excelFileName = System.getProperty("user.home") + "/Desktop/Generation_de_catalogue/"+txt+".xls";//name of excel file= "C:/"+txt+".xlsx";/
        reader.addToConsole("le fichier sera enregistré sous "+excelFileName);



        LinkedHashMap<String, ArrayList<String>> map = getHeaders();

         List<String> keys = new ArrayList<String>();


        Workbook newWorkBook = new HSSFWorkbook();
        String sheetName = "Sheet1";//name of sheet
        HSSFSheet sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);

        /*initialisations*/
        HSSFRow row1 = sheet.createRow( 0);
        int i=1;
        HSSFCell firstHeaderCell = row1.createCell(0);
        firstHeaderCell.setCellValue("Mel Number");
        for (String name: map.keySet()){

            String key =name.toString();
            /*Creation du tableau des clés*/
            keys.add(key);

            /*remplissage du header du excel*/
            HSSFCell cell = row1.createCell( i);
//            System.out.println(name);
            cell.setCellValue(name);
            i++;
        }
        int j=1;
        int cptRow=0;
        int cptSheet=2;
        /*parcourt des mel number recherchés*/
        for (String melNumber : this.rowL) {
            System.out.println("Nouveau Mel Number :"+melNumber);

            BasicDBObject searchQuery = new BasicDBObject();
            searchQuery.put("MEL Number", melNumber);
            DBCursor cursor = collection.find(searchQuery);


            while (cursor.hasNext()) {
                cptRow++;

                /*Initialisation d'une ligne par melNumber*/
                 row1 = sheet.createRow(j);
                 j++;

                DBObject article = cursor.next();
                System.out.println("Nouveau articles :"+article);

                /*Comment recuperer valeur attribut*/
//                String attr = String.valueOf(article.get("Hierarchy Level Image #01"));
//                System.out.println(attr);

//                remplissage d'une ligne
                HSSFCell firstCell = row1.createCell(0);
                firstCell.setCellValue(melNumber);
                i=1;
                for (String name: map.keySet()){
                    HSSFCell cell = row1.createCell(i);


                    /* parcourt de chacune des asset rechercher par header */
                    String key =name.toString();
                    ArrayList<String> values = map.get(name);
                    String valeur="";
                    for(String asset : values ){


                        boolean find = false;
                        /*recherche si le asset existe dans le dbobject*/
                        int cpt = 0;


                            if(String.valueOf(article.get(asset))==null ||
                                    String.valueOf(article.get(asset)).toString().equals("null")){
                                if(valeur.equals("")){
                                    valeur="";
                                }
                            }
                            else{
                                /*Ici, une valeur a été trouvé, il faut donc l'ajouter dans la case du excel*/
                                find=true;
                                System.out.println("nouveau header :"+name);
                                System.out.println("Nouveau attribut : "+asset);
                                valeur=String.valueOf(article.get(asset));
                                System.out.println("trouvé! "+valeur);

                            }

                    }
//                    if(cptRow>2){
//                         sheetName = "Sheet"+cptSheet;//name of sheet
//                        cptSheet++;
//                        sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);
//                        cptRow=0;
//                    }
                    cell.setCellValue(valeur);
                    i++;


                }

            }
        }



        this.reader.addToConsole("Tentative de sauvegarde");
        try  (OutputStream fileOut = new FileOutputStream(excelFileName)) {
            newWorkBook.write(fileOut);
            System.out.println("fichier sauvegardé");
            this.reader.addToConsole("Fichier sauvegardé");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //GETTERS SETTERS
    public String getTxt() {
        return txt;
    }

    public void setTxt(String txt) {
        this.txt = txt;
    }

    public HSSFWorkbook getWb() {
        return wb;
    }

    public void setWb(HSSFWorkbook wb) {
        this.wb = wb;
    }

    public List<String> getRowL() {
        return rowL;
    }

    public void setRowL(List<String> rowL) {
        this.rowL = rowL;
    }

    public Reader getReader() {
        return reader;
    }

    public void setReader(Reader reader) {
        this.reader = reader;
    }

    private Boolean isFloatable(String value){
        value=value.replace('.', ',');
        Float valueF;
        if (value==""){
            return false;
        }
        try{
            valueF = Float.parseFloat(value.replace(",","."));
            return true;
        }catch (NumberFormatException e){
            System.out.println(e);
//            this.reader.addToConsole(e.toString());
            return false;
        }


    }

    public void saveExcel(String excelFileName, Workbook newWorkBook ){
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(excelFileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        //write this workbook to an Outputstream.
        try {
            newWorkBook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public LinkedHashMap<String, ArrayList<String>> getHeaders() {
        try {

            FileInputStream excelFile = new FileInputStream(new File("transMapped.xls"));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            LinkedHashMap<String, ArrayList<String>> map = new LinkedHashMap<>();


            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                int acc=0;

                ArrayList<String> values = new ArrayList<>();
                String tmpheader = null;

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String value = currentCell.getStringCellValue();

                    if (acc==0){
                        tmpheader = value;
                    }
                    else {
                        values.add(value);
                    }
                    acc++;
                }
                map.put(tmpheader, values);

            }

            return map;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

//    public String getValueWithMelAndAsset(String melNumber, String asset){
//        BasicDBObject searchQuery = new BasicDBObject();
//        searchQuery.put("MEL Number", melNumber);
//        DBCursor cursor = collection.find(searchQuery);
//
//        while (cursor.hasNext()) {
//
//            DBObject theObj = cursor.next();
//            System.out.println(theObj);
//
//            if(){
//
//            }
//
//        }
//
//        while (cursor.hasNext()) {
//            DBObject theObj = cursor.next();
//            //How to get the DBObject value to ArrayList of Java Object?
//
//            BasicDBList studentsList = (BasicDBList) theObj.get("students");
//            for (int i = 0; i < studentsList.size(); i++) {
//                BasicDBObject studentObj = (BasicDBObject) studentsList.get(i);
//                String firstName = studentObj.getString("firstName");
//                String lastName = studentObj.getString("lastName");
//                String age = studentObj.getString("age");
//                String gender = studentObj.getString("gender");
//
//                Student student = new Student();
//                student.setFirstName(firstName);
//                student.setLastName(lastName);
//                student.setAge(age);
//                student.setGender(gender);
//
//                students.add(student);
//    }
    



}
