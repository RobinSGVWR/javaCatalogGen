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

        String excelFileName = System.getProperty("user.home") + "/Desktop/Generation_de_catalogue/"+txt;//name of excel file= "C:/"+txt+".xlsx";/
        reader.addToConsole("le fichier sera enregistré sous "+excelFileName);

        final File folder = new File(System.getProperty("user.home") + "/Desktop/Generation_de_catalogue/Mettre_ici_les_fichiers_melNumber");
        listFilesForFolder(folder);

        LinkedHashMap<String, ArrayList<String>> map = getHeaders();

         List<String> keys = new ArrayList<String>();


        Workbook newWorkBook = new HSSFWorkbook();
        String sheetName = "Sheet1";//name of sheet
        HSSFSheet sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);

        /*initialisations*/
        printHeader(sheet,map);
        int j=1;
        int cptRow=0;
        int cptSheet=1;
        int listReader=0;
        /*parcourt des mel number recherchés*/


        int leCpt=0;
        for (String melNumber : this.rowL) {

            System.out.println(listReader);
            BasicDBObject searchQuery = new BasicDBObject();
            searchQuery.put("MEL Number", melNumber);
            DBCursor cursor = collection.find(searchQuery);



            if(cursor==null || cursor.count()<1){
                listReader++;
                System.out.println("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa");

                HSSFRow row1 = sheet.createRow(j);
                j++;
                HSSFCell cell = row1.createCell(0);
                cell.setCellValue(melNumber);
            }
            else{
                while (cursor.hasNext()) {
                    cptRow++;
                    listReader++;
                    DBObject article = cursor.next();
                    System.out.println("||||||||||||||||||||||||");
                    System.out.println(melNumber);
                    System.out.println(String.valueOf(article.get("MEL Number")));
                    System.out.println(melNumber.equals(String.valueOf(article.get("MEL Number"))));
                    leCpt++;


                    /*Initialisation d'une ligne par melNumber*/
                     HSSFRow row1 = sheet.createRow(j);
                     j++;



                    /*Comment recuperer valeur attribut*/
    //                String attr = String.valueOf(article.get("Hierarchy Level Image #01"));
    //                System.out.println(attr);

    //                remplissage d'une ligne
                    HSSFCell firstCell = row1.createCell(0);
                    firstCell.setCellValue(melNumber);
                    int i=1;
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
    //                                System.out.println("nouveau header :"+name);
    //                                System.out.println("Nouveau attribut : "+asset);
                                    valeur=String.valueOf(article.get(asset));
    //                                System.out.println("trouvé! "+valeur);

                                }

                        }
                        cell.setCellValue(valeur);

                        if(cptRow>30000){
                            cptRow=0;
    //                         sheetName = "Sheet"+cptSheet;//name of sheet
    //                        cptSheet++;
    //                        sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);
    //                        cptRow=0;
                            this.reader.addToConsole("Tentative de sauvegarde "+cptSheet);

                            try  (OutputStream fileOut = new FileOutputStream(excelFileName+"-"+cptSheet+".xls")) {
                                newWorkBook.write(fileOut);
                                System.out.println("fichier sauvegardé");
                                this.reader.addToConsole("Fichier sauvegardé");
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                            cptSheet++;
                             newWorkBook = new HSSFWorkbook();

                            sheetName = "Sheet";//name of sheet
                             sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);

                             printHeader(sheet, map);

                            j=1;
                        }

                        i++;


                    }

                }
            }
        }



        this.reader.addToConsole("Tentative de sauvegarde");
        try  (OutputStream fileOut = new FileOutputStream(excelFileName+"-Final.xls")) {
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


    
    private void printHeader(HSSFSheet sheet, LinkedHashMap<String, ArrayList<String>> map){
        HSSFRow row1 = sheet.createRow( 0);
        int i=1;
        HSSFCell firstHeaderCell = row1.createCell(0);
        firstHeaderCell.setCellValue("Mel Number");
        ArrayList<String> keys = new ArrayList();
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
    }



    private void listFilesForFolder(final File folder) {
        for (final File fileEntry : Objects.requireNonNull(folder.listFiles())) {
            if (fileEntry.isDirectory()) {
                listFilesForFolder(fileEntry);
            } else {
                System.out.println(fileEntry.getName());
                reader.addToConsole(fileEntry.getName());
            }
        }
    }



}
