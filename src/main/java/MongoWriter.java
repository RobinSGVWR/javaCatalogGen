import com.mongodb.*;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.net.UnknownHostException;
import java.util.*;

import javax.swing.filechooser.FileSystemView;

public class MongoWriter {
    private Reader reader;

    MongoWriter(Reader reader) {
        this.reader = reader;
    }

    private ArrayList<File> getListFile() {
        File home = FileSystemView.getFileSystemView().getHomeDirectory();
        String currentPath = System.getProperty("user.dir");
        this.reader.addToConsole("getlistFile");
        File folder = new File(currentPath + "/Generation_de_catalogue/Mettre_ici_les_fichiers_melNumber");
        return new ArrayList(Arrays.asList(Objects.requireNonNull(folder.listFiles())));
    }

    void generateMongo(boolean retrieveImages) throws UnknownHostException {
        this.reader.addToConsole("generateMongo");
        ArrayList<File> listFile = getListFile();
        for (File file : listFile) {
            this.reader.addToConsole("Nouveau fichier : " + getFileName(file));
            generateExcelFile(file, retrieveImages);
        }
    }

    void generateExcelFile(File file, boolean retrieveImages) throws UnknownHostException {
        /**
         * le but est de parcourir chaque mel number, et pour chacune de ces reférences,
         * aller dans le tableau associé (map) regardé les attributs disponibles, faire
         * une requête et, si il y a un resultat, l'ajouter dans le excel dans le header
         * correspondant
         */
        reader.addToConsole("Generate Excel File");
        reader.addToConsole("New MongoDB instance");

        MongoClient mongoClient = new MongoClient();
        DB db = mongoClient.getDB("catalog");
        DBCollection collection = db.getCollection("products");

        String currentPath = System.getProperty("user.dir");
        System.out.println(currentPath);
        String excelFileName = currentPath + "/Generation_de_catalogue/resultats/" + getFileName(file);
        reader.addToConsole("le fichier sera enregistré sous " + excelFileName);

        LinkedHashMap<String, ArrayList<String>> headers = getHeaders();

        String sheetName = "Sheet1"; // name of sheet
        Workbook newWorkBook = new HSSFWorkbook();
        HSSFSheet sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);

        // initialisations
        printHeader(sheet, headers);
        int cptRow = 1;
        int cptReadCell = 0;
        int cptSheet = 1;
        int listReader = 0;

        // parcours des mel numbers recherchés
        for (String melNumber : getMelNumberList(file)) {
            System.out.println(listReader);
            BasicDBObject searchQuery = new BasicDBObject();
            searchQuery.put("MEL Number", melNumber);
            DBCursor cursor = collection.find(searchQuery);

            if (cursor == null || cursor.count() < 1) {
                listReader++;
                System.out.println("-------------- Nothing for " + melNumber + " ---------------");

                HSSFRow row = sheet.createRow(cptRow);
                cptRow++;
                HSSFCell cell = row.createCell(0);
                cell.setCellValue(melNumber);
            } else {
                while (cursor.hasNext()) {
                    cptReadCell++;
                    listReader++;
                    DBObject product = cursor.next();
                    System.out.println("#####################################");
                    System.out.println(melNumber);
                    System.out.println(String.valueOf(product.get("MEL Number")));
                    System.out.println(melNumber.equals(String.valueOf(product.get("MEL Number"))));

                    // Initialisation d'une ligne par melNumber
                    HSSFRow row = sheet.createRow(cptRow);
                    cptRow++;

                    // remplissage d'une ligne
                    HSSFCell firstCell = row.createCell(0);
                    firstCell.setCellValue(melNumber);

                    int cptNewRowCell = 1;
                    for (String headerKey : headers.keySet()) {
                        HSSFCell cell = row.createCell(cptNewRowCell);

                        // parcours de chacun des titres possibles avec le header
                        String cellValue = "";
                        Iterator<String> it = headers.get(headerKey).iterator();

                        while (it.hasNext()) {
                            String headerValue = it.next();
                            String value = String.valueOf(product.get(headerValue));

                            // recherche si le titre dans le dbobject
                            if (value != null && !value.equals("") && !value.equals("null")) {
                                // Ici, une valeur a été trouvé, il faut donc l'ajouter dans la case du excel
                                cellValue = value;
                            }

                            if (retrieveImages && headers.get(headerKey).contains("Hierarchy Level Image #01")) {
                                try {
                                    URL url = new URL(String.valueOf(product.get(headerValue)));
                                    String[] tabUrl = url.getFile().split("/");
                                    File outputImg = new File(currentPath
                                            + "/Generation_de_catalogue/resultats/" + getFileName(file)
                                            + " - img/" + tabUrl[tabUrl.length - 1]);
                                    if(!outputImg.exists()) {
                                        FileUtils.copyURLToFile(url, outputImg);
                                    }
                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                        }

                        cell.setCellValue(cellValue);

                        // Sauvegarde toutes les 30000 cellules
                        if (cptReadCell > 30000) {
                            cptReadCell = 0;
                            this.reader.addToConsole("Tentative de sauvegarde " + cptSheet);

                            try (OutputStream fileOut = new FileOutputStream(excelFileName + "-" + cptSheet + ".xls")) {
                                newWorkBook.write(fileOut);
                                System.out.println("fichier sauvegardé");
                                this.reader.addToConsole("Fichier sauvegardé");
                            } catch (IOException e) {
                                e.printStackTrace();
                            }

                            cptSheet++;
                            newWorkBook = new HSSFWorkbook();

                            sheetName = "Sheet";// name of sheet
                            sheet = (HSSFSheet) newWorkBook.createSheet(sheetName);

                            printHeader(sheet, headers);
                            cptRow = 1;
                        }

                        cptNewRowCell++;
                    }
                }
            }
        }

        this.reader.addToConsole("Tentative de sauvegarde");
        try (OutputStream fileOut = new FileOutputStream(excelFileName + "-Final.xls")) {
            newWorkBook.write(fileOut);
            System.out.println("Fichier sauvegardé");
            this.reader.addToConsole("Fichier sauvegardé");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                newWorkBook.close();
                mongoClient.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void saveExcel(String excelFileName, Workbook newWorkBook) {
        FileOutputStream fileOut = null;

        try {
            fileOut = new FileOutputStream(excelFileName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        // write this workbook to an Outputstream.
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

    private ArrayList<String> getMelNumberList(File file) {
        this.reader.addToConsole("getMelNumberLsit");
        ArrayList<String> melNumberList = new ArrayList<String>();

        try {
            FileInputStream flux = new FileInputStream(file);
            InputStreamReader lecture = new InputStreamReader(flux);
            BufferedReader buff = new BufferedReader(lecture);
            String ligne;

            while ((ligne = buff.readLine()) != null) {
                melNumberList.add(ligne);
            }

            buff.close();
        } catch (Exception ex) {
            System.out.println(ex.toString());
        }

        return melNumberList;
    }

    private Boolean isFloatable(String value) {
        value = value.replace('.', ',');
        Float valueF;

        if (value == "") {
            return false;
        }

        try {
            valueF = Float.parseFloat(value.replace(",", "."));
            return true;
        } catch (NumberFormatException e) {
            System.out.println(e);
            // this.reader.addToConsole(e.toString());
            return false;
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
                int acc = 0;

                ArrayList<String> values = new ArrayList<>();
                String tmpheader = null;

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    String value = currentCell.getStringCellValue();

                    if (acc == 0) {
                        tmpheader = value;
                    } else {
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

    private void printHeader(HSSFSheet sheet, LinkedHashMap<String, ArrayList<String>> map) {
        HSSFRow row1 = sheet.createRow(0);
        int i = 1;
        HSSFCell firstHeaderCell = row1.createCell(0);
        firstHeaderCell.setCellValue("Mel Number");
        ArrayList<String> keys = new ArrayList();

        for (String name : map.keySet()) {
            String key = name.toString();
            /* Creation du tableau des clés */
            keys.add(key);

            /* remplissage du header du excel */
            HSSFCell cell = row1.createCell(i);
            // System.out.println(name);
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

    private String getFileName(File file) {
        return file.getName().substring(0, file.getName().length() - 4);
    }

    // GETTERS SETTERS

    public Reader getReader() {
        return reader;
    }

    public void setReader(Reader reader) {
        this.reader = reader;
    }
}