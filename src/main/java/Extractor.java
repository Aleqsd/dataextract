import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

/***
 * Extractor Tools
 */
class Extractor {
    private Connection conn;
    private String currentPath;

    /***
     * Basic constructor ititializing the connect component for the SQLite DB
     */
    Extractor() {
        this.currentPath = System.getProperty("user.dir").replace("\\", "/");
        connect();
    }

    /***
     * Connect using jdbc sqlite
     */
    private void connect() {
        try {
            // db parameters
            String dbFileName = "iot_data.db";
            String url = "jdbc:sqlite:"+currentPath+"/"+ dbFileName;
            // create a connection to the database
            conn = DriverManager.getConnection(url);

            System.out.println("Connection to SQLite has been established.");

        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
    }

    /***
     * Write Lifi DB extract in the XLSX file
     */
    void writeLifiInExcelFile(){
        try {
            String sql = "SELECT * FROM lifi";

            File excelFile = new File((currentPath+"\\data_extract.xlsx"));
            FileInputStream fsIP = new FileInputStream(excelFile);
            //Access the workbook
            XSSFWorkbook wb = new XSSFWorkbook(fsIP);
            //Access the worksheet, so that we can update / modify it.
            XSSFSheet worksheet = wb.getSheetAt(0);

            Statement stmt  = conn.createStatement();
            ResultSet rs    = stmt.executeQuery(sql);

            // loop through the result set
            int i = 1;
            writeInCell(0,0,"Id", worksheet);
            writeInCell(0,1,"Date", worksheet);
            writeInCell(0,2,"Heure", worksheet);
            writeInCell(0,3,"Durée_dessous", worksheet);
            writeInCell(0,4,"Ip", worksheet);
            writeInCell(0,5,"Upload", worksheet);
            writeInCell(0,6,"Download", worksheet);

            while (rs.next()) {
                writeInCell(i,0,rs.getInt("Id"), worksheet);
                writeInCell(i,1,rs.getString("Date"), worksheet);
                writeInCell(i,2,rs.getString("Heure"), worksheet);
                writeInCell(i,3,rs.getInt("Durée_dessous"), worksheet);
                writeInCell(i,4,rs.getString("Ip"), worksheet);
                writeInCell(i,5,rs.getInt("Upload"), worksheet);
                writeInCell(i,6,rs.getInt("Download"), worksheet);
                i++;
            }

            //Close the InputStream
            fsIP.close();
            //Open FileOutputStream to write updates
            FileOutputStream output_file = new FileOutputStream(excelFile);
            //write changes
            wb.write(output_file);
            //close the stream
            output_file.close();
        } catch (SQLException | IOException e1) {
            e1.printStackTrace();
        }
    }

    /***
     * Write Phonebox DB extract in the XLSX file
     */
    void writePhoneBoxInExcelFile(){
        try {
            String sql = "SELECT * FROM phonebox";

            File excelFile = new File((currentPath+"\\data_extract.xlsx"));
            FileInputStream fsIP = new FileInputStream(excelFile);
            //Access the workbook
            XSSFWorkbook wb = new XSSFWorkbook(fsIP);
            //Access the worksheet, so that we can update / modify it.
            XSSFSheet worksheet = wb.getSheetAt(1);

            Statement stmt  = conn.createStatement();
            ResultSet rs    = stmt.executeQuery(sql);


            int i = 1;
            writeInCell(0,0,"id", worksheet);
            writeInCell(0,1,"Durée_Hotspot", worksheet);
            writeInCell(0,2,"Upload", worksheet);
            writeInCell(0,3,"Download", worksheet);
            writeInCell(0,4,"Itinéraire", worksheet);
            writeInCell(0,5,"Durée_Appel", worksheet);
            writeInCell(0,6,"Date", worksheet);
            writeInCell(0,7,"Heure", worksheet);
            writeInCell(0,8,"Appel_ext", worksheet);
            writeInCell(0,9,"Appel_Taxi", worksheet);
            writeInCell(0,10,"Urgence", worksheet);
            writeInCell(0,11,"Durée_Urgence", worksheet);
            writeInCell(0,12,"Conv_enregistrée", worksheet);
            writeInCell(0,13,"Tps_recharge", worksheet);
            writeInCell(0,14,"Model_tel", worksheet);

            // loop through the result set
            while (rs.next()) {
                writeInCell(i,0,rs.getInt("id"), worksheet);
                writeInCell(i,1,rs.getInt("Durée_Hotspot"), worksheet);
                writeInCell(i,2,rs.getInt("Upload"), worksheet);
                writeInCell(i,3,rs.getInt("Download"), worksheet);
                writeInCell(i,4,rs.getString("Itinéraire"), worksheet);
                writeInCell(i,5,rs.getInt("Durée_Appel"), worksheet);
                writeInCell(i,6,rs.getString("Date"), worksheet);
                writeInCell(i,7,rs.getString("Heure"), worksheet);
                writeInCell(i,8,rs.getInt("Appel_ext"), worksheet);
                writeInCell(i,9,rs.getInt("Appel_Taxi"), worksheet);
                writeInCell(i,10,rs.getInt("Urgence"), worksheet);
                writeInCell(i,11,rs.getInt("Durée_Urgence"), worksheet);
                writeInCell(i,12,rs.getInt("Conv_enregistrée"), worksheet);
                writeInCell(i,13,rs.getInt("Tps_recharge"), worksheet);
                writeInCell(i,14,rs.getString("Model_tel"), worksheet);
                i++;
            }

            //Close the InputStream
            fsIP.close();
            //Open FileOutputStream to write updates
            FileOutputStream output_file = new FileOutputStream(excelFile);
            //write changes
            wb.write(output_file);
            //close the stream
            output_file.close();
        } catch (SQLException | IOException e1) {
            e1.printStackTrace();
        }
    }

    /***
     * Write Appli DB extract in the XLSX file
     */
    void writeAppliInExcelFile(){
        try {
            String sql = "SELECT * FROM Appli";

            File excelFile = new File((currentPath+"\\data_extract.xlsx"));
            FileInputStream fsIP = new FileInputStream(excelFile);
            //Access the workbook
            XSSFWorkbook wb = new XSSFWorkbook(fsIP);
            //Access the worksheet, so that we can update / modify it.
            XSSFSheet worksheet = wb.getSheetAt(2);

            Statement stmt  = conn.createStatement();
            ResultSet rs    = stmt.executeQuery(sql);

            int countMusee = 0;
            int countCinema = 0;
            int countTramway = 0;
            int countBus = 0;
            int countOptionEcolo = 0;


            // loop through the result set
            int i = 1;
            writeInCell(0,0,"id", worksheet);
            writeInCell(0,1,"Itineraire", worksheet);
            writeInCell(0,2,"Transport", worksheet);
            writeInCell(0,3,"Option_Ecolo", worksheet);
            writeInCell(0,4,"Lieu_interet", worksheet);
            writeInCell(0,5,"Ticket", worksheet);
            writeInCell(0,6,"Geolocalisation", worksheet);
            writeInCell(0,7,"Mail", worksheet);
            writeInCell(0,8,"N°Tel", worksheet);
            writeInCell(0,9,"Nom", worksheet);
            writeInCell(0,10,"Prenom", worksheet);

            while (rs.next()) {
                writeInCell(i,0,rs.getInt("id"), worksheet);
                writeInCell(i,1,rs.getString("Itineraire"), worksheet);
                writeInCell(i,2,rs.getString("Transport"), worksheet);
                if (rs.getString("Transport").equals("Tramway"))
                    countTramway++;
                if (rs.getString("Transport").equals("Bus"))
                    countBus++;
                writeInCell(i,3,rs.getInt("Option_Ecolo"), worksheet);
                if (rs.getInt("Option_Ecolo") == 1)
                    countOptionEcolo++;
                writeInCell(i,4,rs.getString("Lieu_interet"), worksheet);
                if (rs.getString("Lieu_interet") != null)
                {
                    if (rs.getString("Lieu_interet").equals("Musée"))
                        countMusee++;
                    if (rs.getString("Lieu_interet").equals("Cinéma"))
                        countCinema++;
                }
                writeInCell(i,5,rs.getString("Ticket"), worksheet);
                writeInCell(i,6,rs.getString("Geolocalisation"), worksheet);
                writeInCell(i,7,rs.getString("Mail"), worksheet);
                writeInCell(i,8,rs.getInt("N°Tel"), worksheet);
                writeInCell(i,9,rs.getString("Nom"), worksheet);
                writeInCell(i,10,rs.getString("Prenom"), worksheet);

                i++;
            }

            writeInCell(20,13,countMusee, worksheet);
            writeInCell(20,14,countCinema, worksheet);
            writeInCell(20,15,countTramway, worksheet);
            writeInCell(20,16,countBus, worksheet);
            writeInCell(20,17,countOptionEcolo, worksheet);

            //Close the InputStream
            fsIP.close();
            //Open FileOutputStream to write updates
            FileOutputStream output_file = new FileOutputStream(excelFile);
            //write changes
            wb.write(output_file);
            //close the stream
            output_file.close();
        } catch (SQLException | IOException e1) {
            e1.printStackTrace();
        }
    }

    /***
     * Write any value you want in a specific location in one worksheet
     * @param rowNumber row number of the worksheet
     * @param cellNumber cell number of the worksheet
     * @param value value you want to enter (String)
     * @param worksheet worksheet to use
     */
    private void writeInCell(int rowNumber, int cellNumber, String value, XSSFSheet worksheet)
    {
        if (worksheet.getRow(rowNumber) == null)
            worksheet.createRow(rowNumber);
        worksheet.getRow(rowNumber).createCell(cellNumber);
        worksheet.getRow(rowNumber).getCell(cellNumber).setCellValue(value);
    }

    /***
     * Write any value you want in a specific location in one worksheet
     * @param rowNumber row number of the worksheet
     * @param cellNumber cell number of the worksheet
     * @param value value you want to enter (int)
     * @param worksheet worksheet to use
     */
    private void writeInCell(int rowNumber, int cellNumber, int value, XSSFSheet worksheet)
    {
        if (worksheet.getRow(rowNumber) == null)
            worksheet.createRow(rowNumber);
        worksheet.getRow(rowNumber).createCell(cellNumber);
        worksheet.getRow(rowNumber).getCell(cellNumber).setCellValue(value);
    }
}
