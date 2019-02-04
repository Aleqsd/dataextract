/***
 * Main class for this POC
 */
public class Main {
    public static void main(String[] args) {
        Extractor extractor = new Extractor();
        extractor.writeLifiInExcelFile();
        extractor.writePhoneBoxInExcelFile();
        extractor.writeAppliInExcelFile();
    }
}
