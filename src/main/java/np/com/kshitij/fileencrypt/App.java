package np.com.kshitij.fileencrypt;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.security.GeneralSecurityException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	
	private static final String FILE_PATH = "/tmp/excel.xlsx";
	private static final String PASSWORD = "secret";
	
	public static void main(String[] args) 
			throws IOException, InvalidFormatException, GeneralSecurityException {

        //create a new workbook
        Workbook wb = new XSSFWorkbook();

        //add a new sheet to the workbook
        Sheet sheet1 = wb.createSheet("Player List");

        //add 3 row to the sheet
        Row row1 = sheet1.createRow(0);
        Row row2 = sheet1.createRow(1);
        Row row3 = sheet1.createRow(2);

        //create cells in the row
        Cell row1col1 = row1.createCell(0);
        Cell row1col2 = row1.createCell(1);
        Cell row2col1 = row2.createCell(0);
        Cell row2col2 = row2.createCell(1);
        Cell row3col1 = row3.createCell(0);
        Cell row3col2 = row3.createCell(1);

        //add data to the cells
        row1col1.setCellValue("First Name");
        row1col2.setCellValue("Last Name");
        row2col1.setCellValue("Wayne");
        row2col2.setCellValue("Rooney");
        row3col1.setCellValue("David");
        row3col2.setCellValue("Beckham");

        //write the excel to a file
        try {
            FileOutputStream fileOut = new FileOutputStream(FILE_PATH);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
        }

        //Add password protection and encrypt the file
        POIFSFileSystem fs = new POIFSFileSystem();
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor enc = info.getEncryptor();
        enc.confirmPassword(PASSWORD);

        OPCPackage opc = OPCPackage.open(new File(FILE_PATH), PackageAccess.READ_WRITE);
        OutputStream os = enc.getDataStream(fs);
        opc.save(os);
        opc.close();

        FileOutputStream fos = new FileOutputStream(FILE_PATH);
        fs.writeFilesystem(fos);
        fos.close();    

        System.out.println("File created!!");

    }
}
