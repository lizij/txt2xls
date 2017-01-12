import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;

/**
 * Created by Lizij on 2017/1/12.
 */
public class txttoxls{
    public static void main(String[] args) throws Exception{
        txtToExcel("input\\test3.eml.content.txt", "input\\test3.eml.metadata.txt","output\\test3.xls", "test");
    }

    static void txtToExcel(String inputContentFileName, String inputMetadataFileName, String outputFileName, String sheetName) throws IOException {
        BufferedReader contentReader = new BufferedReader(new FileReader(new File(inputContentFileName)));
        BufferedReader metadataReader = new BufferedReader(new FileReader(new File(inputMetadataFileName)));
        HSSFWorkbook hwb = new HSSFWorkbook();
        HSSFSheet sheet = hwb.createSheet(sheetName);
        HSSFRow row = null;
        HSSFCell[] cell = null;

        String str = null;
        int i = 0;
        while ((str = metadataReader.readLine()) != null) {
            String[] words = str.split(":");
            row = sheet.createRow(i);
            cell = new HSSFCell[2];
            cell[0] = row.createCell(0);
            cell[0].setCellValue(words[0]);
            cell[1] = row.createCell(1);
            cell[1].setCellValue(words[1]);

            i++;
        }

        StringBuffer content = new StringBuffer();
        while ((str = contentReader.readLine()) != null){
            content.append(str + "\r\n");
        }
        row = sheet.createRow(i);
        cell = new HSSFCell[2];
        cell[0] = row.createCell(0);
        cell[0].setCellValue("content");
        cell[1] = row.createCell(1);
        cell[1].setCellValue(content.toString());

        OutputStream out = new FileOutputStream(outputFileName);
        hwb.write(out);
        out.close();
        contentReader.close();
        metadataReader.close();

    }

}
