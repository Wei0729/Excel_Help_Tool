/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package localization;
import java.io.File; 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.Iterator;
import java.util.Vector;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author wei7771
 */
public class excel {
    public static int columnNumber = 32;
    public static int rowNumber = 0;
    public static int langNumber = 0;
 
    /*public static void main(String args[]){
        String excel = "C:\\Users\\wei7771\\Desktop\\ss\\OnlineJobsReport-201507131621532.zip";
        try{
            convert(excel);
        }catch(Exception e){
            
        }
    } */
     
      public static void convert(String filePath){
          Vector<String> zFile;
          if(filePath.endsWith(".zip")){
            zFile = readzipfile(filePath);
            for(String s : zFile){
                if(s.endsWith(".xlsx")){
                    //System.out.println(s);
                    convert(s);
                }
            }
          }
          else if(!filePath.endsWith(".xlsx")){
            return;
          }
          else{
           try{
                FileInputStream file = new FileInputStream(new File(filePath));
                System.out.println(filePath);
                //Get the workbook instance for XLS file 
                XSSFWorkbook workbook = new XSSFWorkbook (file);
                XSSFSheet sheet = workbook.getSheetAt(0);
                XSSFRow row;
                XSSFCell cell;
                rowNumber = sheet.getPhysicalNumberOfRows();
                try{
                    for(int i = 0; i < rowNumber; i++){
                    row = sheet.getRow(i);
                    if(row != null){
                        int columnNum = row.getPhysicalNumberOfCells();
                        //System.out.println(columnNum);
                        for(int j=0; j<columnNum; j++){
                            cell = row.getCell(j);

                            if(j == 0){
                                String name = cell.getRichStringCellValue().getString();
                                if(name.equalsIgnoreCase("Esri")){
                                    langNumber++;
                                }
                                //System.out.println(name);
                            }
                        }
                        if( i == 3){
                           cell = row.getCell(30);
                           XSSFCellStyle cs = cell.getCellStyle();
                           cell = row.createCell(32);
                           cell.setCellValue("Additional Charge per language");
                           cell.setCellStyle(cs);
                        }
                    }
                }
                }catch(Exception e){
                    
                }
                System.out.println(langNumber);
                double total = Double.parseDouble(sheet.getRow(langNumber+3).getCell(29).getRawValue());
               
                double subTotal = total / langNumber;
                DecimalFormat df = new DecimalFormat("#.000");
                for(int i=0; i<langNumber; i++){
                    cell = sheet.getRow(i+4).createCell(32);
                    cell.setCellValue("$"+df.format(subTotal));
                }

                 file.close();
                 FileOutputStream outFile =new FileOutputStream(filePath);
                 workbook.write(outFile);
                 outFile.close();
                 rowNumber = 0;
                 langNumber = 0;
                 System.out.println("Done");
           }catch(Exception e){
               e.printStackTrace();
           }  
          }
          
          
    }
      
      public static Vector<String> readzipfile(String filepath){
        Vector<String> v = new Vector<String>();
        byte[] buffer = new byte[1024];
        String outputFolder = filepath.substring(0,filepath.lastIndexOf("."));
        try{
            File folder = new File(outputFolder);
            if(!folder.exists()){
		folder.mkdir();
	  }
			
                ZipInputStream zis = new ZipInputStream(new FileInputStream(filepath));
		ZipEntry ze = zis.getNextEntry();
                while(ze != null){
                    String fileName = ze.getName();
                    File subFolder = new File(outputFolder + "\\" + fileName.substring(0,fileName.lastIndexOf(".")));
                    if(!subFolder.exists()){
                        subFolder.mkdir();
                    }
		    File newFile = new File(outputFolder + "\\" + fileName.substring(0,fileName.lastIndexOf(".")) + "\\" + fileName);
		    v.addElement(newFile.getAbsolutePath());
		    FileOutputStream fos = new FileOutputStream(newFile);
		    int len;
		    while((len = zis.read(buffer)) > 0){
			fos.write(buffer, 0, len);
		     }
                fos.close();
                ze = zis.getNextEntry();
              }	 
	    zis.closeEntry();
	    zis.close();  
        }catch(Exception e){
            
        }
        return v;
    } 
      
      
      
}
