
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import javafx.scene.control.Cell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author KantipudiChandraMouli
 */
public class AlbumDriver {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        // TODO code application logic here
    try{
            File f = new File("Kantipudi_Excel read write.xlsx");
            Workbook wb = new XSSFWorkbook(f);
            org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(0);
            ArrayList<Album> albums = new ArrayList<Album>();
            
            System.out.println(wb);
            
            for(int i = 1; i<sh.getPhysicalNumberOfRows();i++){
                Row r = sh.getRow(i);
                Album a = new Album(r.getCell(1).getStringCellValue(), 
                        r.getCell(2).getStringCellValue(), 
                        r.getCell(3).getStringCellValue(), 
                        r.getCell(4).getDateCellValue(), 
                        (int)r.getCell(5).getNumericCellValue());
                albums.add(a);
            }
            
            Collections.sort(albums);
            
            Workbook out = new XSSFWorkbook();
            org.apache.poi.ss.usermodel.Sheet s1 = out.createSheet("Output");
            
            Row rName = s1.createRow(1);
            
            org.apache.poi.ss.usermodel.Cell c = rName.createCell(1);
            c.setCellValue("Name");
            Font font = out.createFont();
            font.setBold(true);
            CellStyle style = out.createCellStyle();
            style.setFont(font);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
            
            style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            c.setCellStyle(style);
            
            c = rName.createCell(2);
            font = out.createFont();
            font.setItalic(true);
            font.setUnderline(Font.U_SINGLE);
            style = out.createCellStyle();
            style.setFont(font);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
            
            style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            c.setCellStyle(style);
            c.setCellValue("Kantipudi, Chandra");
            c = rName.createCell(3);
            c.setCellStyle(style);
            
            s1.addMergedRegion(new CellRangeAddress(1, 1, 2, 3));
            
            int rowIdx = 3;
            Row rOutput = s1.createRow(rowIdx++);
            Row header = sh.getRow(0);
            
            font = out.createFont();
            font.setBold(true);
            font.setColor(IndexedColors.WHITE.getIndex());
            
            style = out.createCellStyle();
            style.setFont(font);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
            style.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            
            c = rOutput.createCell(0);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(0).getStringCellValue());
            c = rOutput.createCell(1);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(2).getStringCellValue());
            c = rOutput.createCell(2);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(5).getStringCellValue());
            c = rOutput.createCell(3);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(1).getStringCellValue());
            c = rOutput.createCell(4);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(3).getStringCellValue());
            c = rOutput.createCell(5);
            c.setCellStyle(style);
            c.setCellValue(header.getCell(4).getStringCellValue());
            
            
            
            
            short col[] = new short[2];
            col[0] = IndexedColors.TAN.getIndex();
            col[1] = IndexedColors.LIGHT_GREEN.getIndex();
            
            String temp = albums.get(0).getGenre();
            int sno = 1;
            for(Album ab:albums){
                if(!temp.equals(ab.getGenre())){
                    short t = col[1];
                    col[1] = col[0];
                    col[0] = t;
                    temp = ab.getGenre();
                }
                style = out.createCellStyle();
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                style.setRightBorderColor(IndexedColors.BLACK.getIndex());
                style.setFillForegroundColor(col[0]);
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                
                rOutput = s1.createRow(rowIdx++);
                c = rOutput.createCell(0);
                c.setCellValue(sno++);
                c.setCellStyle(style);
                        
                c = rOutput.createCell(1);
                c.setCellValue(ab.getGenre());
                c.setCellStyle(style);
                
                c = rOutput.createCell(2);
                c.setCellValue(ab.getCriticScore());
                c.setCellStyle(style);
                
                c = rOutput.createCell(3);
                c.setCellValue(ab.getAlbumName());
                c.setCellStyle(style);
                
                c = rOutput.createCell(4);
                c.setCellValue(ab.getArtist());
                c.setCellStyle(style);
                
                CellStyle cellStyle = out.createCellStyle();
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
                cellStyle.setFillForegroundColor(col[0]);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setDataFormat(
                out.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
                c = rOutput.createCell(5);
                c.setCellValue(ab.getReleasedate());
                c.setCellStyle(cellStyle);
            }
            
            s1.autoSizeColumn(0);
            s1.autoSizeColumn(1);
            s1.autoSizeColumn(2);
            s1.autoSizeColumn(3);
            s1.autoSizeColumn(4);
            s1.autoSizeColumn(5);
            out.write(new FileOutputStream("Kantipudi_Output.xlsx"));
        } catch(InvalidFormatException ex) {
            System.out.println("The input file provided is in invalid format or currupted.");
        }
    }
    
    
}
