import java.io.*;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class xlsFormatter {
   HSSFSheet sheet;
   HSSFRow row;
   Cell cell;
   int startDate, endDate;
   
   public void openWorkbook(String path){
      try{
         File file = new File(path);
         FileInputStream inputStream = new FileInputStream(file);
         HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
         this.sheet = workbook.getSheetAt(0);
         System.out.println(workbook.toString());
         inputStream.close();
      }
      catch(IOException e){}
      //catch(FileNotFoundException fnfE){}
   }
   
   public void getSheetMatrix(){
      Iterator <Row> rowIterator = sheet.iterator(); //Row Iterator
      
      while (rowIterator.hasNext()){//Traverse Rows
         row = (HSSFRow) rowIterator.next();
         Iterator <Cell> cellIterator = row.cellIterator(); //Cell Iterator
        
         cell = cellIterator.next();//Skip first cell
        
         if (cell.getStringCellValue().contains("FRANK")){
            while(cellIterator.hasNext()){//Traverse Cells within Row
               cell = cellIterator.next();
               if(cell.getCellType() == Cell.CELL_TYPE_STRING){
                  print(cell.getStringCellValue());
               range(cell.getStringCellValue());
               if (cell.getStringCellValue().contains("FRANK"))
                  print(cell.getStringCellValue());
               }
            }
             System.out.println();
         }
        
      }
   }
    
            
//                   cell = cellIterator.next();//Skip second cell
//             for(int i=0;i<=8;i++){
//                
//                cell = cellIterator.next();
//                print(cell.getStringCellValue());
//             }
//             System.out.println();
            
//         while(cellIterator.hasNext()){//Traverse Cells within Row
//             cell = cellIterator.next();
//             
//             if(cell.getCellType() == Cell.CELL_TYPE_STRING){
//                //print(cell.getStringCellValue());
//                range(cell.getStringCellValue());
//                if (cell.getStringCellValue().contains("FRANK"))
//                   print(cell.getStringCellValue());
//             }
//             if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
//                //print(cell.getNumericCellValue());
//             } 
//         }
        //print("\n");

   //Computes date range based on found numbers in cells
   private void range(String cellVal)throws NumberFormatException{
      try{
         //Convert to integer
         int value = Integer.parseInt(cellVal);
         //set start date if not initialized and skip if defined
         if(startDate == 0)
            startDate = value;
         else if(value>startDate)
            endDate = value;
      }
      catch(NumberFormatException nfE){}
   }
   
   private void print(Object obj){
      if(obj!=null)
         System.out.print(obj.toString() + " ");
   }
   
   public static void main(String[] args){
      xlsFormatter formatter = new xlsFormatter();
      formatter.openWorkbook("attachment.xls");
      formatter.getSheetMatrix();
   }
}