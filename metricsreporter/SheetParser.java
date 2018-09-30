package metricsreporter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class SheetParser {
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    XSSFRow row;
    Cell cell;

    public SheetParser(String workbookPath,int sheetIndex){
        try{
            FileInputStream inputStream = new FileInputStream(new File(workbookPath));
            this.workbook = new XSSFWorkbook(inputStream);
            this.sheet = workbook.getSheetAt(sheetIndex);
            //LOG(workbook.toString());
            //LOG(sheet.getSheetName());
            inputStream.close();
        }
        catch(IOException e){e.printStackTrace();}
    }

    //Returns a collection of rows that match a specified filter
    private ArrayList<Row> rowFilter(String filter){
        ArrayList<Row> filteredRows = new ArrayList<>();
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter formatter = new DataFormatter();
        String cellValue;
        for (Row row:sheet){
            for(Cell cell:row) {
                cellValue = formatter.formatCellValue(cell);
                if(cellValue.contains(filter) && cell.getColumnIndex()==1)//added condition to restrict cell column to index 1
                    filteredRows.add(row);
            }
        }
        for(Row rw:filteredRows){
            for(Cell cl:rw){
                if(cl.getCellType() == Cell.CELL_TYPE_STRING){//If String
                    cellValue=cl.getStringCellValue();
                    LOG(cellValue + " ");
                }
                else if (cl.getCellType() == Cell.CELL_TYPE_NUMERIC){
                    LOG(String.format("%.2f | ",cl.getNumericCellValue()));
                }
            }
            LOG("\n");
        }
        return filteredRows;
    }

    //Deprecated Function. While loop can be replaced with foreach
    public void getSheetMatrix(){
        Iterator <Row> rowIterator = sheet.iterator(); //Row Iterator

        while (rowIterator.hasNext()){//Traverse Rows
            row = (XSSFRow) rowIterator.next();
            Iterator <Cell> cellIterator = row.cellIterator(); //Cell Iterator
            cell = cellIterator.next();//Skip first cell
            if (cell.getStringCellValue().contains("")){
                while(cellIterator.hasNext()){//Traverse Cells within Row
                    cell = cellIterator.next();
                    if(cell.getCellType() == Cell.CELL_TYPE_STRING){
                        LOG(cell.getStringCellValue());
                        if (cell.getStringCellValue().contains(""))
                            LOG(cell.getStringCellValue());
                    }
                }
                System.out.println();
            }
        }
    }

    //Debug logger
    public void LOG(Object arg){ System.out.print(arg.toString());}

    public static void main(String[] args){
        SheetParser sheet = new SheetParser("recentAttachment.xlsx",2);
        sheet.rowFilter("447");
    }
}