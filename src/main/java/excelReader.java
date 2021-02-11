import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;


public class excelReader {
    private FileInputStream inputstream;
    private XSSFWorkbook xssfWorkbook;
    private int maxRow;
    private XSSFSheet sheet;
    private Row titleRow;
    private int maxColumn;
    private Map<String, Integer> mapTitle;
    private ArrayList<ArrayList<String>> collection_list;
    private ArrayList<ArrayList<String>> require_list;
    excelReader(FileInputStream is) {
        this.inputstream = is;

    }

    /**
     *
     * @param is
     */
    public void read(FileInputStream is) {
        try {
            this.xssfWorkbook = new XSSFWorkbook(is);
            this.sheet = xssfWorkbook.getSheetAt(0);
            this.titleRow = this.sheet.getRow(0);
            mappingTitle();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @return maxRow of the sheet
     */
    public int maxRow() {
        this.maxRow = this.sheet.getLastRowNum();
        return this.maxRow;
    }

    /**
     *
     * @return max column number of the sheet
     */

    public int maxColumn() {
        this.maxColumn = this.titleRow.getLastCellNum();
        return this.maxColumn;
    }

    /**
     *
     * @param row_num
     * @param column_num
     * @return cell data
     */

    public String get_Cell(int row_num, int column_num) {
        String cell = new String();
        XSSFRow row = this.sheet.getRow(row_num);
        XSSFCell cell0 = row.getCell(column_num);
        return cell;
    }

    /**
     * mappingTitle to column
     */
    public void mappingTitle(){
         this.mapTitle= new HashMap<String,Integer>();
        XSSFRow row = this.sheet.getRow(0);//Create map
        short minColIx = row.getFirstCellNum(); //get the first column index for a row
        short maxColIx = row.getLastCellNum();
        for(short colIx=minColIx; colIx<maxColIx;) { //loop from first to last index
            XSSFCell cell = row.getCell(colIx); //get the cell
           // System.out.println(cell.getCellType());

            if(cell.getCellType()==CellType.BLANK){
                System.out.println("Cell is null");
                colIx++;
                continue;
            }
            //System.out.println(cell.getStringCellValue()+": "+cell.getColumnIndex());
            mapTitle.put(cell.getStringCellValue(),cell.getColumnIndex());
            colIx++;//add the cell contents (name of column) and cell index to the map
        }//int idx = map.get("ColumnName");

    }

    /**
     *
     * @param columnName
     * @return get map title
     */
    public int getMapTitle(String columnName){
        return this.mapTitle.get(columnName);
    }

    /**
     *
     * @return Store the entire sheet row by row
     */
    public ArrayList storeSheet(){
        this.collection_list = new ArrayList<ArrayList<String>>();
        //process each row
        for (int i=1; i<this.sheet.getPhysicalNumberOfRows();i++){
            Row row =this.sheet.getRow(i);
            int minCells=row.getFirstCellNum();
            int maxCells=row.getLastCellNum();
            ArrayList<String> temp_row_list = new ArrayList<String>();
            for (int k = 1; k < 4; k++) {
                Cell cell=row.getCell(k);
                if(cell == null){
                    temp_row_list.add(null);
                    continue;
                }
                if(cell.getCellType()== CellType.NUMERIC){
                    temp_row_list.add(cell.toString());
                }else if((cell.getCellType() == CellType.STRING)){
                    temp_row_list.add(cell.toString());
                }else if((cell.getCellType()==CellType.BLANK)){
                    temp_row_list.add(null);

                }
            }
            this.collection_list.add(temp_row_list);
        }
        return this.collection_list;
    }

    /**
     *
     * @param rowNum
     * @return Arraylist of the SN,DI,Instrument name
     */
    public ArrayList processRow(int rowNum){
        Row row =this.sheet.getRow(rowNum);
        ArrayList<String> temp_row_list = new ArrayList<String>();
        for (int k = 1; k < 4; k++) {
            Cell cell=row.getCell(k);

            if(cell == null){
                temp_row_list.add(null);
                continue;
            }

            if(cell.getCellType()== CellType.NUMERIC){
                temp_row_list.add(cell.toString());
            }else if((cell.getCellType() == CellType.STRING)){
                temp_row_list.add(cell.toString());
            }else if((cell.getCellType()==CellType.BLANK)){
                temp_row_list.add(null);
            }
        }
        return temp_row_list;
    }


}




