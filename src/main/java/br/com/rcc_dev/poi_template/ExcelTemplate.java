package br.com.rcc_dev.poi_template;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;


public class ExcelTemplate {

  private String ref;
  private String formatRef;
  private String[] headers;
  private List<Object[]> data = new ArrayList<>(200);

  // -----------------------------

  private ExcelTemplate(){}

  public static ExcelTemplate create(){
    return new ExcelTemplate();
  }

  // -----------------------------

  public String[] headers(){ return this.headers; }
  public ExcelTemplate headers(String... headers){
    this.headers = headers;
    return this;
  }

  public String ref(){ return this.ref; }
  public ExcelTemplate ref(String ref){
    this.ref = ref;
    return this;
  }

  public String formatRef(){ return this.formatRef; }
  public ExcelTemplate formatRef(String formatRef){
    this.formatRef = formatRef;
    return this;
  }

  public List<Object[]> data(){ return this.data; }
  public ExcelTemplate data(List<Object[]> data){
    this.data = data;
    return this;
  }
  public ExcelTemplate add(Object... dataLine){
    this.data.add( dataLine );
    return this;
  }
  public ExcelTemplate addAll(List<Object[]> data){
    this.data.addAll( data );
    return this;
  }

  public ExcelTemplate apply(Workbook wb) {
    AreaReference aref = new AreaReference( wb.getName( this.ref ).getRefersToFormula(), SpreadsheetVersion.EXCEL2007);
    CellReference first = aref.getFirstCell();
    int rowIdx = first.getRow();
    Sheet sheet =  wb.getSheet( first.getSheetName() );
    Row row = sheet.getRow( rowIdx );
    Map<String, int[]> mapCols = new HashMap<>(); // map: Header -> int[ Sheet col , Data col ]
    Map<String, CellStyle> mapStyles = new HashMap<>(); // map: Header -> CellStyle
    CellReference formatFirst = null;
    Row formatRow = null;

    // load formats ranges for the cells
    if ( this.formatRef != null && !this.formatRef.isEmpty() ) {
      AreaReference fref = new AreaReference( wb.getName( this.formatRef ).getRefersToFormula(), SpreadsheetVersion.EXCEL2007);
      formatRow = sheet.getRow( fref.getFirstCell().getRow() );
      formatFirst = fref.getFirstCell();
    }

    // find headers positions
    for (CellReference cellRef : aref.getAllReferencedCells()) {
      if ( cellRef.getRow() != rowIdx ) break;
      Cell cell = row.getCell( cellRef.getCol() );
      
      for (int i = 0; i < this.headers.length; i++) {
        Name named = wb.getName( this.headers[i] );
        int colIdx = -1;
        if ( named != null ) {
          AreaReference cellAreaRef = new AreaReference(named.getRefersToFormula(), SpreadsheetVersion.EXCEL2007);
          colIdx = cellAreaRef.getFirstCell().getCol();
        }
        if ( colIdx < 0 && CellType.STRING.equals( cell.getCellType() ) && this.headers[i].equals( cell.getStringCellValue() ) ) {
          colIdx = cell.getColumnIndex();
        }
        if ( colIdx < 0 ) continue; // not found yet
        mapCols.put( this.headers[i] , new int[]{ colIdx , i } );
        // store formats for cells
        if ( formatRow != null ) {
          int styleIdx = formatFirst.getCol() + ( colIdx - aref.getFirstCell().getCol() ) ;
          mapStyles.put(this.headers[i], formatRow.getCell( styleIdx ).getCellStyle() );
        }
        break;
      }
    }

    // put data
    int line = rowIdx + 1;
    for ( Object[] dados : this.data ) {
      row = sheet.getRow( line );
      if ( row == null ) {
        row = sheet.createRow( line );
      }

      for (int i = 0; i < this.headers.length; i++) {
        int[] cols = mapCols.get( this.headers[i] );
        Object obj = dados[ cols[1] ];
        Cell cell = row.getCell( cols[0] );
        if ( cell == null ) {
          cell = row.createCell( cols[0] );
        }
        
        // set data
        if ( obj instanceof Number ) {
          cell.setCellValue( ((Number)obj).doubleValue() );
        } else if ( obj instanceof String ) {
          cell.setCellValue( ((String)obj) );
        } else if ( obj instanceof Date ) {
          cell.setCellValue( ((Date)obj) );
        } else if ( obj instanceof Calendar ) {
          cell.setCellValue( ((Calendar)obj) );
        } else {
          cell.setCellValue( obj.toString() );
        }

        // set format
        CellStyle style = mapStyles.get( this.headers[i] );
        if ( style != null ) {
          cell.setCellStyle(style);
        }
      }
      line++; // next line !
    }
    return this;
  }

}