package br.com.rcc_dev.poi_template;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.formula.BaseFormulaEvaluator;
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
  private List<Object[]> data;
  private boolean checkHeaders = true;
  private boolean formulaRecalculation = true;

  // -----------------------------

  private ExcelTemplate(int initialCapacity) {
    this.data = new ArrayList<>( initialCapacity );
  }

  public static ExcelTemplate create() {
    return new ExcelTemplate( 1000 );
  }
  public static ExcelTemplate create(int initialCapacity) {
    return new ExcelTemplate( initialCapacity );
  }

  // -----------------------------

  public String[] headers() { return this.headers; }
  public ExcelTemplate headers(String... headers) {
    this.headers = headers;
    return this;
  }

  public String ref() { return this.ref; }
  public ExcelTemplate ref(String ref) {
    this.ref = ref;
    return this;
  }

  public String formatRef() { return this.formatRef; }
  public ExcelTemplate formatRef(String formatRef) {
    this.formatRef = formatRef;
    return this;
  }

  public List<Object[]> data() { return this.data; }
  public ExcelTemplate data(List<Object[]> data) {
    this.data = data;
    return this;
  }
  public ExcelTemplate add(Object... dataLine) {
    this.data.add( dataLine );
    return this;
  }
  public ExcelTemplate addAll(List<Object[]> data) {
    this.data.addAll( data );
    return this;
  }

  public boolean checkHeaders() { return this.checkHeaders; }
  public ExcelTemplate checkHeaders(boolean checkHeaders) {
    this.checkHeaders = checkHeaders;
    return this;
  }
  public boolean formulaRecalculation() { return this.formulaRecalculation; }
  public ExcelTemplate formulaRecalculation(boolean formulaRecalculation) {
    this.formulaRecalculation = formulaRecalculation;
    return this;
  }

  // ---------------------------

  public ExcelTemplate apply(Workbook wb) {
    AreaReference aref = new AreaReference( wb.getName( this.ref ).getRefersToFormula(), wb.getSpreadsheetVersion() );
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
      AreaReference fref = new AreaReference( wb.getName( this.formatRef ).getRefersToFormula(), wb.getSpreadsheetVersion() );
      formatRow = sheet.getRow( fref.getFirstCell().getRow() );
      formatFirst = fref.getFirstCell();
    }

    // find headers positions
    CellReference[] headersRefsCells = aref.getAllReferencedCells();
    for (int i = 0; i < this.headers.length; i++) {
      String header = this.headers[i];
      int colIdx = -1;

      // try find one Name Reference for this header first
      Name named = wb.getName( header );
      if ( named != null ) {
        AreaReference cellAreaRef = new AreaReference(named.getRefersToFormula(), wb.getSpreadsheetVersion() );
        colIdx = cellAreaRef.getFirstCell().getCol();
      } else {
        // try find cell by cell
        for (CellReference cellRef : headersRefsCells) {
          Cell cell = row.getCell( cellRef.getCol() );
          CellType type = cell.getCellType();
          String stringValue = cell.getStringCellValue();
          if ( CellType.STRING.equals( type ) && header.equalsIgnoreCase( stringValue ) ) {
            colIdx = cell.getColumnIndex();
            break;
          }
        }
      }
      
      if ( colIdx < 0 ) {
        continue; // header not found, check for missing headers is after all tries to find
      }
      mapCols.put( header , new int[]{ colIdx , i } );
      // store formats for cells
      if ( formatRow != null ) {
        int styleIdx = formatFirst.getCol() + ( colIdx - aref.getFirstCell().getCol() ) ;
        mapStyles.put(header, formatRow.getCell( styleIdx ).getCellStyle() );
      }
    }

    // check that all headers was found
    if( this.checkHeaders ) {
      StringBuilder sb = new StringBuilder( 500 );
      for(String header : this.headers) {
        if( !mapCols.containsKey(header) ) {
          sb.append(header).append(", ");
        }
      }
      if( sb.length() > 0 ) { // some headers not found
        sb.delete( sb.length() -2, sb.length() );
        throw new IllegalArgumentException(
          String.format("The \"%s\" headers was not found in the SpreadSheet! Review your Template!", sb.toString() ));
      }
    }
    

    // put data
    int line = rowIdx + 1;
    for ( Object[] dados : this.data ) {
      row = sheet.getRow( line );
      if ( row == null ) {
        row = sheet.createRow( line );
      }

      for ( String header : mapCols.keySet() ) {
        int[] cols = mapCols.get( header );
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
        } else if ( obj instanceof LocalDate ) { // JDK 1.8
          cell.setCellValue( Date.from( ((LocalDate)obj).atStartOfDay(ZoneId.systemDefault()).toInstant() ) );
        } else if ( obj instanceof LocalDateTime ) { // JDK 1.8
          cell.setCellValue( Date.from( ((LocalDateTime)obj).atZone(ZoneId.systemDefault()).toInstant() ) );
        } else if ( obj instanceof Instant ) { // JDK 1.8
          cell.setCellValue( Date.from( ((Instant)obj) ) );
        } else {
          cell.setCellValue( obj.toString() );
        }

        // set format
        CellStyle style = mapStyles.get( header );
        if ( style != null ) {
          cell.setCellStyle(style);
        }
      }
      line++; // next line !
    }

    wb.setForceFormulaRecalculation( this.formulaRecalculation );
    if( this.formulaRecalculation ) {
      BaseFormulaEvaluator.evaluateAllFormulaCells(wb);
    }
    return this;
  }

}