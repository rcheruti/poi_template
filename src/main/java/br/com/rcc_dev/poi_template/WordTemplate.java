package br.com.rcc_dev.poi_template;


import java.util.Collection;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordTemplate extends HashMap<String, String>{
    
    private static final long serialVersionUID = 1L;
    
    private XWPFDocument docx = null;
    private Pattern pattern = null ;
    private int regexFlags = 0 ;
    private final String regexChars = "([$\\(\\)\\[\\].^])";
    
    public WordTemplate(){}
    public WordTemplate(XWPFDocument docx){ this(docx, 0 ); }
    public WordTemplate(XWPFDocument docx, int flags){ this.docx = docx; this.regexFlags = flags; }
    
    
    public void apply(){
        apply(false);
    }
    public void apply(boolean recompile){
        if( recompile || pattern == null ){
            StringBuilder strB = new StringBuilder(400);
            strB.append("(");
            for( Entry<String,String> entry : this.entrySet() ){
                strB.append( entry.getKey().replaceAll( regexChars , "\\\\$1") );
                strB.append("|");
            }
            strB.deleteCharAt( strB.length()-1 );
            strB.append(")");
            //System.out.println("Compilado: "+ strB );
            pattern = Pattern.compile( strB.toString() , regexFlags );
        }
        for( XWPFHeader x : this.docx.getHeaderList() )  _apply( x.getParagraphs() );
        for( XWPFFooter x : this.docx.getFooterList() )  _apply( x.getParagraphs() );
        for( XWPFTable x : this.docx.getTables() )
          for( XWPFTableRow r : x.getRows() )
            for( XWPFTableCell c : r.getTableCells() )
              _apply( c.getParagraphs() );
        _apply( this.docx.getParagraphs() );
    }
    
    //============== Privates  =========================
    private void _apply(Collection<XWPFParagraph> pars){
        for( XWPFParagraph par : pars ) _apply( par );
    }
    private void _apply(XWPFParagraph par){
        String text = par.getText() ;
        Matcher m = pattern.matcher( text );
        PositionInParagraph pos = new PositionInParagraph();
        XWPFRun run = null ;
        while( m.find() ){
            String key = m.group(1);
            String value = this.get(key);
            TextSegment seg = par.searchText(key, pos );
            
            if( seg.getBeginRun() == seg.getEndRun() ){
                run = par.getRuns().get( seg.getBeginRun() );
                String str = run.getText(0);
                run.setText( str.replace(key, value) ,0);
            }else{
                run = par.getRuns().get( seg.getEndRun() );
                String str = run.getText(0);
                str = value+ str.substring( seg.getEndChar()+1 , str.length() );
                run.setText(str, 0);
                
                run = par.getRuns().get( seg.getBeginRun() );
                run.setText( run.getText(0).substring(0, seg.getBeginChar() ) ,0 );
                
                for(int i = seg.getBeginRun()+1 ; i < seg.getEndRun(); i++ ){
                    // Essa exceção é pq remover Hyperlink ainda não tem suporte na lib POI
                    try{ par.removeRun(i); }catch(IllegalArgumentException ex){  }
                    
                }
                try{ if( run.getText(0).length() < 1 ) par.removeRun( seg.getBeginRun() ); }
                catch(IllegalArgumentException ex){  }
            }
            pos = seg.getEndPos();
        }
    }
    
    
    //============  Setters e Getters  =================
    public WordTemplate docx(XWPFDocument docx){ this.docx = docx; return this; }
    public XWPFDocument docx(){ return this.docx; }
    
    public WordTemplate regexFlags(int flags){ this.regexFlags = flags; return this; }
    public int regexFlags(){ return this.regexFlags; }
    
    
}