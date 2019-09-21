package br.com.rcc_dev.poi_template;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.Test;

public class AppTests {

	// =============================================

	@Test
	public void testarXLSX() throws InvalidFormatException, IOException {
		Workbook wb = WorkbookFactory.create( AppTests.class.getClassLoader().getResourceAsStream("excel.xlsx") );
    ExcelTemplate.create()
        .checkHeaders(false)
				.ref("TabelaPrecos")
				.formatRef("TabelaPrecosFormat")
				.headers("Nome", "Preço", "id", "CabeçalhoErrado")
				.add("Produto Um", 			15.99, 	1, null )
				.add("Produto Dois", 		35.50, 	2, null )
				.add("Produto Três", 	 	 9.99, 	3, null )
				.apply(wb);
		wb.write( Files.newOutputStream( Paths.get("target/final.xlsx") ) );
		wb.close();
	}

	@Test
	public void testarXLSXLibre() throws InvalidFormatException, IOException {
		Workbook wb = WorkbookFactory.create( AppTests.class.getClassLoader().getResourceAsStream("excel-libre.xlsx") );
		ExcelTemplate.create()
				.ref("TabelaPrecos")
				.formatRef("TabelaPrecosFormat")
				.headers("Nome", "Preço", "id", "Data")
				.add("Livro",         15.99, 	1, LocalDate.of(2019, 01, 25) )
				.add("Almoço", 		    35.50, 	2, LocalDate.of(2019, 04, 02) )
				.add("Refrigerante", 	9.99, 	3, LocalDate.of(2019, 06, 17) )
				.add("Refrigerante", 	-8.99, 	4, LocalDate.of(2019, 07, 17) )
				.apply(wb);
		wb.write( Files.newOutputStream( Paths.get("target/final-libre.xlsx") ) );
		wb.close();
	}

	@Test
	public void testarXLS() throws InvalidFormatException, IOException {
		Workbook wb = WorkbookFactory.create( AppTests.class.getClassLoader().getResourceAsStream("excel.xls") );
		ExcelTemplate.create()
				.ref("TabelaPrecos")
				.formatRef("TabelaPrecosFormat")
				.headers("Nome", "Preço", "id")
				.add("Produto Um", 			15.99, 	1)
				.add("Produto Dois", 		35.50, 	2)
				.add("Produto Três", 	 	 9.99, 	3)
				.apply(wb);
		wb.write( Files.newOutputStream( Paths.get("target/final.xls") ) );
		wb.close();
	}

}

