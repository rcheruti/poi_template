package br.com.rcc_dev.poi_template;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class AppTests {

	// =============================================

	@Test
	public void testar() throws InvalidFormatException, IOException {
		Workbook wb = new XSSFWorkbook(Paths.get("excel.xlsx").toFile());

		try {
			ExcelTemplate.create()
				.ref("TabelaPrecos")
				.formatRef("TabelaPrecosFormat")
				.headers("Nome", "Preço", "id")
				.add("Produto Um", 			15.99, 	1)
				.add("Produto Dois", 		35.50, 	2)
				.add("Produto Três", 	 	 9.99, 	3)
				.apply(wb);
		} catch (Exception ex) {
			log.error("Error!", ex);
			throw ex;
		}
		

		wb.write( Files.newOutputStream( Paths.get("target/final.xlsx") ) );
		wb.close();
	}

}

