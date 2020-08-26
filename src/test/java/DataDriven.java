import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class DataDriven {
    public static void main(String[] args) throws IOException {
        String testPath = System.getProperty("user.dir") + "//test.xlsx";
        FileInputStream fis = new FileInputStream(testPath);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        int sheets = workbook.getNumberOfSheets();

        // Encontrando a celula com o texto "Testcase" na primeira célula de cada coluna do XSSF
        // dentro da tabela "test" (das várias do arquivo)
        for (int i = 0; i < sheets; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("test")) {
                XSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();

                Row firstRow = rows.next();

                Iterator<Cell> cells = firstRow.cellIterator();

                int columnIndex = 0;
                while (cells.hasNext()) {
                    Cell c = cells.next();
                    if (c.getStringCellValue().equalsIgnoreCase("Test")) {
                        break;
                    }

                    columnIndex++;
                }

                System.out.println(columnIndex);
            }
        }

        // Printando os valores da linha Purchase
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rows = sheet.iterator();

        DataFormatter dataFormatter = new DataFormatter();

        while (rows.hasNext()) {
            Row row = rows.next();

            //Usando a coluna de indice 4 encontrado no exemplo antigo
            if (row.getCell(0).getStringCellValue().equalsIgnoreCase("purchase")) {
                Iterator<Cell> cells = row.cellIterator();

                while (cells.hasNext()) {
                    System.out.println(cells.next().toString());
                    // O método abaixo tem que tratar o tipo dos dados da célula antes de utilizar o valor
//                    System.out.println(cells.next().getStringCellValue());

                    // Pode-se usar vários métodos como:
                    //  NumberToTextConverter.toText()
                    //  Cell.getCellTypeEnum() == CellType.STRING
                    //  Cell.getStringCellValue(), getNumericCellValue etc...
                }
            }
        }


        workbook.close();
    }
}
