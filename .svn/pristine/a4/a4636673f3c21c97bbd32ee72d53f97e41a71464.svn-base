package cl.intelidata.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelLib {

	/**
	 * Obtiene el valor de una celda según su tipo
	 * @param cell
	 * @return
	 */
	public static String readCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_ERROR:
			return String.valueOf(cell.getErrorCellValue());
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue().toString();
			} else {
				return Double.toString(cell.getNumericCellValue());
			}
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		default:
			return "Unknown type";
		}
	}

	/**
	 * Copy una fila de una posición a otra
	 * @param worksheet
	 * @param sourceRowNum
	 * @param destinationRowNum
	 * 
	 *            Usage: HSSFWorkbook workbook = new HSSFWorkbook(new
	 *            FileInputStream("c:/input.xls")); HSSFSheet sheet =
	 *            workbook.getSheet("Sheet1"); copyRow(workbook, sheet, 0, 1);
	 *            FileOutputStream out = new FileOutputStream("c:/output.xls");
	 *            workbook.write(out); out.close();
	 */
	private static void copyRow(Sheet worksheet, int sourceRowNum, int destinationRowNum) {
		Row newRow = worksheet.getRow(destinationRowNum);
		Row sourceRow = worksheet.getRow(sourceRowNum);

		if (newRow != null) {
			worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
		} else {
			newRow = worksheet.createRow(destinationRowNum);
		}

		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			Cell oldCell = sourceRow.getCell(i);
			Cell newCell = newRow.createCell(i);

			if (oldCell == null) {
				newCell = null;
				continue;
			}

			newCell.setCellStyle(oldCell.getCellStyle());

			if (newCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			newCell.setCellType(oldCell.getCellType());

			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}
		}
	}

}
