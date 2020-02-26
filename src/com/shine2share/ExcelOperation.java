package com.shine2share;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetViews;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import lombok.Getter;
import lombok.Setter;

public class ExcelOperation {

	@Getter
	@Setter
	private XSSFWorkbook workbook;
	@Getter
	@Setter
	private XSSFSheet sheet;

	private static List<String> listFiles(String linkFolder) {
		List<String> listFile = new ArrayList<>();
		try (Stream<Path> paths = Files.walk(Paths.get(linkFolder))) {
			paths.filter(Files::isRegularFile).forEach(filePath -> listFile.add(filePath.toString()));
		} catch (Exception e) {
			System.out.println("Lỗi khi đọc file name trong folder - " + e.getMessage());
		}
		return listFile;
	}

	private List<InputStream> initialize(String linkFi) {
		List<InputStream> files = new ArrayList<>();
		InputStream file = null;
		try {
			List<String> filePath = listFiles(linkFi);
			int i;
			for (i = 0; i < filePath.size(); ++i) {
				file = new FileInputStream(new File(filePath.get(i)));
				files.add(file);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return files;
	}

	public String formatA1(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					this.sheet = this.workbook.getSheetAt(i);
					if (sheet instanceof XSSFSheet) {
						CTWorksheet ctWorksheet = null;
						CTSheetViews ctSheetViews = null;
						CTSheetView ctSheetView = null;
						XSSFSheet tempSheet = (XSSFSheet) sheet;
						// First step is to get at the CTWorksheet bean underlying the worksheet.
						ctWorksheet = tempSheet.getCTWorksheet();
						// From the CTWorksheet, get at the sheet views.
						ctSheetViews = ctWorksheet.getSheetViews();
						// Grab a single sheet view from that array
						ctSheetView = ctSheetViews.getSheetViewArray(ctSheetViews.sizeOfSheetViewArray() - 1);
						// Se the address of the top left hand cell.
						ctSheetView.setTopLeftCell("A1");
						sheet.setActiveCell("A1");
					} else {
						sheet.setActiveCell("A1");
						sheet.showInPane((short) 0, (short) 0);
					}
				}
				this.workbook.setActiveSheet(0);
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Format A1: FAIL, Check lại định dạng file excel (chỉ support .xlsx)";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Format A1: FAIL";
		}

		return "Format A1: SUCCESS";
	}

	public String setSheetColor(String linkFi, String colorValue) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		int color = 0;
		switch (colorValue) {
		case "No_Color":
			color = -1;
			break;
		case "BLACK":
			color = 0;
			break;
		case "WHITE":
			color = 1;
			break;
		case "RED":
			color = 2;
			break;
		case "GREEN":
			color = 3;
			break;
		case "BLUE":
			color = 4;
			break;
		case "YELLOW":
			color = 5;
			break;
		default:
			color = 6;
			break;
		}
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					this.sheet = this.workbook.getSheetAt(i);
					this.sheet.setTabColor(color);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Set Color sheet: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Set Color sheet: FAIL";
		}
		return "Set Color sheet: SUCCESS";
	}

	public String setSheetFont(String linkFi, String fontValue) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				XSSFFont font = this.workbook.createFont();
				font.setFontName(fontValue);

				CellStyle style = this.workbook.createCellStyle();
				style.setFont(font);

				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					for (Row row : this.workbook.getSheetAt(i)) {
						for (Cell cell : row) {
							cell.setCellStyle(style);
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Set sheet font: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Set sheet font: FAIL";
		}
		return "Set sheet font: SUCCESS";
	}

	public String setSheetSize(String linkFi, short sizeValue) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				XSSFFont font = this.workbook.createFont();
				font.setFontHeightInPoints(sizeValue);

				CellStyle style = this.workbook.createCellStyle();
				style.setFont(font);

				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					for (Row row : this.workbook.getSheetAt(i)) {
						for (Cell cell : row) {
							cell.setCellStyle(style);
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Set sheet size: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Set sheet size: FAIL";
		}
		return "Set sheet size: SUCCESS";
	}

	public List<String> searchSpecWord(String linkFi, String searchWord) {
		List<String> response = new ArrayList<>();
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);

				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					for (Row row : this.workbook.getSheetAt(i)) {
						for (Cell cell : row) {
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
								if (cell.getRichStringCellValue().getString().trim().equals(searchWord)) {
									response.add("Hàng thứ: " + (row.getRowNum() + 1) + "; Cột thứ:"
											+ (cell.getColumnIndex() + 1) + "; Sheet: "
											+ this.workbook.getSheetName(i));
								}
							} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								if (searchWord.matches("(.*)/(.*)")) {
									if (cell.getDateCellValue().toString().equals(searchWord)) {
										System.out.println("equlas");
									}
								} else {
									if ((cell.getNumericCellValue()) == Double.parseDouble(searchWord)) {
										response.add("Hàng thứ: " + (row.getRowNum() + 1) + "; Cột thứ:"
												+ (cell.getColumnIndex() + 1) + "; Sheet: "
												+ this.workbook.getSheetName(i));
									}
								}
							} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
								System.out.println("formula: contact thelam92@gmail.com để xử lý");
							} else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
								System.out.println("blank: contact thelam92@gmail.com để xử lý");
							} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
								System.out.println("boolean: contact thelam92@gmail.com để xử lý");
							} else {
								System.out.println("Date: " + cell.getDateCellValue().toString());
								if (cell.getDateCellValue().toString().trim().equals(searchWord)) {
									response.add("Hàng thứ: " + (row.getRowNum() + 1) + "; Cột thứ:"
											+ (cell.getColumnIndex() + 1) + "; Sheet: "
											+ this.workbook.getSheetName(i));
								}
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return response;
	}

	public String setSheetZoom(String linkFi, int zoomValue) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					XSSFSheet sheet = this.workbook.getSheetAt(i);
					sheet.setZoom(zoomValue);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Set sheet zoom: FAIL, Xem lại khoảng zoom (10-400)";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Set sheet zoom: FAIL";
		}
		return "Set sheet zoom: SUCCESS";
	}

	private boolean write2File(String linkFi, List<XSSFWorkbook> workbooks) {
		try {
			List<String> filePath = listFiles(linkFi);
			int i;
			FileOutputStream out = null;
			for (i = 0; i < filePath.size(); ++i) {
				out = new FileOutputStream(new File(filePath.get(i)));
				workbooks.get(i).write(out);
				out.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// new update
	/**
	 * Delete #Value after run GOOGLETRANSLATE in google docs app
	 * 
	 * @param linkFi
	 * @return
	 */
	public String deleteValue(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 0; i < numberOfSheet; ++i) {
					if (this.workbook.getSheetAt(i).getSheetName().contains("Data")) {
						continue;
					}
					for (Row row : this.workbook.getSheetAt(i)) {
						for (Cell cell : row) {
							if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
								switch (cell.getCachedFormulaResultType()) {
								case Cell.CELL_TYPE_STRING:
									if ("#VALUE!".equals(cell.getRichStringCellValue().getString())) {
										cell.setCellType(Cell.CELL_TYPE_STRING);
										cell.setCellValue("");
									}
									break;
								}
							}
						}

					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Delete #Value: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Delete #Value: FAIL";
		}
		return "Delete #Value: SUCCESS";
	}

///////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
/////////////////////////////////insert new column///////////////////////////////////////////////
	/**
	 * insert new column before D column with title テストケース and change title of E
	 * (after insert) to Testcase
	 * 
	 * @param linkFi
	 * @return
	 */
	public String insertJpColumn(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);

				FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
				evaluator.clearAllCachedResultValues();

				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				// the first 2 sheet are not need to insert
				for (i = 2; i < numberOfSheet; ++i) {
					Sheet sheet = workbook.getSheetAt(i);
					int nrRows = getNumberOfRows(i);
					int nrCols = getNrColumns(i);
					System.out.println("Inserting new column...");
					for (int row = 0; row < nrRows; row++) {
						Row r = sheet.getRow(row);

						if (r == null) {
							continue;
						}

						// shift to right
						// columnIndex = 2
						for (int col = nrCols; col > 2; col--) {
							Cell rightCell = r.getCell(col);
							if (rightCell != null) {
								r.removeCell(rightCell);
							}

							Cell leftCell = r.getCell(col - 1);

							if (leftCell != null) {
								Cell newCell = r.createCell(col, leftCell.getCellType());
								cloneCell(newCell, leftCell);
								if (newCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
									newCell.setCellFormula(ExcelHelper.updateFormula(newCell.getCellFormula(), 2));
									evaluator.notifySetFormula(newCell);
									CellValue cellValue = evaluator.evaluate(newCell);
									evaluator.evaluateFormulaCell(newCell);
									System.out.println(cellValue);
								}
							}
						}

						// delete old column
						int cellType = Cell.CELL_TYPE_BLANK;

						Cell currentEmptyWeekCell = r.getCell(2);
						if (currentEmptyWeekCell != null) {
//							cellType = currentEmptyWeekCell.getCellType();
							r.removeCell(currentEmptyWeekCell);
						}

						// create new column
						r.createCell(2, cellType);
					}

					// Adjust the column widths
					for (int col = nrCols; col > 2; col--) {
						sheet.setColumnWidth(col, sheet.getColumnWidth(col - 1));
					}

					// currently updates formula on the last cell of the moved column
					// TODO: update all cells if their formulas contain references to the moved cell
//					Row specialRow = sheet.getRow(nrRows-1);
//					Cell cellFormula = specialRow.createCell(nrCols - 1);
//					cellFormula.setCellType(XSSFCell.CELL_TYPE_FORMULA);
//					cellFormula.setCellFormula(formula);

					XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "Insert JP Column: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "Insert JP Column: FAIL";
		}
		return "Insert JP Column: SUCCESS";
	}

	private int getNumberOfRows(int sheetIndex) {
		assert workbook != null;

		int sheetNumber = workbook.getNumberOfSheets();

		System.out.println("Found " + sheetNumber + " sheets.");

		if (sheetIndex >= sheetNumber) {
			throw new RuntimeException("Sheet index " + sheetIndex + " invalid, we have " + sheetNumber + " sheets");
		}

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		int rowNum = sheet.getLastRowNum() + 1;

		System.out.println("Found " + rowNum + " rows.");

		return rowNum;
	}

	private int getNrColumns(int sheetIndex) {
		assert workbook != null;

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		// get header row
		Row headerRow = sheet.getRow(0);
		int nrCol = headerRow.getLastCellNum();

		// while
		// (!headerRow.getCell(nrCol++).getStringCellValue().equals(LAST_COLUMN_HEADER));

		// while (nrCol <= headerRow.getPhysicalNumberOfCells()) {
		// Cell c = headerRow.getCell(nrCol);
		// nrCol++;
		//
		// if (c!= null && c.getCellType() == Cell.CELL_TYPE_STRING) {
		// if (c.getStringCellValue().equals(LAST_COLUMN_HEADER)) {
		// break;
		// }
		// }
		// }

		System.out.println("Found " + nrCol + " columns.");
		return nrCol;

	}

	/*
	 * Takes an existing Cell and merges all the styles and forumla into the new one
	 */
	private static void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		switch (cOld.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN: {
			cNew.setCellValue(cOld.getBooleanCellValue());
			break;
		}
		case Cell.CELL_TYPE_NUMERIC: {
			cNew.setCellValue(cOld.getNumericCellValue());
			break;
		}
		case Cell.CELL_TYPE_STRING: {
			cNew.setCellValue(cOld.getStringCellValue());
			break;
		}
		case Cell.CELL_TYPE_ERROR: {
			cNew.setCellValue(cOld.getErrorCellValue());
			break;
		}
		case Cell.CELL_TYPE_FORMULA: {
			cNew.setCellFormula(cOld.getCellFormula());
			break;
		}
		}
	}

	public String mergeCell(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			int firstRow = 0;
			int lastRow = 0;
			int firstCol = 0;
			int lastCol = 0;
			for (j = 0; j < files.size(); ++j) {
				int numberOfRow = 0;
				int k;
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				for (i = 2; i < numberOfSheet; ++i) {
					System.out.println("sheet : " + this.workbook.getSheetAt(i).getSheetName());
					if (this.workbook.isSheetHidden(i)) {
						continue;
					}
					if (this.workbook.getSheetAt(i).getSheetName().contains("Data")) {
						continue;
					}
					numberOfRow = 0;
					// xac dinh so row can loop
					for (k = 0; k < 10000; ++k) {
						if (this.workbook.getSheetAt(i).getRow(k + 11) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 11).getCell(1) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 11).getCell(1).getNumericCellValue() == 0.0) {
							break;
						}
						numberOfRow++;
					}
					int rowStepPlus = 0;
					int temp = 0;
					for (int m = 0; m < numberOfRow; ++m) {

						for (Cell cell : this.workbook.getSheetAt(i).getRow(m + 11)) {
							if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
								switch (cell.getCachedFormulaResultType()) {
								case Cell.CELL_TYPE_STRING:
									if (rowStepPlus > 0 && !"".equals(this.workbook.getSheetAt(i).getRow(11 + m)
											.getCell(3).getStringCellValue())) {
										System.out.println("first row: " + firstRow);
										System.out.println("last row: " + (firstRow + rowStepPlus));
										System.out.println("======================");
										this.workbook.getSheetAt(i).addMergedRegion(
												new CellRangeAddress(firstRow, firstRow + rowStepPlus, 3, 3));

										// reset firstRow, lastRow value for next cell merge
										firstRow = 0;
										lastRow = 0;
										rowStepPlus = 0;
									}
									if (rowStepPlus > 0 && this.workbook.getSheetAt(i).getRow(m + 11).getCell(1)
											.getNumericCellValue() == numberOfRow) {
										System.out.println("first row: " + firstRow);
										System.out.println("last row: " + (firstRow + rowStepPlus));
										System.out.println("======================");
										this.workbook.getSheetAt(i).addMergedRegion(
												new CellRangeAddress(firstRow, firstRow + rowStepPlus, 3, 3));
									}
									if (!"".equals(this.workbook.getSheetAt(i).getRow(11 + m).getCell(3)
											.getStringCellValue())) {
										firstRow = 11 + m;
										continue;
									}
									if ("".equals(this.workbook.getSheetAt(i).getRow(11 + m).getCell(3)
											.getStringCellValue().trim()) && m != temp) {
										temp = m;
										rowStepPlus += 1;
										continue;
									}
									break;
								}
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "merge cell: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "merge cell: FAIL";
		}
		return "merge cell: SUCCESS";
	}

	public String insertGGFomular(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				int k;
				int numberOfRow = 0;
				String googleFomular = "=GOOGLETRANSLATE(";
				String googleFomular1 = "=GOOGLETRANSLATE(";
				String googleFomular2 = "=GOOGLETRANSLATE(";
				for (i = 2; i < numberOfSheet; ++i) {

					if (this.workbook.getSheetAt(i).getSheetName().contains("Data")) {
						continue;
					}

					numberOfRow = 0;
					// xac dinh so row can loop
					for (k = 0; k < 10000; ++k) {
						if (this.workbook.getSheetAt(i).getRow(k + 11) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 11).getCell(1) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 11).getCell(1).getNumericCellValue() == 0.0) {
							break;
						}
						numberOfRow++;
					}

					for (int m = 0; m < numberOfRow; ++m) {

						googleFomular += "H" + (12 + m) + ";\"en\";\"ja\")";
						googleFomular1 += "J" + (12 + m) + ";\"en\";\"ja\")";
						googleFomular2 += "E" + (12 + m) + ";\"en\";\"ja\")";
						XSSFCell cell = this.workbook.getSheetAt(i).getRow(11 + m).getCell(6);
						XSSFCell cell1 = this.workbook.getSheetAt(i).getRow(11 + m).getCell(8);
						XSSFCell cell2 = this.workbook.getSheetAt(i).getRow(11 + m).getCell(3);
						cell.setCellValue(googleFomular);
						cell1.setCellValue(googleFomular1);
						cell2.setCellValue(googleFomular2);
						googleFomular = "=GOOGLETRANSLATE(";
						googleFomular1 = "=GOOGLETRANSLATE(";
						googleFomular2 = "=GOOGLETRANSLATE(";
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "insert GG fomular: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "insert GG fomular: FAIL";
		}
		return "insert GG fomular: SUCCESS";
	}

	public String insertGGFomularScreen(String linkFi) {
		List<XSSFWorkbook> workbooks = new ArrayList<>();
		try {
			List<InputStream> files = initialize(linkFi);
			int j;
			for (j = 0; j < files.size(); ++j) {
				this.workbook = new XSSFWorkbook(files.get(j));
				workbooks.add(this.workbook);
				int numberOfSheet = this.workbook.getNumberOfSheets();
				int i;
				int k;
				int numberOfRow = 0;
				String googleFomular = "=GOOGLETRANSLATE(";
				String googleFomular1 = "=GOOGLETRANSLATE(";
				String googleFomular2 = "=GOOGLETRANSLATE(";
				String googleFomular3 = "=GOOGLETRANSLATE(";
				for (i = 2; i < numberOfSheet; ++i) {

					if (this.workbook.getSheetAt(i).getSheetName().contains("Data")) {
						continue;
					}

					numberOfRow = 0;
					// xac dinh so row can loop
					for (k = 0; k < 10000; ++k) {
						if (this.workbook.getSheetAt(i).getRow(k + 7) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 7).getCell(1) == null) {
							break;
						}
						if (this.workbook.getSheetAt(i).getRow(k + 7).getZeroHeight()) {
							continue;
						}

						if (this.workbook.getSheetAt(i).getRow(k + 7).getCell(1)
								.getCellType() == Cell.CELL_TYPE_FORMULA) {
							switch (this.workbook.getSheetAt(i).getRow(k + 7).getCell(1).getCachedFormulaResultType()) {
							case Cell.CELL_TYPE_ERROR:
								if ("#REF!".equals(this.workbook.getSheetAt(i).getRow(k + 7).getCell(1)
										.getErrorCellString())) {
									this.workbook.getSheetAt(i).getRow(k + 7).getCell(1)
											.setCellType(Cell.CELL_TYPE_STRING);
									numberOfRow++;
								}
								continue;
							}
						}

						if (this.workbook.getSheetAt(i).getRow(k + 7).getCell(1).getNumericCellValue() == 0.0) {
							break;
						}
						numberOfRow++;
					}

					for (int m = 0; m < numberOfRow; ++m) {

						googleFomular += "D" + (8 + m) + ";\"en\";\"ja\")";
						googleFomular1 += "F" + (8 + m) + ";\"en\";\"ja\")";
						googleFomular2 += "I" + (8 + m) + ";\"en\";\"ja\")";
						googleFomular3 += "K" + (8 + m) + ";\"en\";\"ja\")";
						XSSFCell cell = this.workbook.getSheetAt(i).getRow(7 + m).getCell(2);
						XSSFCell cell1 = this.workbook.getSheetAt(i).getRow(7 + m).getCell(4);
						XSSFCell cell2 = this.workbook.getSheetAt(i).getRow(7 + m).getCell(7);
						XSSFCell cell3 = this.workbook.getSheetAt(i).getRow(7 + m).getCell(9);
						cell.setCellValue(googleFomular);
						cell1.setCellValue(googleFomular1);
						cell2.setCellValue(googleFomular2);
						cell3.setCellValue(googleFomular3);
						googleFomular = "=GOOGLETRANSLATE(";
						googleFomular1 = "=GOOGLETRANSLATE(";
						googleFomular2 = "=GOOGLETRANSLATE(";
						googleFomular3 = "=GOOGLETRANSLATE(";
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			return "insert GG fomular Screen: FAIL";
		}
		if (!write2File(linkFi, workbooks)) {
			return "insert GG fomular Screen: FAIL";
		}
		return "insert GG fomular Screen: SUCCESS";
	}
}
