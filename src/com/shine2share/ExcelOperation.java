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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
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

}
