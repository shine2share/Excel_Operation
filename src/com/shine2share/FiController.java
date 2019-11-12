package com.shine2share;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import org.apache.commons.lang3.StringUtils;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.text.Text;
import lombok.Getter;
import lombok.Setter;

public class FiController implements Initializable {

	@Getter
	@Setter
	private ExcelOperation excelOperation;
	@FXML
	private Text txtHello;

	@FXML
	private Text txtResult;

	@FXML
	private TextField txtSize;
	@FXML
	private TextField txtZoom;
	@FXML
	private TextField txtNewSpecWord;

	@FXML
	private TextField txtLinkFile;

	@FXML
	private Button btnGetLink;

	@FXML
	private ComboBox<String> cmbColor;

	@FXML
	private ComboBox<String> cmbFont;

	@FXML
	private Button btnColor;

	@FXML
	private Button btnAddSpecWord;

	@FXML
	private Button btnDeleteSpecWord;

	@FXML
	private Button btnSearchSpecWord;

	@FXML
	private Button btnZoom;

	@Getter
	@Setter
	private String linkFi;

	@Getter
	@Setter
	private String colorValue;

	@Getter
	@Setter
	private String fontValue;

	@Getter
	@Setter
	private short sizeValue;
	@Getter
	@Setter
	private int zoomValue;

	@FXML
	private ListView lsWord;

	@FXML
	private ListView lsResult;

	@Override
	public void initialize(URL location, ResourceBundle resources) {
		showSpecWordListView();
	}

	private void showSpecWordListView() {
		ObservableList elements = FXCollections.observableArrayList();
		elements.addAll(getSpecWord());
		lsWord.setItems(elements);
	}

	private void showResultSpecWordListView(List<String> response) {
		ObservableList elements = FXCollections.observableArrayList();
		elements.addAll(getResultSpecWord(response));
		lsResult.setItems(elements);
	}

	private List<String> getResultSpecWord(List<String> response) {
		List<String> wordList = new ArrayList<>();
		int i;
		for (i = 0; i < response.size(); ++i) {
			wordList.add(response.get(i));
		}

		return wordList;
	}

	public void searchSpecWord() {
		this.txtResult.setText("");
		ObservableList items = lsWord.getItems();
		int size = items.size();
		Object searchWordObj = lsWord.getSelectionModel().getSelectedItem();
		if (searchWordObj == null) {
			this.txtResult.setText("Dòng Em muốn search không có nhé");
			return;
		}
		String searchWord = searchWordObj.toString();
		excelOperation = new ExcelOperation();
		List<String> response = excelOperation.searchSpecWord(this.linkFi, searchWord);
		if (response == null || response.size() == 0) {
			this.txtResult.setText("Dòng Em muốn search không có nhé");
			return;
		} else {
			this.txtResult.setText("Dòng Em muốn search có nhé");
		}
		showResultSpecWordListView(response);
	}

	public void getLinkFile() {
		this.linkFi = this.txtLinkFile.getText().trim();
	}

	public void formatA1() {
		if (StringUtils.isEmpty(this.linkFi)) {
			this.txtResult.setText("Nhập link đi cưng");
			return;
		}
		this.txtResult.setText("");
		excelOperation = new ExcelOperation();
		String response = excelOperation.formatA1(this.linkFi);
		this.txtResult.setText(response);
	}

	public void setColor() {
		if (StringUtils.isEmpty(this.linkFi)) {
			this.txtResult.setText("Nhập link đi cưng");
			return;
		}
		if (StringUtils.isEmpty(this.colorValue)) {
			this.txtResult.setText("Vui lòng chọn color");
			return;
		}
		this.txtResult.setText("");
		excelOperation = new ExcelOperation();
		String response = excelOperation.setSheetColor(this.linkFi, this.colorValue);
		this.txtResult.setText(response);
	}

	public void setFont() {
		if (StringUtils.isEmpty(this.linkFi)) {
			this.txtResult.setText("Nhập link đi cưng");
			return;
		}
		if (StringUtils.isEmpty(this.fontValue)) {
			this.txtResult.setText("Vui lòng chọn font");
			return;
		}
		this.txtResult.setText("");
		excelOperation = new ExcelOperation();
		String response = excelOperation.setSheetFont(this.linkFi, this.fontValue);
		this.txtResult.setText(response);
	}

	public void setSize() {
		if (StringUtils.isEmpty(this.linkFi)) {
			this.txtResult.setText("Nhập link đi cưng");
			return;
		}
		this.txtResult.setText("");
		if (StringUtils.isNumeric(this.txtSize.getText().trim())) {
			this.sizeValue = Short.parseShort(this.txtSize.getText().trim());
		} else {
			this.txtResult.setText("Vui lòng nhập số");
			return;
		}
		excelOperation = new ExcelOperation();
		String response = excelOperation.setSheetSize(this.linkFi, this.sizeValue);
		this.txtResult.setText(response);
	}

	public void setZoom() {
		if (StringUtils.isEmpty(this.linkFi)) {
			this.txtResult.setText("Nhập link đi cưng");
			return;
		}
		this.txtResult.setText("");
		if (StringUtils.isNumeric(this.txtZoom.getText().trim())) {
			this.zoomValue = Integer.parseInt(this.txtZoom.getText().trim());
		} else {
			this.txtResult.setText("Vui lòng nhập số");
			return;
		}
		excelOperation = new ExcelOperation();
		String response = excelOperation.setSheetZoom(this.linkFi, this.zoomValue);
		this.txtResult.setText(response);
	}

	public void choiceColor() {
		this.colorValue = this.cmbColor.getValue();
	}

	public void choiceFont() {
		this.fontValue = this.cmbFont.getValue();
	}

	private List<String> getSpecWord() {
		List<String> wordList = new ArrayList<>();
		try (BufferedReader br = new BufferedReader(new FileReader("SpecWord.txt"))) {
			String line;
			while ((line = br.readLine()) != null) {
				if (line.toString().trim().length() == 0) {
					continue;
				}
				wordList.add(line);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wordList;
	}

	public void addSpecWord() {
		this.txtResult.setText("");
		String response = writeSpecWord2TxtFile();
		this.txtResult.setText(response);
	}

	public void removeSpecWord() {
		this.txtResult.setText("");
		Object removeWordObj = lsWord.getSelectionModel().getSelectedItem();
		if (removeWordObj == null) {
			this.txtResult.setText("Dòng Em muốn xóa không có nhé");
			return;
		}
		String removeWord = removeWordObj.toString();
		// String response = deleteSpecWordFromTxtFile(removeWord);
		String response = removeLineFromFile("SpecWord.txt", removeWord);
		this.txtResult.setText(response);
	}

	private String deleteSpecWordFromTxtFile(String lineToRemove) {

		try {
			File inputFile = new File("SpecWord.txt");
			File tempFile = new File("SpecWord_temp.txt");

			BufferedReader reader = new BufferedReader(new FileReader(inputFile));
			BufferedWriter writer = new BufferedWriter(new FileWriter(tempFile));

			String currentLine;

			while ((currentLine = reader.readLine()) != null) {
				// trim newline when comparing with lineToRemove
				String trimmedLine = currentLine.trim();
				if (trimmedLine.equals(lineToRemove))
					continue;
				writer.write(currentLine + System.getProperty("line.separator"));
			}
			writer.close();
			reader.close();
			tempFile.renameTo(inputFile);
		} catch (IOException e) {
			return "Có lỗi xảy ra khi xóa spec word";
		}
		showSpecWordListView();
		return "Xóa spec word thành công";
	}

	private String writeSpecWord2TxtFile() {
		if (StringUtils.isEmpty(this.txtNewSpecWord.getText())) {
			return "Nhập từ cần thêm đi cưng";
		}
		File file = new File("SpecWord.txt");
		FileWriter fr;
		try {
			fr = new FileWriter(file, true);
			fr.write("\n" + this.txtNewSpecWord.getText());
			fr.close();
		} catch (IOException e) {
			return "Thêm SpecWord bị lỗi rồi";
		}
		showSpecWordListView();
		return "Thêm SpecWord Success";
	}

	private String removeLineFromFile(String file, String lineToRemove) {
		try {

			File inFile = new File(file);

			if (!inFile.isFile()) {
				return "Từ muốn xóa không tồn tại";
			}

			// Construct the new file that will later be renamed to the original filename.
			File tempFile = new File(inFile.getAbsolutePath() + ".tmp");

			BufferedReader br = new BufferedReader(new FileReader(file));
			PrintWriter pw = new PrintWriter(new FileWriter(tempFile));

			String line = null;

			// Read from the original file and write to the new
			// unless content matches data to be removed.
			while ((line = br.readLine()) != null) {

				if (!line.trim().equals(lineToRemove)) {

					pw.println(line);
					pw.flush();
				}
			}
			pw.close();
			br.close();

			// Delete the original file
			if (!inFile.delete()) {
				return "Xóa spec word bị lỗi rồi";
			}

			// Rename the new file to the filename the original file had.
			if (!tempFile.renameTo(inFile))
				return "Xóa spec word bị lỗi rồi || ko rename đc file";

		} catch (FileNotFoundException ex) {
			return "Xóa spec word bị lỗi rồi || ko tìm thấy file";
		} catch (IOException ex) {
			return "Xóa spec word bị lỗi rồi";
		}
		showSpecWordListView();
		return "Xóa spec word thành công";
	}
}
