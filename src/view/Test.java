package view;

import javafx.geometry.VPos;
import javafx.geometry.HPos;
import javafx.geometry.Insets;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.*;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.stage.*;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.input.TransferMode;

@SuppressWarnings("restriction")
public class Test extends Application {

	@Override
	public void start(Stage primaryStage) throws Exception {
		VBox box = new VBox(10);
		box.setPadding(new Insets(30, 10, 20, 10));
		box.setAlignment(Pos.CENTER);

		// So the window doesn't auto select the first item
		TextField empty = new TextField();
		empty.setPrefSize(0, 0);
		empty.setMaxSize(0, 0);
		empty.setMinSize(0, 0);

		Label title = new Label("Extract Highlighted Words");
		title.setFont(Font.font("Arial", FontWeight.BOLD, FontPosture.REGULAR, 20));
		title.setStyle("-fx-underline: true ;");

		Label msg = new Label(
				"Drag and drop your Word Document into this window or use the button to select its filepath.");
		Label msg2 = new Label(
				"A new document named \"Output\" will be created and placed where this tool is located.");
		msg.setAlignment(Pos.CENTER);

		HBox filepathBox = new HBox(10);
		filepathBox.setAlignment(Pos.CENTER);

		Label docLabel = new Label("Highlighted Document");
		Label completeLabel = new Label("Transfer Completed!");
		completeLabel.setVisible(false);

		TextField docField = new TextField();
		docField.setPromptText("File Path of Highlighted Document");
		docField.setPrefWidth(300);

		Button docButton = new Button("Select Document");
		docButton.setOnAction(e -> {
			docField.setText(launchFileChooser());
		});

		filepathBox.getChildren().addAll(docLabel, docField, docButton);

		Button start = new Button("Start");
		start.setPrefWidth(100);
		start.setOnAction(e -> {
			completeLabel.setVisible(true);
		});

		Button reset = new Button("Reset");
		reset.setPrefWidth(100);
		reset.setOnAction(e -> {
			completeLabel.setVisible(false);
		});

		HBox buttonBox = new HBox(20);
		buttonBox.setAlignment(Pos.BASELINE_CENTER);
		buttonBox.getChildren().addAll(start, reset);
		buttonBox.setMaxWidth(Double.MAX_VALUE);

		box.getChildren().addAll(title, msg, msg2, empty, filepathBox, buttonBox, completeLabel);

		// Extensions that are valid to be drag-n-dropped
		List<String> validExtensions = Arrays.asList("doc", "docx");

		box.setOnDragOver(event -> {
			// On drag over if the DragBoard has files
			if (event.getGestureSource() != box && event.getDragboard().hasFiles()) {
				// All files on the dragboard must have an accepted extension
				if (!validExtensions.containsAll(event.getDragboard().getFiles().stream()
						.map(file -> getExtension(file.getName())).collect(Collectors.toList()))) {

					event.consume();
					return;
				}

				// Allow for both copying and moving
				event.acceptTransferModes(TransferMode.COPY_OR_MOVE);
			}
			event.consume();
		});

		box.setOnDragDropped(event -> {
			boolean success = false;
			if (event.getGestureSource() != box && event.getDragboard().hasFiles()) {
				// Print files
				event.getDragboard().getFiles().forEach(file -> docField.setText(file.getAbsolutePath()));
				success = true;
			}
			event.setDropCompleted(success);
			event.consume();
		});

		Scene scene = new Scene(box);

		primaryStage.setScene(scene);
		primaryStage.setTitle("Word Transfer Tool");
		primaryStage.getScene().getRoot().setStyle("-fx-base:lavender");
		primaryStage.show();
	}

	// Method to to get extension of a file
	private String getExtension(String fileName) {
		String extension = "";

		int i = fileName.lastIndexOf('.');
		if (i > 0 && i < fileName.length() - 1) // if the name is not empty
			return fileName.substring(i + 1).toLowerCase();

		return extension;
	}

	/**
	 * Launches a file chooser for the user to select their Word documents
	 */
	private String launchFileChooser() {
		Stage stage = new Stage();

		FileChooser chooser = new FileChooser();
		chooser.setTitle("Select Microsoft Word Document");
		stage.setAlwaysOnTop(true);

		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Microsoft Word Document", "*.docx",
				"*.doc");

		chooser.getExtensionFilters().add(extFilter);
		File file = chooser.showOpenDialog(stage);

		if (file == null) {
			System.out.println("User backed out without selecting a file");
			return null;
		}

		return file.getAbsolutePath();

	}

	private boolean processDocuments(String inputPath, String outputPath) {

		if (inputPath == null || outputPath == null || inputPath.equals("") || outputPath.equals("")) {
			invalidDocuments();
			return false;
		}

		if (!inputPath.contains(".doc") || !inputPath.contains(".docx") || !outputPath.contains(".doc")
				|| !outputPath.contains(".docx")) {
			invalidDocuments();
			return false;
		}

		try {
			FileInputStream fis = new FileInputStream(inputPath);
			FileOutputStream fos = new FileOutputStream(outputPath);
			XWPFDocument inDoc = new XWPFDocument(OPCPackage.open(fis));

			XWPFDocument outDoc = new XWPFDocument();
			XWPFParagraph outDocParagraph;
			XWPFRun outDocRun;

			List<XWPFParagraph> paragraphList = inDoc.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {

				for (XWPFRun rn : paragraph.getRuns()) {
					System.out.println(rn.toString());
					System.out.println(rn.isHighlighted());

					if (rn.isHighlighted()) {
						System.out.println("Adding to new document");
						outDocParagraph = outDoc.createParagraph();
						outDocRun = outDocParagraph.createRun();
						outDocRun.setTextHighlightColor(rn.getTextHightlightColor().toString());
						outDocRun.setItalic(rn.isItalic());
						outDocRun.setStyle(rn.getStyle());
						outDocRun.setFontFamily(rn.getFontFamily());
						outDocRun.setColor(rn.getColor());
						if (rn.getFontSize() != -1) {
							outDocRun.setFontSize(rn.getFontSize());
						}
						outDocRun.setText(rn.toString());
						System.out.println("Font size is: " + rn.getFontSize());
					}
				}

				System.out.println("********************************************************************");
			}
			outDoc.write(fos);
			outDoc.close();
			fos.close();
			fis.close();
			inDoc.close();
		} catch (FileNotFoundException ex) {
			ex.printStackTrace();
			if (ex.getMessage()
					.contains("(The process cannot access the file because it is being used by another process)")) {
				documentAlreadyOpened();
				return false;
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return true;
	}

	private void processDocFile(FileInputStream fis) {
		try {
			HWPFDocument doc = new HWPFDocument(fis);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private boolean processDocxFile(FileInputStream fis, FileOutputStream fos) {
		try {
			XWPFDocument inDoc = new XWPFDocument(OPCPackage.open(fis));
			XWPFDocument outDoc = new XWPFDocument();
			XWPFParagraph outDocParagraph;
			XWPFRun outDocRun;

			List<XWPFParagraph> paragraphList = inDoc.getParagraphs();

			for (XWPFParagraph paragraph : paragraphList) {

				for (XWPFRun rn : paragraph.getRuns()) {
					System.out.println(rn.toString());
					System.out.println(rn.isHighlighted());

					if (rn.isHighlighted()) {
						System.out.println("Adding to new document");
						outDocParagraph = outDoc.createParagraph();
						outDocRun = outDocParagraph.createRun();
						outDocRun.setTextHighlightColor(rn.getTextHightlightColor().toString());
						outDocRun.setItalic(rn.isItalic());
						outDocRun.setStyle(rn.getStyle());
						outDocRun.setFontFamily(rn.getFontFamily());
						outDocRun.setColor(rn.getColor());
						if (rn.getFontSize() != -1) {
							outDocRun.setFontSize(rn.getFontSize());
						}
						outDocRun.setText(rn.toString());
						System.out.println("Font size is: " + rn.getFontSize());
					}
				}

				System.out.println("********************************************************************");
			}
			outDoc.write(fos);
			outDoc.close();
			fos.close();
			fis.close();
			inDoc.close();
		} catch (FileNotFoundException ex) {
			ex.printStackTrace();
			if (ex.getMessage()
					.contains("(The process cannot access the file because it is being used by another process)")) {
				documentAlreadyOpened();
				return false;
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}

		return true;
	}

	private void documentAlreadyOpened() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Output document needs to be closed before the tool can process both documents.", ButtonType.OK);
		alert.setTitle("Output Document Open");
		alert.setHeaderText("Output Document is Open");

		alert.showAndWait();
	}

	private void invalidDocuments() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Please ensure that you have selected valid input/output documents", ButtonType.OK);
		alert.setTitle("Invalid Documents");
		alert.setHeaderText("Unable to Process Documents");

		alert.showAndWait();

	}

	public static void main(String[] args) {
		launch(args);
	}
}