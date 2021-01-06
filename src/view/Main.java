package view;

import javafx.geometry.Insets;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.scene.layout.HBox;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.stage.Stage;
import javafx.stage.FileChooser;
import javafx.scene.control.TextField;
import javafx.scene.control.Label;
import javafx.scene.control.Button;
import javafx.scene.input.TransferMode;

/**
 * This class creates the GUI and contains the method to process the input
 * document and create the output document containing the extracted highlighted
 * text.
 * 
 * @author Brenton Haliw
 *
 */
@SuppressWarnings("restriction")
public class Main extends Application {

	/**
	 * Creates the GUI for the Highlighted Text Extractor tool
	 */
	@Override
	public void start(Stage primaryStage) {
		VBox box = new VBox(10);
		box.setPadding(new Insets(30, 10, 5, 10));
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
				"Drag and drop your Word Document (.docx) into this window or use the button to select its filepath.");
		Label msg2 = new Label(
				"A new document named \"Extracted Text\" will be created and placed where this tool is located.");
		msg.setAlignment(Pos.CENTER);

		HBox filepathBox = new HBox(10);
		filepathBox.setAlignment(Pos.CENTER);

		Label docLabel = new Label("Highlighted Document");
		Label completeLabel = new Label("Transfer Completed!");
		completeLabel.setVisible(false);

		TextField docField = new TextField();
		docField.setPromptText("File Path of Highlighted Document");
		docField.setPrefWidth(300);
		docField.setEditable(false);

		Button docButton = new Button("Select Document");
		docButton.setOnAction(e -> {
			docField.setText(launchFileChooser());
		});

		// Adding the document nodes to the file path box
		filepathBox.getChildren().addAll(docLabel, docField, docButton);

		Button start = new Button("Start");
		start.setPrefWidth(100);
		start.setOnAction(e -> {
			if (this.processDocuments(docField.getText())) {
				completeLabel.setVisible(true);
			}
		});

		Button reset = new Button("Reset");
		reset.setPrefWidth(100);
		reset.setOnAction(e -> {
			docField.setText("");
			completeLabel.setVisible(false);
		});

		// Create the button box to add the start/reset buttons to
		HBox buttonBox = new HBox(20);
		buttonBox.setAlignment(Pos.BASELINE_CENTER);
		buttonBox.getChildren().addAll(start, reset);
		buttonBox.setMaxWidth(Double.MAX_VALUE);

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
				completeLabel.setVisible(false);
				success = true;
			}
			event.setDropCompleted(success);
			event.consume();
		});

		// This box contains information about the author and his github account
		HBox info = new HBox();
		info.setAlignment(Pos.CENTER_RIGHT);
		Label infoLabel = new Label("Brenton Haliw | github.com/bjhaliw");
		info.getChildren().addAll(infoLabel);

		box.getChildren().addAll(title, msg, msg2, empty, filepathBox, buttonBox, completeLabel, info);
		Scene scene = new Scene(box);

		primaryStage.setScene(scene);
		primaryStage.setTitle("Word Transfer Tool");
		primaryStage.getScene().getRoot().setStyle("-fx-base:gainsboro");
		primaryStage.show();
	}

	/**
	 * Helper method to that returns the file extension name
	 * 
	 * @param fileName
	 * @return
	 */
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

		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Microsoft Word Document", "*.docx");

		chooser.getExtensionFilters().add(extFilter);
		File file = chooser.showOpenDialog(stage);

		if (file == null) {
			System.out.println("User backed out without selecting a file");
			return null;
		}

		return file.getAbsolutePath();

	}

	/**
	 * This method is responsible for parsing the input Microsoft Word document and
	 * extracting the highlighted text. A new Microsoft Word document named
	 * "Extracted Text.docx" will be created and will contain the extracted text,
	 * keeping its original format (font size, font style, etc)
	 * 
	 * @param inputPath - String representation of the input document's path
	 * @return - true if successfully parsed and created a new document, false if
	 *         otherwise
	 */
	private boolean processDocuments(String inputPath) {

		// Making sure we have a valid path for the input document
		if (inputPath == null || inputPath.equals("") || !getExtension(inputPath).equals("docx")) {
			Alerts.invalidDocuments();
			return false;
		}

		// Making sure we're not using the designated output document
		if (inputPath.equals(System.getProperty("user.dir") + "\\Extracted Text.docx")) {
			Alerts.cannotExtractSelf();
			return false;
		}

		try {
			// Opening and creating the input document
			FileInputStream fis = new FileInputStream(inputPath);
			XWPFDocument inDoc = new XWPFDocument(OPCPackage.open(fis));

			// Opening and creating the output document
			FileOutputStream fos = new FileOutputStream(System.getProperty("user.dir") + "\\Extracted Text.docx");
			XWPFDocument outDoc = new XWPFDocument();
			XWPFParagraph outDocParagraph;
			XWPFRun outDocRun;

			// Getting all of the paragraphs in the input document
			List<XWPFParagraph> paragraphList = inDoc.getParagraphs();

			// Looping through each paragraph in the input document
			for (XWPFParagraph paragraph : paragraphList) {
				// Looping through the individual runs inside the current paragraph of the input
				// document
				for (XWPFRun rn : paragraph.getRuns()) {
					// If the current run is highlighted, then we want to extract it to the output
					// document
					if (rn.isHighlighted()) {
						// Create a new paragraph and a new run in the output document
						outDocParagraph = outDoc.createParagraph();
						outDocRun = outDocParagraph.createRun();

						// Make the output run have the same format as the input run
						outDocRun.setBold(rn.isBold());
						outDocRun.setUnderline(rn.getUnderline());
						outDocRun.setStrikeThrough(rn.isStrikeThrough());
						outDocRun.setTextHighlightColor(rn.getTextHightlightColor().toString());
						outDocRun.setItalic(rn.isItalic());
						outDocRun.setStyle(rn.getStyle());
						outDocRun.setFontFamily(rn.getFontFamily());
						outDocRun.setColor(rn.getColor());
						if (rn.getFontSize() != -1) { // Font size is -1 sometimes which is weird
							outDocRun.setFontSize(rn.getFontSize());
						}
						outDocRun.setText(rn.toString());
					}
				}
			}
			// Write the output document to the specified file path
			outDoc.write(fos);

			// Close everything
			outDoc.close();
			fos.close();
			fis.close();
			inDoc.close();
		} catch (FileNotFoundException ex) {
			System.out.println("The message is: " + ex.getMessage());
			ex.printStackTrace();

			// If the Extracted Text document is currently opened
			if (ex.getMessage()
					.contains("(The process cannot access the file because it is being used by another process)")) {
				Alerts.documentAlreadyOpened();
			}

			// If the file was selected and then moved later
			if (ex.getMessage().contains("(The system cannot find the file specified)")) {
				Alerts.cannotFindFile();
			}

			// If the path was selected and then moved later
			if (ex.getMessage().contains("(The system cannot find the path specified)")) {
				Alerts.cannotFindPath();
			}

			return false;

		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			return false;
		}

		return true;
	}

	/**
	 * Launch the program
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		launch(args);
	}
}