package view;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javafx.geometry.VPos;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.*;
import javafx.stage.*;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;

@SuppressWarnings("restriction")
public class MainGUI extends Application {


	
	@Override
	public void start(Stage primaryStage) throws Exception {		
		VBox box = new VBox(10);
		box.setPadding(new Insets(30,10,20,10));
		box.setAlignment(Pos.CENTER);

		
		TextField empty = new TextField();
		empty.setPrefSize(0, 0);
		empty.setMaxSize(0, 0);
		empty.setMinSize(0,0);
		
		Label firstDocLabel = new Label("Input Document");
		Label secondDocLabel = new Label("Output Document");
		Label completeLabel = new Label("Transfer Completed!");
		completeLabel.setVisible(false);
		
		GridPane pane = new GridPane();
		pane.setAlignment(Pos.CENTER);
		pane.setHgap(10);
		pane.setVgap(10);
		
		TextField firstDocField = new TextField();
		TextField secondDocField = new TextField();
		firstDocField.setPromptText("File Path of Input Document");
		secondDocField.setPromptText("File Path of Output Document");
		firstDocField.setMaxWidth(Double.MAX_VALUE);
		secondDocField.setMaxWidth(Double.MAX_VALUE);
		firstDocField.setPrefWidth(300);
		
		Button firstButton = new Button("Select Document");
		firstButton.setOnAction(e -> {
			firstDocField.setText(launchFileChooser());
		});
		
		Button secondButton = new Button("Select Document");
		secondButton.setOnAction(e -> {
			secondDocField.setText(launchFileChooser());
		});
		
		Button start = new Button("Start");
		start.setOnAction(e -> {
			boolean successful = processDocuments(firstDocField.getText(), secondDocField.getText());
			if (successful) {
				completeLabel.setVisible(true);
			}
		});
		
		Button reset = new Button("Reset");
		reset.setOnAction(e -> {
			completeLabel.setVisible(false);
		});
		
		pane.add(firstDocLabel, 0, 0, 1, 1);
		pane.add(firstDocField, 1, 0, 4, 1);
		pane.add(firstButton, 5, 0, 1, 1);
		
		pane.add(secondDocLabel, 0, 1, 1, 1);
		pane.add(secondDocField, 1, 1, 4, 1);
		pane.add(secondButton, 5, 1, 1, 1);
		
		HBox buttonBox = new HBox(20);
		buttonBox.setAlignment(Pos.CENTER);
		buttonBox.getChildren().addAll(start, reset);
		buttonBox.setMaxWidth(Double.MAX_VALUE);
		
		//pane.add(start, 1, 2);
		//pane.add(reset, 3, 2);
		
		//pane.add(completeLabel, 2, 3, 1, 1);
		
		//pane.setPadding(new Insets(0,0,10,0));

		
		GridPane.setHalignment(firstDocLabel, HPos.RIGHT);
		GridPane.setHalignment(secondDocLabel, HPos.RIGHT);
		
		GridPane.setHgrow(firstDocField, Priority.ALWAYS);
		GridPane.setHgrow(secondDocField, Priority.ALWAYS);
		
		GridPane.setFillWidth(firstDocField, true);
		GridPane.setFillWidth(secondDocField, true);
		
		//GridPane.setHalignment(reset, HPos.CENTER);
		//GridPane.setHalignment(start, HPos.CENTER);
		
		GridPane.setHalignment(completeLabel, HPos.CENTER);
		
		box.getChildren().addAll(empty,pane, buttonBox, completeLabel);
		Scene scene = new Scene(box);
		
		primaryStage.setScene(scene);
		primaryStage.setTitle("Word Transfer Tool");
		//primaryStage.getScene().getRoot().setStyle("-fx-base:black");
		primaryStage.show();
	}
	
	/**
	 * Launches a file chooser for the user to select their Word documents
	 */
	private String launchFileChooser() {
		Stage stage = new Stage();

		FileChooser chooser = new FileChooser();
		chooser.setTitle("Select Microsoft Word Document");
		stage.setAlwaysOnTop(true);
		
		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Microsoft Word Document", "*.docx", "*.doc");
		
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
		
		if (!inputPath.contains(".doc") || !inputPath.contains(".docx") || !outputPath.contains(".doc") || !outputPath.contains(".docx")) {
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
					    if(rn.getFontSize() != -1 ) {
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
			if(ex.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
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
					    if(rn.getFontSize() != -1 ) {
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
			if(ex.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
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
				"Output document needs to be closed before the tool can process both documents.",
				ButtonType.OK);
		alert.setTitle("Output Document Open");
		alert.setHeaderText("Output Document is Open");

		alert.showAndWait();
	}
	
	private void invalidDocuments() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Please ensure that you have selected valid input/output documents",
				ButtonType.OK);
		alert.setTitle("Invalid Documents");
		alert.setHeaderText("Unable to Process Documents");

		alert.showAndWait();
		
	}
	
	public static void main(String[] args) {
		launch(args);
	}
	
	

}
