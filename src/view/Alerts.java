package view;

import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Alert.AlertType;

/**
 * Creates the alerts to be used throughout the Highlighted Text Extractor tool.
 * All alerts are static methods.
 * 
 * @author Brenton Haliw
 *
 */
@SuppressWarnings("restriction")
public class Alerts {

	/**
	 * Alerts the user that the output document is currently opened and will not be
	 * able to be used by the program until it is closed.
	 */
	protected static void documentAlreadyOpened() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Extracted Text.docx needs to be closed before the tool can process the documents.", ButtonType.OK);
		alert.setTitle("Output Document Open");
		alert.setHeaderText("Extracted Text Document is Open");

		alert.showAndWait();
	}

	/**
	 * Alerts the user that the format of the file that they have chosen is invalid
	 * and that it should be .docx only.
	 */
	protected static void invalidDocuments() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Please ensure that you have selected a valid Microsoft Word Document (.docx) only.", ButtonType.OK);
		alert.setTitle("Invalid Document Format");
		alert.setHeaderText("Unable to Process Document");

		alert.showAndWait();

	}

	/**
	 * Alerts the user that they cannot use the Extracted Text.docx file that is in
	 * the same location as the tool. If so, the document gets corrupted.
	 */
	protected static void cannotExtractSelf() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Unable to extract text from the desginated output Microsoft Word document (Extracted Text.docx).\n\n"
						+ "Please select another valid Microsoft Word document to be processed.",
				ButtonType.OK);
		alert.setTitle("Invalid Document Selected");
		alert.setHeaderText("Unable to Use Selected Document");

		alert.showAndWait();
	}

	/**
	 * Alerts the user that the file that they have selected cannot be found. This
	 * can occur if the user selects a valid file and then either moves or deletes
	 * it.
	 */
	protected static void cannotFindFile() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"The system cannot find the file specified.\n\n"
						+ "This may have occured because you have moved or deleted the Microsoft Word document.\n\n"
						+ "Please select another Microsoft Word document to be processed.",
				ButtonType.OK);
		alert.setTitle("Invalid Document Selected");
		alert.setHeaderText("Unable to Find Selected Document");

		alert.showAndWait();
	}

	/**
	 * Alerts the user that the path that they have selected cannot be found. This
	 * can occur if the user selects a valid path and then either moves or deletes
	 * it.
	 */
	protected static void cannotFindPath() {
		Alert alert = new Alert(AlertType.INFORMATION, "The system cannot find the path specified.\n\n"
				+ "This may have occured because you have moved or deleted a folder containing the Microsoft Word document.\n\n"
				+ "Please select another Microsoft Word document to be processed.", ButtonType.OK);
		alert.setTitle("Invalid Path Selected");
		alert.setHeaderText("Unable to Find Selected Path");

		alert.showAndWait();
	}
}
