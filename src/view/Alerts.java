package view;

import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Alert.AlertType;

@SuppressWarnings("restriction")
public class Alerts {

	protected static void documentAlreadyOpened() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Extracted Text.docx needs to be closed before the tool can process the documents.", ButtonType.OK);
		alert.setTitle("Output Document Open");
		alert.setHeaderText("Extracted Text Document is Open");

		alert.showAndWait();
	}

	protected static void invalidDocuments() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Please ensure that you have selected valid input/output documents", ButtonType.OK);
		alert.setTitle("Invalid Documents");
		alert.setHeaderText("Unable to Process Documents");

		alert.showAndWait();

	}

	protected static void cannotExtractSelf() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"Unable to extract text from the desginated output text document.\n\nPlease select another document to be processed.",
				ButtonType.OK);
		alert.setTitle("Invalid Document Selected");
		alert.setHeaderText("Unable to Use Selected Document");

		alert.showAndWait();
	}
	
	protected static void cannotFindFile() {
		Alert alert = new Alert(AlertType.INFORMATION,
				"The system cannot find the file specified.\n\nPlease select another document to be processed.",
				ButtonType.OK);
		alert.setTitle("Invalid Document Selected");
		alert.setHeaderText("Unable to Find Selected Document");

		alert.showAndWait();
	}
}
