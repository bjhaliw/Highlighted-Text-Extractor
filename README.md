# Highlighted Text Extractor
 This tool allows the user to extract all highlighted text from a Word document and transfer it to automatically created Word document.

# How to use
- Drag and drop your .docx into the window or use the selection button to find the .docx with the highlighted text
- Select the start button and the program will run through the supplied document
  - Text features such as font size, font style, font color, and highlight color will be preserved
- After the program has finished parsing through the Word document, a new Word document named "Extracted Text" will be created
  - This Word document will be created in the same location as the program's location

# How it works
This program utilizes Apache POI from the Apache Software Foundation to interact with Microsoft Word documents (.docx only). The program will open and parse through the supplied Word document, looking for highlighted text. The Apache POI software specifically looks at what is called "runs" within a paragraph, meaning text that shares similar characteristics (same font style, font size, etc.).  It is possible for multiple runs to be present in the same sentence, however, such as italicized words and different colored words. There were two different approaches that I could have taken to deal with this:
  - Separate each run into their own separate paragraphs which would cause the output document to be incredibly long and hard to read
  - Concatenate the runs in the paragraphs that they are found in. This would cause sentences with multiple runs to stay intact, but it might cause formatting errors if individual words or characters in a paragraph are highlighted.
  
I went with number 2 in order to preserve the best formatting possible since I assume most people would be highlighting multiple words in a sentence rather than just a character or a single word.

Once the program is done parsing the document, the highlighted text is then written into a new Word document named "Extracted Text.docx" that is created automatically and placed in the same directory as the program.

# Sample Photos
### The Tool
![Tool](https://github.com/bjhaliw/Highlighted-Text-Extractor/blob/main/Sample%20Photos/highlight%20tool.png)
### The text before processing
![Before](https://github.com/bjhaliw/Highlighted-Text-Extractor/blob/main/Sample%20Photos/Before.PNG)
### The extracted text in the new document
![After](https://github.com/bjhaliw/Highlighted-Text-Extractor/blob/main/Sample%20Photos/After.PNG)
