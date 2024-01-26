# The-PDF-Mania
![image](https://github.com/Ali-Hasan-Khan28/The-PDF-Mania/assets/101451471/020ff314-2ef2-42a0-aea6-ebe24712fa27)
This web application is built using Flask, a Python web framework, to facilitate document conversion tasks. Users can convert various types of documents to PDF format and perform other document-related operations. The application includes the following features:

1. Image to PDF Conversion
Route: /index1
Description: Converts an image file (e.g., JPEG, PNG) to a PDF file.
Usage:
Upload an image file through the provided form.
Click the submit button to convert the image to a PDF file.
The resulting PDF file is saved on the desktop with the same filename as the original image.
2. Text to PDF Conversion
Route: /index2
Description: Converts a text file to a PDF file.
Usage:
Upload a text file through the provided form.
Click the submit button to convert the text to a PDF file.
The resulting PDF file is saved on the desktop with the same filename as the original text file.
3. Word to PDF Conversion
Route: /index3
Description: Converts a Microsoft Word document (.docx) to a PDF file.
Usage:
Upload a Word document through the provided form.
Click the submit button to convert the Word document to a PDF file.
The resulting PDF file is saved on the desktop with the same filename as the original Word document.
4. PDF to Word Conversion
Route: /index4
Description: Converts a PDF file to a Microsoft Word document (.docx).
Usage:
Upload a PDF file through the provided form.
Click the submit button to convert the PDF file to a Word document.
The resulting Word document is saved on the desktop with the same filename as the original PDF file.
5. PDF to Image Conversion
Route: /index5
Description: Converts a PDF file to a series of JPEG images.
Usage:
Upload a PDF file through the provided form.
Click the submit button to convert the PDF file to a set of JPEG images.
The resulting images are saved on the desktop with filenames indicating their order.
6. PDF Merging
Route: /index6
Description: Merges two PDF files into a single PDF file.
Usage:
Upload two PDF files through the provided form.
Click the submit button to merge the two PDF files into a single PDF.
The resulting merged PDF file is saved on the desktop.
7. Greeting Page
Route: /docx
Description: Displays a greeting page with the option to enter a name.
Usage:
Enter a name in the form and click the submit button.
The application will display a personalized greeting.
Note
The application uses various Python libraries for document conversion, including PyPDF2, img2pdf, pdf2image, pythoncom, docx2pdf, PIL, FPDF, and pdf2docx.
Make sure to install the required libraries before running the application. You can use the requirements.txt file to install dependencies.
The application runs in debug mode (debug=True), which is suitable for development but not recommended for production use.
Ensure that the necessary tools and libraries, such as poppler, are installed and configured correctly for PDF to image conversion.
Feel free to explore and modify the code according to your requirements. If you encounter any issues or have suggestions for improvement, please refer to the documentation of the used libraries and the Flask framework.
