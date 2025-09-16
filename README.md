# **Secure-PDF-Word-Converter-Flask-WebApp**
This project is a secure web-based PDF to Word and Word to PDF converter built with Flask. It was engineered with data privacy and file safety as the top priority—ensuring that no uploaded or converted files are ever saved, logged, or traceable after the conversion process. Temporary files are automatically cleaned up after each request, minimizing risks of data leaks or sniffing.

The backend conversion logic leverages multiple libraries (PyMuPDF, pdfplumber, pdf2docx, python-docx, docx2pdf) with layered fallback methods to maximize conversion accuracy while preserving formatting such as bullet points, indentation, and layout.

The project is intended as a proof of concept for enterprise use cases where companies handle sensitive or high-value documents and cannot risk exposing them to third-party services.

## **Features**

-Security-first design: No persistent file storage, automatic cleanup of temp files.

-Two-way conversion:

1. PDF → Word (.docx)

2. Word (.docx/.doc) → PDF

-Resilient pipeline: Multi-library fallback (PyMuPDF → pdfplumber → pdf2docx) ensures reliability even if one method fails.

-Formatting preservation: Special handling of bullet points, indentation, and symbols for accurate output.

-Browser-accessible web interface: Upload and download files via a simple HTML form.

-OpenAPI spec (swagger.yaml): Ready for integration with tools like Microsoft Copilot.

## **Future Plans**

Support for additional formats (e.g., images to PDF, cross-image format conversions).

Advanced PDF editing features similar to Adobe Acrobat.

Cloud deployment with enterprise-grade security and team collaboration features.

*This project demonstrates the potential of building privacy-respecting document conversion services for enterprises. With the right funding and a proper development team, it can be expanded into a full-fledged secure document suite.*

## **Getting Started**

Follow these steps to run the server locally. Exact paths, environment names, and tunnel setup may vary depending on your system.

1. Navigate to your project folder

Open a terminal and move into the directory containing your files (webapp.py, conversion_utils.py, swagger.yaml, etc.):

cd path/to/teams-converter

2. Activate your Python environment

If you use Conda (example):

conda activate converter_env


(Replace converter_env with your environment’s name.)

3. Start the Flask app
python webapp.py


This runs the app on http://127.0.0.1:5000 by default.

4. (Optional) Expose the server externally

If you want external access (e.g., for testing with Copilot), open another terminal and start ngrok:

ngrok http 5000


This will give you a public URL, e.g.:

https://xxxx.ngrok-free.app/swagger.yaml
