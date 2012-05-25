punoaslo
========

	Miscelaneous Python UNO Automation Scripts for open/Libre Office



	Before usage of the scripts in this project, you should launch a
headless instance of the Libre Office / Open Office, which will be
ruled by the python script:

		soffice -o --accept='socket,port=2002;urp;StarOffice.ServiceManager' --nologo --nofirststartwizard --norestore --headless


#append_and_make_pdf.py
======================

	The append_and_make_pdf.py script allows *appending*
a custom LibreOffice document (ODT) - a footer,
to another one - a source (any format that OO understands),
and exporting the resulting document as a PDF file.

	Example usage: 
		
    ./append_and_make_pdf.py --source example_documents/source.odt --footer example_documents/footer.odt --pdf results/output.pdf

