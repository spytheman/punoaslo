# PUNOASLO #
============

... or Miscelaneous __P__ython __UNO__ __A__utomation __S__cripts for Open/__L__ibre __O__ffice



Before you use the scripts in this project, you should launch a headless instance of the Libre Office / Open Office (it will be ruled by the python script):

```bash
soffice --accept='socket,port=2002;urp;StarOffice.ServiceManager' \
--nologo --nofirststartwizard --norestore --headless &
```

============================

## append_and_make_pdf.py ##
============================

The append_and_make_pdf.py script allows *appending* a custom LibreOffice document (ODT) - a footer,
to another one - a source (any format that OO understands), and exporting the resulting document as a PDF file.

Example usage: 
		
```bash
./append_and_make_pdf.py --source examples/source.odt --footer examples/footer.odt --pdf results/output.pdf
```