scan
====

`update_dates.py` was written to drive Acrobat 8 Professional scanning with a Pentax DSmobile 600.  It uses Acrobat to do OCR, and adds PDFs to the EagleFiler library.  The scanner only scans a sheet at a time, so it also attempts to simplify scanning multipage documents by prompting for the number of pages to be scanned.

`update_dates_in_place.py` was written to be used from EagleFiler once OCRed PDFs are already in the library.  It still uses Acrobat 8 Professional, but not for OCR, just to reduce file size.  It is more useful with a scanner with an ADF, such as the Canon PIXMA MX922 I am currently using.

These scripts are not designed to be directly used by anyone other than me, but they're freely available in case they're useful as examples.
