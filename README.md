XMLforOJS
=========

Interactive Applescript for fetching info from pdfs and writing xml for import in Open Journal Systems.

On importing in OJS, see: http://oapress.org/wp-content/uploads/2013/02/Import-Export-OJS.pdf

The script rests on the following presuppositions:

- pdfs can be read as text (natively or after OCR)
- skim is installed http://skim-app.sourceforge.net/ 
- title can be grabbed as first paragraph of first page
- first page can be grabbed as last paragraph of first page
- other content like author-info and abstract can occur in many places, and need to be selected by the user when prompted for this.
- the following "sections" are used in OJS: "Editorials", "Articles", "Reviews", "Reports", "Announcements"
