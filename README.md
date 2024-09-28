Inspired by the post by iseahound/malcev ( [url=https://www.autohotkey.com/boards/viewtopic.php?style=7&t=80735]How to view PDF with Windows API? (PDF -> bitmap)[/url] ) I made a simple GUI front end to re-arrange / remove / insert PDF pages using a combination of Listview, Windows Runtime and QPDF.

It has the following basics features:
* Listview icon view with custom sizes to display PDF pages as thumbnails
* Re-arranging Listivew items (to allow re-arranging PDF pages)
* Mouse handling of drag event to re-position Listview items
* Keyboard copy-cut-paste of Listview items
* Displaying underlined button characters (accelerators) at startup (by default hidden by Windows)
* Asking for PDF password (error handling for Windows Runtime)
* Using PdfPageRenderOptions to output BMP (instead of the default PNG)
* Reading from streams
* Using QPDF via DLLCall

Qpdf is available from its github releases page. I only tested it with [url]https://github.com/qpdf/qpdf/releases/download/v11.9.1/qpdf-11.9.1-msvc64.zip[/url]. The qpdf dll was placed in the qpdf\ subfolder.
