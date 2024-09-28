#Requires AutoHotkey v2.0
#SingleInstance Force
; GUI front end for QPDF
; Uses Windows Runtime to render the PDF (requires Windows 10)
For n,arg in A_Args
	pdf(arg)
if !A_Args.length
	pdf()

pdf(fn:="",w:=150,h:=200) {
	static pvGui, LV, PdfDocumentStatics, PdfPageRenderOptions, clsid:=buffer(16), PdfDocs := map()	; PdfDocs to allow preview
	if !isSet(pvGui) {
		; Construct PdfPageRenderOptions so we can set BMP output (default is PNG)
		PdfPageRenderOptions:=CreateInstance("Windows.Data.Pdf.PdfPageRenderOptions","{3C98056F-B7CF-4C29-9A04-52D90267F425}")	
		Numput("Int64", 0x47C8D66D69BE8BB4, "Int64", 0x8237438915ED5A86, clsid)		; BMP {69BE8BB4-D66D-47C8-865A-ED1589433782}
		ComCall(17, PdfPageRenderOptions, "ptr", clsid)		; put_BitmapEncoderId
	; GetFactory("Windows.Graphics.Imaging.BitmapEncoder","{A74356A7-A4E4-4EB9-8E40-564DE7E1CCB2}",&BitmapEncoderStatics:=0)
	; BitmapEncoderId:=Buffer(16)
	; ComCall(6, BitmapEncoderStatics, "ptr",BitmapEncoderId)	; BitmapEncoderId	{69BE8BB4-D66D-47C8-865A-ED1589433782}
	; ComCall(7, IBitmapEncoderStatics,"ptr",BitmapEncoderId)	; JpegEncoderId		{1A34F5C1-4A5A-46DC-B644-1F4567E7A676}
	; ComCall(8, IBitmapEncoderStatics,"ptr",BitmapEncoderId)	; PngEncoderId		{27949969-876A-41D7-9447-568F6A35A4DC}
	; ComCall(9, IBitmapEncoderStatics,"ptr",BitmapEncoderId)	; TiffEncoderId		{0131BE10-2001-4C5F-A9B0-CC88FAB64CE8}
	; ComCall(10,IBitmapEncoderStatics,"ptr",BitmapEncoderId)	; GifEncoderId		{114F5598-0B22-40A0-86A1-C83EA495ADBD}
	; ComCall(11,IBitmapEncoderStatics,"ptr",BitmapEncoderId)	; JXR EncoderId		{AC4CE3CB-E1C1-44CD-8215-5A1665509EC2}
	; ComCall(17, PdfPageRenderOptions,"ptr",BitmapEncoderId)	; put_BitmapEncoderId
	; To show BitmapEncoderId
	; ComCall(16, PdfPageRenderOptions, "ptr", BitmapEncoderId)	; get_BitmapEncoderId
	; DllCall("Ole32\StringFromCLSID", "Ptr", BitmapEncoderId, "Str*", &str:="")
	; msgbox str
		GetFactory("Windows.Data.Pdf.PdfDocument", "{433A0B5F-C007-4788-90F2-08143D922599}", &PdfDocumentStatics:=0)
		pvGui := Gui("+Resize -DPIScale",fn)
        pvGui.AddButton(,"Add &PDFs").OnEvent("Click", AddPDF)
		pvGui.AddButton("x+5","&Open PDF").OnEvent("Click",(*)=>OpenPdf(fn := FileSelect(1,A_ScriptDir,"Open PDF","PDF (*.pdf)"),0))
		pvGui.AddButton("x+5","Save &As").OnEvent("Click",pdfSave)
		; Create PDF thumbnail ListView
		MonitorGetWorkArea(1,&Left, &Top, &Right, &Bottom)
		LV:=pvGui.AddListView("x5 Icon -Hdr LV0x10000 w" (Right-Left)*3//5 " h" Bottom-Top-90,["pdf"])	
		; -90 for title bar & buttons, width 60% for thumbnails
		; Icon to display label below image (Tile displays text right of image)
		; LVS_EX_DOUBLEBUFFER LV0x10000 Paints via double-buffering, which reduces flicker.
		LV.OnEvent("ItemSelect",preview)
 		LV.OnNotify(-109, LVN_BEGINDRAG)
		pvGui.Show("x-4 y-4")	; show on left to allow preview window on right
		PdfDocs.default := "", PdfDocs.CaseSense := 0
		GroupAdd("pvGui","ahk_id " pvGui.Hwnd)
		SendMessage(0x0128,2+2<<16,0,pvGui.Hwnd)	; WM_UPDATEUISTATE := 0x0128
	; WM_UPDATEUISTATE changes the UI state for the specified window and all its child windows.
	; wParam	Low-order word specifies the action to be performed:
	;			UIS_CLEAR		2	UI state element specified by high-order word should be cleared.
	;			UIS_INITIALIZE	3	UI state element should be changed based on the last input event. 
	;			UIS_SET			1	UI state element specified by high-order word should be set.
	;			High-order word specifies which UI state elements are affected or the style of the control.
	;			UISF_ACTIVE		4	A control should be drawn in the style used for active controls.
	;			UISF_HIDEACCEL	2	Keyboard accelerators.
	;			UISF_HIDEFOCUS	1	Focus indicators.
	; lParam	Not used.
	; A window should send this message to change the UI state of all its child windows. 
	; In contrast to the WM_CHANGEUISTATE message, which is a notification, 
	; when DefWindowProc processes the WM_UPDATEUISTATE message it changes the UI state 
	; and propagates the changes to all child windows.
	; The DefWindowProc function updates the UI state according to the wParam value. 
	; If the UI state is modified, the function sends the message to all the immediate child windows. 
	; DefWindowProc also sends this message when it receives a WM_CHANGEUISTATE message notifying 
	; the system that a child window intends to modify the UI state.
	; See:	https://devblogs.microsoft.com/oldnewthing/20130516-00/?p=4343
	;		https://devblogs.microsoft.com/oldnewthing/20130517-00/?p=4323
	}
	Loop Files, fn
		OpenPdf(A_LoopFileFullPath)
	return

	AddPDF(*) {
		if fns:= FileSelect("M3",, "Add PDFs", "PDF (*.pdf)")
			for fn in fns
				OpenPdf(fn)
	}

	OpenPdf(fn,add:=1) {	; add=1 to add pdf file to the pdf thumbnails ListView
		static ImgLst
		if !fn
			return

		t := A_TickCount
		if !PdfDocs.has(fn) {
			Numput("Int64", 0x11DFBC53905A0FE1, "Int64", 0xDA86C64F1E00498C, clsid)		
			; IRandomAccessStream {905A0FE1-BC53-11DF-8C49-001E4FC686DA}
			DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", fn, "uint", Read:=0, "ptr", clsid, "ptr*", &RandomAccessStream:=0, "HRESULT")
			ComCall(LoadFromStreamAsync := 8, PdfDocumentStatics, "ptr", RandomAccessStream, "ptr*", &PdfDocument:=0)	
			Loop {
 				try {
 					Await(&PdfDocument)							; ~ 32 ms
 					break
 				} catch error as e {
 					if e.number = 0x8007052b {		; 0x8007052b	HRESULT_FROM_WIN32(ERROR_WRONG_PASSWORD)
	 					res := InputBox(fn " is protected. Please enter a Document Open Password.","Password","w580 h130")
 						if res.result = "OK" {
				 			DllCall("combase\WindowsCreateString", "wstr", res.value, "uint", strlen(res.value), "ptr*", &hString:=0, "HRESULT")
				 			ComCall(LoadFromStreamWithPasswordAsync := 9, PdfDocumentStatics, "ptr", RandomAccessStream, "ptr", hString, "ptr*", &PdfDocument:=0)	; PdfDocument.LoadFromStreamAsync
				 			DllCall("combase\WindowsDeleteString", "ptr", hString, "HRESULT")
 						} else return
	 				} else if e.number=0x80048040
	 					throw error("0x80048040: PDF version/encryption not supported.`n"
	 								. "Only Acrobat IX encryption or below is supported.`n"
	 								. "Try opening it with another program.",-1,fn)
					; Error code 0x80048040	has no description.
					; Seems to be thrown if a document uses Acrobat 10 encryption.
	 				else if e.number=0x80048014
	 					throw error("0x80048014: PDF format not recognised.",-1,fn)
	 				else throw e
 				}
 			}
			PdfDocs[fn] := [PdfDocument,RandomAccessStream]
		} else PdfDocument := PdfDocs[fn][1]
		ComCall(7, PdfDocument, "uint*", &PageCount:=0)	; Get PageCount
		if !PageCount
			return

		LV.focus()
		if !add && IsSet(ImgLst)
			IL_Destroy(ImgLst), ImgLst := Unset, LV.Delete()

		if !IsSet(ImgLst) {						; Initialize ImageList for listview
			ImgLst := IL_Create(PageCount,5,1)	; Icon view use large icons
			; Sets the dimensions of images in an image list and removes all images from the list.
			DllCall("ComCtl32.dll\ImageList_SetIconSize","Ptr",ImgLst,"Int",w,"Int",h,"Int")	
			wh := SendMessage(0x1035,0,(w+4)|((h+53)<<16), LV.Hwnd)		; guestimate w & h to include label and margin
	; LVM_SETICONSPACING:=0x1035, wParam 0, lParam width (loword) + height (hiword)
	; Returns a DWORD containing the previous x-axis distance in low word, 
	; and previous y-axis distance in high word.
	; Values are relative to upper-left corner of an icon bitmap. 
	; Therefore, to set spacing between icons that do not overlap, 
	; values must include the size of the icon, plus the amount of empty space desired between icons. 
	; Values that do not include the width of the icon will result in overlaps.
	; When defining icon spacing, values must be 4 or larger. Smaller values will not yield the desired layout. 
	; To reset the icons to the default spacing, set lParam to -1.
	; setup rendering options BitmapEncoder to BMP because default is PNG
			LV.SetImageList(ImgLst,0)			; NB set LV ImageList AFTER the iconsize has been set
		} 

		; do two renderPdf threads (~5-10% speedup; no obvious gains beyond 2)
		renderPdf(PdfDocument,0,&StreamOut,&RandomAccessStreamOut,&AsyncInfo,&width:=w,&height:=h)	
		Loop PageCount {			; Add and show the images concurrently for the user
			if A_Index < PageCount	; start rendering next page while waiting for current page to finish 
				renderPdf(PdfDocument,A_Index,&StreamOut2,&RandomAccessStreamOut2,&AsyncInfo2,&width:=w,&height:=h)
			if hBmp := AwaitHBitmap(&StreamOut,&RandomAccessStreamOut,&AsyncInfo,w,h) {
				; Add image to imagelist, and add imagelist item & description to row
				LV.Add("Icon" IL_Add(ImgLst, "HBITMAP:" hBmp), fn ": " A_Index "/" PageCount)	
				pvGui.title := fn " loading... (" A_Index "/" PageCount ")"
			}
			if A_Index < PageCount	
				StreamOut := StreamOut2, RandomAccessStreamOut := RandomAccessStreamOut2, AsyncInfo := AsyncInfo2 
		}
		pvGui.title := fn " (" PageCount " pages, " A_Tickcount-t "ms)"
	}

	renderPdf(PdfDocument,Page,&StreamOut,&RandomAccessStreamOut,&AsyncInfo,&imgw,&imgh) {	; Render PDF page with options
		ComCall(6, PdfDocument, "uint", Page, "ptr*", &PdfPage:=0)		; page index is 0 based
		DllCall("ole32\CreateStreamOnHGlobal", "ptr", 0, "uint", true, "ptr*", &StreamOut:=0)
		DllCall("ShCore\CreateRandomAccessStreamOverStream", "ptr", StreamOut, "uint", BSOS_DEFAULT := 0, "ptr", CLSID, "ptr*", &RandomAccessStreamOut:=0)
		ComCall(10, PdfPage, "ptr", Size:=buffer(8))		; PdfPage.size returns a size structure of float width & float height
		pgw := NumGet(size,"float")
		pgh := NumGet(size,4,"float")
		if imgw || imgh {
			if (imgw && !imgh) || (pgw*imgh/pgh>imgw)
				imgh := pgh * imgw/pgw
			else imgw := pgw * imgh/pgh
		} else imgh:=pgh, imgw:=pgw
		ComCall(9, PdfPageRenderOptions, "uint", imgw:=round(imgw))	; put_DestinationWidth
		ComCall(11, PdfPageRenderOptions, "uint", imgh:=round(imgh))	; put_DestinationHeight
		ComCall(7, PdfPage, "ptr", RandomAccessStreamOut, "ptr", PdfPageRenderOptions, "ptr*", &AsyncInfo:=0)	; RenderWithOptionsToStreamAsync
		; Await(&AsyncInfo)
		; Dispose(&RandomAccessStreamOut)		; Dispose RandomStream to speedup normal IStream operations later
	}

	AwaitHBitmap(&StreamOut,&RandomAccessStreamOut,&AsyncInfo,w?,h?) {	
		; Await separated from renderPdf to allow rendering another page while waiting for the current page
		static BITMAPINFO := Buffer(40), hDC := DllCall('GetDC', 'Ptr', 0, 'Ptr')
		AsyncInfo := ComObjQuery(AsyncInfo, IAsyncInfo := "{00000036-0000-0000-C000-000000000046}")
		while !ComCall(7, AsyncInfo, "uint*", &status:=0) and !status
			sleep 0 ; wait for rendering to complete; sleep is essential, otherwise will keep looping without updating
		ComCall(IAsyncInfo_Close := 10, AsyncInfo)
		if status=1 { ; Copies BMP filestream to hBitmap
			Dispose(&RandomAccessStreamOut)		; Dispose RandomStream to try speedup normal IStream operations later
			DllCall("shlwapi\IStream_Size", "ptr", StreamOut, "uint64*", &size:=0, "hresult")	
			; IStream includes the BMPFILE HEADER (14 bytes), which we will skip
			; Seek (long dlibMove, int dwOrigin, IntPtr plibNewPosition);
			ComCall(Seek := 5, StreamOut, "uint64", 14, "uint", 0, "uint64*", &current:=0)		
			DllCall("shlwapi\IStream_Read","ptr",StreamOut,"ptr",BITMAPINFO, "uint64", 40, "hresult")	; copy BITMAPINFO (40 bytes) 
			if IsSet(w) && IsSet(h) {
				width := NumGet(BITMAPINFO,4,"UInt"), height := NumGet(BITMAPINFO,8,"Int")
				; msgbox width "x" height " in " w "x" h
				if height < h
					NumPut("Int",round(h*width/w),BITMAPINFO,8)	; if landscape maintain aspect ratio
; 				b:=fileopen("bmpinfo","w")
; 				b.rawwrite(BITMAPINFO)
; 				b.close()
			}
			hBitmap := DllCall('CreateDIBSection','Ptr',hDC,'Ptr',BITMAPINFO,'UInt',0,'PtrP', &pBits:=0, 'Ptr', 0, 'UInt', 0, 'Ptr')
			DllCall("shlwapi\IStream_Read", "ptr", StreamOut, "ptr", pBits, "uint64", size-54, "hresult")	; copy the rest of data
			ObjRelease(StreamOut)
			return hBitmap
		}
	}

	pdfSaveDlg() {
		while f := FileSelect(16,A_ScriptDir,"Save As","PDF (*.pdf)") {
			if !RegExMatch(f,"i).pdf$")
				f.=".pdf"
			if PdfDocs[f] {
				msgbox "Cannot save an opened file.  Please enter a new filename."
				continue
			}
			return f
		}
	}

	#DllLoad "qpdf\qpdf29.dll"
	pdfDecrypt(*) {
		if f := pdfSaveDlg() {
 			qpdfJob('--decrypt "' fn '" "' f '"')
		}
	}

	pdfEncrypt(*) {
		if f := pdfSaveDlg() {
 			; qpdfJob('--encrypt "' fn '" "' f '"')
 			; Windows PDF only supports up to Acrobat 9
 			; Acrobat 4: qpdf --allow-weak-crypto --encrypt --owner-password=123456 --bits=40 -- in.pdf out.pdf
			; Acrobat 5: qpdf --allow-weak-crypto --encrypt --owner-password=123456 --bits=128 -- in.pdf out.pdf 
			; Acrobat 7-8: qpdf --allow-weak-crypto --encrypt --owner-password=123456 --bits=128 --use-aes=y -- in.pdf out.pdf
			; Acrobat 9: qpdf --encrypt --owner-password=123456 --bits=256 --force-R5 -- in.pdf out.pdf 
			; Acrobat 10: qpdf --encrypt --owner-password=123456 --bits=256 -- in.pdf out.pdf 
			; Acrobat 10 is not supported by Windows PDF (but is supported by e.g. mupdf) 
		}
	}

	qpdfJob(cmdline) {
		argv := DllCall("Shell32\CommandLineToArgvW", "WStr","qpdfjob " CmdLine,"UInt*",&argc:=0, "Ptr" ) 
		; CommandLineToArgvW is only Unicode so need to convert to Ansi
		Loop argc {											; Convert to Ansi for qpdf
			arg := NumGet(argv,(A_Index-1)*A_PtrSize,"Ptr")	; Get pointer to argument
			StrPut(StrGet(arg), arg, "CP0")					; Convert to Ansi and store back
		}
		DllCall("qpdf29\qpdfjob_run_from_argv","Ptr",argv)
		DllCall("LocalFree", "Ptr", argv)					; Free memory
		msgbox "Done " cmdline
	}

	pdfSave(*) {
		; static hModule := DllCall("LoadLibrary", "Str", "qpdf\qpdf29.dll", "Ptr")
		if f := pdfSaveDlg() {

;			; Commandline method
;			prevfn := "", CmdLine := "--empty --pages"
;			txt := ListViewGetContent("Col1",LV.Hwnd)		; ~10% faster than 
; 			Loop Parse, txt, "`n" { 						; 	Loop LV.GetCount() {
; 				s:=A_LoopField 								;		s:=LV.GetText(A_Index) ...
; 				if RegExMatch(s,"(.+): (\d+)/(\d+)$",&pg) {
; 					if prevfn != pg.1
; 						CmdLine .= ' "' (prevfn:=pg.1) '" '
; 					if A_Index=1
; 						CmdLine .= pg.2
; 					else if (pos:=RegExMatch(CmdLine,"-(\d+)$",&n)) && (n.1=pg.2-1)
; 						CmdLine := SubStr(CmdLine,1,pos-1) "-" pg.2
; 					else if (pos:=RegExMatch(CmdLine,"(\d+)$",&n)) && (n.1=pg.2-1)
; 						CmdLine .= "-" pg.2
; 					else CmdLine .= "," pg.2
; 				}
; 			}
; 			CmdLine .= ' -- "' f '"'
;
;			; Commandline RunWait method
;; 			RunWait("qpdf\qpdf.exe " CmdLine)
;
;			; Commandline DllCall method
;			qpdfJob(cmdLine)
;			msgbox "Done " A_Clipboard := "qpdf\qpdf.exe " CmdLine

;			; DllCall C Method
 			newpdf := DllCall("qpdf29\qpdf_init","Cdecl Ptr")			; Initialize output pdf handle
 			DllCall("qpdf29\qpdf_empty_pdf","Ptr",newpdf,"Cdecl UInt")	; Make it an empty pdf
 			source := map()												; Create mapping array of input pdf handles
			txt := ListViewGetContent("Col1",LV.Hwnd)					; Collect list of pages to output
 			Loop Parse, txt, "`n" {
 				if RegExMatch(A_LoopField,"(.+): (\d+)/(\d+)$",&pg) {
 					if !source.has(pg.1) {
						srcpdf := DllCall("qpdf29\qpdf_init","Cdecl Ptr")
 						source[pg.1] := srcpdf
 						; DllCall("qpdf29\qpdf_read","Ptr",srcpdf,"AStr",pg.1,"Ptr",0,"Cdecl UInt")
 						RandomAccessStream := PdfDocs[pg.1][2]
						ReadRandomAccessStream(RandomAccessStream,&buf,&size)
 						DllCall("qpdf29\qpdf_read_memory","Ptr",srcpdf,"AStr",pg.1,"Ptr",buf,"UInt",size,"Ptr",0,"Cdecl UInt")
 					} else srcpdf := source[pg.1]
 					pageobj := DllCall("qpdf29\qpdf_get_page_n","Ptr",srcpdf,"UInt",pg.2-1,"Cdecl UInt")	; page is 0 based
 					DllCall("qpdf29\qpdf_add_page","Ptr",newpdf,"Ptr",srcpdf,"UInt",pageobj,"UInt",0,"Cdecl UInt")
; 					DllCall("qpdf29\qpdf_oh_release","Ptr",srcpdf,"UInt",pageobj,"Cdecl UInt")
 				}
 			}
 			DllCall("qpdf29\qpdf_init_write","Ptr",newpdf,"Astr",f,"Cdecl UInt")
 			DllCall("qpdf29\qpdf_write","Ptr",newpdf,"Cdecl UInt")
 			DllCall("qpdf29\qpdf_cleanup","Ptr*",&newpdf,"Cdecl")	
 			for k,v in source 
				DllCall("qpdf29\qpdf_cleanup","Ptr*",&v,"Cdecl")
			msgbox "Done"
		}
	}

	ReadRandomAccessStream(RandomAccessStream,&buf,&size) {
		static DataReaderFactory 
		if !IsSet(DataReaderFactory)
			GetFactory("Windows.Storage.Streams.DataReader","{D7527847-57DA-4E15-914C-06806699A098}",&DataReaderFactory:=0)
		ComCall(GetSize:=6, RandomAccessStream,"Int64*",&size:=0)
		ComCall(GetInputStreamAt:=8, RandomAccessStream,"Int64",0,"Ptr*",&stream:=0)
		ComCall(CreateDataReader:=6,DataReaderFactory,"Ptr",stream,"Ptr*",&DataReader:=0)
		ComCall(LoadAsync:=29,DataReader,"UInt",size,"Ptr*",&numBytesLoaded:=0)
		Await(&numBytesLoaded)
		buf:=Buffer(numBytesLoaded)
		ComCall(ReadBytes:=14,DataReader,"UInt",numBytesLoaded,"Ptr", buf)
	}

	preview(LV,row,selected) {
		static pgGui, pgPic
		if !selected
			return
		RegExMatch(title:=LV.GetText(row),"(.+): (\d+)/(\d+)$",&pg)
		MonitorGetWorkArea(1,&Left,,&Right)
		LV.gui.GetClientPos(,,,&lvh)
		LV.gui.GetPos(&x,&y,&guiw)
		x+=guiw-20	; place preview to right of main window
		width := Right-Left-x, height := lvh

		renderPdf(PdfDocs[pg.1][1],pg.2-1,&StreamOut,&RandomAccessStreamOut,&AsyncInfo,&width,&height)
		hBmp := AwaitHBitmap(&StreamOut,&RandomAccessStreamOut,&AsyncInfo) 

		if !isSet(pgGui) {
			pgGui:= Gui("+Resize -DPIScale",title)
			pgGui.OnEvent("Escape",(*)=>pgGui.Hide())
			pgGui.OnEvent("Close",(*)=>pgGui.Hide())
			pgGui.Marginx := pgGui.MarginY := 0
			pgPic := pgGui.Add("Picture",,"HBITMAP:" hBmp)
			pgGui.Show("AutoSize NA x" x " y" y)
		} else {
			pgGui.Title := title " (" width "x" height ")"
			pgPic.Value := "*w" width " *h" height " HBITMAP:" hBmp
			pgGui.Show("AutoSize NA")
		}
	}
}

#HotIf WinActive("ahk_Group pvGui")
LV_Init() {
	global LV_Clip
	if !IsSet(LV_Clip)
		LV_Clip:=[]
	pvGui:=GuiFromHwnd(WinExist())
	return  pvGui["SysListView321"]
}

LV_Refresh(LV) {
	View:=SendMessage(0x108F, 0, 0, LV.Hwnd)
	if !(View&1){	; If not list/report view
		LV.Opt("-Redraw")					; Disable redraw to hide that we are switching view
		SendMessage(0x108E, 3, 0, LV.Hwnd)		; Set to list view to order items, LVM_SETVIEW := 0x108E ; (LVM_FIRST + 142)
		SendMessage(0x108E, View, 0, LV.Hwnd)	; Set back to original view to display items
		LV.Opt("+Redraw 0x2000")			; Enable redraw with LVS_NOSCROLL
		LV.Opt("-0x2000")					; Re-enable scrolling after redraw
   	}
}

^c::
	LV_Copy(*) {
		LV:=LV_Init()
        row := 0
        While row := LV.GetNext(row) 
            LV_Clip.push(LV_GetIcon(row,LV.Hwnd),LV.GetText(row))
	}

^x::
	LV_Cut(*) {
		LV:=LV_Init()
        row := 0
        While row := LV.GetNext(row) {
            LV_Clip.push(LV_GetIcon(row,LV.Hwnd),LV.GetText(row))
            LV.Delete(row--)
        }
        LV_Refresh(LV)
	}

^v::
	LV_Paste(*) {
		LV:=LV_Init()
		row := LV.GetNext(0)
		Loop LV_Clip.length/2 
			LV.Insert(row++,"Icon" LV_Clip[A_Index*2-1], LV_Clip[A_Index*2])
		LV_Refresh(LV)
	}

Del::
    LV_Delete(*) {	; Delete selected rows
		LV:=GuiFromHwnd(WinExist())["SysListView321"]
        row := 0
        While row := LV.GetNext(row)
            LV.Delete(row--)
        LV_Refresh(LV)
    }

	LV_GetIcon(row,Hwnd) {
		static	LVITEM := Buffer(48 + (A_PtrSize * 3), 0)
				; LVM_GETITEM := A_IsUnicode ? 0x104B : 0x1005 ; LVM_GETITEMW : LVM_GETITEMA
	; UINT	mask;
	;		LVIF_TEXT               0x00001	; pszText member is valid or must be set.
	;		LVIF_IMAGE              0x00002	; iImage member is valid or must be set.
	;		LVIF_PARAM              0x00004	; lParam member is valid or must be set.
	;		LVIF_STATE              0x00008	; state member is valid or must be set.
	;		LVIF_INDENT             0x00010	; iIndent member is valid or must be set.
	;		LVIF_NORECOMPUTE        0x00800	; control will not generate LVN_GETDISPINFO to retrieve text information 
											; if it receives an LVM_GETITEM message. Instead, pszText will contain LPSTR_TEXTCALLBACK.
	;		LVIF_DI_SETITEM         0x01000	; OS should store the requested list item information and not ask for it again. 
											; This flag is used only with the LVN_GETDISPINFO notification code.
	;		>= WINXP
	;		LVIF_GROUPID            0x00100	; iGroupId member is valid or must be set. If this flag is not set when LVM_INSERTITEM 
											; message is sent, the value of iGroupId is assumed to be I_GROUPIDCALLBACK.
	;		LVIF_COLUMNS            0x00200	; cColumns member is valid or must be set.
	;		>= VISTA
	;		LVIF_COLFMT             0x10000	; piColFmt member is valid or must be set. If this flag is used, 
	; int	iItem;
	; int	iSubItem;
	; UINT	state;					; Indicates the item's state, state image, and overlay image. 
									; The stateMask member indicates the valid bits of this member.
									; Bits 0 through 7 contain the item state flags
	;		LVIS_FOCUSED            0x0001	; The item has the focus, so it is surrounded by a standard focus rectangle.
											; Although more than one item may be selected, only one item can have the focus.
	;		LVIS_SELECTED           0x0002	; The item is selected. The appearance of a selected item depends on whether it has the focus 
											; and also on the system colors used for selection.
	;		LVIS_CUT                0x0004	; The item is marked for a cut-and-paste operation.
	;		LVIS_DROPHILITED        0x0008	; The item is highlighted as a drag-and-drop target.
	;		LVIS_GLOW               0x0010
	;		LVIS_ACTIVATING         0x0020	; Not currently supported.
									; Bits 8 through 11 specify the one-based overlay image index. 
									; The overlay image is superimposed over the item's icon image.
	;		LVIS_OVERLAYMASK        0x0F00	; The item's one-based overlay image index is retrieved by a mask.
									; Bits 12 through 15 of this member specify the state image index. 
									; The state image is displayed next to an item's icon
	;		LVIS_STATEIMAGEMASK     0xF000	; The item's one-based state image index is retrieved by a mask.
	; UINT   stateMask;				; This member allows you to modify one or more item states without having to retrieve 
									; all of the item states first. For example, setting this member to LVIS_SELECTED and 
									; state to zero will cause the item's selection state to be cleared, 
									; but none of the other states will be affected.
									; To retrieve or modify all of the states, set this member to (UINT)-1.
	; LPSTR  pszText;
	; int    cchTextMax				; Number of TCHARs in the buffer pointed to by pszText, including the terminating NULL.
	; int    iImage					; Index of the item's icon in the control's image list. If this member is the I_IMAGECALLBACK value, 
	;		= 20+(A_PtrSize*2)		; parent window is responsible for storing the index. 
									; In this case, the list-view control sends the parent 
									; an LVN_GETDISPINFO notification code to retrieve the index when it needs to display the image.
	; LPARAM lParam;
	; int    iIndent;
	; int    iGroupId;
	; UINT   cColumns;
	; PUINT  puColumns;
	; int    *piColFmt;
	; int    iGroup;
		NumPut("UInt", 2, "Int", Row-1, LVITEM)
	; LVM_GETITEM 
	; wParam	Must be zero.
	; lParam	Pointer to an LVITEM structure that specifies the information to retrieve and receives information about the list-view item.
	; 	iItem and iSubItem members identify the item or subitem to retrieve information 
	; 	mask member specifies which attributes to retrieve. 
	;	If LVIF_TEXT flag is set in the mask, pszText must point to a valid buffer and cchTextMax member 
	;	must be set to the number of characters in that buffer. 
	;	Applications should not assume that the text will necessarily be placed in the specified buffer. 
	;	The control may instead change the pszText member of the structure to point to the new text, rather than place it in the buffer.
	;	If the mask specifies LVIF_STATE, the stateMask member must specify the item state bits to retrieve. 
	;	On output, the state member contains the values of the specified state bits.
	; Returns TRUE if successful, or FALSE otherwise.
		SendMessage(0x104B, 0, LVITEM, Hwnd)	; LVM_GETITEMW
		return NumGet(LVITEM,A_PtrSize*2+20,"Int")+1
	}

	GetFactory(className, interface, &factory) {	; for static classes e.g. PdfDocumentStatics, BitmapEncoderStatics, DataReaderFactory
	   DllCall("combase\WindowsCreateString", "wstr", className, "uint", StrLen(className), "ptr*", &hString:=0, "HRESULT")
	   DllCall("ole32\CLSIDFromString", "wstr", interface, "ptr", CLSID := Buffer(16), "HRESULT")
	   DllCall("combase\RoGetActivationFactory", "ptr", hString, "ptr", CLSID, "ptr*", &factory:=0, "HRESULT")
	   DllCall("combase\WindowsDeleteString", "ptr", hString, "HRESULT")
	}

	CreateInstance(className, interface) {			; for instance classes (have constructors) e.g. PdfPageRenderOptions Class
		DllCall("combase\WindowsCreateString", "wstr", className, "uint", StrLen(className), "ptr*", &hString:=0, "HRESULT")
		DllCall("combase\RoActivateInstance", "ptr", hString, "ptr*", &Instance:=0, "HRESULT")
		DllCall("combase\WindowsDeleteString", "ptr", hString, "HRESULT")
		return ComObjQuery(Instance, interface)
	}

	Await(&Obj) {
		AsyncInfo := ComObjQuery(Obj, IAsyncInfo := "{00000036-0000-0000-C000-000000000046}")
		while !ComCall(7, AsyncInfo, "uint*", &status:=0) and (!status)		; IAsyncInfo.Status, 0 Started, 1 Completed, 2 Canceled, 3 Error
			Sleep 0
		ComCall(8, Obj, "ptr*", &Obj) ; GetResults
		; if AsyncInfo fails, Obj will be 0 and AHK will throw AsyncInfo error code with description
		ComCall(IAsyncInfo_Close := 10, AsyncInfo)
	}

	Dispose(&Object) {
		if (Close := ComObjQuery(Object, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"))
			ComCall(IClosable_Close := 6, Close)
		return ObjRelease(Object)
	}


LVN_BEGINDRAG(LV, LPARAM) {
	static LVINSERTMARK := Buffer(16,0), item := NumPut("UInt",16,LVINSERTMARK), POINT:=Buffer(8)
	;   UINT  cbSize;
	;	DWORD dwFlags;		// LVIM_AFTER	Insertion point appears after item specified if LVIM_AFTER flag is set; 
	;										otherwise it appears before the specified item.
	;	int   iItem;		// Item next to which the insertion point appears. If this member contains -1, there is no insertion point.
	;	DWORD dwReserved;
	Item := NumGet(LPARAM + (A_PtrSize * 3), "Int")
	DragButton := GetKeyState("LButton") ? "LButton" : "RButton"
	ins := -1		; ? also activates if drag is performed on the window then the user clicks on a listview item
	LV.Opt("0x100")
	; LVS_AUTOARRANGE 0x100 Icons are automatically kept arranged in icon and small icon view. Required for insertmark to work.
	SendMessage(0x100E, Item, RECT:=Buffer(16,0), LV.Hwnd)
	; LVM_GETITEMRECT 0x100E	(LVM_FIRST + 14)		; LVM_FIRST 0x1000      // ListView messages
	; wParam Index of the list-view item.
	; lParam RECT structure that receives the bounding rectangle. 
	;	LONG left;
	;	LONG top;
	;	LONG right;
	;	LONG bottom;
	; When the message is sent, the left member of RECT specifies the portion of the list-view item to retrieve.
	; 	LVIR_BOUNDS			0 bounding rectangle of the entire item, including the icon and label.
	; 	LVIR_ICON			1 bounding rectangle of the icon or small icon.
	; 	LVIR_LABEL			2 bounding rectangle of the item text.
	; 	LVIR_SELECTBOUNDS	3 union of the LVIR_ICON and LVIR_LABEL rectangles, but excludes columns in report view.
	ItemHeight := NumGet(RECT,12,"Int") - NumGet(RECT,4,"Int")						
	LV.GetPos(&LvX,&LvY,&LvW,&LvH)
	While GetKeyState(DragButton, "P") {
	; LVM_SETINSERTMARKCOLOR	0x10AA 
	;	wParam Must be zero.
	;	lParam **COLORREF** structure that specifies the color to set the insertion point.
	;	Returns previous colorref
	;	Default is black (0)
	;	rgbRed   =  0x000000FF;
	;	rgbGreen =  0x0000FF00;
	;	rgbBlue  =  0x00FF0000;
	;	rgbBlack =  0x00000000;
	;	rgbWhite =  0x00FFFFFF;
	; LVM_GETINSERTMARKRECT := 0x10A9 ; (LVM_FIRST + 169)
	;	wParam must be zero.
	;	lParam Pointer to a **RECT** structure that contains the coordinates of a rectangle that bounds the insertion point.
	;	Return 0 if No insertion point found. 1 if Insertion point found.
	    DllCall("User32.dll\GetCursorPos", "Ptr", POINT)    				; Get mouse cursor position in screen coordinates
	    DllCall("User32.dll\ScreenToClient", "Ptr", LV.Hwnd, "Ptr", POINT)	; Convert to client coordinates related to the ListView
	    cy:=NumGet(POINT,4,"Int")
	    scroll := cy<LvY ? -ItemHeight : cy>(LvY+LvH) ? ItemHeight : 0		; scroll up/down if mouse position y is above/below Listview
	    if scroll { 
	; LVM_SCROLL 0x1014
	; wParam	int that specifies the amount of horizontal scrolling, in pixels, 
	;			relative to the current position of the list view content. 
	;			If the list-view control is in list view, this value is 
	;			rounded up to the nearest number of pixels that form a whole column.
	; lParam	int that specifies the amount of vertical scrolling, in pixels, 
	;			relative to the current position of the list view content.
			SendMessage(0x1014, 0, scroll, LV.Hwnd)	; LVM_SCROLL
			Sleep(100)	; ScrollDelay
			ins := -1
	    } else if SendMessage(0x10A8, POINT, LVINSERTMARK, LV.Hwnd) {		; LVM_INSERTMARKHITTEST 0x10A8
	; LVM_INSERTMARKHITTEST 0x10A8 
	; 	wParam	Pointer to a **POINT** structure that contains the hit test coordinates.
	; 	lParam	Pointer to an LVINSERTMARK structure that specifies the insertion point 
	;			closest to the coordinates defined by the *wParam* parameter.
	; 	Returns TRUE if successful, or FALSE if cbSize member of the LVINSERTMARK structure 
	;	does not equal the actual size of the structure (16 bytes),
	;	or when an insertion point does not apply in the current view.
	; 	An insertion point can only appear if the list-view control is in icon view, 
	;	small icon view, or tile view and is not in group-view mode.
	;	If insertion points do not apply for the view, 
	;	the LVINSERTMARK structure contains a -1 in the iItem member.
	;	or if above control (gives closest item if below/left/right of control).
			SendMessage(0x10A6, 0, LVINSERTMARK, LV.Hwnd)
    		ins := NumGet(LVINSERTMARK,8,"Int")
	; LVM_SETINSERTMARK 	 0x10A6
	;	Sets the insertion point to the defined position.
	;	wParam Must be zero.
	;	lParam Pointer to a LVINSERTMARK structure that specifies where to set the insertion point.
	; 	Returns TRUE if successful, or FALSE if the size in the cbSize member of the 
	;	LVINSERTMARK structure does not equal the actual size of the structure,
	;	or when an insertion point does not apply in the current view.
	; 	An insertion point can only appear if the list-view control is in icon view, 
	;	small icon view, or tile view and is not in group-view mode.
	;	Also works in list view. In icon view lvsautoarrange must be set.
		} 
	}
	NumPut("Int",-1,LVINSERTMARK,8)			; Hide InsertMark 
	SendMessage(0x10A6, 0, LVINSERTMARK, LV.Hwnd)

	if ins >= 0 {
		View := SendMessage(0x108F, 0, 0, LV.Hwnd)
		; LVM_GETVIEW 0x108F	; wParam Must be zero.	; lParam Must be zero.
		; Returns a DWORD that specifies the current view.
		; Views := {0x00: "Icon", 0x01: "Report", 0x02: "IconSmall", 0x03: "List", 0x04: "Tile"}
	
	; Positioned vs. non-positioned listview views	Raymond Chen
	; "I inserted an item with LVM_INSERTITEM but it went to the end of the list instead of in the location I inserted it." 
	; Some listview views are "positioned" and others are "non-positioned". 
	; "(Large) icon view", "small icon view", and "tile view" are positioned views. 
	; Each item carries its own coordinates, which you can customize via LVM_SETITEMPOSITION. 
	; When a new item is inserted, it gets an item index based on the insertion point, 
	; but its physical location on the screen is the first available space not already occupied by another item. 
	; Existing items are not moved around to make room for the inserted item. 
	; The other views, "list view" and "report (aka details) view", are non-positioned views. 
	; In these views, items do not get to choose their positions. 
	; Instead, the position of an item is determined by its item index. 
	; In non-positioned views, inserting or deleting an item will indeed cause all subsequent items to shift.
	; https://stackoverflow.com/questions/73648009/items-inserted-into-a-tlistview-in-tile-view-always-appear-at-the-bottom-of-the
	; 2 ways to change item position
	;	1) use LVM_GETITEMPOSITION and LVM_SETITEMPOSITION messages
	;	2) set View to list/report view then set the view back to icon view

	; https://library.thedatadungeon.com/msdn-2000-04/wceui/htm/ctrls_40.htm
	; Setting the List View Item and Scroll Position
	; Every list view item has a position and size, which you can retrieve and set using messages. 
	; You can also determine which item, if any, is at a specified position. 
	; The position of list view items is specified in view coordinates, which are client coordinates offset by the scroll position.
	; To retrieve and set an item's position, use the LVM_GETITEMPOSITION and LVM_SETITEMPOSITION messages.
	; LVM_GETITEMPOSITION works for all views, but LVM_SETITEMPOSITION works only for icon and small icon views.
	; You can determine which item, if any, is at a particular location by using the LVM_HITTEST message. 
	; To get the bounding rectangle for a list item, or for only its icon or label, use the LVM_GETITEMRECT message.
	; Unless the LVS_NOSCROLL window style is specified, you can use messages to perform a variety of scrolling operations. 
	; You can scroll a list view control to show items that do not fit in the client area of the control, 
	; determine a list view control's scroll position, scroll a list view control by a specified amount, 
	; or scroll a list view control so that a specified list item is visible.
	; In icon view or small icon view, the current scroll position is defined by the view origin. 
	; The view origin is the set of coordinates, relative to the visible area of the list view control, 
	; that corresponds to the view coordinates (0, 0).
	; To get the current view origin, use the LVM_GETORIGIN message (0x1029 = LVM_FIRST + 41) 
	;	SendMessage(0x1029, 0, POINT, Lv.Hwnd)
	;	x:=NumGet(POINT,"Int"), y:=NumGet(POINT,4,"Int") 
	; This message should be used only in icon or small icon view; it returns an error in list or report view.
	; In list or report view, the current scroll position is defined by the top index. 
	; The top index is the index of the first visible item in the list view control. 
	; To get the current top index, use the LVM_GETTOPINDEX message. 
	; This message returns a valid result only in list view or report view; it returns zero in icon or small icon view.
	; Use LVM_GETVIEWRECT to get the bounding rectangle of all items in a list view relative to the visible area of the control.
	; The LVM_GETCOUNTPERPAGE message returns the number of items that fit in one page of the list view control. 
	; This message returns a valid result only in list and report views; 
	; in icon and small icon views, it returns the total number of items.
	; To scroll a list view control by a specific amount, use the LVM_SCROLL message. 
	;	SendMessage(0x1029, 0, POINT, Lv.Hwnd)	; LVM_GETORIGIN := 0x1029 ; (LVM_FIRST + 41) 
	;   x:=x-NumGet(POINT,"Int"), y:=y-NumGet(POINT,4,"Int") 
	; 	SendMessage(0x1014, x, y, LV.Hwnd)	; LVM_SCROLL
	; Use LVM_ENSUREVISIBLE to scroll the list view control, if necessary, 
	; to ensure that a specified item is visible.
		row := 0
		if view&1 {
			While row := LV.GetNext(row)
	        	if ins < row
					LV_Move(row,++ins)
				else LV_Move(row--,ins+1)
		} else {
			LV.Opt("-0x100")	; switch off LVS_AUTOARRANGE to allow setting the item position
			; LVM_GETITEMPOSITION = 0x00001010, increment insertion point for LV.Insert (which is 1 based)
			SendMessage(0x00001010, ins++, POINT:=Buffer(8), LV.Hwnd)	
			; Combine view coordinates for LVM_SETITEMPOSITION
			viewpos:=NumGet(POINT,"Int")+(NumGet(POINT,4,"Int")<<16)	
	        ; LVM_SETITEMPOSITION = 0x1000+15
			; if moving to front, increment insertion point afterwards; -1 for SET_ITEMPOS (0 based)
			; if moving to back, decrement selected row so we don't skip past it with LV.GetNext
	        While row := LV.GetNext(row) 								
	        	if ins < row
					SendMessage(0x100F, Lv_Move(row,ins++)-1, viewpos, LV.Hwnd)	
				else SendMessage(0x100F, Lv_Move(row--,ins)-1, viewpos, LV.Hwnd)	
	; LVM_SETITEMPOSITION 0x100F (0x1000+15)
	; Moves an item to a specified position in a list-view control (must be in icon or small icon view). 
	; wParam	Index of the list-view item.
	; lParam	LOWORD specifies the new x-position of the item's upper-left corner, in view coordinates. 
	;			HIWORD specifies the new y-position of the item's upper-left corner, in view coordinates.
	; Returns TRUE if successful, or FALSE otherwise.
	; If the list-view control has the LVS_AUTOARRANGE style, the items in the list-view control 
	; are arranged after the position of the item is set.
	; On Windows Vista, sending this message to a list-view control with 
	; the LVS_AUTOARRANGE style does nothing, and the return value is FALSE.
	; View coordinates can be retrieved with LVM_GETITEMPOSITION, and is the 
	; virtual coordinates within the entire listview (i.e. y can be > screen height)
			LV.Opt("0x2100")	; switch on LVS_NOSCROLL & LVS_AUTOARRANGE to rearrange items afterwards 
								; according to item index without changing the scroll position
			LV.Opt("-0x2000")	; restore scrolling
		}
	}
	return

	Lv_Move(src,dest) {
        i:=LV_GetIcon(src,LV.Hwnd)
   	    t:=LV.GetText(src)
		LV.Delete(src)
		return LV.Insert(dest,"Icon" i, t)
	}
}
