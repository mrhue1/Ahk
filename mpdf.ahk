#Requires AutoHotkey v2.0
#SingleInstance Force
#DllLoad "libmupdf.dll"
SetIcon()
pdf()

pdf(fn:="", thumbw:=150, thumbh:=200) {
	static pvGui, LV, ctx, PdfDocs := map(), Passwords := map(), Loading:=0, arrows := "▲►▼◄", ImgLst
	if !isSet(pvGui) {
		; fz_context *fz_new_context_imp(fz_alloc_context *alloc, fz_locks_context *locks, unsigned int max_store, const char *version);
		if !ctx := DllCall("libmupdf\fz_new_context_imp","Ptr",0,"Ptr",0,"UInt",0,"AStr","1.25.4","Cdecl Ptr")
			throw error("Can't create libmupdf context: Wrong version of dll",-1,"libmupdf.dll")
		pvGui := Gui("+Resize -DPIScale",fn)
		pvGui.OnEvent("Close",ExitPdf)
		pvGui.OnEvent("Escape",(*)=>Loading:=0)
		pvGui.OnEvent("DropFiles",DropPDF)

AboutText := '
(
	Open starts a new list of PDF pages
	Add appends PDF to existing list
	<Esc> to cancel PDF loading
	Rearrange pages with mouse / keyboard (Ctrl-C copy, Ctrl-X cut, Ctrl-V paste)
	Delete page(s) by pressing <Del>
	Select pages using shift/control + arrow keys/mouse-click
)'
		pvMenu := MenuBar()
		pvMenu.Add("&Open",OpenPDF)
		pvMenu.Add("&Add", AddPDF)
		pvMenu.Add("&Save", SavePDF)
		pvMenu.Add("&Rotate", RotatePDF)
		pvMenu.Add("C&ut", LV_Cut)
		pvMenu.Add("&Copy", LV_Copy)
		pvMenu.Add("&Paste", LV_Paste)
		pvMenu.Add("&Delete", LV_Delete)
		pvMenu.Add("&Help", (*) => MsgBox(AboutText, "PDF Help"))
		pvGui.MenuBar := pvMenu
		pvGui.Marginx := pvGui.MarginY := 0

		; Create PDF thumbnail ListView
		MonitorGetWorkArea(1,&Left, &Top, &Right, &Bottom)
		LV:=pvGui.AddListView("Icon -Hdr LV0x10000 w" (Right-Left)*3//5 " h" Bottom-Top-90,["pdf"])	
		; -90 for title bar & buttons, width 60% for thumbnails
		; Icon to display label below image (Tile displays text right of image)
		; LVS_EX_DOUBLEBUFFER LV0x10000 Paints via double-buffering, which reduces flicker.
		LV.OnEvent("ItemSelect",preview)
 		LV.OnNotify(-109, LVN_BEGINDRAG)
		pvGui.Show("x-4 y-4")	; show on left to allow preview window on right
		PdfDocs.default := "", PdfDocs.CaseSense := 0, Passwords.default := ""
		GroupAdd("pvGui","ahk_id " pvGui.Hwnd)
		if !fn {
			if !A_Args.length	; FileOpen dialogue
 				AddPdf()
			else for n,arg in A_Args	; or load commandline filespecs
				Loop Files, arg
					pdfLoad(A_LoopFileFullPath)
			return pvGui
		}
	}
	if IsObject(fn) {
		for f in fn
			pdfLoad(f)
	} else Loop Files, fn
		pdfLoad(A_LoopFileFullPath)
	return pvGui

	ExitPDF(*) {
		For k,v in PdfDocs
			DllCall("libmupdf\pdf_drop_document","Ptr",ctx,"Ptr",v,"Cdecl")
		DllCall("libmupdf\fz_drop_context","Ptr",ctx,"Cdecl")
		pvGui.Destroy()
		ExitApp	; required due to hotkey
	}

	AddPDF(*) {
		if fns:= FileSelect("M3",, "Add PDFs", "PDF (*.pdf)")
			Queue()

		Queue() {	; add PDFs after any existing pdfLoad
			if InStr(pvGui.Title," loading... (")
				SetTimer(Queue,-100)
			else for fn in fns
				pdfLoad(fn)
		}
	}

	OpenPDF(*) {
		if fns:= FileSelect("M3",, "Add PDFs", "PDF (*.pdf)")
			for fn in fns
				pdfLoad(fn,A_Index!=1)
		Loading := 0	; stop any existing pdfLoad from resuming
	}

	DropPdf(GuiObj, GuiCtrl, FileArray, X, Y) {
		NumPut("Int", X, "Int", Y, LVHITTESTINFO := Buffer(24,0))
		SendMessage(0x1012, 0, LVHITTESTINFO, LV.hwnd)	; LVM_HITTEST to find out row being dropped on
		row := NumGet(LVHITTESTINFO,12,"Int")
		for i, DroppedFile in FileArray {
    			row := pdfLoad(DroppedFile,row+1)
			LV_Refresh(LV)
		}
	}

	pdfLoad(fn,row:=2147483647) {	; row>=1 to add pdf file to the pdf thumbnails ListView
		static Pwd
		if !fn || !FileExist(fn)
			return

		Loading := A_TickCount
		if !PdfDocs.has(fn) {
			StrPut(fn, utf := Buffer(StrPut(fn, "UTF-8")), "UTF-8")		; convert to UTF-8 for mupdf
			doc := DllCall("libmupdf\pdf_open_document","Ptr",ctx,"Ptr",utf,"Cdecl Ptr")
			; int pdf_needs_password(fz_context *ctx, pdf_document *doc);
			if DllCall("libmupdf\pdf_needs_password","Ptr",ctx,"Ptr",doc,"Cdecl Int") {
				if Passwords[fn]
					Pwd := Passwords[fn]
				if !IsSet(Pwd) || !DllCall("libmupdf\pdf_authenticate_password","Ptr",ctx,"Ptr",doc,"AStr",pwd,"Cdecl Int")
					Loop {
						res := InputBox(fn " is protected. Please enter a Document Open Password.","Password","w580 h130")
	 					if res.result != "OK"
	 						return
				; int pdf_authenticate_password(fz_context *ctx, pdf_document *doc, const char *pw);
				; Returns 0 for failure, non-zero for success.
				; In the non-zero case:
				; bit 0 set => no password required
				; bit 1 set => user password authenticated
				; bit 2 set => owner password authenticated
					} until DllCall("libmupdf\pdf_authenticate_password","Ptr",ctx,"Ptr",doc,"AStr",Pwd:=res.value,"Cdecl Int")
				Passwords[fn]:=Pwd
			}
			PdfDocs[fn] := doc
		} else doc := PdfDocs[fn]

		if !PageCount := DllCall("libmupdf\pdf_count_pages","Ptr",ctx,"Ptr",doc,"Cdecl UInt")
			return

		LV.focus()
		if !row && IsSet(ImgLst)
			IL_Destroy(ImgLst), ImgLst := Unset, LV.Delete()

		if !IsSet(ImgLst) {					; Initialize ImageList for listview
			ImgLst := IL_Create(PageCount,5,1)	; Icon view use large icons
	; ImageList_SetIconSize sets the dimensions of images in an image list and removes all images from the list.
			DllCall("ComCtl32.dll\ImageList_SetIconSize","Ptr",ImgLst,"Int",thumbw,"Int",thumbh,"Int")	
	; LVM_SETICONSPACING:=0x1035, wParam 0, lParam width (loword) + height (hiword)
	; Returns a DWORD with previous width in low word and height in high word.
	; Values are relative to upper-left corner of an icon bitmap. 
	; Therefore, to set spacing between icons that do not overlap, 
	; values must include the size of the icon, plus the amount of empty space desired between icons. 
	; Values that do not include the width of the icon will result in overlaps.
	; When defining icon spacing, values must be 4 or larger. Smaller values will not yield the desired layout. 
	; To reset the icons to the default spacing, set lParam to -1.
			wh := SendMessage(0x1035,0,(thumbw+4)|((thumbh+53)<<16), LV.Hwnd)		; guestimate w & h to include label and margin
			LV.SetImageList(ImgLst,0)			; NB set LV ImageList AFTER the iconsize has been set
		} 

		Loop PageCount {
			hBMP := renderPage(doc,A_Index,thumbw,thumbh,1)	; do the most time consuming part before checking to abort
			if !Loading		; check if need to abort before we add the image
				return pvGui.title .= " - Cancelled"
			; Add and show the images concurrently for the user
			row:=LV.Insert(row,"Icon" IL_Add(ImgLst, "HBITMAP:" hBMP), fn ": " A_Index "/" PageCount " " SubStr(arrows,mod(GetRotate(doc,A_Index)//90,4)+1,1))+1
			pvGui.title := fn " loading... (" A_Index "/" PageCount ")"
		}
		pvGui.title := fn " (" PageCount " pages, " A_Tickcount-Loading "ms)"
		return row
	}

	renderPage(doc,Page,imgw,imgh,iconview:=0,&fz_rect?,&pgw?,&pgh?) {	
	; iconview=1 places image at bottom of a blank bitmap of size imgw*imgh, otherwise return bitmap without blank padding
		static BITMAPINFO := Buffer(40,0), hDC := DllCall('GetDC', 'Ptr', 0, 'Ptr')
				, fz_colorspace := DllCall("libmupdf\fz_device_bgr","Ptr",ctx)	; bitmap is bgr
		pdf_page := DllCall("libmupdf\pdf_load_page","Ptr",ctx,"Ptr",doc,"int",Page-1,"Cdecl Ptr")	; page is 0 based
		; fz_rect pdf_bound_page(fz_context *ctx, pdf_page *page, fz_box_type box);
		;	FZ_MEDIA_BOX,
		;	FZ_CROP_BOX,
		;	FZ_BLEED_BOX,
		;	FZ_TRIM_BOX,
		;	FZ_ART_BOX,
		;	FZ_UNKNOWN_BOX
		; For functions returning a structure, put the structure as first parameter in DllCall
		DllCall("libmupdf\pdf_bound_page","Ptr",fz_rect := buffer(16),"Ptr",ctx,"Ptr",pdf_page,"UInt",1,"Cdecl Ptr")	; fz_crop_box=1
		pgw := NumGet(fz_rect,8,"Float") - NumGet(fz_rect,"Float")
		pgh := NumGet(fz_rect,12,"Float") - NumGet(fz_rect,4,"Float")
;		msgbox pgw "x" pgh
		if imgw || imgh {		; setup resize scale
			if !imgh || (imgw/pgw<imgh/pgh)
				r := imgw/pgw
			else r := imgh/pgh
		} else r := 1
		; fz_matrix fz_scale(float sx, float sy);
		DllCall("libmupdf\fz_scale","Ptr",fz_matrix:=buffer(24),"Float",r, "Float",-r,"Cdecl Ptr")	
		; -r for y for bottom up bitmap (so bitmap is drawn from bottom of icon)
		; fz_matrix fz_pre_rotate(fz_matrix m, float degrees);
		; pix = fz_new_pixmap_from_page_number(ctx, doc, page_number, ctm, fz_device_rgb(ctx), 0);
		; fz_pixmap *fz_new_pixmap_from_page(fz_context *ctx, fz_page *page, fz_matrix ctm, fz_colorspace *cs, int alpha);
		pix := DllCall("libmupdf\fz_new_pixmap_from_page","Ptr",ctx,"Ptr",pdf_page,"Ptr", fz_matrix, "Ptr", fz_colorspace, "Int", 1, "Cdecl Ptr")	; include alpha for bitmap
		; fz_new_pixmap_from_page_with_separations(ctx, page, ctm, cs, NULL, alpha);
		pixw := DllCall("libmupdf\fz_pixmap_width","Ptr",ctx,"Ptr",pix)
		pixh := DllCall("libmupdf\fz_pixmap_height","Ptr",ctx,"Ptr",pix)
		size := pixw*pixh*4		; save size, because pgh can get altered
		if iconview && pixh < imgh
			pixh := imgh
		NumPut('UInt', 40, 'UInt', pixw, 'Int', pixh, 'UInt', 0x200001, BITMAPINFO)	; 0x200001 = 'UShort', planes := 1, 'UShort', bpp:=32
		hBitmap := DllCall('CreateDIBSection','Ptr',hDC,'Ptr',BITMAPINFO,'UInt',0,'PtrP', &pBits:=0, 'Ptr', 0, 'UInt', 0, 'Ptr')
		; CreateDIBSection to create a blank bitmap
		; msgbox size_t := DllCall("libmupdf\fz_pixmap_size","Ptr",ctx, "Ptr", pix) "`n" pgw*pgh*4
		; unsigned char *fz_pixmap_samples(fz_context *ctx, const fz_pixmap *pix);
		src := DllCall("libmupdf\fz_pixmap_samples", "Ptr",ctx,"Ptr",pix)
		; void *memcpy(void *dest, const void *src, size_t count);
		DllCall("ntdll\memcpy", "Ptr", pBits, "Ptr", src, "UInt", size)
		DllCall("libmupdf\fz_drop_pixmap","Ptr",ctx, "Ptr",pix)
		DllCall("libmupdf\pdf_drop_page","Ptr",ctx,"Ptr",pdf_page)
		return hBitmap
	}

	GetRotate(doc,page,&dict?) {
		; get page dictionary
		dict := DllCall("libmupdf\pdf_lookup_page_obj","Ptr",ctx,"Ptr",doc,"Int",page-1,"Cdecl Ptr")	; page is 0 based
		; get Rotate key from dictionary
		rotation := DllCall("libmupdf\pdf_dict_gets","Ptr",ctx,"Ptr",dict,"AStr","Rotate","Cdecl Ptr")
		; convert to int
		return DllCall("libmupdf\pdf_to_int","Ptr",ctx,"Ptr",rotation,"Cdecl Int")
	}

	RotatePDF(*) {
		row := 0
		focus := LV.GetNext(0,"F")
		While row := LV.GetNext(row) {
			RegExMatch(LV.GetText(row),"(.+): (\d+)/(\d+) (.)$",&pg)	; get pdf name, page, total pages, rotation
			rotation := GetRotate(doc := PdfDocs[pg.1],pg.2,&dict)	
			; add rotate to existing Rotate
			rotate := DllCall("libmupdf\pdf_new_int","Ptr",ctx,"int64",rotation = 270 ? 0 : rotation + 90,"Cdecl Ptr")
			; store back to dictionary
			DllCall("libmupdf\pdf_dict_puts_drop","Ptr",ctx,"Ptr",dict,"AStr","Rotate","Ptr",rotate,"Cdecl")
			hBmp := renderPage(doc,pg.2,thumbw,thumbh,1)
			icon := LV_GetIcon(row,LV.Hwnd)
			DllCall("ImageList_Replace", "Ptr", ImgLst, "int", icon-1, "Ptr", hBmp, "Ptr", 0)
			DllCall("DeleteObject", "Ptr", hBmp)
			LV.Modify(row,, pg.1 ": " pg.2 "/" pg.3 " " SubStr(arrows,mod(rotation//90+1,4)+1,1))
			if focus=row
				preview(LV,row,1)
		}
	}

	SavePDF(*) {
		fn := RegExReplace(LV.GetText(1),".+\\|: .+")
		while f := FileSelect(16,fn,"Save As","Document/Image (*.pdf; *.txt; *.html; *.xml; *.json; *.xhtml; *.jpg; *.png; *.svg)") {
			if !RegExMatch(f,"\.\w+$")
				f:=RegExReplace(f,"\.?$", ".pdf")
			if PdfDocs[f] {
				msgbox "Cannot save to an opened file.  Please enter a new filename."
				continue
			}
			StrPut(f,utf:=Buffer(StrPut(f,"UTF-8")),"UTF-8")	; convert filename to UTF-8 for mupdf
			ext := RegExReplace(f,".+\.")
			dst := DllCall("libmupdf\pdf_create_document","Ptr",ctx,"Cdecl Ptr")
			graft := DllCall("libmupdf\pdf_new_graft_map","Ptr",ctx,"Ptr",dst,"Cdecl Ptr")
 			Loop Parse, ListViewGetContent("Col1",LV.Hwnd), "`n"	; Go through list of pages to output
 				if RegExMatch(A_LoopField,"(.+): (\d+)/(\d+) (.)$",&pg) 	; page is 0 based
					DllCall("libmupdf\pdf_graft_mapped_page","Ptr",ctx,"Ptr",graft,"int",-1,"Ptr",PdfDocs[pg.1],"int",pg.2-1,"Cdecl")
			DllCall("libmupdf\pdf_drop_graft_map","Ptr",ctx,"Ptr",graft,"Cdecl")
			if ext="pdf" {
				Numput("Int",1,"Int",1,"Int",1,"Int",0,"Int",4,"Int",1,options := buffer(20*4+256,0),12)
; int do_incremental; /* Write just the changed objects. */
; int do_pretty; /* Pretty-print dictionaries and arrays. */
; int do_ascii; /* ASCII hex encode binary streams. */
; int do_compress; /* Compress streams. */
; int do_compress_images; /* Compress (or leave compressed) image streams. */
; int do_compress_fonts; /* Compress (or leave compressed) font streams. */
; int do_decompress; /* Decompress streams (except when compressing images/fonts). */
; int do_garbage; /* Garbage collect objects before saving; 1=gc, 2=re-number, 3=de-duplicate. */
; int do_linear; /* Write linearised. */
; int do_clean; /* Clean content streams. */
; int do_sanitize; /* Sanitize content streams. */
; int do_appearance; /* (Re)create appearance streams. */
; int do_encrypt; /* Encryption method to use: keep, none, rc4-40, etc. */
; int dont_regenerate_id; /* Don't regenerate ID if set (used for clean) */
; int permissions; /* Document encryption permissions. */
; char opwd_utf8[128]; /* Owner password. */
; char upwd_utf8[128]; /* User password. */
; int do_snapshot; /* Do not use directly. Use the snapshot functions. */
; int do_preserve_metadata; /* When cleaning, preserve metadata unchanged. */
; int do_use_objstms; /* Use objstms if possible */
; int compression_effort; /* 0 for default. 100 = max, 1 = min. */
				DllCall("libmupdf\pdf_save_document","Ptr",ctx,"Ptr",dst,"Ptr",utf,"Ptr",options,"Cdecl")
				DllCall("libmupdf\pdf_drop_document","Ptr",ctx,"Ptr",dst,"Cdecl")
			} else try {
				if InStr("json,xml", ext)
					ext := "stext." ext
				else if ext = "htm"
					ext := "html"
				wri := DllCall("libmupdf\fz_new_document_writer","Ptr",ctx,"Ptr",utf,"AStr",ext,"Ptr",0,"Cdecl Ptr")
				DllCall("libmupdf\fz_write_document","Ptr",ctx,"Ptr",wri,"Ptr",dst,"Cdecl")
				DllCall("libmupdf\fz_close_document_writer","Ptr",ctx,"Ptr",wri,"Cdecl")
				DllCall("libmupdf\fz_drop_document_writer","Ptr",ctx,"Ptr",wri,"Cdecl")
			}
			msgbox "Done"
			return
		}
	}

	preview(LV,row,selected) {
		static pgGui, hText, vText, pgSel, pgPic, pg, fz_stext_page, fz_stext_options := Buffer(8,0), pgw, pgh, rotate
		if !selected
			return
		RegExMatch(title:=LV.GetText(row),"(.+): (\d+)/(\d+) (.)$",&pg)
		MonitorGetWorkArea(1,&Left,,&Right)
		LV.gui.GetClientPos(,,,&lvh)
		LV.gui.GetPos(&x,&y,&guiw)
		x+=guiw-20	; place preview to right of main window
		width := Right-Left-x, height := lvh

		doc := PdfDocs[pg.1]
		hBmp := renderPage(doc,pg.2,width,height,0,&fz_rect,&pgw,&pgh)
		rotate:=90*(Instr(arrows,pg.4)-1)	; rotate := GetRotate(doc,pg.2)

		if isSet(fz_stext_page)
			DllCall("libmupdf\fz_drop_stext_page","Ptr",ctx, "Ptr",fz_stext_page,"Cdecl")
		fz_stext_page := DllCall("libmupdf\fz_new_stext_page_from_page_number","Ptr",ctx, "Ptr",doc,"Int",pg.2-1,"Ptr",fz_stext_options,"Cdecl Ptr")

		if !isSet(pgGui) {
			pgGui:= Gui("+Resize -DPIScale +Owner" pvGui.Hwnd,title)
			pgGui.BackColor := 0xFFFFFF
			pgGui.Marginx := pgGui.MarginY := 0
			pgSel:=pgGui.AddEdit("x0 y0 BackgroundBlue +0x4000000 -VScroll -E0x200 w1 h1")
			hText:=pgGui.AddEdit("x0 y0 BackgroundLime +0x4000000 -VScroll -E0x200 w1 h1")	; so that cursor changes to I-Beam on hover over text
			vText:=pgGui.AddEdit("x0 y0 BackgroundYellow +0x4000000 -VScroll -E0x200 w1 h1")
			WinSetTransparent(1, hText.Hwnd)	; set to e.g. 80 to see the text region
			WinSetTransparent(1, vText.Hwnd)
			WinSetTransparent(80, pgSel.Hwnd)
			pgPic:=pgGui.Add("Picture","x0 y0","HBITMAP:" hBmp)
			pgGui.OnEvent("Escape",(*)=>pgGui.Hide())
			pgGui.OnEvent("Close",(*)=>pgGui.Hide())
			pgGui.OnEvent("Size",pgSize)
			OnMessage(0x201, WM_LBUTTONDOWN)
			pgGui.Show("AutoSize NA x" x " y" y)
			GroupAdd("pgGui","ahk_id " pgGui.Hwnd)
		} else {
			pgGui.Title := title
			pgPic.Value := "*w0 *h0 HBITMAP:" hBmp
			SetTextRegions()
			pgGui.Show("AutoSize NA")
		}

;typedef struct
;{
;	fz_pool *pool;
;	fz_rect mediabox;
;	fz_stext_block *first_block;
;
;	/* The following fields are only of use to the routines that
;	 * build an fz_stext_page. They change during page construction
;	 * and their meaning is subject to change. These values should
;	 * not be used by anything outside of the stext device. */
;	fz_stext_block *last_block;
;	fz_stext_struct *last_struct;
;} fz_stext_page;
;
;enum
;{
;	FZ_STEXT_BLOCK_TEXT = 0,
;	FZ_STEXT_BLOCK_IMAGE = 1,
;	FZ_STEXT_BLOCK_STRUCT = 2,
;	FZ_STEXT_BLOCK_VECTOR = 3,
;	FZ_STEXT_BLOCK_GRID = 4
;};
;
;/**
;	A text block is a list of lines of text (typically a paragraph),
;	or an image.
;*/
;struct fz_stext_block
;{
;	int type;
;	fz_rect bbox;	// float x0, y0; x1, y1;					// 16 bytes
;	union {
;		struct { fz_stext_line *first_line, *last_line; } t;
;		struct { fz_matrix transform; fz_image *image; } i;		// fz_matrix:=buffer(24)
;		struct { fz_stext_struct *down; int index; } s;
;		struct { uint8_t stroked; uint8_t rgba[4]; } v;			// uint8_t	1 byte unsigned integer
;		struct { fz_stext_grid_positions *xs; fz_stext_grid_positions *ys; } b;
;	} u;
;	fz_stext_block *prev, *next;
;};
;struct fz_stext_line
;{
;	int wmode; /* 0 for horizontal, 1 for vertical */
;	fz_point dir; /* normalized direction of baseline */
;	fz_rect bbox;									// x0 := NumGet(line, 12, "float")
;	fz_stext_char *first_char, *last_char;			// char := NumGet(line, A_PtrSize+24, "Ptr")
;	fz_stext_line *prev, *next;
;};
;struct fz_stext_char
;{
;	int c; /* unicode character value */
;	uint16_t bidi; /* even for LTR, odd for RTL - probably only needs 8 bits? */
;	uint16_t flags;
;	int color; /* sRGB hex color */
;	fz_point origin;
;	fz_quad quad;									//	x0:=NumGet(char, 20, "float")
;	float size;
;	fz_font *font;
;	fz_stext_char *next;
;};

		pgSize(pgGui, MinMax, Width, Height) {
			if !MinMax {
				pgPic.Move(,,width,height)
				SetTextRegions()
			} 
		}

		SetTextRegions() {
			pgPic.GetPos(&x,&y,&w,&h)
			hText.Move(x,y,w,h)
			vText.Move(x,y,w,h)
			pgSel.Move(0,0,1,1)
			wratio := w/pgw, hratio := h/pgh
			if block := NumGet(fz_stext_page, A_PtrSize+16, "Ptr") {
				Loop {	; mark window region occupied by horizontal and vertical text blocks
					if !NumGet(block, "Int") {	; FZ_STEXT_BLOCK_TEXT
						x0 := round(NumGet(block, 4, "float")*wratio)
						y0 := round(NumGet(block, 8, "float")*hratio)
						x1 := round(NumGet(block, 12, "float")*wratio)
						y1 := round(NumGet(block, 16, "float")*hratio)
						line := NumGet(block, A_PtrSize+16, "Ptr")
						if NumGet(line, "Int") {		; vertical text ? fz_rect is off by line height ?
							ht := round((NumGet(line, 24, "float")-NumGet(line, 16, "float"))*hratio)
							switch rotate {
								case 0: y0 -= ht, y1 -= ht
								case 90: x0 += ht, x1 += ht
								case 180: y0 += ht, y1 += ht
								case 270: x0 -= ht, x1 -= ht
							}
							if IsSet(vregion) {
								vregion2 := DllCall("CreateRectRgn","int",x0,"int",y0,"int",x1,"int",y1,"Ptr")
								DllCall("CombineRgn","Ptr",vregion,"Ptr",vregion,"Ptr",vregion2,"int",2)	; RGN_AND = 1, RGN_OR = 2, RGN_XOR = 3, RGN_DIFF = 4, RGN_COPY = 5
							} else vregion := DllCall("CreateRectRgn","int",x0,"int",y0,"int",x1,"int",y1,"Ptr")
						} else if IsSet(hregion) {
							hregion2 := DllCall("CreateRectRgn","int",x0,"int",y0,"int",x1,"int",y1,"Ptr")
							DllCall("CombineRgn","Ptr",hregion,"Ptr",hregion,"Ptr",hregion2,"int",2)	; RGN_AND = 1, RGN_OR = 2, RGN_XOR = 3, RGN_DIFF = 4, RGN_COPY = 5
						} else hregion := DllCall("CreateRectRgn","int",x0,"int",y0,"int",x1,"int",y1,"Ptr")
					}
				} until !block:= NumGet(block, A_PtrSize*3+40, "Ptr")
			}
			if hText.Visible:=IsSet(hregion)
				DllCall("SetWindowRgn","Ptr",hText.Hwnd,"Ptr",hregion,"uint",Redraw:=1)
			if vText.Visible:=IsSet(vregion)
				DllCall("SetWindowRgn","Ptr",vText.Hwnd,"Ptr",vregion,"uint",Redraw:=1)
			WinRedraw(pgGui.Hwnd)
		}

		WM_LBUTTONDOWN(*) {
 			pgPic.GetPos(&x,&y,&w,&h)
			wratio := w/pgw, hratio := h/pgh
 			MouseGetPos(&x1, &y1)
 			if y1<y || y1>y+h || x1<x || x1>x+w
 				return							; clicking on menu / title bar
 			While GetKeyState("LButton", "P") {
 				MouseGetPos(&x2, &y2)
 				pgSel.Move(x:=min(x1,x2),y:=min(y1,y2),abs(x2-x1),abs(y2-y1))
 			}
	    		NumPut("Float",x/wratio,"Float",y/hratio,"float",max(x1,x2)/wratio,"float",max(y1,y2)/hratio,area:=Buffer(16))
			if str := StrGet(DllCall("libmupdf\fz_copy_rectangle","Ptr",ctx,"Ptr",fz_stext_page,"Ptr",area,"int",0,"Cdecl Ptr"),"UTF-8")
				tooltip("Copied: " A_Clipboard := str)
			SetTimer(tooltip,-3000)
		}
	}
}

SaveHBITMAP(hBitmap, sFile) {
	DIBSECTION := Buffer(A_PtrSize*5+64, 0)
	NumPut("UInt",40, DIBSECTION, A_PtrSize*2+16)	;dsBmih.biSize
	DllCall("GetObject", "Ptr", hBitmap, "int", DIBSECTION.size, "Ptr", DIBSECTION)
	pBits := NumGet(DIBSECTION, A_PtrSize+16, "Ptr")
	; Overwrites BITMAP section of DIBSECTION with 14 bytes of BMP file header 
	NumPut("UShort",0x4D42,"UInt",14+40+biSizeImage:=NumGet(DIBSECTION, A_PtrSize*2+36, "UInt"),"Int64",54<<32,DIBSECTION,A_PtrSize*2+2)
	out := FileOpen(sFile, "w")
	out.RawWrite(DIBSECTION.ptr + A_PtrSize*2+2,54)
	out.Rawwrite(pBits, biSizeImage)
	out.Close()
}

#HotIf WinActive("ahk_Group pvGui")
; In icon view left arrow doesn't wrap (move to previous line) at start of a row
; Need to handle the arrow keys to enable autowrap
; This requires setting item focus, selections, and selection mark 
^left::LV_Left("^{left}")
+left::LV_Left("+{left}")
left::LV_Left("{left}")

LV_Left(key) {	
	LV := GuiFromHwnd(WinExist())["SysListView321"]
	if LV.focused {
		View:=SendMessage(0x108F, 0, 0, LV.Hwnd)
		if View&1 { 	; list/report view
			Send(key)
		} else if 1 < row:=LV.GetNext(0,"F") {
			Send(key)	; normal key navigation to handle selections
			if SendMessage(0x102C,row-1,1,LV.Hwnd) { 	; if focus did not change 
; LVM_GETITEMSTATE:= 0x102C
; wParam Index of the list-view item.
; lParam State information to retrieve. 
; Bits 0 through 7 contain the item state flags
;	LVIS_FOCUSED            0x0001	; The item has the focus, so it is surrounded by a standard focus rectangle.
									; Although more than one item may be selected, only one item can have the focus.
;	LVIS_SELECTED           0x0002	; The item is selected. The appearance of a selected item depends on whether it has the focus 
									; and also on the system colors used for selection.
;	LVIS_CUT                0x0004	; The item is marked for a cut-and-paste operation.
;	LVIS_DROPHILITED        0x0008	; The item is highlighted as a drag-and-drop target.
;	LVIS_GLOW               0x0010
;	LVIS_ACTIVATING         0x0020	; Not currently supported.
							; Bits 8 through 11 specify the one-based overlay image index. 
							; The overlay image is superimposed over the item's icon image.
;	LVIS_OVERLAYMASK        0x0F00	; The item's one-based overlay image index is retrieved by a mask.
							; Bits 12 through 15 of this member specify the state image index. 
							; The state image is displayed next to an item's icon
;	LVIS_STATEIMAGEMASK     0xF000	; The item's one-based state image index is retrieved by a mask.
				if GetKeyState("Shift") {
					if SendMessage(0x102C,row-2,2,LV.Hwnd) {	; previous item already selected ?
						LV.Modify(row,"-Select")
						LV.Modify(row-1,"Focus Vis")
					} else LV.Modify(row-1,"Focus Select Vis")
				} else if GetKeyState("Control") {
					LV.Modify(row-1,"Focus Vis")
				} else {
					LV.Modify(row,"-Select")
					LV.Modify(row-1,"Focus Select Vis")
					SendMessage(0x1043,0,row-2,LV.Hwnd)	; LVM_SETSELECTIONMARK := 0x1043 
; LVM_SETSELECTIONMARK := 0x1043 ; (LVM_FIRST + 67) 
; wParam Must be zero.
; lParam Zero-based index of the new selection mark. If set to -1, the selection mark is removed.
; Returns the previous selection mark, or -1 if there is no previous selection mark.
; The selection mark is the item index from which a multiple selection starts. 
; This message does not affect the selection state of the item.
				} 
			} 
		}
	} else Send(key)
}

; In icon view right arrow doesn't wrap (move to next line) at end of a row
; At end of the list, it also jumps to the right icon in the previous row
; Need to handle the arrow keys to enable autowrap and not move to previous row 
; This requires setting item focus, selections, and selection mark
^right::LV_Right("^{right}")	; can't use A_ThisHotkey which is ^right
+right::LV_Right("+{right}")
right::LV_Right("{right}")

LV_Right(key) {	
	LV := GuiFromHwnd(WinExist())["SysListView321"]
	if LV.focused {
		View:=SendMessage(0x108F, 0, 0, LV.Hwnd)
		if View&1 { 	; list/report view
			Send(key)
		} else if LV.GetCount() > row:=LV.GetNext(0,"F") {
			Send(key)	; normal key navigation to handle selections
			if SendMessage(0x102C,row-1,1,LV.Hwnd) { ; if LVM_GETITEMSTATE is still focused
				if GetKeyState("Shift") {
					if SendMessage(0x1042,0,0,LV.Hwnd)>row	; erasing selection?
; tooltip(SendMessage(0x1042,0,0,LV.Hwnd))
; LVM_GETSELECTIONMARK := 0x1042 
; Returns the zero-based selection mark, or -1 if there is no selection mark.
; The selection mark is the item index from which a multiple selection starts.
						LV.Modify(row,"-Select"), LV.Modify(row+1,"Focus Vis")
					else LV.Modify(row+1,"Focus Select Vis")
				} else if GetKeyState("Control") {
					LV.Modify(row+1,"Focus Vis")
				} else {
					LV.Modify(row,"-Select")
					LV.Modify(row+1,"Focus Select Vis")
					SendMessage(0x1043,0,row,LV.Hwnd)	; LVM_SETSELECTIONMARK := 0x1043 
				} 
			} 
		}
	} else Send(key)
}

LV_Init() {
	global LV_Clip
	LV_Clip:=[]
	pvGui:=GuiFromHwnd(WinExist())
	return  pvGui["SysListView321"]
}

LV_Refresh(LV) {
	View:=SendMessage(0x108F, 0, 0, LV.Hwnd)
	if !(View&1){	; If not list/report view
		LV.Opt("-Redraw 0x2000")				; Disable redraw to hide that we are switching view
		SendMessage(0x108E, 3, 0, LV.Hwnd)		; Set to list view to order items, LVM_SETVIEW := 0x108E ; (LVM_FIRST + 142)
		SendMessage(0x108E, View, 0, LV.Hwnd)	; Set back to original view to display items
		LV.Opt("-0x2000")					; Re-enable scrolling 
		LV.Opt("+Redraw")					; Enable redraw
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
		ControlSend("{space}", LV)
		LV_Refresh(LV)
	}

^v::
	LV_Paste(*) {
		LV:=GuiFromHwnd(WinExist())["SysListView321"]
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
		ControlSend("{space}", LV)
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
	;		relative to the current position of the list view content. 
	;		If the list-view control is in list view, this value is 
	;		rounded up to the nearest number of pixels that form a whole column.
	; lParam	int that specifies the amount of vertical scrolling, in pixels, 
	;		relative to the current position of the list view content.
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
	;	x:=x-NumGet(POINT,"Int"), y:=y-NumGet(POINT,4,"Int") 
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
	;		HIWORD specifies the new y-position of the item's upper-left corner, in view coordinates.
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

SetIcon() {
; https://www.autohotkey.com/boards/viewtopic.php?f=82&t=118167
Base64PNG := ' 
(
	iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAA
	Cxjwv8YQUAAAJ6SURBVHjaYhiUoKioiMnY2JgF0H45wMwRRVH4rLe2bdu2jaB2w9p2o9q24tS2
	bbdRbRvTPeneZN9mXn7NTnmTszvzdL+LHy8gT0DeOMpTpUoV16hRoxyIyooUKeKaHA+V1sXHro
	0JYFipwJk7Z8VD+Q4dOpiDJEmSxDMzHgZwcSS1OD56BiCcCI+8rw8NuMAOTfSjaq9evZRMJF0a
	H0fsAlgRHwcA+KXhSJKTE3YKQGrQunfvznoUsx8AWfmTIQAlrDh0W/Wqxqu7d43bO3cqY48uXz
	ZonON7ECC75QBHx4w2vn76pACIc7FzixdFDcBF906fNt4+eaJsfnb7tnGgbx8tAPcRgCB8392m
	tfHxzRuDkZ+cOoXPnNMDyCF6U9MYJkYr88pZ/Kbjd8+eEUoLoNSR0R8dO0ap5bNbtxgFD9HVnx
	AyxlJo1usBJG1ykOJAiVDGNDU+NW0ay6HMXd+4IQoA80biocwIwaQHCMox9gpBuF6pv2REV7Ko
	AH46fPVK4MWRUm8q9F3Src+kKOomJARrzndxrO32S2vXcospADMTI4DwVIZrW/qkSnOJcykRxz
	SljAJAk8pwCWSIcY8yxrpfXLWSj5EBYCMyYkYonS0p5zijlwwwQzECsEv/ARSA+h40+udLUGR+
	PJyxy/mCeDgFIAsBwA++1HVjkF0AjTzoBSA5QiwhgCo9fVgUaee9fZgPoGjgHuKFWPC2kgZAnZ
	pujJwXD2etdswza7sxDEBZAInMrmVOAKkBlAPQEkB7AJ0tUCcA7QA0ZuR0vnTpUvPrWfC24geQ
	DkB+AMUBlIijigHIDSClkvaorEqVKk7eaq1QkyZNtDfjH1FGlIE8CpmmAAAAAElFTkSuQmCC
)'

	size := StrLen( RTrim(Base64PNG, '=') )*3//4, width := 0, height := 0
	DllCall('Crypt32\CryptStringToBinary', 'Str', Base64PNG, 'UInt', StrLen(Base64PNG), 'UInt', 1, 'Ptr', buf := Buffer(size), 'UIntP', &size, 'Ptr', 0, 'Ptr', 0)
	icon := DllCall('CreateIconFromResourceEx', 'Ptr', buf, 'UInt', size, 'UInt', 1, 'UInt', 0x30000, 'Int', width, 'Int', height, 'UInt', 0)
	TraySetIcon('HICON: ' icon)
}
