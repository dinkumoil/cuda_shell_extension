{ **************************************************************************** }
//
// CudaText context menu handler for Windows Explorer
//
// Version: 1.1
// Author:  Andreas Heim, 2019-2020
// Website: https://github.com/dinkumoil/cuda_shell_extension
//
//
// The following websites provided great help to get this piece of software
// working:
//
// The basics (integrating an entry into the Explorer context menu):
//   http://blog.marcocantu.com/blog/2016-03-writing-windows-shell-extension.html
//   http://www.andreanolanusse.com/en/shell-extension-for-windows-32-bit-and-64-bit-with-delphi-xe2/
//
// The really hard stuff (drawing the context menu item's icon, thank you MS!):
//   https://www.nanoant.com/programming/themed-menus-icons-a-complete-vista-xp-solution
//
//
// This program is free software; you can redistribute it and/or modify
// it under the terms of the Mozilla Public License Version 2.0.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
//
{ **************************************************************************** }

unit CudaTextContextMenuHandlerImpl;

{$MODE Delphi}
{$WARN SYMBOL_PLATFORM OFF}


interface

uses
  Windows, UxTheme, Messages, ActiveX, ShellAPI, ShlObj, ComObj, Registry,
  Types, Classes,

  CudaTextContextMenuHandler_TLB;


const
{ ============================================================================ }
{ Missing constants                                                            }
{ ============================================================================ }

  CMF_ITEMMENU          = $00000080;
  CMF_DISABLEDVERBS     = $00000200;
  CMF_ASYNCVERBSTATE    = $00000400;
  CMF_OPTIMIZEFORINVOKE = $00000800;
  CMF_SYNCCASCADEMENU   = $00001000;
  CMF_DONOTPICKDEFAULT  = $00002000;


type
{ ============================================================================ }
{ Missing types                                                                }
{ ============================================================================ }

  ARGB = DWORD;

  // we have to do arithmetics with this type of pointer
  {$POINTERMATH ON}
  PARGB = ^ARGB;
  {$POINTERMATH OFF}

  BP_BUFFERFORMAT = BPBF_COMPATIBLEBITMAP..BPBF_TOPDOWNMONODIB;

  // for dynamically importing UxTheme functions if OS is Vista or newer
  TGetBufferedPaintBits = function(hBufferedPaint: HPAINTBUFFER; out ppbBuffer: PRGBQUAD; out pcxRow: integer): HRESULT; stdcall;
  TBeginBufferedPaint   = function(hdcDest: HDC; const prcTarget: TRect; dwFormat: BP_BUFFERFORMAT; pPaintParams: PBP_PAINTPARAMS; out phdc: HDC): HPAINTBUFFER; stdcall;
  TEndBufferedPaint     = function(hBufferedPaint: HPAINTBUFFER; fUpdateTarget: BOOL): HRESULT; stdcall;


{ ============================================================================ }
{ TCudaTextContextMenuHandler                                                  }
{ ============================================================================ }

  TCudaTextContextMenuHandler = class(TAutoObject, ICudaTextContextMenuHandler, IShellExtInit, IContextMenu, IContextMenu2, IContextMenu3)
  private
    FStatusbarMessage: string;
    FCudaTextPath:     string;
    FFileNames:        TStringList;

    pfnGetBufferedPaintBits: TGetBufferedPaintBits;
    pfnBeginBufferedPaint:   TBeginBufferedPaint;
    pfnEndBufferedPaint:     TEndBufferedPaint;

    procedure OpenWithCudaText(Handle: HWND; SW_Mode: longint; FileNames: TStrings);
    function  Execute(Handle: HWND; SW_Mode: longint; const AFile: string; var Params: string): boolean;

    procedure InitUXThemeFuncs;
    function  IconToBitmapPARGB32(const FilePath: string): HBITMAP;
    function  Create32BitHBITMAP(hdcDest: HDC; const sizBmp: TSize; out pvBits: Pointer; out hBmp: HBITMAP): HRESULT;
    function  ConvertBufferToPARGB32(hPaintBuf: HPAINTBUFFER; hdcDest: HDC; hSrcIcon: HICON; const sizIcon: TSize): HRESULT;
    function  ConvertToPARGB32(hdcDest: HDC; pBufARGB: PARGB; hbmp: HBITMAP; const sizImage: TSize; cxRow: integer): HRESULT;
    function  HasAlpha(pBufARGB: PARGB; const sizImage: TSize; cxRow: integer): boolean;
    procedure InitBitmapInfo(out bmi: BITMAPINFO; cx, cy: LONG; bpp: WORD);

    function  GetCudaTextIcon(const CudaTextPath: string; out IconHandle: HICON): boolean;
    function  GetCudaTextDir: string;

  protected
    { ------------------------------------------------------------------------ }
    { IShellExtInit Methods                                                    }
    { ------------------------------------------------------------------------ }
    // Initialize the context menu if a file was selected
    function IShellExtInit.Initialize = ShellExtInitialize;
    function ShellExtInitialize(pidlFolder: PItemIDList; lpdobj: IDataObject;
                                hKeyProgID: HKEY): HRESULT; stdcall;

    { ------------------------------------------------------------------------ }
    { IContextMenu Methods                                                     }
    { ------------------------------------------------------------------------ }
    // Initializes the context menu and it decides which items appear in it,
    // based on the flags you pass
    function QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst, idCmdLast,
      uFlags: UINT): HRESULT; stdcall;

    // Set help string on the Explorer status bar when the menu item is selected
    function GetCommandString(idCmd: UINT_PTR; uFlags: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HRESULT; stdcall;

    // Execute the command, which will be opening a file in CudaText
    function InvokeCommand(var lpici: TCMInvokeCommandInfo): HRESULT; stdcall;

    { ------------------------------------------------------------------------ }
    { IContextMenu2 Methods                                                    }
    { ------------------------------------------------------------------------ }
    // Handle some window events to draw the menu icon's bitmap
    function HandleMenuMsg(uMsg: UINT; wParam: WPARAM; lParam: LPARAM): HRESULT; stdcall;

    { ------------------------------------------------------------------------ }
    { IContextMenu3 Methods                                                    }
    { ------------------------------------------------------------------------ }
    // Handle some window events to draw the menu icon's bitmap
    function HandleMenuMsg2(uMsg: UINT; wParam: WPARAM; lParam: LPARAM; pResult: PLRESULT): HRESULT; stdcall;

  end;


{ ============================================================================ }
{ TCudaTextContextMenuHandlerObjectFactory, the new class factory              }
{ ============================================================================ }

  TCudaTextContextMenuHandlerObjectFactory = class(TAutoObjectFactory)
  public
    procedure UpdateRegistry(DoRegister: Boolean); override;

  end;



implementation

uses
  SysUtils, Graphics, ComServ;


const
  // Context menu entry text
  MENU_ITEM_CAPTON = 'Open with CudaText';

  // Text to show in Explorer window status bar
  // when hovering over context menu entry
  STATUS_BAR_MSG_ITEM      = 'Open selected item(s) with CudaText';
  STATUS_BAR_MSG_CONTAINER = 'Open this folder with CudaText';


{ ============================================================================ }
{ TCudaTextContextMenuHandler                                                  }
{ ============================================================================ }

// Called by Windows Explorer to inform the context menu handler which items
// are currently selected
function TCudaTextContextMenuHandler.ShellExtInitialize(pidlFolder: PItemIDList;
  lpdobj: IDataObject; hKeyProgID: HKEY): HRESULT;
var
  DataFormat: TFormatEtc;
  StrgMedium: TStgMedium;
  CntFiles:   UINT;
  Idx:        UINT;
  Buffer:     array[0..MAX_PATH] of WideChar;
  FileName:   WideString;

begin
  Result := E_FAIL;

  FreeAndNil(FFileNames);
  FStatusbarMessage := '';

  // check if an object was defined, i.e. user right-clicked a file or folder,
  // a shortcut to a file or folder or a drive icon
  if Assigned(lpdobj) then
  begin
    // prepare to get information about the object
    DataFormat.cfFormat := CF_HDROP;
    DataFormat.ptd      := nil;
    DataFormat.dwAspect := DVASPECT_CONTENT;
    DataFormat.lindex   := -1;
    DataFormat.tymed    := TYMED_HGLOBAL;

    // read data into storage medium backed by a memory buffer
    if lpdobj.GetData(DataFormat, StrgMedium) <> S_OK then exit;

    try
      // get count of data sets in storage medium
      CntFiles := DragQueryFileW(StrgMedium.hGlobal, $FFFFFFFF, nil, 0);
      if CntFiles < 1 then exit;

      FCudaTextPath := GetCudaTextDir();
      FFileNames    := TStringList.Create;

      // read file paths from storage medium and store them in list
      for Idx := 0 to Pred(CntFiles) do
      begin
        FillChar(Buffer, SizeOf(Buffer), 0);
        DragQueryFileW(StrgMedium.hGlobal, Idx, @Buffer, MAX_PATH);
        FileName := PWideChar(Buffer);
        FFileNames.Add(UTF8Encode(FileName));
      end;

      // set statusbar message text
      FStatusbarMessage := STATUS_BAR_MSG_ITEM;

      Result := S_OK;

    finally
      ReleaseStgMedium(StrgMedium);
    end;
  end

  // check if a containing folder was defined, i.e. user right-clicked
  // Explorer window background
  else if Assigned(pidlFolder) then
  begin
    FillChar(Buffer, SizeOf(Buffer), 0);

    // convert provided shell item identifier list to folder path
    // and store it in list
    if SHGetPathFromIDListW(pidlFolder, @Buffer) then
    begin
      FCudaTextPath := GetCudaTextDir();
      FFileNames    := TStringList.Create;

      FileName := PWideChar(Buffer);
      FFileNames.Add(UTF8Encode(FileName));

      // set statusbar message text
      FStatusbarMessage := STATUS_BAR_MSG_CONTAINER;

      Result := S_OK;
    end;
  end;
end;


// Called by Windows Explorer to determine if and how to display the context
// menu entry for the currently selected items
function TCudaTextContextMenuHandler.QueryContextMenu(Menu: HMENU;
  indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HRESULT;
var
  ContextMenuItem: TMenuItemInfoW;
  MenuCaption:     WideString;

begin
  // only adding one menu context menu item, so generate the result code accordingly
  Result := MakeResult(SEVERITY_SUCCESS, FACILITY_NULL, 1);

  // specify what the menu says, depending on where it was spawned
  if (uFlags = CMF_NORMAL) then                            // the default case
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)

  else if (uFlags and CMF_ITEMMENU) = CMF_ITEMMENU  then   // from file/folder
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)

  else if (uFlags and CMF_VERBSONLY) = CMF_VERBSONLY then  // from file/folder shortcut
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)

  else if (uFlags and CMF_CANRENAME ) = CMF_CANRENAME then // from file/folder in XP
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)

  else if (uFlags and CMF_EXPLORE) = CMF_EXPLORE then      // from folder in explorer tree
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)

  else if (uFlags and CMF_NODEFAULT) = CMF_NODEFAULT then  // from Explorer window background
    MenuCaption := UTF8Decode(MENU_ITEM_CAPTON)            // when navigation pane is deactivated

  else
    // fail for any other value
    Result := E_NOTIMPL;

  if Succeeded(Result) then
  begin
    FillChar(ContextMenuItem, SizeOf(ContextMenuItem), 0);

    ContextMenuItem.cbSize     := SizeOf(ContextMenuItem);
    ContextMenuItem.fMask      := MIIM_FTYPE or MIIM_STRING or MIIM_ID or MIIM_BITMAP;
    ContextMenuItem.fType      := MFT_STRING;
    ContextMenuItem.wID        := idCmdFirst;
    ContextMenuItem.dwTypeData := PWideChar(MenuCaption);
    ContextMenuItem.cch        := Length(MenuCaption);

    // use different approaches for drawing the menu icon bitmap depending on
    // the OS version. as the XP approach causes removing of menu themes in
    // Vista and later, there we use UxTheme for creating the icon bitmap.
    if Win32MajorVersion < 6 then
      ContextMenuItem.hbmpItem := HBITMAP(HBMMENU_CALLBACK)
    else
    begin
      InitUXThemeFuncs();
      ContextMenuItem.hbmpItem := IconToBitmapPARGB32(FCudaTextPath);
    end;

    InsertMenuItemW(Menu, indexMenu, True, @ContextMenuItem);
  end;
end;


// Called by Windows Explorer to determine the status bar text related to the
// selected items to be shown in Explorer window
function TCudaTextContextMenuHandler.GetCommandString(idCmd: UINT_PTR; uFlags: UINT;
  pwReserved: PUINT; pszName: LPSTR; cchMax: UINT): HRESULT;
begin
  Result := E_INVALIDARG;

  // process only our one-and-only menu entry
  if idCmd <> 0 then exit;

  // set help string on Explorer status bar when menu item is mouse-hovered
  case uFlags of
    // Explorer requests unicode string
    GCS_HELPTEXTW:
    begin
      StrCopy(PWideChar(pszName), PWideChar(UTF8Decode(FStatusbarMessage)));
      Result := S_OK;
    end;

    // Explorer requests ANSI string
    GCS_HELPTEXTA:
    begin
      WideCharToMultiByte(CP_ACP, 0,
                          PWideChar(UTF8Decode(FStatusbarMessage)), -1,
                          pszName, cchMax * SizeOf(AnsiChar),
                          nil, nil
                         );
      Result := S_OK;
    end;
  end;
end;


// The following two methods are called by Windows Explorer to send Windows
// messages regarding the context menu entry to the context menu handler
function TCudaTextContextMenuHandler.HandleMenuMsg(uMsg: UINT; wParam: WPARAM; lParam: LPARAM): HRESULT;
var
  NotUsed: LRESULT;

begin
  Result := HandleMenuMsg2(uMsg, wParam, lParam, @NotUsed);
end;


function TCudaTextContextMenuHandler.HandleMenuMsg2(uMsg: UINT; wParam: WPARAM; lParam: LPARAM; pResult: PLRESULT): HRESULT;
var
  lpmis:      PMEASUREITEMSTRUCT;
  lpdis:      PDRAWITEMSTRUCT;
  IconHandle: HICON;

begin
  pResult^ := LRESULT(false);
  Result   := S_FALSE;

  // in XP it is suitable to handle the events below to draw the menu icon's bitmap.
  // the OS takes the burden to take care of selection state, transparentness,
  // correct background color and all that weird stuff
  case uMsg of
    WM_MEASUREITEM:
    begin
      lpmis := PMEASUREITEMSTRUCT(lParam);

      if lpmis <> nil then
      begin
        // we don't want the icon to become down-scaled
        lpmis.itemWidth  := 16;
        lpmis.itemHeight := 16;

        Result   := S_OK;
        pResult^ := LRESULT(true);
      end;
    end;

    WM_DRAWITEM:
    begin
      lpdis := PDRAWITEMSTRUCT(lParam);

      if (lpdis = nil) or (lpdis.CtlType <> ODT_MENU) then
      begin
        Result := S_OK; // not for a menu
        exit;
      end;

      GetCudaTextIcon(FCudaTextPath, IconHandle);

      if IconHandle = 0 then
      begin
        Result := S_OK;
        exit;
      end;

      // since the default size of a context menu icon on XP is only 12 pixels
      // we work with a fixed icon size of 16 pixels here to avoid down-scaling.
      // the formula for calculating icon's position is only correct for XP!
      DrawIconEx(lpdis.hDC,
                 lpdis.rcItem.left - 16,
                 lpdis.rcItem.top + (lpdis.rcItem.bottom - lpdis.rcItem.top - 16) div 2,
                 IconHandle,
                 16, 16, 0, 0,
                 DI_NORMAL);

      DestroyIcon(IconHandle);

      Result   := S_OK;
      pResult^ := LRESULT(true);
    end;
  end;
end;


// Called by Windows Explorer when the context menu entry has been clicked
function TCudaTextContextMenuHandler.InvokeCommand(var lpici: TCMInvokeCommandInfo): HRESULT;
begin
  Result := E_FAIL;

  // exit if lpici.lpVerb is a pointer
  if not IS_INTRESOURCE(PChar(lpici.lpVerb)) then exit;

  // retrieve the clicked menu item by its offset to the first added menu item
  case LoWord(NativeInt(lpici.lpVerb)) of
    0:
    begin
      if not Assigned(FFileNames) then exit;

      try
        OpenWithCudaText(lpici.HWND, lpici.nShow, FFileNames);

        FreeAndNil(FFileNames);
        FStatusbarMessage := '';

      except
        on E: Exception do
          MessageBoxW(lpici.hwnd, PWideChar(UTF8Decode(E.Message)), nil, MB_ICONERROR);
      end;

      Result := S_OK;
    end;
  end
end;


// Internal method to open all items selected in Explorer window with CudaText.
// To improve performance, CudaText is started with packets of 125 file paths
// at maximum in order to comply with the command line length limit of 32768
// characters.
procedure TCudaTextContextMenuHandler.OpenWithCudaText(Handle: HWND; SW_Mode: longint; FileNames: TStrings);
var
  Idx:          integer;
  FileNameList: string;

begin
  Idx          := 0;
  FileNameList := '';

  // since the lpCommandLine parameter of CreateProcess has a length limit of
  // 32768 characters, we call the Execute function with packets of 125 files
  // maximum (260 chrs for CudaText path + 125 file paths * 260 chrs = 32760 chrs)
  while Idx < FileNames.Count do
  begin
    FileNameList := Format('%s "%s"', [FileNameList, FileNames[Idx]]);
    Inc(Idx);

    if (Idx mod 126 = 0) or (Idx = FileNames.Count) then
    begin
      Execute(Handle, SW_Mode, FCudaTextPath, FileNameList);
      FileNameList := '';
    end;
  end;
end;


// Internal method for wrapping Win32 CreateProcess API
function TCudaTextContextMenuHandler.Execute(Handle: HWND; SW_Mode: longint; const AFile: string; var Params: string): boolean;
var
  si: Windows.TStartupInfoW;
  pi: TProcessInformation;

begin
  Result := false;

  FillChar(si, SizeOf(si), 0);
  FillChar(pi, SizeOf(pi), 0);

  si.cb          := SizeOf(si);
  si.dwFlags     := STARTF_USESHOWWINDOW;
  si.wShowWindow := WORD(SW_Mode);

  if not CreateProcessW(nil,
                        PWideChar('"' + UTF8Decode(AFile) + '" ' + UTF8Decode(Params)),
                        nil, nil,
                        false,
                        CREATE_DEFAULT_ERROR_MODE,
                        nil,
                        PWideChar(UTF8Decode(ExtractFileDir(AFile))),
                        si, pi) then
    MessageBoxW(Handle, PWideChar(UTF8Decode(SysErrorMessage(GetLastError()))), nil, MB_ICONERROR)
  else
    Result := true;
end;


// Internal method to dynamically load certain UxTheme functions
procedure TCudaTextContextMenuHandler.InitUXThemeFuncs;
var
  hUxTheme: HMODULE;

begin
  hUxTheme := GetModuleHandle('UXTHEME.DLL');

  @pfnGetBufferedPaintBits := GetProcAddress(hUxTheme, 'GetBufferedPaintBits');
  @pfnBeginBufferedPaint   := GetProcAddress(hUxTheme, 'BeginBufferedPaint'  );
  @pfnEndBufferedPaint     := GetProcAddress(hUxTheme, 'EndBufferedPaint'    );
end;


// Internal method to convert the icon of an executable file to 32 bpp ARGB
// (i.e. RGB with alpha channel) bitmap format, the format required for Explorer
// context menu icons in Windows Vista and later
function TCudaTextContextMenuHandler.IconToBitmapPARGB32(const FilePath: string): HBITMAP;
var
  hr:          HRESULT;
  hBmp:        HBITMAP;
  sizIcon:     TSize;
  hdIcon:      HICON;
  rcIcon:      TRect;
  hdcDest:     HDC;
  hbmpOld:     HBITMAP;
  pvBits:      pointer;
  bfAlpha:     BLENDFUNCTION;
  paintParams: BP_PAINTPARAMS;
  hdcBuffer:   HDC;
  hPaintBuf:   HPAINTBUFFER;

begin
  Result := 0;
  hBmp   := 0;
  pvBits := nil;
  hr     := E_OUTOFMEMORY;

  if not GetCudaTextIcon(FilePath, hdIcon) then exit;

  sizIcon.cx := GetSystemMetrics(SM_CXSMICON);
  sizIcon.cy := GetSystemMetrics(SM_CYSMICON);

  SetRect(rcIcon, 0, 0, sizIcon.cx, sizIcon.cy);

  hdcDest := CreateCompatibleDC(0);

  if hdcDest <> 0 then
  begin
    hr := Create32BitHBITMAP(hdcDest, sizIcon, pvBits, hbmp);

    if SUCCEEDED(hr) then
    begin
      hr := E_FAIL;

      hbmpOld := SelectObject(hdcDest, hbmp);

      if hbmpOld <> 0 then
      begin
        bfAlpha.BlendOp             := AC_SRC_OVER;
        bfAlpha.BlendFlags          := 0;
        bfAlpha.SourceConstantAlpha := 255;
        bfAlpha.AlphaFormat         := AC_SRC_ALPHA;

        FillChar(paintParams, SizeOf(paintParams), 0);

        paintParams.cbSize         := SizeOf(paintParams);
        paintParams.dwFlags        := BPPF_ERASE;
        paintParams.pBlendFunction := @bfAlpha;

        hPaintBuf := pfnBeginBufferedPaint(hdcDest, rcIcon, BPBF_DIB, @paintParams, hdcBuffer);

        if hPaintBuf <> 0 then
        begin
          if DrawIconEx(hdcBuffer, 0, 0, hdIcon, sizIcon.cx, sizIcon.cy, 0, 0, DI_NORMAL) then
          begin
            // if icon did not have an alpha channel, we need to convert buffer to PARGB.
            hr := ConvertBufferToPARGB32(hPaintBuf, hdcDest, hdIcon, sizIcon);
          end;

          // this will write the buffer contents to the destination bitmap.
          pfnEndBufferedPaint(hPaintBuf, true);
        end;

        SelectObject(hdcDest, hbmpOld);
      end;
    end;

    DeleteDC(hdcDest);
  end;

  DestroyIcon(hdIcon);

  if not SUCCEEDED(hr) then
    DeleteObject(hBmp)
  else
    Result := hBmp;
end;


// Internal method to create a device independent bitmap (DIB) where the context
// menu entry's icon can be drawn to
function TCudaTextContextMenuHandler.Create32BitHBITMAP(hdcDest: HDC; const sizBmp: TSize; out pvBits: Pointer; out hBmp: HBITMAP): HRESULT;
var
  bmi:     BITMAPINFO;
  hdcUsed: HDC;

begin
  hBmp := 0;

  InitBitmapInfo(bmi, sizBmp.cx, sizBmp.cy, 32);

  if hdcDest <> 0 then
    hdcUsed := hdcDest
  else
    hdcUsed := GetDC(0);

  if hdcUsed <> 0 then
  begin
    hBmp := CreateDIBSection(hdcUsed, bmi, DIB_RGB_COLORS, pvBits, 0, 0);

    if hdcDest <> hdcUsed then
      ReleaseDC(0, hdcUsed);
  end;

  if hBmp = 0 then
    Result := E_OUTOFMEMORY
  else
    Result := S_OK;
end;


// Check if the bitmap in a given UxTheme paint buffer has an alpha channel.
// If not convert it to 32 bpp ARGB format.
function TCudaTextContextMenuHandler.ConvertBufferToPARGB32(hPaintBuf: HPAINTBUFFER; hdcDest: HDC; hSrcIcon: HICON; const sizIcon: TSize): HRESULT;
var
  pBufPix:  PRGBQUAD;
  cxRow:    integer;
  hr:       HRESULT;
  pBufARGB: PARGB;
  info:     ICONINFO;

begin
  hr := pfnGetBufferedPaintBits(hPaintBuf, pBufPix, cxRow);

  if SUCCEEDED(hr) then
  begin
    pBufARGB := PARGB(pBufPix);

    if not HasAlpha(pBufARGB, sizIcon, cxRow) then
    begin
      if GetIconInfo(hSrcIcon, info) then
      begin
        if info.hbmMask <> 0 then
          hr := ConvertToPARGB32(hdcDest, pBufARGB, info.hbmMask, sizIcon, cxRow);

        DeleteObject(info.hbmColor);
        DeleteObject(info.hbmMask);
      end;
    end;
  end;

  Result := hr;
end;


// Internal method to convert the pixels of a 32 bpp bitmap to ARGB format
function TCudaTextContextMenuHandler.ConvertToPARGB32(hdcDest: HDC; pBufARGB: PARGB; hbmp: HBITMAP; const sizImage: TSize; cxRow: integer): HRESULT;
var
  bmi:       BITMAPINFO;
  hr:        HRESULT;
  hHeap:     THandle;
  pvBits:    pointer;
  cxDelta:   ULONG;
  pargbMask: PARGB;
  x:         ULONG;
  y:         ULONG;

begin
  InitBitmapInfo(bmi, sizImage.cx, sizImage.cy, 32);

  hr     := E_OUTOFMEMORY;
  hHeap  := GetProcessHeap();

  pvBits := HeapAlloc(hHeap, 0, bmi.bmiHeader.biWidth * 4 * bmi.bmiHeader.biHeight);

  if pvBits <> nil then
  begin
    hr := E_UNEXPECTED;

    if GetDIBits(hdcDest, hbmp, 0, bmi.bmiHeader.biHeight, pvBits, bmi, DIB_RGB_COLORS) = bmi.bmiHeader.biHeight then
    begin
      cxDelta   := cxRow - bmi.bmiHeader.biWidth;
      pargbMask := PARGB(pvBits);

      for y := Pred(bmi.bmiHeader.biHeight) downto 0 do
      begin
        for x := Pred(bmi.bmiHeader.biWidth) downto 0 do
        begin
          if pargbMask^ <> 0 then
          begin
            // transparent pixel
            pBufARGB^ := 0;
            pBufARGB  := pBufARGB + SizeOf(ARGB);
          end
          else
          begin
            // opaque pixel
            pBufARGB^ := pBufARGB^ or $FF000000;
            pBufARGB  := pBufARGB + SizeOf(ARGB);
          end;

          pargbMask := pargbMask + SizeOf(ARGB);
        end;

        pBufARGB := pBufARGB + cxDelta;
      end;

      hr := S_OK;
    end;

    HeapFree(hHeap, 0, pvBits);
  end;

  Result := hr;
end;


// Internal method to check if the pixels of a 32 bpp bitmap have an alpha channel
function TCudaTextContextMenuHandler.HasAlpha(pBufARGB: PARGB; const sizImage: TSize; cxRow: integer): boolean;
var
  cxDelta: ULONG;
  x:       ULONG;
  y:       ULONG;

begin
  cxDelta := cxRow - sizImage.cx;

  for y := Pred(sizImage.cy) downto 0 do
  begin
    for x := Pred(sizImage.cx) downto 0 do
    begin
      if (pBufARGB^ and $FF000000) <> 0 then
      begin
        Result := true;
        exit;
      end;

      pBufARGB := pBufARGB + SizeOf(ARGB);
    end;

    pBufARGB := pBufARGB + cxDelta;
  end;

  Result := false;
end;


// Internal method to initialize a BITMAPINFO structure
procedure TCudaTextContextMenuHandler.InitBitmapInfo(out bmi: BITMAPINFO; cx, cy: LONG; bpp: WORD);
begin
  FillChar(bmi, SizeOf(bmi), 0);

  bmi.bmiHeader.biSize        := SizeOf(BITMAPINFOHEADER);
  bmi.bmiHeader.biPlanes      := 1;
  bmi.bmiHeader.biCompression := BI_RGB;

  bmi.bmiHeader.biWidth    := cx;
  bmi.bmiHeader.biHeight   := cy;
  bmi.bmiHeader.biBitCount := bpp;
end;


// Internal method to retrieve the small version of an executable's icon
function TCudaTextContextMenuHandler.GetCudaTextIcon(const CudaTextPath: string; out IconHandle: HICON): boolean;
var
  ShellFileInfo: TSHFileInfoW;

begin
  IconHandle := 0;
  Result     := false;

  FillChar(ShellFileInfo, SizeOf(ShellFileInfo), 0);

  if SHGetFileInfoW(PWideChar(UTF8Decode(CudaTextPath)), 0, ShellFileInfo, SizeOf(ShellFileInfo),
                    SHGFI_ICON or SHGFI_SHELLICONSIZE or SHGFI_SMALLICON) = 0 then
    exit;

  IconHandle := ShellFileInfo.hIcon;
  Result     := true;
end;


// Internal method to retrive the path to the directory of the context menu
// handler's DLL file. This should be the same as CudaText's directory.
function TCudaTextContextMenuHandler.GetCudaTextDir: string;
var
  szBuf:    array of WideChar;
  szBuf2:   WideString;
  dwBufLen: DWORD;
  dwRet:    DWORD;

begin
  Result := '';

  dwBufLen := MAX_PATH;

  repeat
    SetLength(szBuf, dwBufLen);
    dwRet := GetModuleFileNameW(HInstance, PWideChar(szBuf), dwBufLen);

    // if dwRet is 0 there was an error
    //   => leave loop
    // if dwRet is less than dwBufLen the buffer size was sufficient
    //   => leave loop
    // if dwRet is equal to dwBufLen the buffer size was too small
    //   => loop and retry with double sized buffer
    // dwRet greater than dwBufLen is a non-existing case
    if dwRet < dwBufLen then break;
    dwBufLen := dwBufLen * 2;
  until false;

  if dwRet > 0 then
  begin
    SetLength(szBuf, dwRet + 1);
    szBuf2 := PWideChar(szBuf);
    Result := ExtractFileDir(UTF8Encode(szBuf2)) + PathDelim  + 'cudatext.exe';
  end;
end;



{ ============================================================================ }
{ TCudaTextContextMenuHandlerObjectFactory                                     }
{ ============================================================================ }

// Overridden method of the context menu handler's class factory.
// On installing the context menu handler, create the required registry keys
// to show the context menu handler's menu entries in Explorer.
// On uninstalling the context menu handler, delete these registry keys.
procedure TCudaTextContextMenuHandlerObjectFactory.UpdateRegistry(DoRegister: Boolean);
const
  FileContextKey  = '*\shellex\ContextMenuHandlers\%s';
  DirContextKey   = 'Directory\shellex\ContextMenuHandlers\%s';
  DirBgContextKey = 'Directory\Background\shellex\ContextMenuHandlers\%s';
  DriveContextKey = 'Drive\shellex\ContextMenuHandlers\%s';

begin
  // perform normal registration
  inherited UpdateRegistry(DoRegister);

  // if this server is being registered, register the required key/values
  // to expose it to Explorer
  if DoRegister then
  begin
    CreateRegKey(Format(FileContextKey,  [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT);
    CreateRegKey(Format(DirContextKey,   [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT);
    CreateRegKey(Format(DirBgContextKey, [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT);
    CreateRegKey(Format(DriveContextKey, [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT);
  end
  else
  begin
    DeleteRegKey(Format(FileContextKey,  [ClassName]));
    DeleteRegKey(Format(DirContextKey,   [ClassName]));
    DeleteRegKey(Format(DirBgContextKey, [ClassName]));
    DeleteRegKey(Format(DriveContextKey, [ClassName]));
  end;
end;



initialization
  // on DLL initialization, create an instance of the class factory required to
  // instantiate the context menu handler's COM server
  TCudaTextContextMenuHandlerObjectFactory.Create(ComServer, TCudaTextContextMenuHandler,
                                                  CLASS_CudaTextContextMenuHandler,
                                                  ciMultiInstance, tmApartment);

end.

