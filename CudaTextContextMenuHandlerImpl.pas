{ **************************************************************************** }
//
// CudaText context menu handler for Windows Explorer
//
// Author: Andreas Heim, 2019
//
//
// The following web sites provided great help to get this piece of software
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
  Windows, ActiveX, ShellAPI, ShlObj, ComObj, Registry, Classes,

  CudaTextContextMenuHandler_TLB;


type
{ ============================================================================ }
{ TCudaTextContextMenuHandler                                                  }
{ ============================================================================ }

  TCudaTextContextMenuHandler = class(TAutoObject, ICudaTextContextMenuHandler, IShellExtInit, IContextMenu)
  private
    FMenuItemIndex: UINT;
    FCudaTextPath:  string;
    FFileNames:     TStringList;

    function  Execute(Handle: HWND; const AFile: string; var Params: string): boolean;
    procedure OpenWithCudaText(Handle: HWND; FileNames: TStrings);
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

    // Execute the command, which will be opening a file in CudaText
    function InvokeCommand(var lpici: TCMInvokeCommandInfo): HRESULT; stdcall;

    // Set help string on the Explorer status bar when the menu item is selected
    function GetCommandString(idCmd: UINT_PTR; uFlags: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HRESULT; stdcall;

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


{ ============================================================================ }
{ TCudaTextContextMenuHandler                                                  }
{ ============================================================================ }

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

  // Check if an object was defined
  if lpdobj = nil then exit;

  // Prepare to get information about the object
  DataFormat.cfFormat := CF_HDROP;
  DataFormat.ptd      := nil;
  DataFormat.dwAspect := DVASPECT_CONTENT;
  DataFormat.lindex   := -1;
  DataFormat.tymed    := TYMED_HGLOBAL;

  if lpdobj.GetData(DataFormat, StrgMedium) <> S_OK then exit;

  try
    CntFiles := DragQueryFileW(StrgMedium.hGlobal, $FFFFFFFF, nil, 0);
    if CntFiles < 1 then exit;

    FCudaTextPath := GetCudaTextDir();
    FFileNames    := TStringList.Create;

    for Idx := 0 to Pred(CntFiles) do
    begin
      FillChar(Buffer, SizeOf(Buffer), 0);
      DragQueryFileW(StrgMedium.hGlobal, Idx, @Buffer, MAX_PATH);
      FileName := PWideChar(Buffer);
      FFileNames.Add(UTF8Encode(FileName));
    end;

    Result := NOERROR;

  finally
    ReleaseStgMedium(StrgMedium);
  end;
end;


function TCudaTextContextMenuHandler.QueryContextMenu(Menu: HMENU;
  indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HRESULT;
const
  CMF_ITEMMENU     = $00000080;
  MENU_ITEM_CAPTON = 'Open with CudaText';

var
  ContextMenuItem: TMenuItemInfoW;
  MenuCaption:     string;

begin
  // only adding one menu context menu item, so generate the result code accordingly
  Result := MakeResult(SEVERITY_SUCCESS, 0, 1);

  // store the context menu item index
  FMenuItemIndex := indexMenu;

  // specify what the menu says, depending on where it was spawned
  if (uFlags = CMF_NORMAL) then // from the desktop
    MenuCaption := MENU_ITEM_CAPTON

  else if (uFlags and CMF_ITEMMENU) = CMF_ITEMMENU  then  // from desktop
    MenuCaption := MENU_ITEM_CAPTON

  else if (uFlags and CMF_VERBSONLY) = CMF_VERBSONLY then // from a shortcut
    MenuCaption := MENU_ITEM_CAPTON

  else if (uFlags and CMF_EXPLORE) = CMF_EXPLORE then // from explorer
    MenuCaption := MENU_ITEM_CAPTON

  else if (uFlags and CMF_CANRENAME ) = CMF_CANRENAME  then // important for XP
    MenuCaption := MENU_ITEM_CAPTON

  else
    // fail for any other value
    Result := E_FAIL;

  if Result <> E_FAIL then
  begin
    FillChar(ContextMenuItem, SizeOf(ContextMenuItem), 0);

    ContextMenuItem.cbSize     := SizeOf(ContextMenuItem);
    ContextMenuItem.fMask      := MIIM_FTYPE or MIIM_STRING or MIIM_ID;
    ContextMenuItem.fType      := MFT_STRING;
    ContextMenuItem.wID        := idCmdFirst;
    ContextMenuItem.dwTypeData := PWideChar(UTF8Decode(MenuCaption));
    ContextMenuItem.cch        := Length(MenuCaption);

    InsertMenuItemW(Menu, FMenuItemIndex, True, @ContextMenuItem);
  end;
end;


function TCudaTextContextMenuHandler.GetCommandString(idCmd: UINT_PTR; uFlags: UINT;
  pwReserved: PUINT; pszName: LPSTR; cchMax: UINT): HRESULT;
const
  StatusbarMessage = 'Open selected file(s) with CudaText';

begin
  Result := E_INVALIDARG;

  // Set help string on the Explorer status bar when the menu item is selected
  if (idCmd = 0) and (uFlags = GCS_HELPTEXT) then
  begin
    WideCharToMultiByte(CP_ACP, 0,
                        PWideChar(UTF8Decode(StatusbarMessage)), Length(StatusbarMessage),
                        PChar(pszName), cchMax * SizeOf(AnsiChar),
                        nil, nil
                       );

    Result := NOERROR;
  end;
end;


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
        OpenWithCudaText(lpici.HWND, FFileNames);
        FreeAndNil(FFileNames);

      except
        on E: Exception do
          MessageBoxW(lpici.hwnd, PWideChar(UTF8Decode(E.Message)), nil, MB_ICONERROR);
      end;

      Result := NOERROR;
    end;
  end
end;


procedure TCudaTextContextMenuHandler.OpenWithCudaText(Handle: HWND; FileNames: TStrings);
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
      Execute(Handle, FCudaTextPath, FileNameList);
      FileNameList := '';
    end;
  end;
end;


function TCudaTextContextMenuHandler.Execute(Handle: HWND; const AFile: string; var Params: string): boolean;
var
  si: Windows.TStartupInfoW;
  pi: TProcessInformation;

begin
  Result := false;

  FillChar(si, SizeOf(si), 0);
  FillChar(pi, SizeOf(pi), 0);

  si.cb          := SizeOf(si);
  si.dwFlags     := STARTF_USESHOWWINDOW;
  si.wShowWindow := SW_SHOWNORMAL;

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

    // If dwRet is 0 there was an error
    //   => leave loop
    // if dwRet is less than dwBufLen the buffer size was sufficient
    //   => leave loop
    // If dwRet is equal to dwBufLen the buffer size was too small
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

procedure TCudaTextContextMenuHandlerObjectFactory.UpdateRegistry(DoRegister: Boolean);
const
  ContextKey = '*\shellex\ContextMenuHandlers\%s';

begin
  // perform normal registration
  inherited UpdateRegistry(DoRegister);

  // if this server is being registered, register the required key/values
  // to expose it to Explorer
  if DoRegister then
    CreateRegKey(Format(ContextKey, [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT)
  else
    DeleteRegKey(Format(ContextKey, [ClassName]));
end;



initialization
  TCudaTextContextMenuHandlerObjectFactory.Create(ComServer, TCudaTextContextMenuHandler,
                                                  CLASS_CudaTextContextMenuHandler,
                                                  ciMultiInstance, tmApartment);

end.

