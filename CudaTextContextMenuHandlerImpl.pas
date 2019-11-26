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
                                hKeyProgID: HKEY): HResult; stdcall;

    { ------------------------------------------------------------------------ }
    { IContextMenu Methods                                                     }
    { ------------------------------------------------------------------------ }
    // Initializes the context menu and it decides which items appear in it,
    // based on the flags you pass
    function QueryContextMenu(Menu: HMENU; indexMenu, idCmdFirst, idCmdLast,
      uFlags: UINT): HResult; stdcall;

    // Execute the command, which will be opening a file in CudaText
    function InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult; stdcall;

    // Set help string on the Explorer status bar when the menu item is selected
    function GetCommandString(idCmd: UINT_PTR; uFlags: UINT; pwReserved: PUINT;
      pszName: LPSTR; cchMax: UINT): HResult; stdcall;

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
  SysUtils, Math, Graphics, ComServ;


{ ============================================================================ }
{ TCudaTextContextMenuHandler                                                  }
{ ============================================================================ }
function TCudaTextContextMenuHandler.Execute(Handle: HWND; const AFile: string; var Params: string): boolean;
var
  si: TStartupInfo;
  pi: TProcessInformation;

begin
  Result := false;

  ZeroMemory(@si, SizeOf(si));
  ZeroMemory(@pi, SizeOf(pi));

  si.cb          := SizeOf(si);
  si.dwFlags     := STARTF_USESHOWWINDOW;
  si.wShowWindow := SW_SHOWNORMAL;

  if not CreateProcess(nil, PChar(Format('"%s" %s', [AFile, Params])), nil, nil,
                       false, CREATE_DEFAULT_ERROR_MODE, nil, PChar(ExtractFileDir(AFile)), si, pi) then
    MessageBox(Handle, PChar(SysErrorMessage(GetLastError())), nil, MB_ICONERROR)
  else
    Result := true;
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


function TCudaTextContextMenuHandler.GetCommandString(idCmd: UINT_PTR; uFlags: UINT;
  pwReserved: PUINT; pszName: LPSTR; cchMax: UINT): HResult;
begin
  Result := E_INVALIDARG;

  // Set help string on the Explorer status bar when the menu item is selected
  if (idCmd = 0) and (uFlags = GCS_HELPTEXT) then
  begin
    StrLCopy(PChar(pszName),
             PChar('Edit selected file(s) with CudaText'),
             cchMax);

    Result := NOERROR;
  end;
end;


function TCudaTextContextMenuHandler.InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult;
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
          MessageBox(lpici.hwnd, PChar(E.Message), nil, MB_ICONERROR);
      end;

      Result := NOERROR;
    end;
  end
end;


function TCudaTextContextMenuHandler.QueryContextMenu(Menu: HMENU;
  indexMenu, idCmdFirst, idCmdLast, uFlags: UINT): HResult;
const
  CMF_ITEMMENU     = $00000080;
  MENU_ITEM_CAPTON = 'Edit with CudaText';

var
  ContextMenuItem: TMenuItemInfo;
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

  else
    // fail for any other value
    Result := E_FAIL;

  if Result <> E_FAIL then
  begin
    FillChar(ContextMenuItem, SizeOf(ContextMenuItem), #0);

    ContextMenuItem.cbSize     := SizeOf(ContextMenuItem);
    ContextMenuItem.fMask      := MIIM_FTYPE or MIIM_STRING or MIIM_ID;
    ContextMenuItem.fType      := MFT_STRING;
    ContextMenuItem.wID        := idCmdFirst;
    ContextMenuItem.dwTypeData := PChar(MenuCaption);
    ContextMenuItem.cch        := Length(MenuCaption);

    InsertMenuItem(Menu, FMenuItemIndex, True, ContextMenuItem);
  end;
end;


function TCudaTextContextMenuHandler.ShellExtInitialize(pidlFolder: PItemIDList;
  lpdobj: IDataObject; hKeyProgID: HKEY): HResult;
var
  DataFormat: TFormatEtc;
  StrgMedium: TStgMedium;
  CntFiles:   UINT;
  Idx:        UINT;
  Buffer:     array[0..MAX_PATH] of Char;
  FileName:   string;

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
    CntFiles := DragQueryFile(StrgMedium.hGlobal, $FFFFFFFF, nil, 0);
    if CntFiles < 1 then exit;

    FCudaTextPath := GetCudaTextDir();
    FFileNames    := TStringList.Create;

    for Idx := 0 to Pred(CntFiles) do
    begin
      DragQueryFile(StrgMedium.hGlobal, Idx, @Buffer, SizeOf(Buffer));
      SetString(FileName, Buffer, Min(StrLen(Buffer), Pred(SizeOf(Buffer))));
      FFileNames.Add(FileName);
    end;

    Result := NOERROR;

  finally
    ReleaseStgMedium(StrgMedium);
  end;
end;


function TCudaTextContextMenuHandler.GetCudaTextDir: string;
var
  szBuf:    string;
  dwBufLen: DWORD;
  dwRet:    DWORD;

begin
  Result := '';

  dwBufLen := MAX_PATH;

  repeat
    SetLength(szBuf, dwBufLen);
    dwRet := GetModuleFileName(HInstance, PChar(szBuf), dwBufLen);

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
    SetString(Result, PChar(szBuf), dwRet);
    Result := ExtractFileDir(Result) + PathDelim  + 'cudatext.exe';
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
