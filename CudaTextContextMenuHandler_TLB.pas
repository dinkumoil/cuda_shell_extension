unit CudaTextContextMenuHandler_TLB;

{$MODE Delphi}

// ************************************************************************ //
// WARNING
// -------
// The types declared in this file were generated from data read from a
// Type Library. If this type library is explicitly or indirectly (via
// another type library referring to this type library) re-imported, or the
// 'Refresh' command of the Type Library Editor activated while editing the
// Type Library, the contents of this file will be regenerated and all
// manual modifications will be lost.
// ************************************************************************ //

// $Rev: 45604 $
// File generated on 11/24/2019 12:59:37 AM from Type Library described below.

// ************************************************************************  //
// Type Lib: S:\Dev\_src\CudaTextContextMenuHandler\CudaTextContextMenuHandler (1)
// LIBID: {9F5022BF-4CDB-4F5A-B9BE-7971D91B954F}
// LCID: 0
// Helpfile:
// HelpString:
// DepndLst:
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}


interface

uses
  Windows, Classes, Variants, Graphics;


const
// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
  // TypeLibrary Major and minor versions
  CudaTextContextMenuHandlerMajorVersion = 1;
  CudaTextContextMenuHandlerMinorVersion = 0;

  LIBID_CudaTextContextMenuHandler: TGUID = '{9F5022BF-4CDB-4F5A-B9BE-7971D91B954F}';

  IID_ICudaTextContextMenuHandler: TGUID = '{2C7271F7-C5F5-4E97-BE65-93552CEB96D1}';
  CLASS_CudaTextContextMenuHandler: TGUID = '{424D212F-385A-45BD-B844-12DE48079799}';


type
// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  ICudaTextContextMenuHandler = interface;
  ICudaTextContextMenuHandlerDisp = dispinterface;


// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  CudaTextContextMenuHandler = ICudaTextContextMenuHandler;


// *********************************************************************//
// Interface: ICudaTextContextMenuHandler
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2C7271F7-C5F5-4E97-BE65-93552CEB96D1}
// *********************************************************************//
  ICudaTextContextMenuHandler = interface(IDispatch)
    ['{2C7271F7-C5F5-4E97-BE65-93552CEB96D1}']
  end;


// *********************************************************************//
// DispIntf:  ICudaTextContextMenuHandlerDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {2C7271F7-C5F5-4E97-BE65-93552CEB96D1}
// *********************************************************************//
  ICudaTextContextMenuHandlerDisp = dispinterface
    ['{2C7271F7-C5F5-4E97-BE65-93552CEB96D1}']
  end;


// *********************************************************************//
// The Class CoCudaTextContextMenuHandler_ provides a Create and CreateRemote method to
// create instances of the default interface ICudaTextContextMenuHandler exposed by
// the CoClass CudaTextContextMenuHandler_. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//

  CoCudaTextContextMenuHandler = class
    class function Create: ICudaTextContextMenuHandler;
    class function CreateRemote(const MachineName: WideString): ICudaTextContextMenuHandler;
  end;



implementation

uses
  ComObj;


class function CoCudaTextContextMenuHandler.Create: ICudaTextContextMenuHandler;
begin
  Result := CreateComObject(CLASS_CudaTextContextMenuHandler) as ICudaTextContextMenuHandler;
end;


class function CoCudaTextContextMenuHandler.CreateRemote(const MachineName: WideString): ICudaTextContextMenuHandler;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CudaTextContextMenuHandler) as ICudaTextContextMenuHandler;
end;


end.

