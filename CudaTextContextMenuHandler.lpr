library CudaTextContextMenuHandler;

{$MODE Delphi}


uses
  ComServ,
  CudaTextContextMenuHandler_TLB in 'CudaTextContextMenuHandler_TLB.pas',
  CudaTextContextMenuHandlerImpl in 'CudaTextContextMenuHandlerImpl.pas' {CudaTextContextMenuHandler: CoClass};


exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;


{$R *.tlb}
{$R *.res}


begin
end.
