program docxReplacer;

{$mode DelphiUnicode}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Classes,
  SysUtils,
  docxreplacerapp;

var
  Application: TDocxReplacerApp;

begin
  Application := TDocxReplacerApp.Create(nil);
  Application.Title := 'docxReplacer';
  Application.Run;
  Application.Free;
end.
