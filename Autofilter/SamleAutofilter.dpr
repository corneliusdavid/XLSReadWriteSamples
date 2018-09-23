program SamleAutofilter;

uses
  Forms,
  Main in 'Main.pas' {frmMain};

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
