program SampleDefinedNames;

uses
  Forms,
  Main in 'Main.pas' {frmMain},
  EditName in 'EditName.pas' {frmEditName};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
