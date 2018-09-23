unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XLSReadWriteII5, StdCtrls, ExtCtrls, IniFiles, XLSComment5,
  XLSDrawing5, Xc12Utils5, XLSSheetData5, ExtDlgs, Spin, XPMan;

type
  TfrmMain = class(TForm)
    Panel: TPanel;
    btnRead: TButton;
    btnWrite: TButton;
    edReadFilename: TEdit;
    edWriteFilename: TEdit;
    btnDlgOpen: TButton;
    btnDlgSave: TButton;
    btnAdd1: TButton;
    btnClose: TButton;
    dlgSave: TSaveDialog;
    dlgOpen: TOpenDialog;
    XLS: TXLSReadWriteII5;
    btnAdd3: TButton;
    ScrollBar: TScrollBar;
    PBox: TPaintBox;
    edImageFile: TEdit;
    dlgImage: TOpenPictureDialog;
    seCol: TSpinEdit;
    seRow: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    edWidth: TEdit;
    lblUnitsW: TLabel;
    Label5: TLabel;
    edHeight: TEdit;
    lblUnitsH: TLabel;
    cbAspect: TCheckBox;
    cbUnits: TCheckBox;
    btnUpdateWidth: TButton;
    btnUpdateHeight: TButton;
    seScale: TSpinEdit;
    Label4: TLabel;
    Label6: TLabel;
    btnUpdateScale: TButton;
    XPManifest1: TXPManifest;
    Label7: TLabel;
    cbPositioning: TComboBox;
    procedure btnCloseClick(Sender: TObject);
    procedure btnReadClick(Sender: TObject);
    procedure btnWriteClick(Sender: TObject);
    procedure btnDlgOpenClick(Sender: TObject);
    procedure btnDlgSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ScrollBarChange(Sender: TObject);
    procedure PBoxPaint(Sender: TObject);
    procedure btnAdd3Click(Sender: TObject);
    procedure btnAdd1Click(Sender: TObject);
    procedure cbUnitsClick(Sender: TObject);
    procedure btnUpdateHeightClick(Sender: TObject);
    procedure btnUpdateWidthClick(Sender: TObject);
    procedure btnUpdateScaleClick(Sender: TObject);
  private
    FBitmaps: array of TBitmap;

    procedure FreeImages;
    procedure ReadPictures;
    procedure PaintImage(const AIndex: integer);
    procedure EnableControls(const AEnabled: boolean);
    procedure UpdateControls;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnAdd1Click(Sender: TObject);
var
  Image: TXLSDrawingImage;
begin
//function InsertImage(const AFilename: AxUCString;  Image filename
//                     const ACol: integer;          Column
//                     const ARow: integer;          Row
//                     const AColOffset: double;     Column offset, in percent
//                     const ARowOffset: double;     Row offset, in percent
//                     const AScale: double          Scale image by this value, percent
//                     ): TXLSDrawingImage;
  Image := XLS[0].Drawing.InsertImage(edImageFile.Text,seCol.Value - 1,seRow.Value,0,0,1.25);
  Image.Positioning := TDrwOptImagePositioning(cbPositioning.ItemIndex);

  ReadPictures;

  UpdateControls;
end;

procedure TfrmMain.btnAdd3Click(Sender: TObject);
begin
  dlgImage.InitialDir := ExtractFilePath(Application.ExeName);
  dlgImage.FileName := edImageFile.Text;
  if dlgImage.Execute then
    edImageFile.Text := dlgImage.FileName;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMain.btnReadClick(Sender: TObject);
begin
  XLS.Filename := edReadFilename.Text;
  XLS.Read;

  EnableControls(False);
  
  ReadPictures;
end;

procedure TfrmMain.btnWriteClick(Sender: TObject);
begin
  XLS.Filename := edWriteFilename.Text;
  XLS.Write;
end;

procedure TfrmMain.btnDlgOpenClick(Sender: TObject);
begin
  dlgOpen.InitialDir := ExtractFilePath(Application.ExeName);
  dlgOpen.FileName := edReadFilename.Text;
  if dlgOpen.Execute then
    edReadFilename.Text := dlgOpen.FileName;
end;

procedure TfrmMain.btnDlgSaveClick(Sender: TObject);
begin
  dlgSave.InitialDir := ExtractFilePath(Application.ExeName);
  dlgSave.FileName := edWriteFilename.Text;
  if dlgSave.Execute then
    edWriteFilename.Text := dlgSave.FileName;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    edReadFilename.Text := Ini.ReadString('Files','Read','');
    edWriteFilename.Text := Ini.ReadString('Files','Write','');
    edImageFile.Text := Ini.ReadString('Files','Image','');
  finally
    Ini.Free;
  end;

//  type TDrwOptImagePositioning = (doipMoveButDoNotSize,doipMoveAndSize,doipDoNotMoveOrSize);
  cbPositioning.Items.Add('Move but do not size');
  cbPositioning.Items.Add('Move and size');
  cbPositioning.Items.Add('Do not move or size');
  cbPositioning.ItemIndex := 0;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
var
  S: string;
  Ini: TIniFile;
begin
  S := ChangeFileExt(Application.ExeName,'.ini');
  Ini := TIniFile.Create(S);
  try
    Ini.WriteString('Files','Read',edReadFilename.Text);
    Ini.WriteString('Files','Write',edWriteFilename.Text);
    Ini.WriteString('Files','Image',edImageFile.Text);
  finally
    Ini.Free;
  end;

  FreeImages;
end;

procedure TfrmMain.FreeImages;
var
  i: integer;
begin
  for i := 0 to High(FBitmaps) do begin
    if FBitmaps[i] <> Nil then
      FBitmaps[i].Free;
  end;
end;

procedure TfrmMain.PaintImage(const AIndex: integer);
var
  Scale: double;
begin
  if FBitmaps[AIndex] <> Nil then begin
    if FBitmaps[AIndex].Width > PBox.Width then begin
      Scale := PBox.Width / FBitmaps[AIndex].Width;
      PBox.Canvas.StretchDraw(Rect(0,0,PBox.Width,Round(FBitmaps[AIndex].Height * Scale)),FBitmaps[AIndex]);
    end
    else
      PBox.Canvas.Draw(0,0,FBitmaps[AIndex]);
  end;
end;

procedure TfrmMain.PBoxPaint(Sender: TObject);
begin
  if ScrollBar.Enabled then
    PaintImage(ScrollBar.Position);
end;

procedure TfrmMain.ReadPictures;
var
  i: integer;
  Image: TXLSDrawingImage;
begin
  FreeImages;

  SetLength(FBitmaps,XLS[0].Drawing.Images.Count);
  for i := 0 to XLS[0].Drawing.Images.Count - 1 do begin
    Image := XLS[0].Drawing.Images[i];
    FBitmaps[i] := Image.CreateBitmap;
  end;
  ScrollBar.Max := High(FBitmaps);
  ScrollBar.Enabled := High(FBitmaps) >= 0;
  EnableControls(High(FBitmaps) >= 0);
  if High(FBitmaps) >= 0 then
    UpdateControls;
  PBox.Invalidate;
end;

procedure TfrmMain.ScrollBarChange(Sender: TObject);
begin
  UpdateControls;
  PBox.Invalidate;
end;

procedure TfrmMain.EnableControls(const AEnabled: boolean);
var
  i: integer;
begin
  for i := 0 to Panel.ControlCount - 1 do begin
    if Panel.Controls[i].Tag = 100 then
      Panel.Controls[i].Enabled := AEnabled;
  end;
end;

procedure TfrmMain.UpdateControls;
var
  Image: TXLSDrawingImage;
  Editor: TXLSDrawingEditorImage;
begin
  if XLS[0].Drawing.Images.Count > 0 then begin
    Image := XLS[0].Drawing.Images[ScrollBar.Position];
    if Image <> Nil then begin
      Editor := XLS[0].Drawing.EditImage(Image);
      try
        if cbUnits.Checked then begin
          edWidth.Text := Format('%.2f',[Editor.CmWidth]);
          edHeight.Text := Format('%.2f',[Editor.CmHeight]);
        end
        else begin
          edWidth.Text := Format('%.2f',[Editor.InchWidth]);
          edHeight.Text := Format('%.2f',[Editor.InchHeight]);
        end;
      finally
        Editor.Free;
      end;
    end;
  end;
end;

procedure TfrmMain.cbUnitsClick(Sender: TObject);
begin
  if cbUnits.Checked then begin
    lblUnitsW.Caption := 'cm';
    lblUnitsH.Caption := 'cm';
  end
  else begin
    lblUnitsW.Caption := 'inch';
    lblUnitsH.Caption := 'inch';
  end;

  UpdateControls;
end;

procedure TfrmMain.btnUpdateHeightClick(Sender: TObject);
var
  Image: TXLSDrawingImage;
  Editor: TXLSDrawingEditorImage;
  V: double;
begin
  if TryStrToFloat(edHeight.Text,V) then begin
    Image := XLS[0].Drawing.Images[ScrollBar.Position];
    if Image <> Nil then begin
      Editor := XLS[0].Drawing.EditImage(Image);
      try
        Editor.KeepAspect := cbAspect.Checked;
        if cbUnits.Checked then
          Editor.CmHeight := V
        else
          Editor.InchHeight := V;
      finally
        Editor.Free;
      end;
    end;
  end;

  UpdateControls;
end;

procedure TfrmMain.btnUpdateWidthClick(Sender: TObject);
var
  Image: TXLSDrawingImage;
  Editor: TXLSDrawingEditorImage;
  V: double;
begin
  if TryStrToFloat(edWidth.Text,V) then begin
    Image := XLS[0].Drawing.Images[ScrollBar.Position];
    if Image <> Nil then begin
      Editor := XLS[0].Drawing.EditImage(Image);
      try
        Editor.KeepAspect := cbAspect.Checked;
        if cbUnits.Checked then
          Editor.CmWidth := V
        else
          Editor.InchWidth := V;
      finally
        Editor.Free;
      end;
    end;
  end;

  UpdateControls;
end;

procedure TfrmMain.btnUpdateScaleClick(Sender: TObject);
var
  Image: TXLSDrawingImage;
  Editor: TXLSDrawingEditorImage;
begin
  Image := XLS[0].Drawing.Images[ScrollBar.Position];
  if Image <> Nil then begin
    Editor := XLS[0].Drawing.EditImage(Image);
    try
      Editor.Scale(seScale.Value / 100);
    finally
      Editor.Free;
    end;
  end;

  UpdateControls;
end;

end.
