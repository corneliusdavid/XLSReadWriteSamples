object frmMain: TfrmMain
  Left = 313
  Top = 125
  Caption = 'Format cells'
  ClientHeight = 315
  ClientWidth = 558
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 240
    Top = 188
    Width = 46
    Height = 13
    Caption = 'Password'
  end
  object btnWrite: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Write'
    TabOrder = 0
    OnClick = btnWriteClick
  end
  object edWriteFilename: TEdit
    Left = 89
    Top = 9
    Width = 408
    Height = 21
    TabOrder = 1
    OnChange = edWriteFilenameChange
  end
  object btnDlgSave: TButton
    Left = 503
    Top = 7
    Width = 28
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = btnDlgSaveClick
  end
  object btnSingleCell: TButton
    Left = 9
    Top = 90
    Width = 224
    Height = 25
    Caption = 'Single cell'
    TabOrder = 3
    OnClick = btnSingleCellClick
  end
  object btnAreas: TButton
    Left = 10
    Top = 122
    Width = 223
    Height = 25
    Caption = 'Areas of cells'
    TabOrder = 4
    OnClick = btnAreasClick
  end
  object btnClose: TButton
    Left = 8
    Top = 42
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 5
    OnClick = btnCloseClick
  end
  object btnDefaultFormat: TButton
    Left = 10
    Top = 154
    Width = 223
    Height = 25
    Caption = 'Use default formats'
    TabOrder = 6
    OnClick = btnDefaultFormatClick
  end
  object btnExcel: TButton
    Left = 92
    Top = 42
    Width = 141
    Height = 25
    Caption = 'Open file in Excel'
    TabOrder = 7
    OnClick = btnExcelClick
  end
  object btnProtect: TButton
    Left = 12
    Top = 184
    Width = 221
    Height = 25
    Caption = 'Password protect sheet'
    TabOrder = 8
    OnClick = btnProtectClick
  end
  object btnColumns: TButton
    Left = 12
    Top = 216
    Width = 221
    Height = 25
    Caption = 'Format colummns'
    TabOrder = 9
    OnClick = btnColumnsClick
  end
  object btnRows: TButton
    Left = 12
    Top = 248
    Width = 221
    Height = 25
    Caption = 'Format rows'
    TabOrder = 10
    OnClick = btnRowsClick
  end
  object edPassword: TEdit
    Left = 292
    Top = 184
    Width = 121
    Height = 21
    TabOrder = 11
    Text = 'quokka'
  end
  object Button1: TButton
    Left = 12
    Top = 280
    Width = 221
    Height = 25
    Caption = 'Cell Borders'
    TabOrder = 12
    OnClick = Button1Click
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 408
    Top = 40
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 472
    Top = 40
  end
  object XPManifest1: TXPManifest
    Left = 332
    Top = 40
  end
end
