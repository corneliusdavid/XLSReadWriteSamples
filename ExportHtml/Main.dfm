object frmMain: TfrmMain
  Left = 374
  Top = 144
  Caption = 'Export HTML'
  ClientHeight = 393
  ClientWidth = 574
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 16
    Top = 104
    Width = 80
    Height = 14
    Caption = 'Sheets to export'
  end
  object Label2: TLabel
    Left = 184
    Top = 128
    Width = 356
    Height = 14
    Caption = 
      'Will only write cell values without any formatting, pictures, co' +
      'mments, etc.'
  end
  object Label3: TLabel
    Left = 184
    Top = 228
    Width = 141
    Height = 14
    Caption = 'Include in exported HTML file:'
  end
  object Label4: TLabel
    Left = 204
    Top = 376
    Width = 69
    Height = 14
    Caption = 'Cells to export'
  end
  object Label5: TLabel
    Left = 184
    Top = 180
    Width = 304
    Height = 14
    Caption = 'The separate files will have the same name as the sheet name.'
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 574
    Height = 97
    Align = alTop
    TabOrder = 6
    object btnRead: TButton
      Left = 8
      Top = 6
      Width = 75
      Height = 25
      Caption = 'Read'
      TabOrder = 6
      OnClick = btnReadClick
    end
    object btnWrite: TButton
      Left = 8
      Top = 36
      Width = 75
      Height = 25
      Caption = 'Export'
      TabOrder = 2
      OnClick = btnWriteClick
    end
    object edReadFilename: TEdit
      Left = 89
      Top = 10
      Width = 408
      Height = 22
      TabOrder = 0
    end
    object edWriteFilename: TEdit
      Left = 89
      Top = 37
      Width = 408
      Height = 22
      TabOrder = 3
    end
    object btnDlgOpen: TButton
      Left = 503
      Top = 8
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 1
      OnClick = btnDlgOpenClick
    end
    object btnDlgSave: TButton
      Left = 503
      Top = 35
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 4
      OnClick = btnDlgSaveClick
    end
    object btnClose: TButton
      Left = 8
      Top = 67
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 5
      OnClick = btnCloseClick
    end
    object Progress: TProgressBar
      Left = 92
      Top = 64
      Width = 405
      Height = 13
      TabOrder = 7
      Visible = False
    end
  end
  object clbSheets: TCheckListBox
    Left = 8
    Top = 120
    Width = 169
    Height = 273
    ItemHeight = 14
    TabOrder = 0
  end
  object cbSimpleExport: TCheckBox
    Left = 184
    Top = 144
    Width = 97
    Height = 17
    Caption = 'Simple export. '
    TabOrder = 1
    OnClick = cbSimpleExportClick
  end
  object cbComments: TCheckBox
    Left = 184
    Top = 244
    Width = 97
    Height = 17
    Caption = 'Comments'
    Checked = True
    State = cbChecked
    TabOrder = 3
    OnClick = cbCommentsClick
  end
  object cbImages: TCheckBox
    Left = 184
    Top = 268
    Width = 97
    Height = 17
    Caption = 'Images'
    Checked = True
    State = cbChecked
    TabOrder = 4
    OnClick = cbImagesClick
  end
  object cbHyperlinks: TCheckBox
    Left = 184
    Top = 316
    Width = 97
    Height = 17
    Caption = 'Hyperlinks'
    Checked = True
    State = cbChecked
    TabOrder = 5
    OnClick = cbHyperlinksClick
  end
  object cbSeparateFiles: TCheckBox
    Left = 184
    Top = 196
    Width = 229
    Height = 17
    Caption = 'Write sheets to separate files.'
    TabOrder = 2
    OnClick = cbSeparateFilesClick
  end
  object cbWriteImages: TCheckBox
    Left = 204
    Top = 292
    Width = 113
    Height = 17
    Caption = 'Write image files'
    Checked = True
    State = cbChecked
    TabOrder = 7
    OnClick = cbWriteImagesClick
  end
  object cbExportAll: TCheckBox
    Left = 184
    Top = 352
    Width = 97
    Height = 17
    Caption = 'Export all cells'
    Checked = True
    State = cbChecked
    TabOrder = 8
    OnClick = cbExportAllClick
  end
  object edArea: TEdit
    Left = 276
    Top = 372
    Width = 77
    Height = 22
    Enabled = False
    TabOrder = 9
    Text = 'A1:H20'
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 328
    Top = 228
  end
  object dlgSave: TSaveDialog
    Filter = 'HTML files (*.htm)|*.htm|All files (*.*)|*.*'
    Left = 396
    Top = 228
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    OnProgress = XLSProgress
    Left = 464
    Top = 228
  end
  object ExportHTML: TXLSExportHTML5
    Options = []
    Col1 = -1
    Col2 = -1
    FileExtension = 'htm'
    Row1 = -1
    Row2 = -1
    XLS = XLS
    WriteOnlyTables = False
    SimpleExport = False
    HTMLOPtions = [xeohComments, xeohImages, xeohWriteImages, xeohHyperlinks]
    TABLE.BordeWidth = 0
    TABLE.CellPadding = 0
    TABLE.CellSpacing = 0
    Left = 528
    Top = 228
  end
  object XPManifest1: TXPManifest
    Left = 464
    Top = 288
  end
end
