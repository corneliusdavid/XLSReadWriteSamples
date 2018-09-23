object frmMain: TfrmMain
  Left = 497
  Top = 179
  Caption = 'Read and write hyperlinks'
  ClientHeight = 642
  ClientWidth = 582
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Calibri'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 582
    Height = 113
    Align = alTop
    TabOrder = 0
    object btnRead: TButton
      Left = 10
      Top = 10
      Width = 92
      Height = 27
      Caption = 'Read'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Calibri'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = btnReadClick
    end
    object btnWrite: TButton
      Left = 10
      Top = 44
      Width = 92
      Height = 25
      Caption = 'Write'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Calibri'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = btnWriteClick
    end
    object edReadFilename: TEdit
      Left = 110
      Top = 12
      Width = 415
      Height = 21
      TabOrder = 2
    end
    object edWriteFilename: TEdit
      Left = 110
      Top = 46
      Width = 415
      Height = 21
      TabOrder = 3
    end
    object btnDlgOpen: TButton
      Left = 531
      Top = 10
      Width = 26
      Height = 23
      Caption = '...'
      TabOrder = 4
      OnClick = btnDlgOpenClick
    end
    object btnDlgSave: TButton
      Left = 531
      Top = 44
      Width = 26
      Height = 25
      Caption = '...'
      TabOrder = 5
      OnClick = btnDlgSaveClick
    end
    object Button1: TButton
      Left = 10
      Top = 79
      Width = 92
      Height = 26
      Caption = 'Close'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Calibri'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      OnClick = Button1Click
    end
    object btnAddCells: TButton
      Left = 108
      Top = 79
      Width = 130
      Height = 26
      Caption = 'Add hyperlinks'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Calibri'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
      OnClick = btnAddCellsClick
    end
    object Button2: TButton
      Left = 248
      Top = 80
      Width = 129
      Height = 25
      Caption = 'Open in Excel'
      TabOrder = 8
      OnClick = Button2Click
    end
  end
  object Grid: TStringGrid
    Left = 0
    Top = 113
    Width = 582
    Height = 529
    Align = alClient
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 101
    TabOrder = 1
    ColWidths = (
      38
      169
      69
      124
      148)
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 140
    Top = 148
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 76
    Top = 148
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 216
    Top = 148
  end
  object XPManifest1: TXPManifest
    Left = 296
    Top = 148
  end
end
