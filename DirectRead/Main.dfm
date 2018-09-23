object frmMain: TfrmMain
  Left = 313
  Top = 124
  Width = 682
  Height = 229
  Caption = 'Direct Read'
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
  DesignSize = (
    666
    191)
  PixelsPerInch = 96
  TextHeight = 13
  object lblSum: TLabel
    Left = 16
    Top = 96
    Width = 73
    Height = 13
    Caption = 'Sum of all cells:'
  end
  object lblString: TLabel
    Left = 16
    Top = 120
    Width = 641
    Height = 13
    Anchors = [akLeft, akTop, akRight]
    AutoSize = False
    Caption = 'Longest string:'
  end
  object lblFormulas: TLabel
    Left = 16
    Top = 144
    Width = 98
    Height = 13
    Caption = 'Number of formulas:'
  end
  object lblTotal: TLabel
    Left = 16
    Top = 168
    Width = 103
    Height = 13
    Caption = 'Total number of cells:'
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 666
    Height = 81
    Align = alTop
    TabOrder = 0
    object lblStatus: TLabel
      Left = 92
      Top = 52
      Width = 35
      Height = 13
      Caption = 'Status:'
    end
    object btnRead: TButton
      Left = 8
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Read'
      TabOrder = 0
      OnClick = btnReadClick
    end
    object edReadFilename: TEdit
      Left = 89
      Top = 10
      Width = 408
      Height = 21
      TabOrder = 1
    end
    object btnDlgOpen: TButton
      Left = 503
      Top = 8
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 2
      OnClick = btnDlgOpenClick
    end
    object btnClose: TButton
      Left = 8
      Top = 46
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 3
      OnClick = btnCloseClick
    end
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx;*.xls|All files (*.*)|*.*'
    Left = 320
    Top = 140
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '5.10.06a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    OnReadCell = XLSReadCell
    Left = 394
    Top = 140
  end
  object XPManifest1: TXPManifest
    Left = 468
    Top = 140
  end
end
