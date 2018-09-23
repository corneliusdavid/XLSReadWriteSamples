object frmMain: TfrmMain
  Left = 555
  Top = 163
  Width = 576
  Height = 171
  Caption = 'Direct Write Sample'
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 560
    Height = 81
    Align = alTop
    TabOrder = 0
    object btnClose: TButton
      Left = 8
      Top = 46
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 1
      OnClick = btnCloseClick
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
      Top = 10
      Width = 408
      Height = 21
      TabOrder = 2
      OnChange = edWriteFilenameChange
    end
    object btnDlgOpen: TButton
      Left = 503
      Top = 8
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 3
      OnClick = btnDlgOpenClick
    end
    object btnExcel: TButton
      Left = 88
      Top = 46
      Width = 121
      Height = 25
      Caption = 'Open file in Excel'
      Enabled = False
      TabOrder = 4
      OnClick = btnExcelClick
    end
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 120
    Top = 84
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '5.20.55'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    OnWriteCell = XLSWriteCell
    Left = 188
    Top = 84
  end
  object XPManifest1: TXPManifest
    Left = 264
    Top = 84
  end
end
