object frmMain: TfrmMain
  Left = 513
  Top = 177
  Caption = 'Copy, Move and Delete'
  ClientHeight = 151
  ClientWidth = 581
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 581
    Height = 97
    Align = alTop
    TabOrder = 0
    object btnRead: TButton
      Left = 8
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Read'
      TabOrder = 0
      OnClick = btnReadClick
    end
    object btnWrite: TButton
      Left = 8
      Top = 36
      Width = 75
      Height = 25
      Caption = 'Write'
      TabOrder = 1
      OnClick = btnWriteClick
    end
    object edReadFilename: TEdit
      Left = 89
      Top = 10
      Width = 408
      Height = 22
      TabOrder = 2
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
      TabOrder = 4
      OnClick = btnDlgOpenClick
    end
    object btnDlgSave: TButton
      Left = 503
      Top = 35
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 5
      OnClick = btnDlgSaveClick
    end
    object Button1: TButton
      Left = 8
      Top = 64
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 6
      OnClick = Button1Click
    end
    object btnCopy: TButton
      Left = 88
      Top = 64
      Width = 193
      Height = 25
      Caption = 'Copy, Move and Delete cells'
      TabOrder = 7
      OnClick = btnCopyClick
    end
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 164
    Top = 104
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 104
    Top = 104
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 47
    Top = 104
  end
  object XPManifest1: TXPManifest
    Left = 232
    Top = 104
  end
end
