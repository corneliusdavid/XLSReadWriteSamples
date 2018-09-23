object frmMain: TfrmMain
  Left = 744
  Top = 127
  Caption = 'Sample Named Cells'
  ClientHeight = 691
  ClientWidth = 736
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
    Width = 736
    Height = 121
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 95
      Top = 66
      Width = 420
      Height = 16
      Caption = 'THIS DOES NOT WORK! IT WILL MESS UP YOUR SPREADSHEET!!!'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold, fsItalic]
      ParentFont = False
    end
    object btnRead: TButton
      Left = 9
      Top = 9
      Width = 80
      Height = 20
      Caption = 'Read'
      TabOrder = 0
      OnClick = btnReadClick
    end
    object btnWrite: TButton
      Left = 9
      Top = 38
      Width = 80
      Height = 22
      Caption = 'Write'
      TabOrder = 3
      OnClick = btnWriteClick
    end
    object edReadFilename: TEdit
      Left = 96
      Top = 10
      Width = 438
      Height = 22
      TabOrder = 1
    end
    object edWriteFilename: TEdit
      Left = 95
      Top = 38
      Width = 438
      Height = 22
      TabOrder = 4
    end
    object btnDlgOpen: TButton
      Left = 540
      Top = 9
      Width = 31
      Height = 20
      Caption = '...'
      TabOrder = 2
      OnClick = btnDlgOpenClick
    end
    object btnDlgSave: TButton
      Left = 540
      Top = 38
      Width = 31
      Height = 22
      Caption = '...'
      TabOrder = 5
      OnClick = btnDlgSaveClick
    end
    object btnAddName: TButton
      Left = 96
      Top = 88
      Width = 80
      Height = 25
      Caption = 'Edit Name'
      TabOrder = 7
      OnClick = btnAddNameClick
    end
    object btnEditName: TButton
      Left = 182
      Top = 88
      Width = 82
      Height = 25
      Caption = 'Add Name'
      TabOrder = 8
      OnClick = btnEditNameClick
    end
    object btnClose: TButton
      Left = 9
      Top = 88
      Width = 80
      Height = 25
      Caption = 'Close'
      TabOrder = 6
      OnClick = btnCloseClick
    end
    object Button1: TButton
      Left = 522
      Top = 88
      Width = 115
      Height = 27
      Caption = 'Open in Excel'
      TabOrder = 10
      OnClick = Button1Click
    end
    object btnNameValues: TButton
      Left = 284
      Top = 88
      Width = 97
      Height = 25
      Caption = 'Name Values'
      TabOrder = 9
      OnClick = btnNameValuesClick
    end
  end
  object lvNames: TListView
    Left = 0
    Top = 121
    Width = 736
    Height = 570
    Align = alClient
    Columns = <
      item
        Caption = 'Name'
        Width = 161
      end
      item
        Caption = 'Value'
        Width = 161
      end
      item
        Caption = 'Referes To'
        Width = 161
      end
      item
        Caption = 'Scope'
        Width = 161
      end
      item
        Caption = 'Comment'
        Width = 107
      end>
    ReadOnly = True
    RowSelect = True
    TabOrder = 1
    ViewStyle = vsReport
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 592
    Top = 24
  end
  object XPManifest: TXPManifest
    Left = 652
    Top = 24
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 592
    Top = 68
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 652
    Top = 68
  end
end
