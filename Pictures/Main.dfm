object frmMain: TfrmMain
  Left = 412
  Top = 195
  Caption = 'Pictures'
  ClientHeight = 640
  ClientWidth = 595
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
  object PBox: TPaintBox
    Left = 0
    Top = 197
    Width = 578
    Height = 443
    Align = alClient
    OnPaint = PBoxPaint
    ExplicitWidth = 580
    ExplicitHeight = 452
  end
  object Panel: TPanel
    Left = 0
    Top = 0
    Width = 595
    Height = 197
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 104
      Top = 96
      Width = 35
      Height = 13
      Caption = 'Column'
    end
    object Label2: TLabel
      Left = 205
      Top = 96
      Width = 21
      Height = 13
      Caption = 'Row'
    end
    object Label3: TLabel
      Tag = 100
      Left = 12
      Top = 140
      Width = 28
      Height = 13
      Caption = 'Width'
      Enabled = False
    end
    object lblUnitsW: TLabel
      Tag = 100
      Left = 112
      Top = 140
      Width = 13
      Height = 13
      Caption = 'cm'
      Enabled = False
    end
    object Label5: TLabel
      Tag = 100
      Left = 192
      Top = 140
      Width = 31
      Height = 13
      Caption = 'Height'
      Enabled = False
    end
    object lblUnitsH: TLabel
      Tag = 100
      Left = 292
      Top = 140
      Width = 13
      Height = 13
      Caption = 'cm'
      Enabled = False
    end
    object Label4: TLabel
      Tag = 100
      Left = 384
      Top = 140
      Width = 25
      Height = 13
      Caption = 'Scale'
    end
    object Label6: TLabel
      Tag = 100
      Left = 480
      Top = 140
      Width = 11
      Height = 13
      Caption = '%'
    end
    object Label7: TLabel
      Left = 296
      Top = 96
      Width = 88
      Height = 13
      Caption = 'Image positioning:'
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
      Height = 21
      TabOrder = 2
    end
    object edWriteFilename: TEdit
      Left = 89
      Top = 37
      Width = 408
      Height = 21
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
    object btnAdd1: TButton
      Left = 8
      Top = 64
      Width = 75
      Height = 25
      Caption = 'Add image'
      TabOrder = 6
      OnClick = btnAdd1Click
    end
    object btnClose: TButton
      Left = 8
      Top = 91
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 7
      OnClick = btnCloseClick
    end
    object btnAdd3: TButton
      Left = 503
      Top = 62
      Width = 28
      Height = 25
      Caption = '...'
      TabOrder = 8
      OnClick = btnAdd3Click
    end
    object edImageFile: TEdit
      Left = 89
      Top = 64
      Width = 408
      Height = 21
      TabOrder = 9
    end
    object seCol: TSpinEdit
      Left = 144
      Top = 91
      Width = 53
      Height = 22
      MaxValue = 256
      MinValue = 1
      TabOrder = 10
      Value = 1
    end
    object seRow: TSpinEdit
      Left = 232
      Top = 91
      Width = 53
      Height = 22
      MaxValue = 65536
      MinValue = 1
      TabOrder = 11
      Value = 1
    end
    object edWidth: TEdit
      Tag = 100
      Left = 52
      Top = 136
      Width = 57
      Height = 21
      Enabled = False
      TabOrder = 12
    end
    object edHeight: TEdit
      Tag = 100
      Left = 232
      Top = 136
      Width = 57
      Height = 21
      Enabled = False
      TabOrder = 14
    end
    object cbAspect: TCheckBox
      Tag = 100
      Left = 12
      Top = 168
      Width = 85
      Height = 17
      Caption = 'Lock aspect'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 18
    end
    object cbUnits: TCheckBox
      Tag = 100
      Left = 96
      Top = 168
      Width = 97
      Height = 17
      Caption = 'Centimeters'
      Checked = True
      Enabled = False
      State = cbChecked
      TabOrder = 19
      OnClick = cbUnitsClick
    end
    object btnUpdateWidth: TButton
      Tag = 100
      Left = 132
      Top = 136
      Width = 53
      Height = 21
      Caption = 'Update'
      Enabled = False
      TabOrder = 13
      OnClick = btnUpdateWidthClick
    end
    object btnUpdateHeight: TButton
      Tag = 100
      Left = 312
      Top = 136
      Width = 53
      Height = 21
      Caption = 'Update'
      Enabled = False
      TabOrder = 15
      OnClick = btnUpdateHeightClick
    end
    object seScale: TSpinEdit
      Tag = 100
      Left = 416
      Top = 136
      Width = 61
      Height = 22
      MaxValue = 500
      MinValue = 1
      TabOrder = 16
      Value = 100
    end
    object btnUpdateScale: TButton
      Tag = 100
      Left = 500
      Top = 136
      Width = 53
      Height = 21
      Caption = 'Update'
      Enabled = False
      TabOrder = 17
      OnClick = btnUpdateScaleClick
    end
    object cbPositioning: TComboBox
      Left = 388
      Top = 92
      Width = 145
      Height = 21
      Style = csDropDownList
      TabOrder = 20
    end
  end
  object ScrollBar: TScrollBar
    Left = 578
    Top = 197
    Width = 17
    Height = 443
    Align = alRight
    Enabled = False
    Kind = sbVertical
    PageSize = 0
    TabOrder = 1
    OnChange = ScrollBarChange
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 384
    Top = 240
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 324
    Top = 240
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 442
    Top = 240
  end
  object dlgImage: TOpenPictureDialog
    Filter = 
      'All (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|Portable' +
      ' Network Graphics (*.png)|*.png|JPEG Image File (*.jpg)|*.jpg|JP' +
      'EG Image File (*.jpeg)|*.jpeg|Bitmaps (*.bmp)|*.jpeg'
    Left = 264
    Top = 240
  end
  object XPManifest1: TXPManifest
    Left = 508
    Top = 240
  end
end
