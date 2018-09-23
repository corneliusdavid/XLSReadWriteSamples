object frmMain: TfrmMain
  Left = 1426
  Top = 1055
  Caption = 'Conditional Formats Sample'
  ClientHeight = 343
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
    Left = 16
    Top = 104
    Width = 76
    Height = 13
    Caption = 'Cell ref. or area'
  end
  object Label2: TLabel
    Left = 16
    Top = 144
    Width = 44
    Height = 13
    Caption = 'Operator'
  end
  object Label3: TLabel
    Left = 16
    Top = 180
    Width = 35
    Height = 13
    Caption = 'Value 1'
  end
  object Label4: TLabel
    Left = 16
    Top = 208
    Width = 35
    Height = 13
    Caption = 'Value 2'
  end
  object Label5: TLabel
    Left = 16
    Top = 248
    Width = 34
    Height = 13
    Caption = 'Format'
  end
  object shpColor: TShape
    Left = 184
    Top = 240
    Width = 37
    Height = 25
  end
  object Label6: TLabel
    Left = 100
    Top = 280
    Width = 22
    Height = 13
    Caption = 'Font'
  end
  object lblCFCount: TLabel
    Left = 256
    Top = 324
    Width = 100
    Height = 13
    Caption = '0 conditional formats'
  end
  object Label7: TLabel
    Left = 256
    Top = 104
    Width = 234
    Height = 13
    Caption = 'Cell reference or area for the conditional format.'
  end
  object Label8: TLabel
    Left = 256
    Top = 140
    Width = 199
    Height = 13
    Caption = 'Operator to compare the cell values with.'
  end
  object Label9: TLabel
    Left = 252
    Top = 180
    Width = 255
    Height = 13
    Caption = 'Value that the operator compares the cell value with.'
  end
  object Label10: TLabel
    Left = 252
    Top = 208
    Width = 231
    Height = 13
    Caption = 'Second value. Used with the Between operator.'
  end
  object Label11: TLabel
    Left = 252
    Top = 244
    Width = 235
    Height = 13
    Caption = 'Formatting for the cell when the condition is met.'
  end
  object btnDlgOpen: TButton
    Left = 503
    Top = 8
    Width = 28
    Height = 25
    Caption = '...'
    TabOrder = 0
    OnClick = btnDlgOpenClick
  end
  object edFilename: TEdit
    Left = 89
    Top = 10
    Width = 408
    Height = 21
    TabOrder = 1
  end
  object btnExcel: TButton
    Left = 8
    Top = 46
    Width = 121
    Height = 25
    Caption = 'Open file in Excel'
    TabOrder = 2
    OnClick = btnExcelClick
  end
  object btnAddCondFmt: TButton
    Left = 96
    Top = 318
    Width = 149
    Height = 25
    Caption = 'Add Conditional formats'
    TabOrder = 3
    OnClick = btnAddCondFmtClick
  end
  object btnClose: TButton
    Left = 468
    Top = 318
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 4
    OnClick = btnCloseClick
  end
  object btnWrite: TButton
    Left = 8
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Write'
    TabOrder = 5
    OnClick = btnWriteClick
  end
  object edRef: TEdit
    Left = 100
    Top = 100
    Width = 149
    Height = 21
    TabOrder = 6
    Text = 'A1'
  end
  object cbOperator: TComboBox
    Left = 100
    Top = 136
    Width = 149
    Height = 21
    Style = csDropDownList
    TabOrder = 7
    Items.Strings = (
      '< Less than'
      '<= Less than or equal'
      '= Equal'
      '<> Not equal'
      '>= Greater than or equal'
      '> Greater than'
      'Between'
      'Not between')
  end
  object edValue1: TEdit
    Left = 100
    Top = 176
    Width = 145
    Height = 21
    TabOrder = 8
  end
  object edValue2: TEdit
    Left = 100
    Top = 204
    Width = 145
    Height = 21
    TabOrder = 9
  end
  object btnColor: TButton
    Left = 100
    Top = 240
    Width = 75
    Height = 25
    Caption = 'Color'
    TabOrder = 10
    OnClick = btnColorClick
  end
  object cbBold: TCheckBox
    Left = 132
    Top = 280
    Width = 45
    Height = 17
    Caption = 'Bold'
    TabOrder = 11
  end
  object cbItalic: TCheckBox
    Left = 184
    Top = 280
    Width = 57
    Height = 17
    Caption = 'Italic'
    TabOrder = 12
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 380
    Top = 44
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 432
    Top = 44
  end
  object XPManifest1: TXPManifest
    Left = 488
    Top = 44
  end
  object dlgColor: TColorDialog
    Left = 328
    Top = 44
  end
end
