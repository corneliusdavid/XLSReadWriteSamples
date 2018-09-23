object frmMain: TfrmMain
  Left = 214
  Top = 302
  Caption = 'Autofilter sample'
  ClientHeight = 285
  ClientWidth = 734
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    734
    285)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 220
    Top = 126
    Width = 54
    Height = 13
    Caption = 'Condition 1'
  end
  object Label2: TLabel
    Left = 296
    Top = 126
    Width = 35
    Height = 13
    Caption = 'Value 1'
  end
  object Label3: TLabel
    Left = 579
    Top = 126
    Width = 35
    Height = 13
    Caption = 'Value 2'
  end
  object Label4: TLabel
    Left = 160
    Top = 126
    Width = 35
    Height = 13
    Caption = 'Column'
  end
  object Label5: TLabel
    Left = 148
    Top = 97
    Width = 28
    Height = 13
    Caption = 'Sheet'
  end
  object Label6: TLabel
    Left = 500
    Top = 124
    Width = 54
    Height = 13
    Caption = 'Condition 2'
  end
  object Button1: TButton
    Left = 16
    Top = 260
    Width = 117
    Height = 25
    Caption = 'Close'
    TabOrder = 17
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 16
    Top = 92
    Width = 117
    Height = 25
    Caption = 'Add Autofilter'
    TabOrder = 6
    OnClick = Button2Click
  end
  object btnAddCol: TButton
    Left = 16
    Top = 142
    Width = 117
    Height = 25
    Caption = 'Add column'
    Enabled = False
    TabOrder = 8
    OnClick = btnAddColClick
  end
  object cbCondition1: TComboBox
    Left = 216
    Top = 144
    Width = 69
    Height = 21
    Style = csDropDownList
    Enabled = False
    TabOrder = 10
    Items.Strings = (
      '='
      '>'
      '>='
      '<'
      '<='
      '<>')
  end
  object edValue1: TEdit
    Left = 291
    Top = 144
    Width = 121
    Height = 21
    Enabled = False
    TabOrder = 11
  end
  object edValue2: TEdit
    Left = 571
    Top = 144
    Width = 125
    Height = 21
    Enabled = False
    TabOrder = 14
  end
  object lvColumns: TListView
    Left = 148
    Top = 171
    Width = 585
    Height = 114
    Anchors = [akLeft, akTop, akRight, akBottom]
    Columns = <
      item
        Caption = 'Column'
        Width = 67
      end
      item
        Caption = 'Condition 1'
        Width = 76
      end
      item
        Caption = 'Value 1'
        Width = 126
      end
      item
        Caption = 'Operator'
        Width = 78
      end
      item
        Caption = 'Condition 2'
        Width = 76
      end
      item
        Caption = 'Value 2'
        Width = 125
      end>
    TabOrder = 18
    ViewStyle = vsReport
    OnClick = lvColumnsClick
  end
  object cbAND: TCheckBox
    Left = 422
    Top = 146
    Width = 63
    Height = 17
    Caption = 'AND op'
    Checked = True
    Enabled = False
    State = cbChecked
    TabOrder = 12
  end
  object btnDelCol: TButton
    Left = 16
    Top = 173
    Width = 117
    Height = 25
    Caption = 'Delete column'
    Enabled = False
    TabOrder = 15
    OnClick = btnDelColClick
  end
  object btnRead: TButton
    Left = 16
    Top = 8
    Width = 117
    Height = 25
    Caption = 'Read'
    TabOrder = 0
    OnClick = btnReadClick
  end
  object btnWrite: TButton
    Left = 16
    Top = 44
    Width = 117
    Height = 25
    Caption = 'Write'
    TabOrder = 3
    OnClick = btnWriteClick
  end
  object edFilenameRead: TEdit
    Left = 148
    Top = 10
    Width = 328
    Height = 21
    TabOrder = 1
  end
  object edFilenameWrite: TEdit
    Left = 148
    Top = 46
    Width = 328
    Height = 21
    TabOrder = 4
  end
  object Button6: TButton
    Left = 482
    Top = 8
    Width = 31
    Height = 25
    Caption = '...'
    TabOrder = 2
    OnClick = Button6Click
  end
  object Button7: TButton
    Left = 482
    Top = 44
    Width = 31
    Height = 25
    Caption = '...'
    TabOrder = 5
    OnClick = Button7Click
  end
  object cbColumn: TComboBox
    Left = 156
    Top = 144
    Width = 54
    Height = 21
    Style = csDropDownList
    Enabled = False
    TabOrder = 9
    Items.Strings = (
      'A'
      'B'
      'C'
      'D'
      'E')
  end
  object cbSheet: TComboBox
    Left = 182
    Top = 94
    Width = 95
    Height = 21
    Style = csDropDownList
    TabOrder = 7
  end
  object cbCondition2: TComboBox
    Left = 496
    Top = 144
    Width = 69
    Height = 21
    Style = csDropDownList
    Enabled = False
    TabOrder = 13
    Items.Strings = (
      '='
      '>'
      '>='
      '<'
      '<='
      '<>')
  end
  object btnApply: TButton
    Left = 16
    Top = 220
    Width = 117
    Height = 25
    Caption = 'Apply filter'
    TabOrder = 16
    OnClick = btnApplyClick
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 540
    Top = 12
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files|*.xlsx;*.xlsm;*.xls|All files|*.*'
    Left = 580
    Top = 12
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files|*.xlsx;*.xlsm|All files|*.*'
    Left = 624
    Top = 12
  end
end
