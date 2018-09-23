object frmEditName: TfrmEditName
  Left = 379
  Top = 285
  Width = 337
  Height = 250
  Caption = 'Edit Name'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 12
    Top = 16
    Width = 31
    Height = 13
    Caption = 'Name:'
  end
  object Label2: TLabel
    Left = 12
    Top = 40
    Width = 34
    Height = 13
    Caption = 'Scope:'
  end
  object Label3: TLabel
    Left = 12
    Top = 64
    Width = 47
    Height = 13
    Caption = 'Comment:'
  end
  object Label4: TLabel
    Left = 12
    Top = 156
    Width = 46
    Height = 13
    Caption = 'Refers to:'
  end
  object edName: TEdit
    Left = 72
    Top = 12
    Width = 229
    Height = 21
    TabOrder = 0
  end
  object cbScope: TComboBox
    Left = 72
    Top = 36
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 1
  end
  object memComment: TMemo
    Left = 72
    Top = 60
    Width = 229
    Height = 89
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object edRefersTo: TEdit
    Left = 72
    Top = 152
    Width = 233
    Height = 21
    TabOrder = 3
  end
  object Button1: TButton
    Left = 140
    Top = 180
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 4
  end
  object Button2: TButton
    Left = 228
    Top = 180
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 5
  end
end
