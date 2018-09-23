object frmMain: TfrmMain
  Left = 563
  Top = 129
  Caption = 'XLSGrid sample'
  ClientHeight = 718
  ClientWidth = 1011
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1011
    Height = 41
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 172
      Top = 16
      Width = 42
      Height = 14
      Caption = 'Filename'
    end
    object Button1: TButton
      Left = 8
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 92
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Read'
      TabOrder = 1
      OnClick = Button2Click
    end
    object edFilename: TEdit
      Left = 220
      Top = 12
      Width = 341
      Height = 22
      TabOrder = 2
    end
    object Button3: TButton
      Left = 564
      Top = 12
      Width = 29
      Height = 21
      Caption = '...'
      TabOrder = 3
      OnClick = Button3Click
    end
  end
  object XLSGrid: TXLSGrid
    Left = 0
    Top = 41
    Width = 1011
    Height = 656
    HeaderColor = 16248036
    GridlineColor = 15062992
    Align = alClient
    ColCount = 32
    DefaultColWidth = 68
    DefaultRowHeight = 20
    RowCount = 255
    Options = [goFixedVertLine, goFixedHorzLine, goRangeSelect, goRowSizing, goColSizing, goEditing]
    TabOrder = 1
    ColWidths = (
      21
      63
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68
      68)
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files|*.xls;*.xlsx;*.xlsm|All files|*.*'
    Left = 44
    Top = 88
  end
end
