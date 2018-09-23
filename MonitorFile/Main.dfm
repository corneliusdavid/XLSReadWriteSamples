object frmMain: TfrmMain
  Left = 563
  Top = 129
  Caption = 'Monitor File Sample'
  ClientHeight = 618
  ClientWidth = 949
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 949
    Height = 57
    Align = alTop
    TabOrder = 0
    ExplicitWidth = 738
    object Label1: TLabel
      Left = 268
      Top = 16
      Width = 66
      Height = 14
      Caption = 'File to monitor'
    end
    object Label2: TLabel
      Left = 8
      Top = 38
      Width = 687
      Height = 14
      Caption = 
        'Open a file in Excel, select the same file here. Click Start.  D' +
        'o some changes in Excel and save.. The grid will be updated here' +
        ' at the same time.'
    end
    object Button1: TButton
      Left = 8
      Top = 10
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 0
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 104
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Start'
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 184
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Stop'
      TabOrder = 2
      OnClick = Button3Click
    end
    object edFilename: TEdit
      Left = 340
      Top = 12
      Width = 329
      Height = 22
      TabOrder = 3
    end
    object Button4: TButton
      Left = 675
      Top = 11
      Width = 31
      Height = 21
      Caption = '...'
      TabOrder = 4
      OnClick = Button4Click
    end
  end
  object XLSGrid: TXLSGrid
    Left = 0
    Top = 57
    Width = 949
    Height = 540
    HeaderColor = 16248036
    GridlineColor = 15062992
    Align = alClient
    ColCount = 32
    DefaultColWidth = 68
    DefaultRowHeight = 20
    RowCount = 255
    Options = [goFixedVertLine, goFixedHorzLine, goRangeSelect, goRowSizing, goColSizing, goEditing]
    TabOrder = 1
    ExplicitWidth = 738
    ExplicitHeight = 439
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
    Left = 96
    Top = 124
  end
end
