object frmMain: TfrmMain
  Left = 965
  Top = 592
  Caption = 'Create Charts'
  ClientHeight = 467
  ClientWidth = 643
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
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 176
    Top = 12
    Width = 42
    Height = 13
    Caption = 'Filename'
  end
  object shpChart: TShape
    Left = 10
    Top = 72
    Width = 608
    Height = 389
  end
  object shpPlotArea: TShape
    Left = 108
    Top = 132
    Width = 369
    Height = 285
  end
  object shpTitle: TShape
    Left = 108
    Top = 84
    Width = 369
    Height = 42
  end
  object Label3: TLabel
    Left = 25
    Top = 79
    Width = 70
    Height = 16
    Caption = 'ChartSpace'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label4: TLabel
    Left = 115
    Top = 132
    Width = 56
    Height = 16
    Caption = 'PlotArea'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Shape1: TShape
    Left = 483
    Top = 132
    Width = 109
    Height = 285
  end
  object lblLegendPos: TLabel
    Left = 495
    Top = 232
    Width = 37
    Height = 13
    Caption = 'Position'
    Enabled = False
  end
  object Label5: TLabel
    Left = 127
    Top = 162
    Width = 52
    Height = 13
    Caption = 'Chart type'
  end
  object Shape2: TShape
    Left = 26
    Top = 132
    Width = 83
    Height = 285
  end
  object Shape3: TShape
    Left = 108
    Top = 372
    Width = 369
    Height = 45
  end
  object Label2: TLabel
    Left = 32
    Top = 134
    Width = 49
    Height = 16
    Caption = 'ValAxis'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label6: TLabel
    Left = 115
    Top = 376
    Width = 35
    Height = 16
    Caption = 'CatAx'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label7: TLabel
    Left = 223
    Top = 136
    Width = 52
    Height = 13
    Caption = 'Chart data'
  end
  object Label8: TLabel
    Left = 127
    Top = 269
    Width = 38
    Height = 13
    Caption = 'Markers'
  end
  object lblTitleRef: TLabel
    Left = 318
    Top = 96
    Width = 52
    Height = 13
    Caption = 'Text or ref'
    Enabled = False
  end
  object Label9: TLabel
    Left = 30
    Top = 218
    Width = 72
    Height = 13
    Caption = 'Number format'
  end
  object edFilename: TEdit
    Left = 223
    Top = 7
    Width = 370
    Height = 21
    TabOrder = 0
    Text = 'd:\xtemp\wtest.xlsx'
  end
  object Button3: TButton
    Left = 597
    Top = 7
    Width = 21
    Height = 21
    Caption = '...'
    TabOrder = 1
    OnClick = Button3Click
  end
  object Button1: TButton
    Left = 7
    Top = 7
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button4: TButton
    Left = 97
    Top = 7
    Width = 75
    Height = 25
    Caption = 'Save'
    TabOrder = 3
    OnClick = Button4Click
  end
  object cbTitle: TCheckBox
    Left = 116
    Top = 96
    Width = 53
    Height = 17
    Caption = 'Title'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    OnClick = cbTitleClick
  end
  object btnTitleStyle: TButton
    Left = 170
    Top = 91
    Width = 66
    Height = 25
    Caption = 'Style...'
    Enabled = False
    TabOrder = 5
    OnClick = btnTitleStyleClick
  end
  object btnTitleText: TButton
    Left = 242
    Top = 91
    Width = 70
    Height = 25
    Caption = 'Font...'
    Enabled = False
    TabOrder = 6
    OnClick = btnTitleTextClick
  end
  object Button5: TButton
    Left = 127
    Top = 341
    Width = 75
    Height = 25
    Caption = 'Style...'
    TabOrder = 7
    OnClick = Button5Click
  end
  object cbLegend: TCheckBox
    Left = 491
    Top = 140
    Width = 97
    Height = 17
    Caption = 'Legend'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Courier New'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 8
    OnClick = cbLegendClick
  end
  object btnLegendStyle: TButton
    Left = 491
    Top = 168
    Width = 93
    Height = 25
    Caption = 'Style...'
    Enabled = False
    TabOrder = 9
    OnClick = btnLegendStyleClick
  end
  object btnLegendText: TButton
    Left = 491
    Top = 199
    Width = 93
    Height = 25
    Caption = 'Font...'
    Enabled = False
    TabOrder = 10
    OnClick = btnLegendTextClick
  end
  object cbLegendPos: TComboBox
    Left = 495
    Top = 251
    Width = 89
    Height = 21
    Style = csDropDownList
    Enabled = False
    TabOrder = 11
    Items.Strings = (
      'Right'
      'Left'
      'Top'
      'Bottom')
  end
  object cbChart: TComboBox
    Left = 127
    Top = 180
    Width = 75
    Height = 21
    Style = csDropDownList
    TabOrder = 12
    Items.Strings = (
      'Bar'
      'Line'
      'Area'
      'Bubble'
      'Doughnut'
      'Pie'
      'Radar'
      'Scatter')
  end
  object Button2: TButton
    Left = 35
    Top = 156
    Width = 66
    Height = 25
    Caption = 'Style...'
    TabOrder = 13
    OnClick = Button2Click
  end
  object Button6: TButton
    Left = 35
    Top = 187
    Width = 67
    Height = 25
    Caption = 'Font...'
    TabOrder = 14
    OnClick = Button6Click
  end
  object Button7: TButton
    Left = 230
    Top = 383
    Width = 66
    Height = 25
    Caption = 'Style...'
    TabOrder = 15
    OnClick = Button7Click
  end
  object Button8: TButton
    Left = 302
    Top = 383
    Width = 67
    Height = 25
    Caption = 'Font...'
    TabOrder = 16
    OnClick = Button8Click
  end
  object sgData: TStringGrid
    Left = 223
    Top = 156
    Width = 246
    Height = 125
    ColCount = 1
    DefaultColWidth = 50
    DefaultRowHeight = 16
    FixedCols = 0
    RowCount = 6
    FixedRows = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
    TabOrder = 17
    ColWidths = (
      50)
    RowHeights = (
      16
      16
      16
      16
      16
      16)
  end
  object btnAddSerie: TButton
    Left = 127
    Top = 207
    Width = 75
    Height = 25
    Caption = 'Add serie'
    TabOrder = 18
    OnClick = btnAddSerieClick
  end
  object BtnRemoveSerie: TButton
    Left = 127
    Top = 238
    Width = 75
    Height = 25
    Caption = 'Remove serie'
    Enabled = False
    TabOrder = 19
    OnClick = BtnRemoveSerieClick
  end
  object Button9: TButton
    Left = 8
    Top = 37
    Width = 164
    Height = 25
    Caption = 'Create the chart'
    TabOrder = 20
    OnClick = Button9Click
  end
  object cbMarkers: TComboBox
    Left = 127
    Top = 287
    Width = 75
    Height = 21
    Style = csDropDownList
    TabOrder = 21
    Items.Strings = (
      'None'
      'Circle'
      'Dash'
      'Diamond'
      'Dot'
      'Plus'
      'Square'
      'Star'
      'Triangle'
      'X')
  end
  object edTitleRef: TEdit
    Left = 376
    Top = 93
    Width = 93
    Height = 21
    Enabled = False
    TabOrder = 22
    Text = 'Sheet1!$D$1'
  end
  object Button10: TButton
    Left = 32
    Top = 101
    Width = 66
    Height = 25
    Caption = 'Style...'
    TabOrder = 23
    OnClick = Button10Click
  end
  object edValAxisNumFmt: TEdit
    Left = 29
    Top = 238
    Width = 78
    Height = 21
    TabOrder = 24
  end
  object cbValAxisLog: TCheckBox
    Left = 30
    Top = 265
    Width = 72
    Height = 17
    Caption = 'Log scale'
    TabOrder = 25
  end
  object btbPos: TButton
    Left = 222
    Top = 37
    Width = 223
    Height = 25
    Caption = 'Set size and position (after create)'
    TabOrder = 26
    OnClick = btbPosClick
  end
  object cbCatAxValues: TCheckBox
    Left = 115
    Top = 391
    Width = 103
    Height = 17
    Caption = 'Values (months)'
    TabOrder = 27
  end
  object dlgSave: TSaveDialog
    Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    Left = 512
    Top = 32
  end
  object dlgFont: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Left = 556
    Top = 32
  end
end
