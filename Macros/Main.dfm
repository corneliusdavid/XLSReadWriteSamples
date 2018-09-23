object frmMain: TfrmMain
  Left = 728
  Top = 185
  Caption = 'Macro Sample'
  ClientHeight = 554
  ClientWidth = 747
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 137
    Top = 89
    Height = 465
    ExplicitHeight = 458
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 747
    Height = 89
    Align = alTop
    TabOrder = 0
    DesignSize = (
      747
      89)
    object Label1: TLabel
      Left = 163
      Top = 8
      Width = 42
      Height = 13
      Caption = 'Filename'
    end
    object btnClose: TButton
      Left = 4
      Top = 5
      Width = 49
      Height = 21
      Caption = 'Close'
      TabOrder = 0
      OnClick = btnCloseClick
    end
    object btnWrite: TButton
      Left = 59
      Top = 5
      Width = 49
      Height = 21
      Caption = 'Write'
      TabOrder = 1
      OnClick = btnWriteClick
    end
    object btnRead: TButton
      Left = 108
      Top = 5
      Width = 49
      Height = 21
      Caption = 'Read'
      TabOrder = 2
      OnClick = btnReadClick
    end
    object edFilename: TEdit
      Left = 207
      Top = 5
      Width = 514
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 3
      Text = 'D:\XTemp\XLS\Backup_ABCbil.xls'
    end
    object btnFilename: TButton
      Left = 724
      Top = 5
      Width = 23
      Height = 21
      Anchors = [akTop, akRight]
      Caption = '...'
      TabOrder = 4
      OnClick = btnFilenameClick
    end
    object Memo1: TMemo
      Left = 4
      Top = 32
      Width = 529
      Height = 51
      BevelInner = bvNone
      BevelOuter = bvNone
      BorderStyle = bsNone
      Color = clBtnFace
      Lines.Strings = (
        
          'NOTE! There are neither any syntax check, nor any compilations o' +
          'f the macros, so make sure that the code '
        'you writes will run correctly.'
        
          'Excel will compile the macros when the file is opened. If there ' +
          'are any errors, there will be an error message.')
      TabOrder = 5
    end
  end
  object Panel3: TPanel
    Left = 140
    Top = 89
    Width = 607
    Height = 465
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    ExplicitHeight = 449
    DesignSize = (
      607
      465)
    object memModule: TMemo
      Left = 3
      Top = 25
      Width = 603
      Height = 446
      Anchors = [akLeft, akTop, akRight, akBottom]
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Courier New'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssBoth
      TabOrder = 0
      ExplicitHeight = 430
    end
    object Button1: TButton
      Left = 3
      Top = 6
      Width = 162
      Height = 19
      Caption = 'Update module source code'
      TabOrder = 1
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 176
      Top = 6
      Width = 109
      Height = 19
      Caption = 'Read from file...'
      TabOrder = 2
      OnClick = Button2Click
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 89
    Width = 137
    Height = 465
    Align = alLeft
    BevelOuter = bvNone
    TabOrder = 2
    ExplicitHeight = 449
    DesignSize = (
      137
      465)
    object Label2: TLabel
      Left = 4
      Top = 3
      Width = 39
      Height = 13
      Caption = 'Modules'
    end
    object lbModules: TListBox
      Left = 0
      Top = 25
      Width = 134
      Height = 420
      Anchors = [akLeft, akTop, akRight, akBottom]
      ItemHeight = 13
      TabOrder = 0
      OnDblClick = lbModulesDblClick
      ExplicitHeight = 404
    end
    object btnOpenMod: TButton
      Left = 4
      Top = 453
      Width = 37
      Height = 18
      Anchors = [akLeft, akBottom]
      Caption = 'Open'
      TabOrder = 1
      OnClick = btnOpenModClick
      ExplicitTop = 437
    end
    object btnNewMod: TButton
      Left = 47
      Top = 453
      Width = 37
      Height = 18
      Anchors = [akLeft, akBottom]
      Caption = 'New'
      TabOrder = 2
      OnClick = btnNewModClick
      ExplicitTop = 437
    end
    object btnDeleteMod: TButton
      Left = 90
      Top = 453
      Width = 37
      Height = 18
      Anchors = [akLeft, akBottom]
      Caption = 'Delete'
      TabOrder = 3
      OnClick = btnDeleteModClick
      ExplicitTop = 437
    end
  end
  object dlgOpen: TOpenDialog
    Filter = 'Excel files (*.xls)|*.xls|All files (*.*)|*.*'
    Left = 532
    Top = 64
  end
  object dlgOpenTxt: TOpenDialog
    Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    Left = 596
    Top = 56
  end
  object XLS: TXLSReadWriteII5
    ComponentVersion = '6.00.43'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 660
    Top = 56
  end
end
