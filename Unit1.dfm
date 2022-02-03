object mainform: Tmainform
  Left = 686
  Top = 309
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'PDSAP - 2022-01'
  ClientHeight = 513
  ClientWidth = 519
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  GlassFrame.Enabled = True
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = FormCreate
  DesignSize = (
    519
    513)
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 502
    Height = 65
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Seznam staveb'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 497
    DesignSize = (
      502
      65)
    object ComboBox1: TComboBox
      Left = 16
      Top = 24
      Width = 447
      Height = 26
      Margins.Left = 30
      Margins.Top = 30
      Margins.Right = 30
      Margins.Bottom = 30
      Style = csOwnerDrawVariable
      Anchors = [akLeft, akTop, akRight]
      DropDownCount = 25
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Tahoma'
      Font.Style = []
      ItemHeight = 20
      ParentFont = False
      TabOrder = 0
      OnSelect = ComboBox1Select
      ExplicitWidth = 442
    end
    object BitBtn1: TBitBtn
      Left = 469
      Top = 24
      Width = 25
      Height = 25
      Glyph.Data = {
        0E060000424D0E06000000000000360000002800000016000000160000000100
        180000000000D8050000C40E0000C40E00000000000000000000FFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC5C5C55858580808080000000000000808
        08585858C5C5C5FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000FFFF
        FFFFFFFFFFFFFFFFFFFFE3E3E37A7A7A0707071E1E1E797979DDDDDDFFFFFFFF
        FFFFDDDDDD7979791D1D1D060606797979E3E3E3FFFFFFFFFFFFFFFFFFFFFFFF
        0000FFFFFFFFFFFFFFFFFFC2C2C23131313E3E3EC4C4C4FFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC4C4C43D3D3D313131C2C2C2FFFFFFFFFF
        FFFFFFFF0000FFFFFFFFFFFFC2C2C22B2B2B626262F7F7F7FFFFFFFFFFFFF7F7
        F7E4E4E4E1E1E1E1E1E1E5E5E5F7F7F7FFFFFFFFFFFFF7F7F76262622B2B2BC2
        C2C2FFFFFFFFFFFF0000FFFFFFE4E4E4323232626262FCFCFCFFFFFFFDFDFDCA
        CACA8E8E8E5D5D5D6767676767675F5F5F8F8F8FCACACAFDFDFDFFFFFFFCFCFC
        626262323232E4E4E4FFFFFF0000FFFFFF7A7A7A3D3D3DF6F6F6FFFFFFEDEDED
        9C9C9C5656566A6A6A878787AAAAAAAAAAAA878787696969555555A1A1A1F3F3
        F3FFFFFFF7F7F73D3D3D7A7A7AFFFFFF0000FFFFFF070707C5C5C5FFFFFFFBFB
        FBA1A1A1525252A0A0A0E5E5E5FFFFFFFFFFFFFFFFFFFFFFFFE5E5E59F9F9F52
        5252A2A2A2FDFDFDFFFFFFC5C5C5070707FFFFFF0000C5C5C51D1D1DFFFFFFFF
        FFFFF9F9F9C7C7C7C4C4C4FBFBFBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFF9F9F9F565656CBCBCBFFFFFFFFFFFF1D1D1DC5C5C500005757577B7B7B
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFECECEC6A6A6A878787FEFEFEFFFFFF7D7D7D58585800000707
        07DFDFDFFFFFFFFFFFFFDDDDDDE7E7E7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFF0F0F0D9D9D99898989D9D9DB4B4B4B1B1B1DDDDDD0B0B0B
        0000000000FFFFFFFFFFFF9A9A9A0000005E5E5EFDFDFDFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFDBDBDB2D2D2D0E0E0E242424000000989898FFFF
        FF0000000000000000FFFFFF909090000000303030000000313131DDDDDDFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F8F85353530000009D9D9DFF
        FFFFFFFFFF00000000000B0B0BDBDBDBA3A3A3B5B5B59C9C9C919191D6D6D6F4
        F4F4FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEAEAEADFDFDF
        FFFFFFFFFFFFDEDEDE08080800005757577E7E7EFFFFFFFCFCFC8787876D6D6D
        EDEDEDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFF7A7A7A5858580000C5C5C51D1D1DFFFFFFFFFFFFCBCB
        CB5555559F9F9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFBFBFBC6
        C6C6CACACAF9F9F9FFFFFFFFFFFF1E1E1EC5C5C50000FFFFFF070707C6C6C6FF
        FFFFFDFDFDA1A1A1525252A0A0A0E5E5E5FFFFFFFFFFFFFFFFFFFFFFFFE5E5E5
        A0A0A0545454A3A3A3FBFBFBFFFFFFC5C5C5070707FFFFFF0000FFFFFF797979
        3E3E3EF7F7F7FFFFFFF2F2F2A1A1A1565656696969878787AAAAAAAAAAAA8787
        876A6A6A5555559C9C9CECECECFFFFFFF7F7F73D3D3D7A7A7AFFFFFF0000FFFF
        FFE3E3E3323232626262FCFCFCFFFFFFFDFDFDCACACA8D8D8D5D5D5D67676767
        67675E5E5E8E8E8ECACACAFEFEFEFFFFFFFCFCFC626262323232E4E4E4FFFFFF
        0000FFFFFFFFFFFFC2C2C22B2B2B636363F7F7F7FFFFFFFFFFFFF7F7F7E4E4E4
        E1E1E1E1E1E1E4E4E4F7F7F7FFFFFFFFFFFFF7F7F76262622B2B2BC2C2C2FFFF
        FFFFFFFF0000FFFFFFFFFFFFFFFFFFC2C2C23131313E3E3EC4C4C4FFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC4C4C43E3E3E313131C2C2C2FF
        FFFFFFFFFFFFFFFF0000FFFFFFFFFFFFFFFFFFFFFFFFE3E3E37979790707071E
        1E1E797979DDDDDDFFFFFFFFFFFFDDDDDD7979791E1E1E070707797979E3E3E3
        FFFFFFFFFFFFFFFFFFFFFFFF0000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFC5C5C5585858080808000000000000080808585858C5C5C5FFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000}
      TabOrder = 1
      OnClick = BitBtn1Click
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 79
    Width = 502
    Height = 250
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Z'#225'kladn'#237' '#250'daje PD'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentBackground = False
    ParentColor = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 497
    DesignSize = (
      502
      250)
    object Label1: TLabel
      Left = 16
      Top = 77
      Width = 33
      Height = 18
      Caption = #268#237'slo:'
    end
    object Label2: TLabel
      Left = 16
      Top = 109
      Width = 46
      Height = 18
      Caption = 'N'#225'zev:'
    end
    object Label3: TLabel
      Left = 16
      Top = 141
      Width = 39
      Height = 18
      Caption = 'Obec:'
    end
    object Label4: TLabel
      Left = 16
      Top = 173
      Width = 52
      Height = 18
      Caption = 'Katastr:'
    end
    object Label5: TLabel
      Left = 176
      Top = 15
      Width = 152
      Height = 13
      Caption = 'kliknut'#237'm zkop'#237'ruje'#353' do schr'#225'nky'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 16
      Top = 45
      Width = 76
      Height = 18
      Caption = 'Cel'#253' n'#225'zev:'
    end
    object Label7: TLabel
      Left = 16
      Top = 205
      Width = 74
      Height = 18
      Caption = 'Term'#237'n PD:'
    end
    object Edit1: TEdit
      Left = 120
      Top = 74
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ReadOnly = True
      ShowHint = True
      TabOrder = 0
      TextHint = #268#237'slo stavby'
      OnClick = Edit1Click
      ExplicitWidth = 361
    end
    object Edit2: TEdit
      Left = 120
      Top = 106
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ReadOnly = True
      ShowHint = True
      TabOrder = 1
      TextHint = 'N'#225'zev stavby'
      OnClick = Edit2Click
      ExplicitWidth = 361
    end
    object Edit3: TEdit
      Left = 120
      Top = 138
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      TextHint = 'M'#237'sto stavby / obec'
      ExplicitWidth = 361
    end
    object Edit4: TEdit
      Left = 120
      Top = 170
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      TextHint = 'Katastr'#225'ln'#237' '#250'zem'#237
      ExplicitWidth = 361
    end
    object Edit5: TEdit
      Left = 120
      Top = 42
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ReadOnly = True
      ShowHint = True
      TabOrder = 4
      TextHint = #268#237'slo stavby'
      OnClick = Edit5Click
      ExplicitWidth = 361
    end
    object Edit6: TEdit
      Left = 120
      Top = 202
      Width = 366
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ReadOnly = True
      ShowHint = True
      TabOrder = 5
      TextHint = 'N'#225'zev stavby'
      OnClick = Edit2Click
      ExplicitWidth = 361
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 335
    Width = 503
    Height = 106
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Dal'#353#237' dokumenty PD'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    ExplicitWidth = 498
    object Button4: TButton
      Left = 17
      Top = 34
      Width = 75
      Height = 25
      Caption = 'DUR'
      TabOrder = 0
      OnClick = Button4Click
    end
    object Button5: TButton
      Left = 98
      Top = 34
      Width = 75
      Height = 25
      Caption = 'DPS'
      TabOrder = 1
      OnClick = Button5Click
    end
    object Button6: TButton
      Left = 176
      Top = 34
      Width = 75
      Height = 25
      Caption = 'POV'
      TabOrder = 2
      OnClick = Button6Click
    end
    object Button8: TButton
      Left = 257
      Top = 65
      Width = 75
      Height = 25
      Caption = 'Tabulky'
      TabOrder = 3
      OnClick = Button8Click
    end
    object Button9: TButton
      Left = 338
      Top = 65
      Width = 75
      Height = 25
      Caption = 'NN Geo'
      TabOrder = 4
      OnClick = Button9Click
    end
    object Button10: TButton
      Left = 419
      Top = 65
      Width = 75
      Height = 25
      Caption = 'VN Geo'
      TabOrder = 5
      OnClick = Button10Click
    end
    object Button11: TButton
      Left = 419
      Top = 34
      Width = 75
      Height = 25
      Caption = 'CD'
      TabOrder = 6
      OnClick = Button11Click
    end
    object Button12: TButton
      Left = 257
      Top = 34
      Width = 75
      Height = 25
      Caption = 'Desky'
      TabOrder = 7
      OnClick = Button12Click
    end
    object Button13: TButton
      Left = 338
      Top = 34
      Width = 75
      Height = 25
      Caption = 'Listy'
      TabOrder = 8
      OnClick = Button13Click
    end
    object Button16: TButton
      Left = 17
      Top = 65
      Width = 75
      Height = 25
      Caption = 'DUR odstr'
      TabOrder = 9
      OnClick = Button16Click
    end
    object Button15: TButton
      Left = 176
      Top = 65
      Width = 75
      Height = 25
      Caption = 'Opdady'
      TabOrder = 10
      OnClick = Button15Click
    end
  end
  object Button1: TButton
    Left = 8
    Top = 479
    Width = 503
    Height = 26
    Caption = 'Nov'#225' stavba'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 8
    Top = 447
    Width = 139
    Height = 26
    Caption = 'Otev'#345#237't slo'#382'ku PD'
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 153
    Top = 447
    Width = 183
    Height = 26
    Caption = 'Otev'#345#237't slo'#382'ku PD na s'#237'ti'
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    OnClick = Button3Click
  end
  object Button14: TButton
    Left = 342
    Top = 447
    Width = 169
    Height = 26
    Caption = 'Tabulka PD'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = Button14Click
  end
  object FileOpenDialog1_old: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <
      item
        DisplayName = 'PDF'
        FileMask = '*.pdf'
      end
      item
        DisplayName = 'XLSx'
        FileMask = '*.xlsx'
      end
      item
        DisplayName = 'XLS'
        FileMask = '*.xls'
      end>
    Options = []
    Left = 376
    Top = 216
  end
  object ExcelApplication1: TExcelApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 376
    Top = 175
  end
  object TrayIcon1: TTrayIcon
    PopupMenu = PopupMenu1
    OnClick = TrayIcon1Click
    OnDblClick = TrayIcon1DblClick
    Left = 376
    Top = 96
  end
  object ApplicationEvents1: TApplicationEvents
    OnMinimize = ApplicationEvents1Minimize
    Left = 376
    Top = 135
  end
  object FileOpenDialog1: TOpenDialog
    Options = []
    Left = 376
    Top = 271
  end
  object Taskbar1: TTaskbar
    TaskBarButtons = <>
    ProgressState = Paused
    ProgressMaxValue = 100
    ProgressValue = 100
    TabProperties = []
    Left = 376
    Top = 327
  end
  object PopupMenu1: TPopupMenu
    Left = 216
    Top = 311
    object Zobrazitaplikaci1: TMenuItem
      Caption = 'Zobrazit aplikaci'
      OnClick = Zobrazitaplikaci1Click
    end
    object Zobrazitseznamstaveb1: TMenuItem
      Caption = 'Zobrazit seznam staveb'
      OnClick = Zobrazitseznamstaveb1Click
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Konec1: TMenuItem
      Caption = 'Konec'
      OnClick = Konec1Click
    end
  end
end
