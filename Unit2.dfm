object novyform: Tnovyform
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'Nov'#225' stavba'
  ClientHeight = 328
  ClientWidth = 658
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  DesignSize = (
    658
    328)
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox2: TGroupBox
    Left = 8
    Top = 8
    Width = 636
    Height = 202
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Z'#225'kladn'#237' '#250'daje PD'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    DesignSize = (
      636
      202)
    object Label1: TLabel
      Left = 16
      Top = 37
      Width = 33
      Height = 18
      Caption = #268#237'slo:'
    end
    object Label2: TLabel
      Left = 16
      Top = 69
      Width = 46
      Height = 18
      Caption = 'N'#225'zev:'
    end
    object Label3: TLabel
      Left = 16
      Top = 101
      Width = 39
      Height = 18
      Caption = 'Obec:'
    end
    object Label4: TLabel
      Left = 16
      Top = 133
      Width = 52
      Height = 18
      Caption = 'Katastr:'
    end
    object Label5: TLabel
      Left = 16
      Top = 165
      Width = 130
      Height = 18
      Caption = 'Popis (jednoduch'#253'):'
    end
    object Edit1: TEdit
      Left = 176
      Top = 34
      Width = 444
      Height = 26
      HelpType = htKeyword
      Anchors = [akLeft, akTop, akRight]
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      TextHint = #268#237'slo stavby'
    end
    object Edit2: TEdit
      Left = 176
      Top = 66
      Width = 444
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 1
      TextHint = 'N'#225'zev stavby'
      OnChange = Edit2Change
    end
    object Edit3: TEdit
      Left = 176
      Top = 98
      Width = 444
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 2
      TextHint = 'Obec, kde se stavba nach'#225'z'#237
    end
    object Edit4: TEdit
      Left = 176
      Top = 130
      Width = 444
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 3
      TextHint = 'Katastr'#225'ln'#237' '#250'zem'#237
    end
    object Edit5: TEdit
      Left = 176
      Top = 162
      Width = 444
      Height = 26
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 4
      TextHint = 'Jednoduch'#253' popis stavby: Jedn'#225' se o ........'
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 216
    Width = 636
    Height = 57
    Anchors = [akLeft, akTop, akRight]
    Caption = 'Dal'#353#237' dokumenty PD'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object CheckBox1: TCheckBox
      Left = 24
      Top = 24
      Width = 81
      Height = 17
      Caption = 'DUR'
      TabOrder = 0
    end
    object CheckBox2: TCheckBox
      Left = 152
      Top = 23
      Width = 81
      Height = 17
      Caption = 'DPS'
      TabOrder = 1
    end
  end
  object Button1: TButton
    Left = 361
    Top = 279
    Width = 137
    Height = 33
    Caption = 'Vytvo'#345'it a ulo'#382'it'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button3: TButton
    Left = 504
    Top = 279
    Width = 137
    Height = 34
    Caption = 'Konec bez ulo'#382'en'#237
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = Button3Click
  end
end
