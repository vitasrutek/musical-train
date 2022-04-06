object Form5: TForm5
  Left = 0
  Top = 0
  Caption = 'Form5'
  ClientHeight = 447
  ClientWidth = 932
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    932
    447)
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 8
    Top = 8
    Width = 903
    Height = 401
    Anchors = [akLeft, akTop, akRight, akBottom]
    DefaultColWidth = 150
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing, goFixedRowDefAlign]
    TabOrder = 0
  end
  object Button1: TButton
    Left = 8
    Top = 415
    Width = 89
    Height = 25
    Caption = 'nacist tabulku'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 103
    Top = 415
    Width = 113
    Height = 25
    Caption = 'termin PD'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 224
    Top = 416
    Width = 75
    Height = 25
    Caption = 'cena PD'
    TabOrder = 3
  end
  object Button4: TButton
    Left = 312
    Top = 416
    Width = 75
    Height = 25
    Caption = 'technik'
    TabOrder = 4
  end
end
