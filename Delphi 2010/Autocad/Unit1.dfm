object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Delphi: Autocad Demo'
  ClientHeight = 310
  ClientWidth = 268
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 69
    Height = 13
    Caption = 'Select a Page:'
  end
  object PageCBX: TComboBox
    Left = 8
    Top = 27
    Width = 169
    Height = 21
    TabOrder = 0
  end
  object Button1: TButton
    Left = 102
    Top = 277
    Width = 75
    Height = 25
    Caption = 'Ok'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 185
    Top = 277
    Width = 75
    Height = 25
    Caption = 'Quit'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Processtxt: TMemo
    Left = 8
    Top = 54
    Width = 252
    Height = 217
    TabOrder = 3
  end
end
