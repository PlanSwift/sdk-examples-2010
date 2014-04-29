object Form2: TForm2
  Left = 823
  Top = 386
  BorderStyle = bsDialog
  Caption = 'Report Type'
  ClientHeight = 95
  ClientWidth = 235
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  PixelsPerInch = 96
  TextHeight = 13
  object Button1: TButton
    Left = 39
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Ok'
    ModalResult = 1
    TabOrder = 0
  end
  object Button2: TButton
    Left = 120
    Top = 64
    Width = 75
    Height = 25
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 1
  end
  object ComboBox1: TComboBox
    Left = 39
    Top = 24
    Width = 156
    Height = 21
    TabOrder = 2
    Text = 'Digitized Items Only'
    Items.Strings = (
      'Digitized Items Only'
      'Parts Only'
      'Digitized Items W/Parts')
  end
end
