object Form1: TForm1
  Left = 754
  Top = 386
  BorderStyle = bsDialog
  Caption = 'Delphi: 2010 Microsoft Office Examples'
  ClientHeight = 171
  ClientWidth = 373
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 167
    Top = 101
    Width = 182
    Height = 26
    Caption = 
      'Demonstrates exporting basic job information to an MS Excel Work' +
      'sheet'
    WordWrap = True
  end
  object Label2: TLabel
    Left = 167
    Top = 16
    Width = 179
    Height = 26
    Caption = 
      'Demonstrates exporting basic job information to an MS Word Docum' +
      'ent'
    WordWrap = True
  end
  object Label3: TLabel
    Left = 167
    Top = 55
    Width = 166
    Height = 26
    Caption = 
      'Demonstrates exporting basic job information to an MS Outlook Em' +
      'ail'
    WordWrap = True
  end
  object MSWord: TButton
    Left = 8
    Top = 8
    Width = 153
    Height = 41
    Caption = 'MS Word Demo'
    TabOrder = 0
    OnClick = MSWordClick
  end
  object Button2: TButton
    Left = 8
    Top = 55
    Width = 153
    Height = 41
    Caption = 'MS Outlook Demo'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 8
    Top = 101
    Width = 153
    Height = 41
    Caption = 'MS Excel  Demo'
    TabOrder = 2
    OnClick = Button3Click
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 154
    Width = 373
    Height = 17
    Align = alBottom
    TabOrder = 3
    ExplicitTop = 243
  end
end
