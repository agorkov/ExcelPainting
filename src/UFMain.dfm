object FMain: TFMain
  Left = 0
  Top = 0
  Caption = #1050#1072#1088#1090#1080#1085#1082#1072' '#1074' Excel'
  ClientHeight = 385
  ClientWidth = 288
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCanResize = FormCanResize
  PixelsPerInch = 96
  TextHeight = 13
  object Img: TImage
    Left = 8
    Top = 8
    Width = 273
    Height = 265
    Center = True
    Proportional = True
    Stretch = True
    OnDblClick = ImgDblClick
  end
  object BCreateExcel: TButton
    Left = 8
    Top = 311
    Width = 273
    Height = 25
    Caption = 'BCreateExcel'
    TabOrder = 0
    OnClick = BCreateExcelClick
  end
  object pbCurrentTask: TProgressBar
    Left = 8
    Top = 366
    Width = 273
    Height = 16
    TabOrder = 1
  end
  object pbAll: TProgressBar
    Left = 8
    Top = 342
    Width = 273
    Height = 16
    TabOrder = 2
  end
  object tbImageQuality: TTrackBar
    Left = 0
    Top = 279
    Width = 281
    Height = 26
    Max = 6
    PageSize = 1
    Position = 2
    TabOrder = 3
    OnChange = tbImageQualityChange
  end
  object OPD: TOpenPictureDialog
    Left = 8
    Top = 8
  end
end
