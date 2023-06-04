object MainForm: TMainForm
  Left = 194
  Top = 111
  Caption = #1056#1077#1076#1072#1082#1090#1086#1088' '#1090#1072#1073#1083#1080#1094
  ClientHeight = 434
  ClientWidth = 671
  Color = clAppWorkSpace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clBlack
  Font.Height = -11
  Font.Name = 'Default'
  Font.Style = []
  FormStyle = fsMDIForm
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  WindowMenu = Window1
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object BottomPanel: TPanel
    Left = 0
    Top = 388
    Width = 671
    Height = 46
    Align = alBottom
    AutoSize = True
    TabOrder = 0
    ExplicitTop = 368
    object Pages: TPageControl
      Left = 1
      Top = 1
      Width = 669
      Height = 25
      Align = alBottom
      Images = DataMod.ButtonImages
      Style = tsFlatButtons
      TabOrder = 0
      TabWidth = 100
      OnChange = PagesChange
    end
    object StatusBar: TStatusBar
      Left = 1
      Top = 26
      Width = 669
      Height = 19
      AutoHint = True
      Panels = <>
      SimplePanel = True
    end
  end
  object ActionToolBar1: TActionToolBar
    Left = 0
    Top = 0
    Width = 671
    Height = 26
    ActionManager = ActionManager1
    Caption = 'ActionToolBar1'
    ColorMap.HighlightColor = clWhite
    ColorMap.BtnSelectedColor = clBtnFace
    ColorMap.UnusedColor = clWhite
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Default'
    Font.Style = []
    ParentFont = False
    Spacing = 0
  end
  object MainMenu1: TMainMenu
    Left = 8
    Top = 32
    object File1: TMenuItem
      Caption = #1060#1072#1081#1083
      object FileExitItem: TMenuItem
        Action = FileExit1
      end
    end
    object N1: TMenuItem
      Caption = #1054'&'#1090#1095#1077#1090#1099
      Hint = #1057#1090#1072#1090#1080#1089#1090#1080#1095#1077#1089#1082#1080#1077' '#1086#1090#1095#1077#1090#1099
      object N3: TMenuItem
        Action = MainReport
      end
      object N4: TMenuItem
        Action = MonthFoProduction
      end
      object N5: TMenuItem
        Action = MonthForCode
      end
    end
    object Window1: TMenuItem
      Caption = '&'#1054#1082#1085#1086
      Hint = 'Window related commands'
      object WindowCascadeItem: TMenuItem
        Action = WindowCascade2
      end
      object WindowTileItem: TMenuItem
        Action = WindowTileHorizontal2
      end
      object WindowTileItem2: TMenuItem
        Action = WindowTileVertical2
      end
      object WindowMinimizeItem: TMenuItem
        Action = WindowMinimizeAll2
      end
      object WindowArrangeItem: TMenuItem
        Action = WindowArrange1
      end
      object N6: TMenuItem
        Caption = '-'
      end
      object Close1: TMenuItem
        Action = WindowClose1
      end
    end
    object Help1: TMenuItem
      Caption = '&'#1055#1086#1084#1086#1097#1100
      Hint = 'Help topics'
      object HelpAboutItem: TMenuItem
        Caption = '&'#1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
        Hint = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
        ImageIndex = 66
        OnClick = HelpContents1Execute
      end
      object N2: TMenuItem
        Caption = #1059#1076#1072#1083#1080#1090#1100' '#1087#1088#1086#1073#1077#1083#1099' '#1074' '#1085#1072#1079#1074#1072#1085#1080#1103#1093' '#1084#1086#1076#1077#1083#1080
        OnClick = N2Click
      end
    end
  end
  object ActionManager1: TActionManager
    ActionBars = <
      item
        Items.CaptionOptions = coAll
        Items = <>
      end
      item
        Items.CaptionOptions = coAll
        Items = <
          item
            Action = MainReport
            ImageIndex = 25
          end
          item
            Action = MonthFoProduction
            ImageIndex = 25
          end
          item
            Action = MonthForCode
            ImageIndex = 25
          end>
        ActionBar = ActionToolBar1
      end>
    Images = DataMod.ButtonImages
    Left = 40
    Top = 32
    StyleName = 'XP Style'
    object MainReport: TAction
      Category = 'TableEditors'
      Caption = #1055#1088#1086#1080#1079#1074#1086#1083#1100#1085#1099#1081' '#1086#1090#1095#1077#1090' '#1087#1086' '#1079#1072#1087#1080#1089#1103#1084
      ImageIndex = 25
      OnExecute = MainReportExecute
    end
    object MonthFoProduction: TAction
      Category = 'TableEditors'
      Caption = #1055#1086#1084#1077#1089#1103#1095#1085#1086' '#1087#1086' '#1087#1088#1086#1076#1091#1082#1094#1080#1080
      ImageIndex = 25
      OnExecute = MonthFoProductionExecute
    end
    object MonthForCode: TAction
      Category = 'TableEditors'
      Caption = #1055#1086#1084#1077#1089#1103#1095#1085#1086' '#1087#1086' '#1082#1086#1076#1072#1084
      ImageIndex = 25
      OnExecute = MonthForCodeExecute
    end
    object WindowClose1: TWindowClose
      Category = 'Window'
      Caption = #1047#1072#1082#1088#1099#1090#1100' '#1074#1089#1077
      Enabled = False
      Hint = #1047#1072#1082#1088#1099#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072
    end
    object WindowCascade2: TWindowCascade
      Category = 'Window'
      Caption = '&'#1050#1072#1089#1082#1072#1076#1086#1084
      Enabled = False
      Hint = #1056#1072#1089#1087#1086#1083#1086#1078#1080#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072' '#1082#1072#1089#1082#1072#1076#1086#1084
      ImageIndex = 61
    end
    object WindowTileHorizontal2: TWindowTileHorizontal
      Category = 'Window'
      Caption = #1043#1086#1088#1080#1079#1086#1085#1090#1072#1083#1100#1085#1086
      Enabled = False
      Hint = #1056#1072#1089#1087#1086#1083#1086#1078#1080#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072' '#1075#1086#1088#1080#1079#1086#1085#1090#1072#1083#1100#1085#1086
      ImageIndex = 62
    end
    object WindowTileVertical2: TWindowTileVertical
      Category = 'Window'
      Caption = #1042#1077#1088#1090#1080#1082#1072#1083#1100#1085#1086
      Enabled = False
      Hint = #1056#1072#1089#1087#1086#1083#1086#1078#1080#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072' '#1074#1077#1088#1090#1080#1082#1072#1083#1100#1085#1086
      ImageIndex = 63
    end
    object WindowMinimizeAll2: TWindowMinimizeAll
      Category = 'Window'
      Caption = #1057#1074#1077#1088#1085#1091#1090#1100' '#1074#1089#1077
      Enabled = False
      Hint = #1057#1074#1077#1088#1085#1091#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072
    end
    object WindowArrange1: TWindowArrange
      Category = 'Window'
      Caption = #1056#1072#1079#1074#1077#1088#1085#1091#1090#1100' '#1074#1089#1077
      Enabled = False
      Hint = #1056#1072#1079#1074#1077#1088#1085#1091#1090#1100' '#1074#1089#1077' '#1086#1082#1085#1072
    end
    object FileExit1: TFileExit
      Category = 'File'
      Caption = #1042#1099#1093#1086#1076
      Hint = #1047#1072#1082#1088#1099#1090#1100' '#1087#1088#1080#1083#1086#1078#1077#1085#1080#1077
      ImageIndex = 65
      OnHint = FileExit1Hint
    end
  end
end
