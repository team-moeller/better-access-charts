Version =20
VersionRequired =20
PublishOption =1
Checksum =-428580874
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6803
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =11745
    Top =300
    Right =18555
    Bottom =5970
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x0b9b3c5fb88ee540
    End
    GUID = Begin
        0x8fbfb473ad5a95409cc09d846ace2264
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =851
            BackColor =14347005
            Name ="Formularkopf"
            GUID = Begin
                0x158bafc97aa438439162a6ab5b99481a
            End
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =9
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =570
                    Top =165
                    Width =4140
                    Height =630
                    FontSize =24
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblHeadline"
                    Caption ="Better Access Charts"
                    GUID = Begin
                        0x28737549c63605478d6498941e9bb75c
                    End
                    GridlineColor =10921638
                End
                Begin Label
                    OverlapFlags =85
                    Left =5010
                    Top =390
                    Width =1485
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblVersion"
                    Caption ="Version x.xx.xx"
                    GUID = Begin
                        0xab82bb028b84c345b8a5fc8d84059457
                    End
                    GridlineColor =10921638
                End
            End
        End
        Begin Section
            Height =3968
            Name ="Detailbereich"
            GUID = Begin
                0xa0b25f451006a74a9c62272f5d3e2509
            End
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =2265
                    Top =3060
                    Width =2286
                    Height =613
                    ForeColor =4210752
                    Name ="cmdDemo"
                    Caption ="Open Demo"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x827b6ae5d48568458b90b656fbdef127
                    End
                    GridlineColor =10921638
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin Label
                    OverlapFlags =85
                    Left =288
                    Top =340
                    Width =6240
                    Height =2490
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblInfo"
                    Caption ="Microsoft Access urgently needs modern charts. The original charts in MS Access "
                        "are from the 90s of the previous century. Microsoft has given the charts in Acce"
                        "ss a lift. They call it \"Modern Charts\".\015\012There are many solutions for c"
                        "harts based on Java Script available on the web. This project makes use of this."
                        " We create charts using the Chart.js library. We display these in the web browse"
                        "r control. We hide the whole logic in a class module.\015\012Take a look at the "
                        "demo and let yourself be inspired by the possibilities."
                    GUID = Begin
                        0x440c4618f80864438f6e4d861f599bfc
                    End
                    GridlineColor =10921638
                End
            End
        End
        Begin FormFooter
            Height =851
            BackColor =14347005
            Name ="Formularfuß"
            GUID = Begin
                0x786aa0223c250a418658fa3d6510c14d
            End
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =9
            BackTint =20.0
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =345
                    Width =6180
                    Height =330
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblLink"
                    Caption ="https://github.com/team-moeller/better-access-charts"
                    HyperlinkAddress ="https://github.com/team-moeller/better-access-charts"
                    GUID = Begin
                        0x27768a9d14fa8249b6b5e2d9afcb3433
                    End
                    GridlineColor =10921638
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Load()
    Me.lblVersion.Caption = "Version: " & DMax("V_Number", "tbl_VersionHistory")
End Sub
Private Sub cmdDemo_Click()
    DoCmd.Close acForm, Me.Name
    DoEvents
    DoCmd.OpenForm "frm_Demo"
End Sub
