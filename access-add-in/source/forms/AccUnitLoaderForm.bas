Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10604
    DatasheetFontHeight =11
    ItemSuffix =154
    Left =6285
    Top =3915
    Right =14978
    Bottom =12750
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x212b6fd80e9ce340
    End
    Caption ="ACLib - AccUnit Loader"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =255
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =5142
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9908
                    Top =120
                    Width =570
                    Height =495
                    TabIndex =3
                    Name ="cmdSelectAccUnitDllPath"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad00000000000adada ,
                        0x003333333330adad0b03333333330ada0fb03333333330ad0bfb03333333330a ,
                        0x0fbfb000000000000bfbfbfbfb0adada0fbfbfbfbf0dadad0bfb0000000adada ,
                        0xa000adadadad000ddadadadadadad00aadadadad0dad0d0ddadadadad000dada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Select AccUnit directory of the dll files"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =9908
                    LayoutCachedTop =120
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =615
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2805
                    Top =120
                    Width =7035
                    Height =495
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    LeftMargin =29
                    Name ="txtAccUnitDllPath"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =120
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =615
                    ColumnStart =2
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =120
                            Width =2618
                            Height =495
                            FontSize =10
                            Name ="Label5"
                            Caption ="Location of AccUnit dll files:"
                            GroupTable =2
                            BottomPadding =150
                            GridlineWidthLeft =0
                            GridlineWidthTop =0
                            GridlineWidthRight =0
                            GridlineWidthBottom =0
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =2738
                            LayoutCachedHeight =615
                            ColumnEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    Name ="sysFirst"

                End
                Begin CommandButton
                    Transparent = NotDefault
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =10575
                    Width =29
                    Height =29
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Schließen"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10575
                    LayoutCachedWidth =10604
                    LayoutCachedHeight =29
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =1395
                    Width =4980
                    Height =405
                    TabIndex =5
                    Name ="cmdSetAccUnitTlbReferenz"
                    Caption ="Set reference to AccUnit.tlb"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =1395
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =1800
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =1988
                    Width =4980
                    Height =405
                    TabIndex =6
                    Name ="cmdRemoveAccUnitTlbReferenz"
                    Caption ="Remove reference to AccUnit.tlb"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =1988
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =2393
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =803
                    Width =4980
                    Height =405
                    TabIndex =4
                    Name ="cmdExportAccUnitFiles"
                    Caption ="Export DLL files from add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =803
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =1208
                    PictureCaptionArrangement =5
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =803
                    Width =1995
                    Height =405
                    Name ="EmptyCell73"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =803
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1208
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =1395
                    Width =2633
                    Height =405
                    Name ="EmptyCell78"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =1395
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =1800
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =1988
                    Width =1995
                    Height =405
                    Name ="EmptyCell83"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =1988
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =2393
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =4508
                    Width =5175
                    Height =465
                    TabIndex =10
                    Name ="cmdInsertFactoryModule"
                    Caption ="Insert/update AccUnit Factory module in application"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =4508
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =4973
                    PictureCaptionArrangement =5
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9908
                    Top =803
                    Width =570
                    Height =405
                    Name ="EmptyCell93"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9908
                    LayoutCachedTop =803
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =1208
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9908
                    Top =1988
                    Width =570
                    Height =405
                    Name ="EmptyCell95"
                    GroupTable =2
                    BottomPadding =86
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9908
                    LayoutCachedTop =1988
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =2393
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9908
                    Top =4508
                    Width =570
                    Height =465
                    TabIndex =11
                    Name ="cmdOpenMenu"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="More commands ..."
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =9908
                    LayoutCachedTop =4508
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =4973
                    PictureCaptionArrangement =5
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9908
                    Top =3750
                    Width =570
                    Height =570
                    TabIndex =9
                    Name ="cmdUserSettings"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="AccUnit Settings"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727287727272f3727272f3 ,
                        0x72727287000000000000000000000000727272b7727272b70000000000000000 ,
                        0x0000000000000000000000000000000000000000727272fc727272ff727272ff ,
                        0x727272f9000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000072727203727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ba727272ff727272ff ,
                        0x7272729c000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x0000000000000000727272067272728d727272ff727272ff7272728d72727206 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x000000000000000072727287727272ff727272ff727272ff727272ff7272728a ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x0000000000000000727272ea727272ff727272bd727272bd727272ff727272ea ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff00000000727272ea727272ff0000000000000000727272ff727272ea ,
                        0x000000000000000000000000000000000000000072727230727272ff727272ff ,
                        0x727272300000000072727287727272ff0000000000000000727272ff72727287 ,
                        0x00000000000000000000000000000000000000007272720f727272ff727272ff ,
                        0x7272720f00000000727272067272728400000000000000007272728472727206 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9908
                    LayoutCachedTop =3750
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =4320
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =3750
                    Width =562
                    Height =570
                    Name ="EmptyCell113"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =3750
                    LayoutCachedWidth =682
                    LayoutCachedHeight =4320
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =4508
                    Width =1995
                    Height =465
                    Name ="EmptyCell119"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =4508
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =4973
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =5355
                    Top =4508
                    Width =2430
                    Height =465
                    Name ="EmptyCell121"
                    GroupTable =2
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =5355
                    LayoutCachedTop =4508
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =4973
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =743
                    Top =3750
                    Width =8887
                    Height =570
                    Name ="labInfo"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =743
                    LayoutCachedTop =3750
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =4320
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9690
                    Top =3750
                    Width =150
                    Height =570
                    Name ="EmptyCell127"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9690
                    LayoutCachedTop =3750
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =4320
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =2573
                    Width =4980
                    Height =405
                    TabIndex =7
                    Name ="cmdInstallVbeAddIn"
                    Caption ="Install VBE Add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =2573
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =2978
                    PictureCaptionArrangement =5
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9690
                    Top =2573
                    Width =150
                    Height =405
                    Name ="EmptyCell135"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9690
                    LayoutCachedTop =2573
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =2978
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9908
                    Top =2573
                    Width =570
                    Height =405
                    Name ="EmptyCell136"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9908
                    LayoutCachedTop =2573
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =2978
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =2573
                    Width =1785
                    Height =405
                    Name ="EmptyCell137"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =150
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =2573
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =2978
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =3165
                    Width =4980
                    Height =405
                    TabIndex =8
                    Name ="cmdLoadVbeAddIn"
                    Caption ="Load VBE Add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0

                    LayoutCachedLeft =2805
                    LayoutCachedTop =3165
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =3570
                    PictureCaptionArrangement =5
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =9690
                    Top =3165
                    Width =150
                    Height =405
                    Name ="EmptyCell145"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9690
                    LayoutCachedTop =3165
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =3570
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9908
                    Top =3165
                    Width =570
                    Height =405
                    Name ="EmptyCell146"
                    GroupTable =2
                    BottomPadding =86
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =9908
                    LayoutCachedTop =3165
                    LayoutCachedWidth =10478
                    LayoutCachedHeight =3570
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7845
                    Top =3165
                    Width =1785
                    Height =405
                    Name ="EmptyCell147"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =3165
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =3570
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =743
                    Top =803
                    Width =1995
                    Height =1590
                    Name ="Label148"
                    Caption ="AccUnit (Framework)"
                    GroupTable =2
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =743
                    LayoutCachedTop =803
                    LayoutCachedWidth =2738
                    LayoutCachedHeight =2393
                    RowStart =1
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =743
                    Top =2573
                    Width =1995
                    Height =997
                    Name ="Label149"
                    Caption ="AccUnit VBE Add-in"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =86
                    GridlineStyleBottom =1
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =743
                    LayoutCachedTop =2573
                    LayoutCachedWidth =2738
                    LayoutCachedHeight =3570
                    RowStart =4
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =2573
                    Width =562
                    Height =997
                    Name ="EmptyCell151"
                    GroupTable =2
                    TopPadding =86
                    BottomPadding =86
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =2573
                    LayoutCachedWidth =682
                    LayoutCachedHeight =3570
                    RowStart =4
                    RowEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =803
                    Width =562
                    Height =1590
                    Name ="EmptyCell153"
                    GroupTable =2
                    BottomPadding =86
                    GridlineWidthLeft =0
                    GridlineWidthTop =0
                    GridlineWidthRight =0
                    GridlineWidthBottom =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =803
                    LayoutCachedWidth =682
                    LayoutCachedHeight =2393
                    RowStart =1
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
            End
        End
    End
End
CodeBehindForm
' See "AccUnitLoaderForm.cls"
