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
    ItemSuffix =125
    Left =5640
    Top =3570
    Right =15765
    Bottom =10695
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x212b6fd80e9ce340
    End
    Caption ="ACLib - AccUnit Loader"
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
            Height =4025
            Name ="Detail"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9915
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

                    LayoutCachedLeft =9915
                    LayoutCachedTop =120
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =615
                    ColumnStart =5
                    ColumnEnd =5
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
                    Width =7043
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

                    LayoutCachedLeft =2805
                    LayoutCachedTop =120
                    LayoutCachedWidth =9848
                    LayoutCachedHeight =615
                    ColumnStart =2
                    ColumnEnd =4
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
                    Width =4988
                    Height =405
                    TabIndex =5
                    Name ="cmdSetAccUnitTlbReferenz"
                    Caption ="Set reference to AccUnit.tlb"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =2805
                    LayoutCachedTop =1395
                    LayoutCachedWidth =7793
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
                    Width =4988
                    Height =405
                    TabIndex =6
                    Name ="cmdRemoveAccUnitTlbReferenz"
                    Caption ="Remove reference to AccUnit.tlb"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =2805
                    LayoutCachedTop =1988
                    LayoutCachedWidth =7793
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
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2805
                    Top =803
                    Width =4988
                    Height =405
                    TabIndex =4
                    Name ="cmdExportAccUnitFiles"
                    Caption ="Export DLL files from add-in"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =2805
                    LayoutCachedTop =803
                    LayoutCachedWidth =7793
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
                    Left =120
                    Top =803
                    Width =2618
                    Height =405
                    Name ="EmptyCell71"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =120
                    LayoutCachedTop =803
                    LayoutCachedWidth =2738
                    LayoutCachedHeight =1208
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7853
                    Top =803
                    Width =1995
                    Height =405
                    Name ="EmptyCell73"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =7853
                    LayoutCachedTop =803
                    LayoutCachedWidth =9848
                    LayoutCachedHeight =1208
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =1395
                    Width =2618
                    Height =405
                    Name ="EmptyCell76"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =120
                    LayoutCachedTop =1395
                    LayoutCachedWidth =2738
                    LayoutCachedHeight =1800
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7853
                    Top =1395
                    Width =2632
                    Height =405
                    Name ="EmptyCell78"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =7853
                    LayoutCachedTop =1395
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =1800
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =120
                    Top =1988
                    Width =2618
                    Height =405
                    Name ="EmptyCell81"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =120
                    LayoutCachedTop =1988
                    LayoutCachedWidth =2738
                    LayoutCachedHeight =2393
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7853
                    Top =1988
                    Width =1995
                    Height =405
                    Name ="EmptyCell83"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =7853
                    LayoutCachedTop =1988
                    LayoutCachedWidth =9848
                    LayoutCachedHeight =2393
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =3338
                    Width =5182
                    Height =458
                    TabIndex =8
                    Name ="cmdInsertFactoryModule"
                    Caption ="Insert/update AccUnit Factory module in application"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =120
                    LayoutCachedTop =3338
                    LayoutCachedWidth =5302
                    LayoutCachedHeight =3796
                    PictureCaptionArrangement =5
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9915
                    Top =803
                    Width =570
                    Height =405
                    Name ="EmptyCell93"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =9915
                    LayoutCachedTop =803
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =1208
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =9915
                    Top =1988
                    Width =570
                    Height =405
                    Name ="EmptyCell95"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =9915
                    LayoutCachedTop =1988
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =2393
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9915
                    Top =3338
                    Width =570
                    Height =458
                    TabIndex =9
                    Name ="cmdOpenMenu"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="More commands ..."
                    GroupTable =2
                    BottomPadding =150

                    LayoutCachedLeft =9915
                    LayoutCachedTop =3338
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =3796
                    PictureCaptionArrangement =5
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    ThemeFontIndex =1
                    GroupTable =2
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9915
                    Top =2580
                    Width =570
                    Height =570
                    TabIndex =7
                    Name ="cmdUserSettings"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="AccUnit Settings"
                    GroupTable =2
                    BottomPadding =150
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

                    LayoutCachedLeft =9915
                    LayoutCachedTop =2580
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =3150
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =743
                            Top =2580
                            Width =9105
                            Height =570
                            Name ="labInfo"
                            GroupTable =2
                            BottomPadding =150
                            LayoutCachedLeft =743
                            LayoutCachedTop =2580
                            LayoutCachedWidth =9848
                            LayoutCachedHeight =3150
                            RowStart =4
                            RowEnd =4
                            ColumnStart =1
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            ForeThemeColorIndex =4
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =120
                    Top =2580
                    Width =562
                    Height =570
                    Name ="EmptyCell113"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =120
                    LayoutCachedTop =2580
                    LayoutCachedWidth =682
                    LayoutCachedHeight =3150
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7853
                    Top =3338
                    Width =1995
                    Height =458
                    Name ="EmptyCell119"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =7853
                    LayoutCachedTop =3338
                    LayoutCachedWidth =9848
                    LayoutCachedHeight =3796
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =5363
                    Top =3338
                    Width =2430
                    Height =458
                    Name ="EmptyCell121"
                    GroupTable =2
                    BottomPadding =150
                    LayoutCachedLeft =5363
                    LayoutCachedTop =3338
                    LayoutCachedWidth =7793
                    LayoutCachedHeight =3796
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
            End
        End
    End
End
CodeBehindForm
' See "AccUnitLoaderForm.cls"
