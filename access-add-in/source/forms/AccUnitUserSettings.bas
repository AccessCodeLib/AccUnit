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
    Width =11508
    DatasheetFontHeight =11
    ItemSuffix =209
    Left =3855
    Top =3030
    Right =17078
    Bottom =13695
    RecSrcDt = Begin
        0x212b6fd80e9ce340
    End
    Caption ="ACLib - AccUnit: Settings"
    DatasheetFontName ="Calibri"
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
            Height =5442
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    Name ="sysFirst"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2318
                    Top =90
                    Width =9072
                    Height =405
                    TabIndex =1
                    Name ="txtTestClassNameFormat"
                    GroupTable =2
                    BottomPadding =150
                    HorizontalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2318
                    LayoutCachedTop =90
                    LayoutCachedWidth =11390
                    LayoutCachedHeight =495
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =90
                            Width =2161
                            Height =405
                            Name ="Label124"
                            Caption ="TestClassNameFormat:"
                            GroupTable =2
                            RightPadding =0
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =90
                            LayoutCachedWidth =2281
                            LayoutCachedHeight =495
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2318
                    Top =683
                    Width =9072
                    Height =405
                    TabIndex =2
                    Name ="txtImportExportFolder"
                    GroupTable =2
                    BottomPadding =150
                    HorizontalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2318
                    LayoutCachedTop =683
                    LayoutCachedWidth =11390
                    LayoutCachedHeight =1088
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =683
                            Width =2161
                            Height =405
                            Name ="Label130"
                            Caption ="ImportExportFolder:"
                            GroupTable =2
                            RightPadding =0
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =683
                            LayoutCachedWidth =2281
                            LayoutCachedHeight =1088
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2318
                    Top =1275
                    Width =9072
                    Height =405
                    TabIndex =3
                    Name ="txtTemplateFolder"
                    GroupTable =2
                    BottomPadding =150
                    HorizontalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2318
                    LayoutCachedTop =1275
                    LayoutCachedWidth =11390
                    LayoutCachedHeight =1680
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =1275
                            Width =2161
                            Height =405
                            Name ="Label132"
                            Caption ="TemplateFolder:"
                            GroupTable =2
                            RightPadding =0
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =1275
                            LayoutCachedWidth =2281
                            LayoutCachedHeight =1680
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Left =2318
                    Top =1868
                    Width =9072
                    Height =2835
                    FontSize =10
                    TabIndex =5
                    Name ="txtTestMethodTemplate"
                    FontName ="Consolas"
                    GroupTable =2
                    BottomPadding =0
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    ShowDatePicker =0

                    LayoutCachedLeft =2318
                    LayoutCachedTop =1868
                    LayoutCachedWidth =11390
                    LayoutCachedHeight =4703
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =1868
                            Width =2161
                            Height =2835
                            Name ="Label134"
                            Caption ="TestMethodTemplate:"
                            GroupTable =2
                            RightPadding =0
                            BottomPadding =0
                            VerticalAnchor =2
                            LayoutCachedLeft =120
                            LayoutCachedTop =1868
                            LayoutCachedWidth =2281
                            LayoutCachedHeight =4703
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9689
                    Top =4860
                    Height =454
                    TabIndex =6
                    Name ="cmdCommit"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =9689
                    LayoutCachedTop =4860
                    LayoutCachedWidth =11390
                    LayoutCachedHeight =5314
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =113
                    Top =4860
                    Height =454
                    TabIndex =4
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =1

                    LayoutCachedLeft =113
                    LayoutCachedTop =4860
                    LayoutCachedWidth =1814
                    LayoutCachedHeight =5314
                End
            End
        End
    End
End
CodeBehindForm
' See "AccUnitUserSettings.cls"
