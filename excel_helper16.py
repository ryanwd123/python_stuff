#Microsoft Excel 16.0 Object Library

from enum import IntEnum

class AboveAverage:
    AboveBelow: 'XlAboveBelow'
    Application: 'Application'
    AppliesTo: 'Range'
    Borders: 'Borders'
    CalcFor: 'XlCalcFor'
    Creator: 'XlCreator'
    Font: 'Font'
    Interior: 'Interior'
    NumberFormat = None
    NumStdDev: float
    Parent = None
    Priority: float
    PTCondition: bool
    ScopeType: 'XlPivotConditionScope'
    StopIfTrue: bool
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: 'Range'): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class Action:
    Application: 'Application'
    Caption: str
    Content: str
    Coordinate: str
    Creator: 'XlCreator'
    Name: str
    Parent = None
    Type: 'XlActionType'
    def Execute(self, ): ...

class Actions:
    Application: 'Application'
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Action': ...
    @property
    def Item(self, Index) -> Action: ...

class AddIn:
    Application: 'Application'
    CLSID: str
    Creator: 'XlCreator'
    FullName: str
    Installed: bool
    IsOpen: bool
    Name: str
    Parent = None
    Path: str
    progID: str

class AddIns:
    Application: 'Application'
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'AddIn': ...
    def Add(self, Filename: str, CopyFile = None) -> AddIn: ...
    @property
    def Item(self, Index) -> AddIn: ...

class AddIns2:
    Application: 'Application'
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'AddIn': ...
    def Add(self, Filename: str, CopyFile = None) -> AddIn: ...
    @property
    def Item(self, Index) -> AddIn: ...

class Adjustments:
    Application = None
    Count: float
    Creator: float
    Parent = None
    def __call__(self, Index) -> 'Single': ...
    @property
    def Item(self, Index: float) -> 'Single': ...

class AllowEditRange:
    Range: 'Range'
    Title: str
    Users: 'UserAccessList'
    def ChangePassword(self, Password: str): ...
    def Delete(self, ): ...
    def Unprotect(self, Password = None): ...

class AllowEditRanges:
    Count: float
    def __call__(self, Index) -> 'AllowEditRange': ...
    def Add(self, Title: str, Range: 'Range', Password = None) -> AllowEditRange: ...
    @property
    def Item(self, Index) -> AllowEditRange: ...

class Application:
    ActiveCell: 'Range'
    ActiveChart: 'Chart'
    ActiveEncryptionSession: float
    ActivePrinter: str
    ActiveProtectedViewWindow: 'ProtectedViewWindow'
    ActiveSheet: 'Worksheet'
    ActiveWindow: 'Window'
    ActiveWorkbook: 'Workbook'
    AddIns: AddIns
    AddIns2: AddIns2
    AlertBeforeOverwriting: bool
    AltStartupPath: str
    AlwaysUseClearType: bool
    Application: 'Application'
    ArbitraryXMLSupportAvailable: bool
    AskToUpdateLinks: bool
    Assistance: 'IAssistance'
    AutoCorrect: 'AutoCorrect'
    AutoFormatAsYouTypeReplaceHyperlinks: bool
    AutomationSecurity: 'MsoAutomationSecurity'
    AutoPercentEntry: bool
    AutoRecover: 'AutoRecover'
    Build: float
    CalculateBeforeSave: bool
    Calculation: 'XlCalculation'
    CalculationInterruptKey: 'XlCalculationInterruptKey'
    CalculationState: 'XlCalculationState'
    CalculationVersion: float
    CanPlaySounds: bool
    CanRecordSounds: bool
    Caption: str
    CellDragAndDrop: bool
    Cells: 'Range'
    ChartDataPointTrack: bool
    Charts: 'Sheets'
    ClusterConnector: str
    Columns: 'Range'
    COMAddIns: 'COMAddIns'
    CommandBars: 'CommandBars'
    CommandUnderlines: 'XlCommandUnderlines'
    ConstrainNumeric: bool
    ControlCharacters: bool
    CopyObjectsWithCells: bool
    Creator: 'XlCreator'
    Cursor: 'XlMousePointer'
    CursorMovement: float
    CustomListCount: float
    CutCopyMode: 'XlCutCopyMode'
    DataEntryMode: float
    DDEAppReturnCode: float
    DecimalSeparator: str
    DefaultFilePath: str
    DefaultPivotTableLayoutOptions: 'DefaultPivotTableLayoutOptions'
    DefaultSaveFormat: 'XlFileFormat'
    DefaultSheetDirection: float
    DefaultWebOptions: 'DefaultWebOptions'
    DeferAsyncQueries: bool
    Dialogs: 'Dialogs'
    DisplayAlerts: bool
    DisplayClipboardWindow: bool
    DisplayCommentIndicator: 'XlCommentDisplayMode'
    DisplayDocumentActionTaskPane: bool
    DisplayDocumentInformationPanel: bool
    DisplayExcel4Menus: bool
    DisplayFormulaAutoComplete: bool
    DisplayFormulaBar: bool
    DisplayFullScreen: bool
    DisplayFunctionToolTips: bool
    DisplayInsertOptions: bool
    DisplayNoteIndicator: bool
    DisplayPasteOptions: bool
    DisplayRecentFiles: bool
    DisplayScrollBars: bool
    DisplayStatusBar: bool
    EditDirectlyInCell: bool
    EnableAnimations: bool
    EnableAutoComplete: bool
    EnableCancelKey: 'XlEnableCancelKey'
    EnableCheckFileExtensions: bool
    EnableEvents: bool
    EnableLargeOperationAlert: bool
    EnableLivePreview: bool
    EnableMacroAnimations: bool
    EnableSound: bool
    ErrorCheckingOptions: 'ErrorCheckingOptions'
    Excel4IntlMacroSheets: 'Sheets'
    Excel4MacroSheets: 'Sheets'
    ExtendList: bool
    FeatureInstall: 'MsoFeatureInstall'
    FileExportConverters: 'FileExportConverters'
    FileValidation: 'MsoFileValidationMode'
    FileValidationPivot: 'XlFileValidationPivotMode'
    FindFormat: 'CellFormat'
    FixedDecimal: bool
    FixedDecimalPlaces: float
    FlashFill: bool
    FlashFillMode: bool
    FormulaBarHeight: float
    GenerateGetPivotData: bool
    GenerateTableRefs: 'XlGenerateTableRefs'
    Height: float
    HighQualityModeForGraphics: bool
    Hinstance: float
    HinstancePtr = None
    Hwnd: float
    IgnoreRemoteRequests: bool
    Interactive: bool
    IsSandboxed: bool
    Iteration: bool
    LanguageSettings: 'LanguageSettings'
    LargeOperationCellThousandCount: float
    Left: float
    LibraryPath: str
    MailSession = None
    MailSystem: 'XlMailSystem'
    MapPaperSize: bool
    MathCoprocessorAvailable: bool
    MaxChange: float
    MaxIterations: float
    MeasurementUnit: float
    MergeInstances: bool
    MouseAvailable: bool
    MoveAfterReturn: bool
    MoveAfterReturnDirection: 'XlDirection'
    MultiThreadedCalculation: 'MultiThreadedCalculation'
    Name: str
    Names: 'Names'
    NetworkTemplatesPath: str
    NewWorkbook: 'NewFile'
    ODBCErrors: 'ODBCErrors'
    ODBCTimeout: float
    OLEDBErrors: 'OLEDBErrors'
    OnWindow: str
    OperatingSystem: str
    OrganizationName: str
    Parent: 'Application'
    Path: str
    PathSeparator: str
    PivotTableSelection: bool
    PrintCommunication: bool
    ProductCode: str
    PromptForSummaryInfo: bool
    ProtectedViewWindows: 'ProtectedViewWindows'
    QuickAnalysis: 'QuickAnalysis'
    Ready: bool
    RecentFiles: 'RecentFiles'
    RecordRelative: bool
    ReferenceStyle: 'XlReferenceStyle'
    ReplaceFormat: 'CellFormat'
    RollZoom: bool
    Rows: 'Range'
    RTD: 'RTD'
    ScreenUpdating: bool
    Selection = None
    SensitivityLabelPolicy: 'SensitivityLabelPolicy'
    Sheets: 'Sheets'
    SheetsInNewWorkbook: float
    ShowChartTipNames: bool
    ShowChartTipValues: bool
    ShowConvertToDataType: bool
    ShowDevTools: bool
    ShowMenuFloaties: bool
    ShowQuickAnalysis: bool
    ShowSelectionFloaties: bool
    ShowStartupDialog: bool
    ShowToolTips: bool
    SmartArtColors: 'SmartArtColors'
    SmartArtLayouts: 'SmartArtLayouts'
    SmartArtQuickStyles: 'SmartArtQuickStyles'
    Speech: 'Speech'
    SpellingOptions: 'SpellingOptions'
    StandardFont: str
    StandardFontSize: float
    StartupPath: str
    StatusBar = None
    TemplatesPath: str
    ThisCell: 'Range'
    ThisWorkbook: 'Workbook'
    ThousandsSeparator: str
    Top: float
    TransitionMenuKey: str
    TransitionMenuKeyAction: float
    TransitionNavigKeys: bool
    UsableHeight: float
    UsableWidth: float
    UseClusterConnector: bool
    UsedObjects: 'UsedObjects'
    UserControl: bool
    UserLibraryPath: str
    UserName: str
    UseSystemSeparators: bool
    Value: str
    VBE: 'VBE'
    Version: str
    Visible: bool
    WarnOnFunctionNameConflict: bool
    Watches: 'Watches'
    Width: float
    Windows: 'Windows'
    WindowsForPens: bool
    WindowState: 'XlWindowState'
    Workbooks: 'Workbooks'
    WorksheetFunction: 'WorksheetFunction'
    Worksheets: 'Sheets'
    def ActivateMicrosoftApp(self, Index: 'XlMSApplication'): ...
    def AddCustomList(self, ListArray, ByRow = None): ...
    def Calculate(self, ): ...
    def CalculateFull(self, ): ...
    def CalculateFullRebuild(self, ): ...
    def CalculateUntilAsyncQueriesDone(self, ): ...
    @property
    def Caller(self, Index = None): ...
    def CentimetersToPoints(self, Centimeters: float) -> float: ...
    def CheckAbort(self, KeepAbort = None): ...
    def CheckSpelling(self, Word: str, CustomDictionary = None, IgnoreUppercase = None) -> bool: ...
    @property
    def ClipboardFormats(self, Index = None): ...
    def ConvertFormula(self, Formula, FromReferenceStyle: 'XlReferenceStyle', ToReferenceStyle = None, ToAbsolute = None, RelativeTo = None): ...
    def DDEExecute(self, Channel: float, String: str): ...
    def DDEInitiate(self, App: str, Topic: str) -> float: ...
    def DDEPoke(self, Channel: float, Item, Data): ...
    def DDERequest(self, Channel: float, Item: str): ...
    def DDETerminate(self, Channel: float): ...
    def DeleteCustomList(self, ListNum: float): ...
    def DisplayXMLSourcePane(self, XmlMap = None): ...
    def DoubleClick(self, ): ...
    def Evaluate(self, Name): ...
    def ExecuteExcel4Macro(self, String: str): ...
    @property
    def FileConverters(self, Index1 = None, Index2 = None): ...
    @property
    def FileDialog(self, fileDialogType: 'MsoFileDialogType') -> 'FileDialog': ...
    def FindFile(self, ) -> bool: ...
    def GetCustomListContents(self, ListNum: float): ...
    def GetCustomListNum(self, ListArray) -> float: ...
    def GetOpenFilename(self, FileFilter = None, FilterIndex = None, Title = None, ButtonText = None, MultiSelect = None): ...
    def GetPhonetic(self, Text = None) -> str: ...
    def GetSaveAsFilename(self, InitialFilename = None, FileFilter = None, FilterIndex = None, Title = None, ButtonText = None): ...
    def Goto(self, Reference = None, Scroll = None): ...
    def Help(self, HelpFile = None, HelpContextID = None): ...
    def InchesToPoints(self, Inches: float) -> float: ...
    def InputBox(self, Prompt: str, Title = None, Default = None, Left = None, Top = None, HelpFile = None, HelpContextID = None, Type = None): ...
    @property
    def International(self, Index = None): ...
    def Intersect(self, Arg1: 'Range', Arg2: 'Range', Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> 'Range': ...
    def MacroOptions(self, Macro = None, Description = None, HasMenu = None, MenuText = None, HasShortcutKey = None, ShortcutKey = None, Category = None, StatusBar = None, HelpContextID = None, HelpFile = None, ArgumentDescriptions = None): ...
    def MailLogoff(self, ): ...
    def MailLogon(self, Name = None, Password = None, DownloadNewMail = None): ...
    def NextLetter(self, ) -> 'Workbook': ...
    def OnKey(self, Key: str, Procedure = None): ...
    def OnRepeat(self, Text: str, Procedure: str): ...
    def OnTime(self, EarliestTime, Procedure: str, LatestTime = None, Schedule = None): ...
    def OnUndo(self, Text: str, Procedure: str): ...
    @property
    def PreviousSelections(self, Index = None): ...
    def Quit(self, ): ...
    @property
    def Range(self, Cell1, Cell2 = None) -> 'Range': ...
    def RecordMacro(self, BasicCode = None, XlmCode = None): ...
    @property
    def RegisteredFunctions(self, Index1 = None, Index2 = None): ...
    def RegisterXLL(self, Filename: str) -> bool: ...
    def Repeat(self, ): ...
    def Run(self, Macro = None, Arg1 = None, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None): ...
    def SendKeys(self, Keys, Wait = None): ...
    def SharePointVersion(self, bstrUrl: str) -> float: ...
    def Undo(self, ): ...
    def Union(self, Arg1: 'Range', Arg2: 'Range', Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> 'Range': ...
    def Volatile(self, Volatile = None): ...
    def Wait(self, Time) -> bool: ...

class Areas:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Range': ...
    @property
    def Item(self, Index: float) -> 'Range': ...

class Author:
    Application: Application
    Creator: 'XlCreator'
    Name: str
    Parent = None
    ProviderID: str
    UserID: str

class AutoCorrect:
    Application: Application
    AutoExpandListRange: bool
    AutoFillFormulasInLists: bool
    CapitalizeNamesOfDays: bool
    CorrectCapsLock: bool
    CorrectSentenceCap: bool
    Creator: 'XlCreator'
    DisplayAutoCorrectOptions: bool
    Parent = None
    ReplaceText: bool
    TwoInitialCapitals: bool
    def AddReplacement(self, What: str, Replacement: str): ...
    def DeleteReplacement(self, What: str): ...
    @property
    def ReplacementList(self, Index = None): ...

class AutoFilter:
    Application: Application
    Creator: 'XlCreator'
    FilterMode: bool
    Filters: 'Filters'
    Parent = None
    Range: 'Range'
    Sort: 'Sort'
    def ApplyFilter(self, ): ...
    def ShowAllData(self, ): ...

class AutoRecover:
    Application: Application
    Creator: 'XlCreator'
    Enabled: bool
    Parent = None
    Path: str
    Time: float

class Axes:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Axis': ...
    def Item(self, Type: 'XlAxisType', AxisGroup: 'XlAxisGroup' = None) -> 'Axis': ...

class Axis:
    Application: Application
    AxisBetweenCategories: bool
    AxisGroup: 'XlAxisGroup'
    AxisTitle: 'AxisTitle'
    BaseUnit: 'XlTimeUnit'
    BaseUnitIsAuto: bool
    Border: 'Border'
    CategoryNames = None
    CategorySortOrder: 'XlCategorySortOrder'
    CategoryType: 'XlCategoryType'
    Creator: 'XlCreator'
    Crosses: 'XlAxisCrosses'
    CrossesAt: float
    DisplayUnit: 'XlDisplayUnit'
    DisplayUnitCustom: float
    DisplayUnitLabel: 'DisplayUnitLabel'
    Format: 'ChartFormat'
    HasDisplayUnitLabel: bool
    HasMajorGridlines: bool
    HasMinorGridlines: bool
    HasTitle: bool
    Height: float
    Left: float
    LogBase: float
    MajorGridlines: 'Gridlines'
    MajorTickMark: 'XlTickMark'
    MajorUnit: float
    MajorUnitIsAuto: bool
    MajorUnitScale: 'XlTimeUnit'
    MaximumScale: float
    MaximumScaleIsAuto: bool
    MinimumScale: float
    MinimumScaleIsAuto: bool
    MinorGridlines: 'Gridlines'
    MinorTickMark: 'XlTickMark'
    MinorUnit: float
    MinorUnitIsAuto: bool
    MinorUnitScale: 'XlTimeUnit'
    Parent = None
    ReversePlotOrder: bool
    ScaleType: 'XlScaleType'
    TickLabelPosition: 'XlTickLabelPosition'
    TickLabels: 'TickLabels'
    TickLabelSpacing: float
    TickLabelSpacingIsAuto: bool
    TickMarkSpacing: float
    Top: float
    Type: 'XlAxisType'
    Width: float
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class AxisTitle:
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    Format: 'ChartFormat'
    Formula: str
    FormulaLocal: str
    FormulaR1C1: str
    FormulaR1C1Local: str
    Height: float
    HorizontalAlignment = None
    IncludeInLayout: bool
    Left: float
    Name: str
    Orientation = None
    Parent = None
    Position: 'XlChartElementPosition'
    ReadingOrder: float
    Shadow: bool
    Text: str
    Top: float
    VerticalAlignment = None
    Width: float
    @property
    def Characters(self, Start = None, Length = None) -> 'Characters': ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class Border:
    Application: Application
    Color = None
    ColorIndex = None
    Creator: 'XlCreator'
    LineStyle = None
    Parent = None
    ThemeColor = None
    TintAndShade = None
    Weight = None

class Borders:
    Application: Application
    Color = None
    ColorIndex = None
    Count: float
    Creator: 'XlCreator'
    LineStyle = None
    Parent = None
    ThemeColor = None
    TintAndShade = None
    Value = None
    Weight = None
    def __call__(self, Index) -> 'Border': ...
    @property
    def Item(self, Index: 'XlBordersIndex') -> Border: ...

class CalculatedFields:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotField': ...
    def Add(self, Name: str, Formula: str, UseStandardFormula = None) -> 'PivotField': ...
    def Item(self, Index) -> 'PivotField': ...

class CalculatedItems:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotItem': ...
    def Add(self, Name: str, Formula: str, UseStandardFormula = None) -> 'PivotItem': ...
    def Item(self, Index) -> 'PivotItem': ...

class CalculatedMember:
    Application: Application
    Creator: 'XlCreator'
    DisplayFolder: str
    Dynamic: bool
    FlattenHierarchies: bool
    Formula: str
    HierarchizeDistinct: bool
    IsValid: bool
    MeasureGroup: str
    Name: str
    NumberFormat: 'XlCalcMemNumberFormatType'
    Parent = None
    ParentHierarchy: str
    ParentMember: str
    SolveOrder: float
    SourceName: str
    Type: 'XlCalculatedMemberType'
    def Delete(self, ): ...

class CalculatedMembers:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'CalculatedMember': ...
    def Add(self, Name: str, Formula, SolveOrder = None, Type = None, Dynamic = None, DisplayFolder = None, HierarchizeDistinct = None) -> CalculatedMember: ...
    def AddCalculatedMember(self, Name: str, Formula, SolveOrder = None, Type = None, DisplayFolder = None, MeasureGroup = None, ParentHierarchy = None, ParentMember = None, NumberFormat = None) -> CalculatedMember: ...
    @property
    def Item(self, Index) -> CalculatedMember: ...

class CalloutFormat:
    Accent: 'MsoTriState'
    Angle: 'MsoCalloutAngleType'
    Application = None
    AutoAttach: 'MsoTriState'
    AutoLength: 'MsoTriState'
    Border: 'MsoTriState'
    Creator: float
    Drop: 'Single'
    DropType: 'MsoCalloutDropType'
    Gap: 'Single'
    Length: 'Single'
    Parent = None
    Type: 'MsoCalloutType'
    def AutomaticLength(self, ): ...
    def CustomDrop(self, Drop: 'Single'): ...
    def CustomLength(self, Length: 'Single'): ...
    def PresetDrop(self, DropType: 'MsoCalloutDropType'): ...

class CategoryCollection:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ChartCategory': ...
    def Item(self, Index) -> 'ChartCategory': ...

class CellFormat:
    AddIndent = None
    Application: Application
    Borders: Borders
    Creator: 'XlCreator'
    Font: 'Font'
    FormulaHidden = None
    HorizontalAlignment = None
    IndentLevel = None
    Interior: 'Interior'
    Locked = None
    MergeCells = None
    NumberFormat = None
    NumberFormatLocal = None
    Orientation = None
    Parent = None
    ShrinkToFit = None
    VerticalAlignment = None
    WrapText = None
    def Clear(self, ): ...

class Characters:
    Application: Application
    Caption: str
    Count: float
    Creator: 'XlCreator'
    Font: 'Font'
    Parent = None
    PhoneticCharacters: str
    Text: str
    def Delete(self, ): ...
    def Insert(self, String: str): ...

class Chart:
    Application: Application
    AutoScaling: bool
    BackWall: 'Walls'
    BarShape: 'XlBarShape'
    CategoryLabelLevel: 'XlCategoryLabelLevel'
    ChartArea: 'ChartArea'
    ChartColor = None
    ChartStyle = None
    ChartTitle: 'ChartTitle'
    ChartType: 'XlChartType'
    CodeName: str
    Creator: 'XlCreator'
    DataTable: 'DataTable'
    DepthPercent: float
    DisplayBlanksAs: 'XlDisplayBlanksAs'
    DisplayValueNotAvailableAsBlank: bool
    Elevation: float
    Floor: 'Floor'
    GapDepth: float
    HasDataTable: bool
    HasLegend: bool
    HasTitle: bool
    HeightPercent: float
    Hyperlinks: 'Hyperlinks'
    Index: float
    Legend: 'Legend'
    MailEnvelope: 'MsoEnvelope'
    Name: str
    Next = None
    PageSetup: 'PageSetup'
    Parent = None
    Perspective: float
    PivotLayout: 'PivotLayout'
    PlotArea: 'PlotArea'
    PlotBy: 'XlRowCol'
    PlotVisibleOnly: bool
    Previous = None
    PrintedCommentPages: float
    ProtectContents: bool
    ProtectData: bool
    ProtectDrawingObjects: bool
    ProtectFormatting: bool
    ProtectionMode: bool
    ProtectSelection: bool
    RightAngleAxes = None
    Rotation = None
    SeriesNameLevel: 'XlSeriesNameLevel'
    Shapes: 'Shapes'
    ShowAllFieldButtons: bool
    ShowAxisFieldButtons: bool
    ShowDataLabelsOverMaximum: bool
    ShowExpandCollapseEntireFieldButtons: bool
    ShowLegendFieldButtons: bool
    ShowReportFilterFieldButtons: bool
    ShowValueFieldButtons: bool
    SideWall: 'Walls'
    Tab: 'Tab'
    Visible: 'XlSheetVisibility'
    Walls: 'Walls'
    def Activate(self, ): ...
    def ApplyChartTemplate(self, Filename: str): ...
    def ApplyDataLabels(self, Type: 'XlDataLabelsType' = None, LegendKey = None, AutoText = None, HasLeaderLines = None, ShowSeriesName = None, ShowCategoryName = None, ShowValue = None, ShowPercentage = None, ShowBubbleSize = None, Separator = None): ...
    def ApplyLayout(self, Layout: float, ChartType = None): ...
    def Axes(self, Type = None, AxisGroup: 'XlAxisGroup' = None): ...
    def ChartGroups(self, Index = None): ...
    def ChartObjects(self, Index = None): ...
    def ChartWizard(self, Source = None, Gallery = None, Format = None, PlotBy = None, CategoryLabels = None, SeriesLabels = None, HasLegend = None, Title = None, CategoryTitle = None, ValueTitle = None, ExtraTitle = None): ...
    def CheckSpelling(self, CustomDictionary = None, IgnoreUppercase = None, AlwaysSuggest = None, SpellLang = None): ...
    def ClearToMatchColorStyle(self, ): ...
    def ClearToMatchStyle(self, ): ...
    def Copy(self, Before = None, After = None): ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None, Size: 'XlPictureAppearance' = None): ...
    def Delete(self, ): ...
    def Evaluate(self, Name): ...
    def Export(self, Filename: str, FilterName = None, Interactive = None) -> bool: ...
    def ExportAsFixedFormat(self, Type: 'XlFixedFormatType', Filename = None, Quality = None, IncludeDocProperties = None, IgnorePrintAreas = None, From = None, To = None, OpenAfterPublish = None, FixedFormatExtClassPtr = None, WorkIdentity = None): ...
    def FullSeriesCollection(self, Index = None): ...
    def GetChartElement(self, x: float, y: float, ElementID: float, Arg1: float, Arg2: float): ...
    def GetProperty(self, ID: str): ...
    @property
    def HasAxis(self, Index1 = None, Index2 = None): ...
    def Location(self, Where: 'XlChartLocation', Name = None) -> 'Chart': ...
    def Move(self, Before = None, After = None): ...
    def OLEObjects(self, Index = None): ...
    def Paste(self, Type = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Protect(self, Password = None, DrawingObjects = None, Contents = None, Scenarios = None, UserInterfaceOnly = None): ...
    def Refresh(self, ): ...
    def SaveAs(self, Filename: str, FileFormat = None, Password = None, WriteResPassword = None, ReadOnlyRecommended = None, CreateBackup = None, AddToMru = None, TextCodepage = None, TextVisualLayout = None, Local = None): ...
    def SaveChartTemplate(self, Filename: str): ...
    def Select(self, Replace = None): ...
    def SeriesCollection(self, Index = None): ...
    def SetBackgroundPicture(self, Filename: str): ...
    def SetDefaultChart(self, Name): ...
    def SetElement(self, Element: 'MsoChartElementType'): ...
    def SetProperty(self, ID: str, Value): ...
    def SetSourceData(self, Source: 'Range', PlotBy = None): ...
    def Unprotect(self, Password = None): ...

class ChartArea:
    Application: Application
    Creator: 'XlCreator'
    Format: 'ChartFormat'
    Height: float
    Left: float
    Name: str
    Parent = None
    RoundedCorners: bool
    Shadow: bool
    Top: float
    Width: float
    def Clear(self, ): ...
    def ClearContents(self, ): ...
    def ClearFormats(self, ): ...
    def Copy(self, ): ...
    def Select(self, ): ...

class ChartCategory:
    Application: Application
    Creator: 'XlCreator'
    IsFiltered: bool
    Name: str
    Parent = None

class ChartFormat:
    Adjustments: Adjustments
    Application: Application
    AutoShapeType: 'MsoAutoShapeType'
    Creator: 'XlCreator'
    Fill: 'FillFormat'
    Glow: 'GlowFormat'
    Line: 'LineFormat'
    Parent = None
    PictureFormat: 'PictureFormat'
    Shadow: 'ShadowFormat'
    SoftEdge: 'SoftEdgeFormat'
    TextFrame2: 'TextFrame2'
    ThreeD: 'ThreeDFormat'

class ChartGroup:
    Application: Application
    AxisGroup: 'XlAxisGroup'
    BinsCountValue: float
    BinsOverflowEnabled: bool
    BinsOverflowValue: float
    BinsType: 'XlBinsType'
    BinsUnderflowEnabled: bool
    BinsUnderflowValue: float
    BinWidthValue: float
    BubbleScale: float
    Creator: 'XlCreator'
    DoughnutHoleSize: float
    DownBars: 'DownBars'
    DropLines: 'DropLines'
    FirstSliceAngle: float
    GapWidth: float
    Has3DShading: bool
    HasDropLines: bool
    HasHiLoLines: bool
    HasRadarAxisLabels: bool
    HasSeriesLines: bool
    HasUpDownBars: bool
    HiLoLines: 'HiLoLines'
    Index: float
    Overlap: float
    Parent = None
    RadarAxisLabels: 'TickLabels'
    SecondPlotSize: float
    SeriesLines: 'SeriesLines'
    ShowNegativeBubbles: bool
    SizeRepresents: 'XlSizeRepresents'
    SplitType: 'XlChartSplitType'
    SplitValue = None
    UpBars: 'UpBars'
    VaryByCategories: bool
    def CategoryCollection(self, Index = None): ...
    def FullCategoryCollection(self, Index = None): ...
    def SeriesCollection(self, Index = None): ...

class ChartGroups:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ChartGroup': ...
    def Item(self, Index) -> ChartGroup: ...

class ChartObject:
    Application: Application
    BottomRightCell: 'Range'
    Chart: Chart
    Creator: 'XlCreator'
    Height: float
    Index: float
    Left: float
    Locked: bool
    Name: str
    Parent = None
    Placement = None
    PrintObject: bool
    ProtectChartObject: bool
    RoundedCorners: bool
    Shadow: bool
    ShapeRange: 'ShapeRange'
    Top: float
    TopLeftCell: 'Range'
    Visible: bool
    Width: float
    ZOrder: float
    def Activate(self, ): ...
    def BringToFront(self, ): ...
    def Copy(self, ): ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None): ...
    def Cut(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, ): ...
    def Select(self, Replace = None): ...
    def SendToBack(self, ): ...

class ChartObjects:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Height: float
    Left: float
    Locked: bool
    Parent = None
    Placement = None
    PrintObject: bool
    ProtectChartObject: bool
    ShapeRange: 'ShapeRange'
    Top: float
    Visible: bool
    Width: float
    def __call__(self, Index) -> 'ChartObject': ...
    def Add(self, Left: float, Top: float, Width: float, Height: float) -> ChartObject: ...
    def Copy(self, ): ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None): ...
    def Cut(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, ): ...
    def Item(self, Index) -> ChartObject: ...
    def Select(self, Replace = None): ...

class Charts:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    HPageBreaks: 'HPageBreaks'
    Parent = None
    Visible = None
    VPageBreaks: 'VPageBreaks'
    def __call__(self, Index) -> 'None': ...
    def Add2(self, Before = None, After = None, Count = None, NewLayout = None) -> Chart: ...
    def Copy(self, Before = None, After = None): ...
    def Delete(self, ): ...
    @property
    def Item(self, Index): ...
    def Move(self, Before = None, After = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Select(self, Replace = None): ...

class ChartSeriesGradientStopData:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    StopColor: 'SeriesGradientStopColorFormat'
    StopPositionType: 'XlGradientStopPositionType'
    StopValue: str

class ChartTitle:
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    Format: ChartFormat
    Formula: str
    FormulaLocal: str
    FormulaR1C1: str
    FormulaR1C1Local: str
    Height: float
    HorizontalAlignment = None
    IncludeInLayout: bool
    Left: float
    Name: str
    Orientation = None
    Parent = None
    Position: 'XlChartElementPosition'
    ReadingOrder: float
    Shadow: bool
    Text: str
    Top: float
    VerticalAlignment = None
    Width: float
    @property
    def Characters(self, Start = None, Length = None) -> Characters: ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class ChartView:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    Sheet = None

class ColorFormat:
    Application = None
    Brightness: 'Single'
    Creator: float
    ObjectThemeColor: 'MsoThemeColorIndex'
    Parent = None
    RGB: 'MsoRGBType'
    SchemeColor: float
    TintAndShade: 'Single'
    Type: 'MsoColorType'

class ColorScale:
    Application: Application
    AppliesTo: 'Range'
    ColorScaleCriteria: 'ColorScaleCriteria'
    Creator: 'XlCreator'
    Formula: str
    Parent = None
    Priority: float
    PTCondition: bool
    ScopeType: 'XlPivotConditionScope'
    StopIfTrue: bool
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: 'Range'): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class ColorScaleCriteria:
    Count: float
    def __call__(self, Index) -> 'ColorScaleCriterion': ...
    @property
    def Item(self, Index) -> 'ColorScaleCriterion': ...

class ColorScaleCriterion:
    FormatColor: 'FormatColor'
    Index: float
    Type: 'XlConditionValueTypes'
    Value = None

class ColorStop:
    Application: Application
    Color = None
    Creator: 'XlCreator'
    Parent = None
    Position: float
    ThemeColor: float
    TintAndShade = None
    def Delete(self, ): ...

class ColorStops:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ColorStop': ...
    def Add(self, Position: float) -> ColorStop: ...
    def Clear(self, ): ...
    def Item(self, Index) -> ColorStop: ...

class Comment:
    Application: Application
    Author: str
    Creator: 'XlCreator'
    Parent = None
    Shape: 'Shape'
    Visible: bool
    def Delete(self, ): ...
    def Next(self, ) -> 'Comment': ...
    def Previous(self, ) -> 'Comment': ...
    def Text(self, Text = None, Start = None, Overwrite = None) -> str: ...

class Comments:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Comment': ...
    def Item(self, Index: float) -> Comment: ...

class CommentsThreaded:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'CommentThreaded': ...
    def Item(self, Index: float) -> 'CommentThreaded': ...

class CommentThreaded:
    Application: Application
    Author: Author
    Creator: 'XlCreator'
    Date = None
    Parent = None
    Replies: CommentsThreaded
    Resolved: bool
    def AddReply(self, Text = None) -> 'CommentThreaded': ...
    def Delete(self, ): ...
    def Next(self, ) -> 'CommentThreaded': ...
    def Previous(self, ) -> 'CommentThreaded': ...
    def Text(self, Text = None, Start = None, Overwrite = None) -> str: ...

class ConditionValue:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    Type: 'XlConditionValueTypes'
    Value = None
    def Modify(self, newtype: 'XlConditionValueTypes', newvalue = None): ...

class Connections:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'WorkbookConnection': ...
    def Add2(self, Name: str, Description: str, ConnectionString, CommandText, lCmdtype = None, CreateModelConnection = None, ImportRelationships = None) -> 'WorkbookConnection': ...
    def AddFromFile(self, Filename: str, CreateModelConnection = None, ImportRelationships = None) -> 'WorkbookConnection': ...
    def Item(self, Index) -> 'WorkbookConnection': ...

class ConnectorFormat:
    Application: Application
    BeginConnected: 'MsoTriState'
    BeginConnectedShape: 'Shape'
    BeginConnectionSite: float
    Creator: 'XlCreator'
    EndConnected: 'MsoTriState'
    EndConnectedShape: 'Shape'
    EndConnectionSite: float
    Parent = None
    Type: 'MsoConnectorType'
    def BeginConnect(self, ConnectedShape: 'Shape', ConnectionSite: float): ...
    def BeginDisconnect(self, ): ...
    def EndConnect(self, ConnectedShape: 'Shape', ConnectionSite: float): ...
    def EndDisconnect(self, ): ...

class Constants(IntEnum):
    xl3DBar = -4099 
    xl3DEffects1 = 13
    xl3DEffects2 = 14
    xl3DSurface = -4103 
    xlAbove = 0
    xlAccounting1 = 4
    xlAccounting2 = 5
    xlAccounting3 = 6
    xlAccounting4 = 17 
    xlAdd = 2
    xlAll = -4104 
    xlAllExceptBorders = 7
    xlAutomatic = -4105 
    xlBar = 2
    xlBelow = 1
    xlBidi = -5000 
    xlBidiCalendar = 3
    xlBoth = 1
    xlBottom = -4107 
    xlCascade = 7
    xlCenter = -4108 
    xlCenterAcrossSelection = 7
    xlChart4 = 2
    xlChartSeries = 17 
    xlChartShort = 6
    xlChartTitles = 18 
    xlChecker = 9
    xlCircle = 8
    xlClassic1 = 1
    xlClassic2 = 2
    xlClassic3 = 3
    xlClosed = 3
    xlColor1 = 7
    xlColor2 = 8
    xlColor3 = 9
    xlColumn = 3
    xlCombination = -4111 
    xlComplete = 4
    xlConstants = 2
    xlContents = 2
    xlContext = -5002 
    xlCorner = 2
    xlCrissCross = 16 
    xlCross = 4
    xlCustom = -4114 
    xlDebugCodePane = 13
    xlDefaultAutoFormat = -1 
    xlDesktop = 9
    xlDiamond = 2
    xlDirect = 1
    xlDistributed = -4117 
    xlDivide = 5
    xlDoubleAccounting = 5
    xlDoubleClosed = 5
    xlDoubleOpen = 4
    xlDoubleQuote = 1
    xlDrawingObject = 14
    xlEntireChart = 20 
    xlExcelMenus = 1
    xlExtended = 3
    xlFill = 5
    xlFirst = 0
    xlFixedValue = 1
    xlFloating = 5
    xlFormats = -4122 
    xlFormula = 5
    xlFullScript = 1
    xlGeneral = 1
    xlGray16 = 17 
    xlGray25 = -4124 
    xlGray50 = -4125 
    xlGray75 = -4126 
    xlGray8 = 18 
    xlGregorian = 2
    xlGrid = 15
    xlGridline = 22 
    xlHigh = -4127 
    xlHindiNumerals = 3
    xlIcons = 1
    xlImmediatePane = 12
    xlInside = 2
    xlInteger = 2
    xlJustify = -4130 
    xlLast = 1
    xlLastCell = 11
    xlLatin = -5001 
    xlLeft = -4131 
    xlLeftToRight = 2
    xlLightDown = 13
    xlLightHorizontal = 11
    xlLightUp = 14
    xlLightVertical = 12
    xlList1 = 10
    xlList2 = 11
    xlList3 = 12
    xlLocalFormat1 = 15
    xlLocalFormat2 = 16 
    xlLogicalCursor = 1
    xlLong = 3
    xlLotusHelp = 2
    xlLow = -4134 
    xlLTR = -5003 
    xlMacrosheetCell = 7
    xlManual = -4135 
    xlMaximum = 2
    xlMinimum = 4
    xlMinusValues = 3
    xlMixed = 2
    xlMixedAuthorizedScript = 4
    xlMixedScript = 3
    xlModule = -4141 
    xlMultiply = 4
    xlNarrow = 1
    xlNextToAxis = 4
    xlNoDocuments = 3
    xlNone = -4142 
    xlNotes = -4144 
    xlOff = -4146 
    xlOn = 1
    xlOpaque = 3
    xlOpen = 2
    xlOutside = 3
    xlPartial = 3
    xlPartialScript = 2
    xlPercent = 2
    xlPlus = 9
    xlPlusValues = 2
    xlReference = 4
    xlRight = -4152 
    xlRTL = -5004 
    xlScale = 3
    xlSemiautomatic = 2
    xlSemiGray75 = 10
    xlShort = 1
    xlShowLabel = 4
    xlShowLabelAndPercent = 5
    xlShowPercent = 3
    xlShowValue = 2
    xlSimple = -4154 
    xlSingle = 2
    xlSingleAccounting = 4
    xlSingleQuote = 2
    xlSolid = 1
    xlSquare = 1
    xlStar = 5
    xlStError = 4
    xlStrict = 2
    xlSubtract = 3
    xlSystem = 1
    xlTextBox = 16 
    xlTiled = 1
    xlTitleBar = 8
    xlToolbar = 1
    xlToolbarButton = 2
    xlTop = -4160 
    xlTopToBottom = 1
    xlTransparent = 2
    xlTriangle = 3
    xlVeryHidden = 2
    xlVisible = 12
    xlVisualCursor = 2
    xlWatchPane = 11
    xlWide = 3
    xlWorkbookTab = 6
    xlWorksheet4 = 1
    xlWorksheetCell = 3
    xlWorksheetShort = 5

class ControlFormat:
    Application: Application
    Creator: 'XlCreator'
    DropDownLines: float
    Enabled: bool
    LargeChange: float
    LinkedCell: str
    ListCount: float
    ListFillRange: str
    ListIndex: float
    LockedText: bool
    Max: float
    Min: float
    MultiSelect: float
    Parent = None
    PrintObject: bool
    SmallChange: float
    Value: float
    def AddItem(self, Text: str, Index = None): ...
    def List(self, Index = None): ...
    def RemoveAllItems(self, ): ...
    def RemoveItem(self, Index: float, Count = None): ...

class CubeField:
    AllItemsVisible: bool
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    CubeFieldSubType: 'XlCubeFieldSubType'
    CubeFieldType: 'XlCubeFieldType'
    CurrentPageName: str
    DragToColumn: bool
    DragToData: bool
    DragToHide: bool
    DragToPage: bool
    DragToRow: bool
    EnableMultiplePageItems: bool
    FlattenHierarchies: bool
    HasMemberProperties: bool
    HierarchizeDistinct: bool
    IncludeNewItemsInFilter: bool
    IsDate: bool
    LayoutForm: 'XlLayoutFormType'
    LayoutSubtotalLocation: 'XlSubtototalLocationType'
    Name: str
    Orientation: 'XlPivotFieldOrientation'
    Parent = None
    PivotFields: 'PivotFields'
    Position: float
    ShowInFieldList: bool
    TreeviewControl: 'TreeviewControl'
    Value: str
    def AddMemberPropertyField(self, Property: str, PropertyOrder = None, PropertyDisplayedIn = None): ...
    def AutoGroup(self, Orientation = None, Position = None): ...
    def ClearManualFilter(self, ): ...
    def CreatePivotFields(self, ): ...
    def Delete(self, ): ...

class CubeFields:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'CubeField': ...
    def AddSet(self, Name: str, Caption: str) -> CubeField: ...
    def GetMeasure(self, AttributeHierarchy, Function: 'XlConsolidationFunction', Caption = None) -> CubeField: ...
    @property
    def Item(self, Index) -> CubeField: ...

class CustomProperties:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'CustomProperty': ...
    def Add(self, Name: str, Value) -> 'CustomProperty': ...
    @property
    def Item(self, Index) -> 'CustomProperty': ...

class CustomProperty:
    Application: Application
    Creator: 'XlCreator'
    Name: str
    Parent = None
    Value = None
    def Delete(self, ): ...

class CustomView:
    Application: Application
    Creator: 'XlCreator'
    Name: str
    Parent = None
    PrintSettings: bool
    RowColSettings: bool
    def Delete(self, ): ...
    def Show(self, ): ...

class CustomViews:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'CustomView': ...
    def Add(self, ViewName: str, PrintSettings = None, RowColSettings = None) -> CustomView: ...
    def Item(self, ViewName) -> CustomView: ...

class Databar:
    Application: Application
    AppliesTo: 'Range'
    AxisColor = None
    AxisPosition: 'XlDataBarAxisPosition'
    BarBorder: 'DataBarBorder'
    BarColor = None
    BarFillType: 'XlDataBarFillType'
    Creator: 'XlCreator'
    Direction: float
    Formula: str
    MaxPoint: ConditionValue
    MinPoint: ConditionValue
    NegativeBarFormat: 'NegativeBarFormat'
    Parent = None
    PercentMax: float
    PercentMin: float
    Priority: float
    PTCondition: bool
    ScopeType: 'XlPivotConditionScope'
    ShowValue: bool
    StopIfTrue: bool
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: 'Range'): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class DataBarBorder:
    Application: Application
    Color = None
    Creator: 'XlCreator'
    Parent = None
    Type: 'XlDataBarBorderType'

class DataFeedConnection:
    AlwaysUseConnectionFile: bool
    Application: Application
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    EnableRefresh: bool
    Parent = None
    RefreshDate: 'Date'
    Refreshing: bool
    RefreshOnFileOpen: bool
    RefreshPeriod: float
    SavePassword: bool
    ServerCredentialsMethod: 'XlCredentialsMethod'
    SourceConnectionFile: str
    SourceDataFile: str
    def CancelRefresh(self, ): ...
    def Refresh(self, ): ...
    def SaveAsODC(self, ODCFileName: str, Description = None, Keywords = None): ...

class DataLabel:
    Application: Application
    AutoText: bool
    Caption: str
    Creator: 'XlCreator'
    Format: ChartFormat
    Formula: str
    FormulaLocal: str
    FormulaR1C1: str
    FormulaR1C1Local: str
    Height: float
    HorizontalAlignment = None
    Left: float
    Name: str
    NumberFormat: str
    NumberFormatLinked: bool
    NumberFormatLocal = None
    Orientation = None
    Parent = None
    Position: 'XlDataLabelPosition'
    ReadingOrder: float
    Separator = None
    Shadow: bool
    ShowBubbleSize: bool
    ShowCategoryName: bool
    ShowLegendKey: bool
    ShowPercentage: bool
    ShowRange: bool
    ShowSeriesName: bool
    ShowValue: bool
    Text: str
    Top: float
    VerticalAlignment = None
    Width: float
    @property
    def Characters(self, Start = None, Length = None) -> Characters: ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class DataLabels:
    Application: Application
    AutoText: bool
    Count: float
    Creator: 'XlCreator'
    Format: ChartFormat
    HorizontalAlignment = None
    Name: str
    NumberFormat: str
    NumberFormatLinked: bool
    NumberFormatLocal = None
    Orientation = None
    Parent = None
    Position: 'XlDataLabelPosition'
    ReadingOrder: float
    Separator = None
    Shadow: bool
    ShowBubbleSize: bool
    ShowCategoryName: bool
    ShowLegendKey: bool
    ShowPercentage: bool
    ShowRange: bool
    ShowSeriesName: bool
    ShowValue: bool
    VerticalAlignment = None
    def __call__(self, Index) -> 'DataLabel': ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Item(self, Index) -> DataLabel: ...
    def Propagate(self, Index): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class DataTable:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Font: 'Font'
    Format: ChartFormat
    HasBorderHorizontal: bool
    HasBorderOutline: bool
    HasBorderVertical: bool
    Parent = None
    ShowLegendKey: bool
    def Delete(self, ): ...
    def Select(self, ): ...

class DefaultPivotTableLayoutOptions:
    AllowMultipleFilters: bool
    Application: Application
    CalculatedMembersInFilters: bool
    ColumnGrand: bool
    CompactRowIndent: float
    Creator: 'XlCreator'
    DisplayContextTooltips: bool
    DisplayEmptyColumn: bool
    DisplayEmptyRow: bool
    DisplayErrorString: bool
    DisplayFieldCaptions: bool
    DisplayImmediateItems: bool
    DisplayMemberPropertyTooltips: bool
    DisplayNullString: bool
    EnableDrilldown: bool
    EnableWriteback: bool
    ErrorString: str
    FieldListSortAscending: bool
    HasAutoFormat: bool
    InGridDropZones: bool
    LayoutBlankLine: bool
    MergeLabels: bool
    NullString: str
    PageFieldOrder: bool
    PageFieldWrapCount: float
    Parent = None
    PreserveFormatting: bool
    PrintDrillIndicators: bool
    PrintTitles: bool
    RefreshOnFileOpen: bool
    RepeatAllLabels: 'XlPivotFieldRepeatLabels'
    RepeatItemsOnEachPrintedPage: bool
    RowAxisLayout: 'XlLayoutRowType'
    RowGrand: bool
    SaveData: bool
    ShowDrillIndicators: bool
    ShowValuesRow: bool
    SortUsingCustomLists: bool
    SubtotalHiddenPageItems: bool
    SubtotalLocation: bool
    Subtotals: bool
    TotalsAnnotation: bool
    ViewCalculatedMembers: bool
    VisualTotals: bool
    VisualTotalsForSets: bool
    xlMissingItemsNone: float

class DefaultWebOptions:
    AllowPNG: bool
    AlwaysSaveInDefaultEncoding: bool
    Application: Application
    CheckIfOfficeIsHTMLEditor: bool
    Creator: 'XlCreator'
    DownloadComponents: bool
    Encoding: 'MsoEncoding'
    FolderSuffix: str
    Fonts: 'WebPageFonts'
    LoadPictures: bool
    LocationOfComponents: str
    OrganizeInFolder: bool
    Parent = None
    PixelsPerInch: float
    RelyOnCSS: bool
    RelyOnVML: bool
    SaveHiddenData: bool
    SaveNewWebPagesAsWebArchives: bool
    ScreenSize: 'MsoScreenSize'
    TargetBrowser: 'MsoTargetBrowser'
    UpdateLinksOnSave: bool
    UseLongFileNames: bool

class Dialog:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def Show(self, Arg1 = None, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> bool: ...

class Dialogs:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Dialog': ...
    @property
    def Item(self, Index: 'XlBuiltInDialog') -> Dialog: ...

class DialogSheetView:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    Sheet = None

class DisplayFormat:
    AddIndent = None
    Application: Application
    Borders: Borders
    Creator: 'XlCreator'
    Font: 'Font'
    FormulaHidden = None
    HorizontalAlignment = None
    IndentLevel = None
    Interior: 'Interior'
    Locked = None
    MergeCells = None
    NumberFormat = None
    NumberFormatLocal = None
    Orientation = None
    Parent = None
    ReadingOrder: float
    ShrinkToFit = None
    Style = None
    VerticalAlignment = None
    WrapText = None
    @property
    def Characters(self, Start = None, Length = None) -> Characters: ...

class DisplayUnitLabel:
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    Format: ChartFormat
    Formula: str
    FormulaLocal: str
    FormulaR1C1: str
    FormulaR1C1Local: str
    Height: float
    HorizontalAlignment = None
    Left: float
    Name: str
    Orientation = None
    Parent = None
    Position: 'XlChartElementPosition'
    ReadingOrder: float
    Shadow: bool
    Text: str
    Top: float
    VerticalAlignment = None
    Width: float
    @property
    def Characters(self, Start = None, Length = None) -> Characters: ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class DownBars:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class DropLines:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def Select(self, ): ...

class Error:
    Application: Application
    Creator: 'XlCreator'
    Ignore: bool
    Parent = None
    Value: bool

class ErrorBars:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    EndStyle: 'XlEndStyleCap'
    Format: ChartFormat
    Name: str
    Parent = None
    def ClearFormats(self, ): ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class ErrorCheckingOptions:
    Application: Application
    BackgroundChecking: bool
    Creator: 'XlCreator'
    EmptyCellReferences: bool
    EvaluateToError: bool
    InconsistentFormula: bool
    InconsistentTableFormula: bool
    IndicatorColorIndex: 'XlColorIndex'
    ListDataValidation: bool
    MisleadingNumberFormats: bool
    NumberAsText: bool
    OmittedCells: bool
    Parent = None
    TextDate: bool
    UnlockedFormulaCells: bool

class Errors:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Error': ...
    @property
    def Item(self, Index) -> Error: ...

class FileExportConverter:
    Application: Application
    Creator: 'XlCreator'
    Description: str
    Extensions: str
    FileFormat: float
    Parent = None

class FileExportConverters:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'FileExportConverter': ...
    @property
    def Item(self, Index) -> FileExportConverter: ...

class FillFormat:
    Application = None
    BackColor: ColorFormat
    Creator: float
    ForeColor: ColorFormat
    GradientAngle: 'Single'
    GradientColorType: 'MsoGradientColorType'
    GradientDegree: 'Single'
    GradientStops: 'GradientStops'
    GradientStyle: 'MsoGradientStyle'
    GradientVariant: float
    Parent = None
    Pattern: 'MsoPatternType'
    PictureEffects: 'PictureEffects'
    PresetGradientType: 'MsoPresetGradientType'
    PresetTexture: 'MsoPresetTexture'
    RotateWithObject: 'MsoTriState'
    TextureAlignment: 'MsoTextureAlignment'
    TextureHorizontalScale: 'Single'
    TextureName: str
    TextureOffsetX: 'Single'
    TextureOffsetY: 'Single'
    TextureTile: 'MsoTriState'
    TextureType: 'MsoTextureType'
    TextureVerticalScale: 'Single'
    Transparency: 'Single'
    Type: 'MsoFillType'
    Visible: 'MsoTriState'
    def OneColorGradient(self, Style: 'MsoGradientStyle', Variant: float, Degree: 'Single'): ...
    def Patterned(self, Pattern: 'MsoPatternType'): ...
    def PresetGradient(self, Style: 'MsoGradientStyle', Variant: float, PresetGradientType: 'MsoPresetGradientType'): ...
    def PresetTextured(self, PresetTexture: 'MsoPresetTexture'): ...
    def Solid(self, ): ...
    def TwoColorGradient(self, Style: 'MsoGradientStyle', Variant: float): ...
    def UserPicture(self, PictureFile: str): ...
    def UserTextured(self, TextureFile: str): ...

class Filter:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Criteria1 = None
    Criteria2 = None
    On: bool
    Operator: 'XlAutoFilterOperator'
    Parent = None

class Filters:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Filter': ...
    @property
    def Item(self, Index: float) -> Filter: ...

class Floor:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    PictureType = None
    Thickness: float
    def ClearFormats(self, ): ...
    def Paste(self, ): ...
    def Select(self, ): ...

class Font:
    Application: Application
    Background = None
    Bold = None
    Color = None
    ColorIndex = None
    Creator: 'XlCreator'
    FontStyle = None
    Italic = None
    Name = None
    Parent = None
    Size = None
    Strikethrough = None
    Subscript = None
    Superscript = None
    ThemeColor = None
    ThemeFont: 'XlThemeFont'
    TintAndShade = None
    Underline = None

class FormatColor:
    Application: Application
    Color = None
    ColorIndex: 'XlColorIndex'
    Creator: 'XlCreator'
    Parent = None
    ThemeColor = None
    TintAndShade = None

class FormatCondition:
    Application: Application
    AppliesTo: 'Range'
    Borders: Borders
    Creator: 'XlCreator'
    DateOperator: 'XlTimePeriods'
    Font: Font
    Formula1: str
    Formula2: str
    Interior: 'Interior'
    NumberFormat = None
    Operator: float
    Parent = None
    Priority: float
    PTCondition: bool
    ScopeType: 'XlPivotConditionScope'
    StopIfTrue: bool
    Text: str
    TextOperator: 'XlContainsOperator'
    Type: float
    def Delete(self, ): ...
    def Modify(self, Type: 'XlFormatConditionType', Operator = None, Formula1 = None, Formula2 = None, String = None, Operator2 = None): ...
    def ModifyAppliesToRange(self, Range: 'Range'): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class FormatConditions:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'FormatCondition': ...
    def Add(self, Type: 'XlFormatConditionType', Operator = None, Formula1 = None, Formula2 = None, String = None, TextOperator = None, DateOperator = None, ScopeType = None): ...
    def AddAboveAverage(self, ): ...
    def AddColorScale(self, ColorScaleType: float): ...
    def AddDatabar(self, ): ...
    def AddIconSetCondition(self, ): ...
    def AddTop10(self, ): ...
    def AddUniqueValues(self, ): ...
    def Delete(self, ): ...
    def Item(self, Index) -> FormatCondition: ...

class FreeformBuilder:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def AddNodes(self, SegmentType: 'MsoSegmentType', EditingType: 'MsoEditingType', X1: 'Single', Y1: 'Single', X2 = None, Y2 = None, X3 = None, Y3 = None): ...
    def ConvertToShape(self, ) -> 'Shape': ...

class FullSeriesCollection:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Series': ...
    def Item(self, Index) -> 'Series': ...

class Graphic:
    Application: Application
    Brightness: 'Single'
    ColorType: 'MsoPictureColorType'
    Contrast: 'Single'
    Creator: 'XlCreator'
    CropBottom: 'Single'
    CropLeft: 'Single'
    CropRight: 'Single'
    CropTop: 'Single'
    Filename: str
    Height: 'Single'
    LockAspectRatio: 'MsoTriState'
    Parent = None
    Width: 'Single'

class Gridlines:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class GroupShapes:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Shape': ...
    def Item(self, Index) -> 'Shape': ...
    @property
    def Range(self, Index) -> 'ShapeRange': ...

class HeaderFooter:
    Picture: Graphic
    Text: str

class HiLoLines:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def Select(self, ): ...

class HPageBreak:
    Application: Application
    Creator: 'XlCreator'
    Extent: 'XlPageBreakExtent'
    Location: 'Range'
    Parent: 'Worksheet'
    Type: 'XlPageBreak'
    def Delete(self, ): ...
    def DragOff(self, Direction: 'XlDirection', RegionIndex: float): ...

class HPageBreaks:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'HPageBreak': ...
    def Add(self, Before) -> HPageBreak: ...
    @property
    def Item(self, Index: float) -> HPageBreak: ...

class Hyperlink:
    Address: str
    Application: Application
    Creator: 'XlCreator'
    EmailSubject: str
    Name: str
    Parent = None
    Range: 'Range'
    ScreenTip: str
    Shape: 'Shape'
    SubAddress: str
    TextToDisplay: str
    Type: float
    def AddToFavorites(self, ): ...
    def CreateNewDocument(self, Filename: str, EditNow: bool, Overwrite: bool): ...
    def Delete(self, ): ...
    def Follow(self, NewWindow = None, AddHistory = None, ExtraInfo = None, Method = None, HeaderInfo = None): ...

class Hyperlinks:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Hyperlink': ...
    def Add(self, Anchor, Address: str, SubAddress = None, ScreenTip = None, TextToDisplay = None): ...
    def Delete(self, ): ...
    @property
    def Item(self, Index) -> Hyperlink: ...

class Icon:
    Application: Application
    Creator: 'XlCreator'
    Index: float
    Parent: 'IconSet'

class IconCriteria:
    Count: float
    def __call__(self, Index) -> 'IconCriterion': ...
    @property
    def Item(self, Index) -> 'IconCriterion': ...

class IconCriterion:
    Icon: 'XlIcon'
    Index: float
    Operator: float
    Type: 'XlConditionValueTypes'
    Value = None

class IconSet:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    ID: 'XlIconSet'
    Parent = None
    def __call__(self, Index) -> 'Icon': ...
    @property
    def Item(self, Index) -> Icon: ...

class IconSetCondition:
    Application: Application
    AppliesTo: 'Range'
    Creator: 'XlCreator'
    Formula: str
    IconCriteria: IconCriteria
    IconSet = None
    Parent = None
    PercentileValues: bool
    Priority: float
    PTCondition: bool
    ReverseOrder: bool
    ScopeType: 'XlPivotConditionScope'
    ShowIconOnly: bool
    StopIfTrue: bool
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: 'Range'): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class IconSets:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'None': ...
    @property
    def Item(self, Index): ...

class Interior:
    Application: Application
    Color = None
    ColorIndex = None
    Creator: 'XlCreator'
    Gradient = None
    InvertIfNegative = None
    Parent = None
    Pattern = None
    PatternColor = None
    PatternColorIndex = None
    PatternThemeColor = None
    PatternTintAndShade = None
    ThemeColor = None
    TintAndShade = None

class IRtdServer:
    def ConnectData(self, TopicID: float, Strings, GetNewValues: bool): ...
    def DisconnectData(self, TopicID: float): ...
    def Heartbeat(self, ) -> float: ...
    def RefreshData(self, TopicCount: 'Long)'): ...
    def ServerStart(self, CallbackObject: 'IRTDUpdateEvent') -> float: ...
    def ServerTerminate(self, ): ...

class IRTDUpdateEvent:
    HeartbeatInterval: float
    def Disconnect(self, ): ...
    def UpdateNotify(self, ): ...

class LeaderLines:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Format: ChartFormat
    Parent = None
    def Delete(self, ): ...
    def Select(self, ): ...

class Legend:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Height: float
    IncludeInLayout: bool
    Left: float
    Name: str
    Parent = None
    Position: 'XlLegendPosition'
    Shadow: bool
    Top: float
    Width: float
    def Clear(self, ): ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def LegendEntries(self, Index = None): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class LegendEntries:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'LegendEntry': ...
    def Item(self, Index) -> 'LegendEntry': ...

class LegendEntry:
    Application: Application
    Creator: 'XlCreator'
    Font: Font
    Format: ChartFormat
    Height: float
    Index: float
    Left: float
    LegendKey: 'LegendKey'
    Parent = None
    Top: float
    Width: float
    def Delete(self, ): ...
    def Select(self, ): ...

class LegendKey:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Height: float
    InvertIfNegative: bool
    Left: float
    MarkerBackgroundColor: float
    MarkerBackgroundColorIndex: 'XlColorIndex'
    MarkerForegroundColor: float
    MarkerForegroundColorIndex: 'XlColorIndex'
    MarkerSize: float
    MarkerStyle: 'XlMarkerStyle'
    Parent = None
    PictureType: float
    PictureUnit2: float
    Shadow: bool
    Smooth: bool
    Top: float
    Width: float
    def ClearFormats(self, ): ...
    def Delete(self, ): ...

class LinearGradient:
    Application: Application
    ColorStops: ColorStops
    Creator: 'XlCreator'
    Degree: float
    Parent = None

class LineFormat:
    Application = None
    BackColor: ColorFormat
    BeginArrowheadLength: 'MsoArrowheadLength'
    BeginArrowheadStyle: 'MsoArrowheadStyle'
    BeginArrowheadWidth: 'MsoArrowheadWidth'
    Creator: float
    DashStyle: 'MsoLineDashStyle'
    EndArrowheadLength: 'MsoArrowheadLength'
    EndArrowheadStyle: 'MsoArrowheadStyle'
    EndArrowheadWidth: 'MsoArrowheadWidth'
    ForeColor: ColorFormat
    InsetPen: 'MsoTriState'
    Parent = None
    Pattern: 'MsoPatternType'
    Style: 'MsoLineStyle'
    Transparency: 'Single'
    Visible: 'MsoTriState'
    Weight: 'Single'

class LinkFormat:
    Application: Application
    AutoUpdate: bool
    Creator: 'XlCreator'
    Locked: bool
    Parent = None
    def Update(self, ): ...

class ListColumn:
    Application: Application
    Creator: 'XlCreator'
    DataBodyRange: 'Range'
    Index: float
    Name: str
    Parent = None
    Range: 'Range'
    Total: 'Range'
    TotalsCalculation: 'XlTotalsCalculation'
    XPath: 'XPath'
    def Delete(self, ): ...

class ListColumns:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ListColumn': ...
    def Add(self, Position = None) -> ListColumn: ...
    @property
    def Item(self, Index) -> ListColumn: ...

class ListDataFormat:
    AllowFillIn: bool
    Application: Application
    Choices = None
    Creator: 'XlCreator'
    DecimalPlaces: float
    DefaultValue = None
    IsPercent: bool
    lcid: float
    MaxCharacters: float
    MaxNumber = None
    MinNumber = None
    Parent = None
    ReadOnly: bool
    Required: bool
    Type: 'XlListDataType'

class ListObject:
    Active: bool
    AlternativeText: str
    Application: Application
    AutoFilter: AutoFilter
    Comment: str
    Creator: 'XlCreator'
    DataBodyRange: 'Range'
    DisplayName: str
    DisplayRightToLeft: bool
    HeaderRowRange: 'Range'
    InsertRowRange: 'Range'
    ListColumns: ListColumns
    ListRows: 'ListRows'
    Name: str
    Parent = None
    QueryTable: 'QueryTable'
    Range: 'Range'
    SharePointURL: str
    ShowAutoFilter: bool
    ShowAutoFilterDropDown: bool
    ShowHeaders: bool
    ShowTableStyleColumnStripes: bool
    ShowTableStyleFirstColumn: bool
    ShowTableStyleLastColumn: bool
    ShowTableStyleRowStripes: bool
    ShowTotals: bool
    Slicers: 'Slicers'
    Sort: 'Sort'
    SourceType: 'XlListObjectSourceType'
    Summary: str
    TableObject: 'TableObject'
    TableStyle = None
    TotalsRowRange: 'Range'
    XmlMap: 'XmlMap'
    def Delete(self, ): ...
    def ExportToVisio(self, ): ...
    def Publish(self, Target, LinkSource: bool) -> str: ...
    def Refresh(self, ): ...
    def Resize(self, Range: 'Range'): ...
    def Unlink(self, ): ...
    def Unlist(self, ): ...

class ListObjects:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ListObject': ...
    def Add(self, SourceType: 'XlListObjectSourceType' = None, Source = None, LinkSource = None, XlListObjectHasHeaders: 'XlYesNoGuess' = None, Destination = None, TableStyleName = None) -> ListObject: ...
    @property
    def Item(self, Index) -> ListObject: ...

class ListRow:
    Application: Application
    Creator: 'XlCreator'
    Index: float
    Parent = None
    Range: 'Range'
    def Delete(self, ): ...

class ListRows:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ListRow': ...
    def Add(self, Position = None, AlwaysInsert = None) -> ListRow: ...
    @property
    def Item(self, Index) -> ListRow: ...

class Mailer:
    Application: Application
    BCCRecipients = None
    CCRecipients = None
    Creator: 'XlCreator'
    Enclosures = None
    Parent = None
    Received: bool
    SendDateTime: 'Date'
    Sender: str
    Subject: str
    ToRecipients = None
    WhichAddress = None

class Model:
    Application: Application
    Creator: 'XlCreator'
    DataModelConnection: 'WorkbookConnection'
    ModelFormatBoolean: 'ModelFormatBoolean'
    ModelFormatGeneral: 'ModelFormatGeneral'
    ModelMeasures: 'ModelMeasures'
    ModelRelationships: 'ModelRelationships'
    ModelTables: 'ModelTables'
    Name: str
    Parent = None
    def AddConnection(self, ConnectionToDataSource: 'WorkbookConnection') -> 'WorkbookConnection': ...
    def CreateModelWorkbookConnection(self, ModelTable) -> 'WorkbookConnection': ...
    def Initialize(self, ): ...
    @property
    def ModelFormatCurrency(self, Symbol = None, DecimalPlaces = None) -> 'ModelFormatCurrency': ...
    @property
    def ModelFormatDate(self, FormatString = None) -> 'ModelFormatDate': ...
    @property
    def ModelFormatDecimalNumber(self, UseThousandSeparator = None, DecimalPlaces = None) -> 'ModelFormatDecimalNumber': ...
    @property
    def ModelFormatPercentageNumber(self, UseThousandSeparator = None, DecimalPlaces = None) -> 'ModelFormatPercentageNumber': ...
    @property
    def ModelFormatScientificNumber(self, DecimalPlaces = None) -> 'ModelFormatScientificNumber': ...
    @property
    def ModelFormatWholeNumber(self, UseThousandSeparator = None) -> 'ModelFormatWholeNumber': ...
    def Refresh(self, ): ...

class Model3DFormat:
    Application = None
    AutoFit: 'MsoTriState'
    CameraPositionX: 'Single'
    CameraPositionY: 'Single'
    CameraPositionZ: 'Single'
    Creator: float
    FieldOfView: 'Single'
    LookAtPointX: 'Single'
    LookAtPointY: 'Single'
    LookAtPointZ: 'Single'
    Parent = None
    RotationX: 'Single'
    RotationY: 'Single'
    RotationZ: 'Single'
    def IncrementRotationX(self, Increment: 'Single'): ...
    def IncrementRotationY(self, Increment: 'Single'): ...
    def IncrementRotationZ(self, Increment: 'Single'): ...
    def ResetModel(self, ResetSize: bool = None): ...

class ModelChanges:
    Application: Application
    ColumnsAdded: 'ModelColumnNames'
    ColumnsChanged: 'ModelColumnChanges'
    ColumnsDeleted: 'ModelColumnNames'
    Creator: 'XlCreator'
    MeasuresAdded: 'ModelMeasureNames'
    Parent = None
    RelationshipChange: bool
    Source: 'XlModelChangeSource'
    TableNamesChanged: 'ModelTableNameChanges'
    TablesAdded: 'ModelTableNames'
    TablesDeleted: 'ModelTableNames'
    TablesModified: 'ModelTableNames'
    UnknownChange: bool

class ModelColumnChange:
    Application: Application
    ColumnName: str
    Creator: 'XlCreator'
    Parent = None
    TableName: str

class ModelColumnChanges:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelColumnChange': ...
    def Item(self, Index) -> ModelColumnChange: ...

class ModelColumnName:
    Application: Application
    ColumnName: str
    Creator: 'XlCreator'
    Parent = None
    TableName: str

class ModelColumnNames:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelColumnName': ...
    def Item(self, Index) -> ModelColumnName: ...

class ModelConnection:
    ADOConnection = None
    Application: Application
    CalculatedMembers: CalculatedMembers
    CommandText = None
    CommandType: 'XlCmdType'
    Creator: 'XlCreator'
    Parent = None

class ModelFormatBoolean:
    Application: Application
    Creator: 'XlCreator'
    Parent = None

class ModelFormatCurrency:
    Application: Application
    Creator: 'XlCreator'
    DecimalPlaces: float
    Parent = None
    Symbol: str

class ModelFormatDate:
    Application: Application
    Creator: 'XlCreator'
    FormatString: str
    Parent = None

class ModelFormatDecimalNumber:
    Application: Application
    Creator: 'XlCreator'
    DecimalPlaces: float
    Parent = None
    UseThousandSeparator: bool

class ModelFormatGeneral:
    Application: Application
    Creator: 'XlCreator'
    Parent = None

class ModelFormatPercentageNumber:
    Application: Application
    Creator: 'XlCreator'
    DecimalPlaces: float
    Parent = None
    UseThousandSeparator: bool

class ModelFormatScientificNumber:
    Application: Application
    Creator: 'XlCreator'
    DecimalPlaces: float
    Parent = None

class ModelFormatWholeNumber:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    UseThousandSeparator: bool

class ModelMeasure:
    Application: Application
    AssociatedTable: 'ModelTable'
    Creator: 'XlCreator'
    Description: str
    FormatInformation = None
    Formula: str
    Name: str
    Parent = None
    def Delete(self, ): ...

class ModelMeasureName:
    Application: Application
    Creator: 'XlCreator'
    MeasureName: str
    Parent = None
    TableName: str

class ModelMeasureNames:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelMeasureName': ...
    def Item(self, Index) -> ModelMeasureName: ...

class ModelMeasures:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelMeasure': ...
    def Add(self, MeasureName: str, AssociatedTable: 'ModelTable', Formula: str, FormatInformation, Description = None) -> ModelMeasure: ...
    def Item(self, Index) -> ModelMeasure: ...

class ModelRelationship:
    Active: bool
    Application: Application
    Creator: 'XlCreator'
    ForeignKeyColumn: 'ModelTableColumn'
    ForeignKeyTable: 'ModelTable'
    Parent = None
    PrimaryKeyColumn: 'ModelTableColumn'
    PrimaryKeyTable: 'ModelTable'
    def Delete(self, ): ...

class ModelRelationships:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelRelationship': ...
    def Add(self, ForeignKeyColumn: 'ModelTableColumn', PrimaryKeyColumn: 'ModelTableColumn') -> ModelRelationship: ...
    def DetectRelationships(self, PivotTable: 'PivotTable'): ...
    def Item(self, Index) -> ModelRelationship: ...

class ModelTable:
    Application: Application
    Creator: 'XlCreator'
    ModelTableColumns: 'ModelTableColumns'
    Name: str
    Parent = None
    RecordCount: float
    SourceName: str
    SourceWorkbookConnection: 'WorkbookConnection'
    def Refresh(self, ): ...

class ModelTableColumn:
    Application: Application
    Creator: 'XlCreator'
    DataType: float
    Name: str
    Parent = None

class ModelTableColumns:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelTableColumn': ...
    def Item(self, Index) -> ModelTableColumn: ...

class ModelTableNameChange:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    TableNameNew: str
    TableNameOld: str

class ModelTableNameChanges:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelTableNameChange': ...
    def Item(self, Index) -> ModelTableNameChange: ...

class ModelTableNames:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'str': ...
    def Item(self, Index) -> str: ...

class ModelTables:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ModelTable': ...
    def Item(self, Index) -> ModelTable: ...

class ModuleView:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    Sheet = None

class MultiThreadedCalculation:
    Application: Application
    Creator: 'XlCreator'
    Enabled: bool
    Parent = None
    ThreadCount: float
    ThreadMode: 'XlThreadMode'

class Name:
    Application: Application
    Category: str
    CategoryLocal: str
    Comment: str
    Creator: 'XlCreator'
    Index: float
    MacroType: 'XlXLMMacroType'
    Name: str
    NameLocal: str
    Parent = None
    RefersTo = None
    RefersToLocal = None
    RefersToR1C1 = None
    RefersToR1C1Local = None
    RefersToRange: 'Range'
    ShortcutKey: str
    ValidWorkbookParameter: bool
    Value: str
    Visible: bool
    WorkbookParameter: bool
    def Delete(self, ): ...

class NamedSheetView:
    Application: Application
    Creator: 'XlCreator'
    Name: str
    Parent = None
    def Activate(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, Name = None) -> 'NamedSheetView': ...

class NamedSheetViewCollection:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def Add(self, Name: str) -> NamedSheetView: ...
    def EnterTemporary(self, ) -> NamedSheetView: ...
    def Exit(self, ): ...
    def GetActive(self, ) -> NamedSheetView: ...
    def GetItem(self, Name: str) -> NamedSheetView: ...
    def GetItemAt(self, Index: float) -> NamedSheetView: ...

class Names:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Name': ...
    def Add(self, Name = None, RefersTo = None, Visible = None, MacroType = None, ShortcutKey = None, Category = None, NameLocal = None, RefersToLocal = None, CategoryLocal = None, RefersToR1C1 = None, RefersToR1C1Local = None) -> Name: ...
    def Item(self, Index = None, IndexLocal = None, RefersTo = None) -> Name: ...

class NegativeBarFormat:
    Application: Application
    BorderColor = None
    BorderColorType: 'XlDataBarNegativeColorType'
    Color = None
    ColorType: 'XlDataBarNegativeColorType'
    Creator: 'XlCreator'
    Parent = None

class ODBCConnection:
    AlwaysUseConnectionFile: bool
    Application: Application
    BackgroundQuery: bool
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    EnableRefresh: bool
    Parent = None
    RefreshDate: 'Date'
    Refreshing: bool
    RefreshOnFileOpen: bool
    RefreshPeriod: float
    RobustConnect: 'XlRobustConnect'
    SavePassword: bool
    ServerCredentialsMethod: 'XlCredentialsMethod'
    ServerSSOApplicationID: str
    SourceConnectionFile: str
    SourceData = None
    SourceDataFile: str
    def CancelRefresh(self, ): ...
    def Refresh(self, ): ...
    def SaveAsODC(self, ODCFileName: str, Description = None, Keywords = None): ...

class ODBCError:
    Application: Application
    Creator: 'XlCreator'
    ErrorString: str
    Parent = None
    SqlState: str

class ODBCErrors:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ODBCError': ...
    def Item(self, Index: float) -> ODBCError: ...

class OLEDBConnection:
    ADOConnection = None
    AlwaysUseConnectionFile: bool
    Application: Application
    BackgroundQuery: bool
    CalculatedMembers: CalculatedMembers
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    EnableRefresh: bool
    IsConnected: bool
    LocalConnection = None
    LocaleID: float
    MaintainConnection: bool
    MaxDrillthroughRecords: float
    OLAP: bool
    Parent = None
    RefreshDate: 'Date'
    Refreshing: bool
    RefreshOnFileOpen: bool
    RefreshPeriod: float
    RetrieveInOfficeUILang: bool
    RobustConnect: 'XlRobustConnect'
    SavePassword: bool
    ServerCredentialsMethod: 'XlCredentialsMethod'
    ServerFillColor: bool
    ServerFontStyle: bool
    ServerNumberFormat: bool
    ServerSSOApplicationID: str
    ServerTextColor: bool
    SourceConnectionFile: str
    SourceDataFile: str
    UseLocalConnection: bool
    def CancelRefresh(self, ): ...
    def MakeConnection(self, ): ...
    def Reconnect(self, ): ...
    def Refresh(self, ): ...
    def SaveAsODC(self, ODCFileName: str, Description = None, Keywords = None): ...

class OLEDBError:
    Application: Application
    Creator: 'XlCreator'
    ErrorString: str
    Native: float
    Number: float
    Parent = None
    SqlState: str
    Stage: float

class OLEDBErrors:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'OLEDBError': ...
    def Item(self, Index: float) -> OLEDBError: ...

class OLEFormat:
    Application: Application
    Creator: 'XlCreator'
    Object = None
    Parent = None
    progID: str
    def Activate(self, ): ...
    def Verb(self, Verb = None): ...

class OLEObject:
    Application: Application
    AutoLoad: bool
    AutoUpdate: bool
    Border: Border
    BottomRightCell: 'Range'
    Creator: 'XlCreator'
    Enabled: bool
    Height: float
    Index: float
    Interior: Interior
    Left: float
    LinkedCell: str
    ListFillRange: str
    Locked: bool
    Name: str
    Object = None
    OLEType = None
    Parent = None
    Placement = None
    PrintObject: bool
    progID: str
    Shadow: bool
    ShapeRange: 'ShapeRange'
    SourceName: str
    Top: float
    TopLeftCell: 'Range'
    Visible: bool
    Width: float
    ZOrder: float
    def Activate(self, ): ...
    def BringToFront(self, ): ...
    def Copy(self, ): ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None): ...
    def Cut(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, ): ...
    def Select(self, Replace = None): ...
    def SendToBack(self, ): ...
    def Update(self, ): ...
    def Verb(self, Verb: 'XlOLEVerb' = None): ...

class OLEObjects:
    Application: Application
    AutoLoad: bool
    Border: Border
    Count: float
    Creator: 'XlCreator'
    Enabled: bool
    Height: float
    Interior: Interior
    Left: float
    Locked: bool
    Parent = None
    Placement = None
    PrintObject: bool
    Shadow: bool
    ShapeRange: 'ShapeRange'
    SourceName: str
    Top: float
    Visible: bool
    Width: float
    ZOrder: float
    def __call__(self, Index) -> 'None': ...
    def Add(self, ClassType = None, Filename = None, Link = None, DisplayAsIcon = None, IconFileName = None, IconIndex = None, IconLabel = None, Left = None, Top = None, Width = None, Height = None) -> OLEObject: ...
    def BringToFront(self, ): ...
    def Copy(self, ): ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None): ...
    def Cut(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, ): ...
    def Item(self, Index): ...
    def Select(self, Replace = None): ...
    def SendToBack(self, ): ...

class Outline:
    Application: Application
    AutomaticStyles: bool
    Creator: 'XlCreator'
    Parent = None
    SummaryColumn: 'XlSummaryColumn'
    SummaryRow: 'XlSummaryRow'
    def ShowLevels(self, RowLevels = None, ColumnLevels = None): ...

class Page:
    CenterFooter: HeaderFooter
    CenterHeader: HeaderFooter
    LeftFooter: HeaderFooter
    LeftHeader: HeaderFooter
    RightFooter: HeaderFooter
    RightHeader: HeaderFooter

class Pages:
    Count: float
    def __call__(self, Index) -> 'Page': ...
    @property
    def Item(self, Index) -> Page: ...

class PageSetup:
    AlignMarginsHeaderFooter: bool
    Application: Application
    BlackAndWhite: bool
    BottomMargin: float
    CenterFooter: str
    CenterFooterPicture: Graphic
    CenterHeader: str
    CenterHeaderPicture: Graphic
    CenterHorizontally: bool
    CenterVertically: bool
    Creator: 'XlCreator'
    DifferentFirstPageHeaderFooter: bool
    Draft: bool
    EvenPage: Page
    FirstPage: Page
    FirstPageNumber: float
    FitToPagesTall = None
    FitToPagesWide = None
    FooterMargin: float
    HeaderMargin: float
    LeftFooter: str
    LeftFooterPicture: Graphic
    LeftHeader: str
    LeftHeaderPicture: Graphic
    LeftMargin: float
    OddAndEvenPagesHeaderFooter: bool
    Order: 'XlOrder'
    Orientation: 'XlPageOrientation'
    Pages: Pages
    PaperSize: 'XlPaperSize'
    Parent = None
    PrintArea: str
    PrintComments: 'XlPrintLocation'
    PrintErrors: 'XlPrintErrors'
    PrintGridlines: bool
    PrintHeadings: bool
    PrintNotes: bool
    PrintTitleColumns: str
    PrintTitleRows: str
    RightFooter: str
    RightFooterPicture: Graphic
    RightHeader: str
    RightHeaderPicture: Graphic
    RightMargin: float
    ScaleWithDocHeaderFooter: bool
    TopMargin: float
    Zoom = None
    @property
    def PrintQuality(self, Index = None): ...

class Pane:
    Application: Application
    Creator: 'XlCreator'
    Index: float
    Parent = None
    ScrollColumn: float
    ScrollRow: float
    VisibleRange: 'Range'
    def Activate(self, ) -> bool: ...
    def LargeScroll(self, Down = None, Up = None, ToRight = None, ToLeft = None): ...
    def PointsToScreenPixelsX(self, Points: float) -> float: ...
    def PointsToScreenPixelsY(self, Points: float) -> float: ...
    def ScrollIntoView(self, Left: float, Top: float, Width: float, Height: float, Start = None): ...
    def SmallScroll(self, Down = None, Up = None, ToRight = None, ToLeft = None): ...

class Panes:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Pane': ...
    @property
    def Item(self, Index: float) -> Pane: ...

class Parameter:
    Application: Application
    Creator: 'XlCreator'
    DataType: 'XlParameterDataType'
    Name: str
    Parent = None
    PromptString: str
    RefreshOnChange: bool
    SourceRange: 'Range'
    Type: 'XlParameterType'
    Value = None
    def SetParam(self, Type: 'XlParameterType', Value): ...

class Parameters:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Parameter': ...
    def Add(self, Name: str, iDataType = None) -> Parameter: ...
    def Delete(self, ): ...
    def Item(self, Index) -> Parameter: ...

class Phonetic:
    Alignment: float
    Application: Application
    CharacterType: float
    Creator: 'XlCreator'
    Font: Font
    Parent = None
    Text: str
    Visible: bool

class Phonetics:
    Alignment: float
    Application: Application
    CharacterType: float
    Count: float
    Creator: 'XlCreator'
    Font: Font
    Length: float
    Parent = None
    Start: float
    Text: str
    Visible: bool
    def __call__(self, Index) -> 'None': ...
    def Add(self, Start: float, Length: float, Text: str): ...
    def Delete(self, ): ...
    @property
    def Item(self, Index: float): ...

class PictureFormat:
    Application = None
    Brightness: 'Single'
    ColorType: 'MsoPictureColorType'
    Contrast: 'Single'
    Creator: float
    Crop: 'Crop'
    CropBottom: 'Single'
    CropLeft: 'Single'
    CropRight: 'Single'
    CropTop: 'Single'
    Parent = None
    TransparencyColor: 'MsoRGBType'
    TransparentBackground: 'MsoTriState'
    def IncrementBrightness(self, Increment: 'Single'): ...
    def IncrementContrast(self, Increment: 'Single'): ...

class PivotAxis:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    PivotLines: 'PivotLines'

class PivotCache:
    ADOConnection = None
    Application: Application
    BackgroundQuery: bool
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    EnableRefresh: bool
    Index: float
    IsConnected: bool
    LocalConnection = None
    MaintainConnection: bool
    MemoryUsed: float
    MissingItemsLimit: 'XlPivotTableMissingItems'
    OLAP: bool
    OptimizeCache: bool
    Parent = None
    QueryType: 'XlQueryType'
    RecordCount: float
    Recordset = None
    RefreshDate: 'Date'
    RefreshName: str
    RefreshOnFileOpen: bool
    RefreshPeriod: float
    RobustConnect: 'XlRobustConnect'
    SavePassword: bool
    SourceConnectionFile: str
    SourceData = None
    SourceDataFile: str
    SourceType: 'XlPivotTableSourceType'
    UpgradeOnRefresh: bool
    UseLocalConnection: bool
    Version: 'XlPivotTableVersionList'
    WorkbookConnection: 'WorkbookConnection'
    def CreatePivotChart(self, ChartDestination, XlChartType = None, Left = None, Top = None, Width = None, Height = None) -> 'Shape': ...
    def CreatePivotTable(self, TableDestination, TableName = None, ReadData = None, DefaultVersion = None) -> 'PivotTable': ...
    def MakeConnection(self, ): ...
    def Refresh(self, ): ...
    def ResetTimer(self, ): ...
    def SaveAsODC(self, ODCFileName: str, Description = None, Keywords = None): ...

class PivotCaches:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotCache': ...
    def Create(self, SourceType: 'XlPivotTableSourceType', SourceData = None, Version = None) -> PivotCache: ...
    def Item(self, Index) -> PivotCache: ...

class PivotCell:
    Application: Application
    CellChanged: 'XlCellChangedState'
    ColumnItems: 'PivotItemList'
    Creator: 'XlCreator'
    CustomSubtotalFunction: 'XlConsolidationFunction'
    DataField: 'PivotField'
    DataSourceValue = None
    MDX: str
    Parent = None
    PivotCellType: 'XlPivotCellType'
    PivotColumnLine: 'PivotLine'
    PivotField: 'PivotField'
    PivotItem: 'PivotItem'
    PivotRowLine: 'PivotLine'
    PivotTable: 'PivotTable'
    Range: 'Range'
    RowItems: 'PivotItemList'
    ServerActions: Actions
    def AllocateChange(self, ): ...
    def DiscardChange(self, ): ...

class PivotField:
    AllItemsVisible: bool
    Application: Application
    AutoShowCount: float
    AutoShowField: str
    AutoShowRange: float
    AutoShowType: float
    AutoSortCustomSubtotal: float
    AutoSortField: str
    AutoSortOrder: float
    AutoSortPivotLine: 'PivotLine'
    BaseField = None
    BaseItem = None
    Calculation: 'XlPivotFieldCalculation'
    Caption: str
    ChildField: 'PivotField'
    Creator: 'XlCreator'
    CubeField: CubeField
    CurrentPage = None
    CurrentPageList = None
    CurrentPageName: str
    DatabaseSort: bool
    DataRange: 'Range'
    DataType: 'XlPivotFieldDataType'
    DisplayAsCaption: bool
    DisplayAsTooltip: bool
    DisplayInReport: bool
    DragToColumn: bool
    DragToData: bool
    DragToHide: bool
    DragToPage: bool
    DragToRow: bool
    DrilledDown: bool
    EnableItemSelection: bool
    EnableMultiplePageItems: bool
    Formula: str
    Function: 'XlConsolidationFunction'
    GroupLevel = None
    Hidden: bool
    HiddenItemsList = None
    IncludeNewItemsInFilter: bool
    IsCalculated: bool
    IsMemberProperty: bool
    LabelRange: 'Range'
    LayoutBlankLine: bool
    LayoutCompactRow: bool
    LayoutForm: 'XlLayoutFormType'
    LayoutPageBreak: bool
    LayoutSubtotalLocation: 'XlSubtototalLocationType'
    MemberPropertyCaption: str
    MemoryUsed: float
    Name: str
    NumberFormat: str
    Orientation: 'XlPivotFieldOrientation'
    Parent = None
    ParentField: 'PivotField'
    PivotFilters: 'PivotFilters'
    Position = None
    PropertyOrder: float
    PropertyParentField: 'PivotField'
    RepeatLabels: bool
    ServerBased: bool
    ShowAllItems: bool
    ShowDetail: bool
    ShowingInAxis: bool
    SourceCaption: str
    SourceName: str
    StandardFormula: str
    SubtotalName: str
    TotalLevels = None
    UseMemberPropertyAsCaption: bool
    Value: str
    VisibleItemsList = None
    def AddPageItem(self, Item: str, ClearList = None): ...
    def AutoGroup(self, ): ...
    def AutoShow(self, Type: float, Range: float, Count: float, Field: str): ...
    def AutoSort(self, Order: float, Field: str, PivotLine = None, CustomSubtotal = None): ...
    def CalculatedItems(self, ) -> CalculatedItems: ...
    @property
    def ChildItems(self, Index = None): ...
    def ClearAllFilters(self, ): ...
    def ClearLabelFilters(self, ): ...
    def ClearManualFilter(self, ): ...
    def ClearValueFilters(self, ): ...
    def Delete(self, ): ...
    def DrillTo(self, Field: str): ...
    @property
    def HiddenItems(self, Index = None): ...
    @property
    def ParentItems(self, Index = None): ...
    def PivotItems(self, Index = None): ...
    @property
    def Subtotals(self, Index = None): ...
    @property
    def VisibleItems(self, Index = None): ...

class PivotFields:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent: 'PivotTable'
    def __call__(self, Index) -> 'PivotField': ...
    def Item(self, Index) -> PivotField: ...

class PivotFilter:
    Active: bool
    Application: Application
    Creator: 'XlCreator'
    DataCubeField: CubeField
    DataField: PivotField
    Description: str
    FilterType: 'XlPivotFilterType'
    IsMemberPropertyFilter: bool
    MemberPropertyField: PivotField
    Name: str
    Order: float
    Parent = None
    PivotField: PivotField
    Value1 = None
    Value2 = None
    WholeDayFilter: bool
    def Delete(self, ): ...

class PivotFilters:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotFilter': ...
    def Add2(self, Type: 'XlPivotFilterType', DataField = None, Value1 = None, Value2 = None, Order = None, Name = None, Description = None, MemberPropertyField = None, WholeDayFilter = None) -> PivotFilter: ...
    @property
    def Item(self, Index) -> PivotFilter: ...

class PivotFormula:
    Application: Application
    Creator: 'XlCreator'
    Formula: str
    Index: float
    Parent = None
    StandardFormula: str
    Value: str
    def Delete(self, ): ...

class PivotFormulas:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotFormula': ...
    def Add(self, Formula: str, UseStandardFormula = None) -> PivotFormula: ...
    def Item(self, Index) -> PivotFormula: ...

class PivotItem:
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    DataRange: 'Range'
    DrilledDown: bool
    Formula: str
    IsCalculated: bool
    LabelRange: 'Range'
    Name: str
    Parent: PivotField
    ParentItem: 'PivotItem'
    ParentShowDetail: bool
    Position: float
    RecordCount: float
    ShowDetail: bool
    SourceName = None
    SourceNameStandard: str
    StandardFormula: str
    Value: str
    Visible: bool
    @property
    def ChildItems(self, Index = None): ...
    def Delete(self, ): ...
    def DrillTo(self, Field: str): ...

class PivotItemList:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotItem': ...
    def Item(self, Index) -> PivotItem: ...

class PivotItems:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent: PivotField
    def __call__(self, Index) -> 'PivotItem': ...
    def Add(self, Name: str): ...
    def Item(self, Index) -> PivotItem: ...

class PivotLayout:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    PivotTable: 'PivotTable'

class PivotLine:
    Application: Application
    Creator: 'XlCreator'
    LineType: 'XlPivotLineType'
    Parent = None
    PivotLineCells: 'PivotLineCells'
    PivotLineCellsFull: 'PivotLineCells'
    Position: float

class PivotLineCells:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Full: bool
    Parent = None
    def __call__(self, Index) -> 'PivotCell': ...
    @property
    def Item(self, Index) -> PivotCell: ...

class PivotLines:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotLine': ...
    @property
    def Item(self, Index) -> PivotLine: ...

class PivotTable:
    ActiveFilters: PivotFilters
    Allocation: 'XlAllocation'
    AllocationMethod: 'XlAllocationMethod'
    AllocationValue: 'XlAllocationValue'
    AllocationWeightExpression: str
    AllowMultipleFilters: bool
    AlternativeText: str
    Application: Application
    CacheIndex: float
    CalculatedMembers: CalculatedMembers
    CalculatedMembersInFilters: bool
    ChangeList: 'PivotTableChangeList'
    ColumnGrand: bool
    ColumnRange: 'Range'
    CompactLayoutColumnHeader: str
    CompactLayoutRowHeader: str
    CompactRowIndent: float
    Creator: 'XlCreator'
    CubeFields: CubeFields
    DataBodyRange: 'Range'
    DataLabelRange: 'Range'
    DataPivotField: PivotField
    DisplayContextTooltips: bool
    DisplayEmptyColumn: bool
    DisplayEmptyRow: bool
    DisplayErrorString: bool
    DisplayFieldCaptions: bool
    DisplayImmediateItems: bool
    DisplayMemberPropertyTooltips: bool
    DisplayNullString: bool
    EnableDataValueEditing: bool
    EnableDrilldown: bool
    EnableFieldDialog: bool
    EnableFieldList: bool
    EnableWizard: bool
    EnableWriteback: bool
    ErrorString: str
    FieldListSortAscending: bool
    GrandTotalName: str
    HasAutoFormat: bool
    Hidden: bool
    InGridDropZones: bool
    InnerDetail: str
    LayoutRowDefault: 'XlLayoutRowType'
    Location: str
    ManualUpdate: bool
    MDX: str
    MergeLabels: bool
    Name: str
    NullString: str
    PageFieldOrder: float
    PageFieldStyle: str
    PageFieldWrapCount: float
    PageRange: 'Range'
    PageRangeCells: 'Range'
    Parent = None
    PivotChart: 'Shape'
    PivotColumnAxis: PivotAxis
    PivotFormulas: PivotFormulas
    PivotRowAxis: PivotAxis
    PivotSelection: str
    PivotSelectionStandard: str
    PreserveFormatting: bool
    PrintDrillIndicators: bool
    PrintTitles: bool
    RefreshDate: 'Date'
    RefreshName: str
    RepeatItemsOnEachPrintedPage: bool
    RowGrand: bool
    RowRange: 'Range'
    SaveData: bool
    SelectionMode: 'XlPTSelectionMode'
    ShowDrillIndicators: bool
    ShowPageMultipleItemLabel: bool
    ShowTableStyleColumnHeaders: bool
    ShowTableStyleColumnStripes: bool
    ShowTableStyleLastColumn: bool
    ShowTableStyleRowHeaders: bool
    ShowTableStyleRowStripes: bool
    ShowValuesRow: bool
    Slicers: 'Slicers'
    SmallGrid: bool
    SortUsingCustomLists: bool
    SourceData = None
    SubtotalHiddenPageItems: bool
    Summary: str
    TableRange1: 'Range'
    TableRange2: 'Range'
    TableStyle2 = None
    Tag: str
    TotalsAnnotation: bool
    VacatedStyle: str
    Value: str
    Version: 'XlPivotTableVersionList'
    ViewCalculatedMembers: bool
    VisualTotals: bool
    VisualTotalsForSets: bool
    def AddDataField(self, Field, Caption = None, Function = None) -> PivotField: ...
    def AddFields(self, RowFields = None, ColumnFields = None, PageFields = None, AddToTable = None): ...
    def AllocateChanges(self, ): ...
    def ApplyLayout(self, ): ...
    def CalculatedFields(self, ) -> CalculatedFields: ...
    def ChangeConnection(self, conn: 'WorkbookConnection'): ...
    def ChangePivotCache(self, PivotCache): ...
    def ClearAllFilters(self, ): ...
    def ClearTable(self, ): ...
    @property
    def ColumnFields(self, Index = None): ...
    def CommitChanges(self, ): ...
    def ConvertToFormulas(self, ConvertFilters: bool): ...
    def CreateCubeFile(self, File: str, Measures = None, Levels = None, Members = None, Properties = None) -> str: ...
    @property
    def DataFields(self, Index = None): ...
    def DiscardChanges(self, ): ...
    def DrillDown(self, PivotItem: PivotItem, PivotLine = None): ...
    def DrillTo(self, PivotItem: PivotItem, CubeField: CubeField, PivotLine = None): ...
    def DrillUp(self, PivotItem: PivotItem, PivotLine = None, LevelUniqueName = None): ...
    def GetData(self, Name: str) -> float: ...
    def GetPivotData(self, DataField = None, Field1 = None, Item1 = None, Field2 = None, Item2 = None, Field3 = None, Item3 = None, Field4 = None, Item4 = None, Field5 = None, Item5 = None, Field6 = None, Item6 = None, Field7 = None, Item7 = None, Field8 = None, Item8 = None, Field9 = None, Item9 = None, Field10 = None, Item10 = None, Field11 = None, Item11 = None, Field12 = None, Item12 = None, Field13 = None, Item13 = None, Field14 = None, Item14 = None) -> 'Range': ...
    @property
    def HiddenFields(self, Index = None): ...
    def ListFormulas(self, ): ...
    @property
    def PageFields(self, Index = None): ...
    def PivotCache(self, ) -> PivotCache: ...
    def PivotFields(self, Index = None): ...
    def PivotSelect(self, Name: str, Mode: 'XlPTSelectionMode' = None, UseStandardName = None): ...
    def PivotTableWizard(self, SourceType = None, SourceData = None, TableDestination = None, TableName = None, RowGrand = None, ColumnGrand = None, SaveData = None, HasAutoFormat = None, AutoPage = None, Reserved = None, BackgroundQuery = None, OptimizeCache = None, PageFieldOrder = None, PageFieldWrapCount = None, ReadData = None, Connection = None): ...
    def PivotValueCell(self, rowline = None, columnline = None) -> 'PivotValueCell': ...
    def RefreshDataSourceValues(self, ): ...
    def RefreshTable(self, ) -> bool: ...
    def RepeatAllLabels(self, Repeat: 'XlPivotFieldRepeatLabels'): ...
    def RowAxisLayout(self, RowLayout: 'XlLayoutRowType'): ...
    @property
    def RowFields(self, Index = None): ...
    def ShowPages(self, PageField = None): ...
    def SubtotalLocation(self, Location: 'XlSubtototalLocationType'): ...
    def Update(self, ): ...
    @property
    def VisibleFields(self, Index = None): ...

class PivotTableChangeList:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ValueChange': ...
    def Add(self, Tuple: str, Value: float, AllocationValue = None, AllocationMethod = None, AllocationWeightExpression = None) -> 'ValueChange': ...
    @property
    def Item(self, Index) -> 'ValueChange': ...

class PivotTables:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotTable': ...
    def Add(self, PivotCache: PivotCache, TableDestination, TableName = None, ReadData = None, DefaultVersion = None) -> PivotTable: ...
    def Item(self, Index) -> PivotTable: ...

class PivotValueCell:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    PivotCell: PivotCell
    ServerActions: Actions
    Value = None
    def ShowDetail(self, ): ...

class PlotArea:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Height: float
    InsideHeight: float
    InsideLeft: float
    InsideTop: float
    InsideWidth: float
    Left: float
    Name: str
    Parent = None
    Position: 'XlChartElementPosition'
    Top: float
    Width: float
    def ClearFormats(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class Point:
    Application: Application
    ApplyPictToEnd: bool
    ApplyPictToFront: bool
    ApplyPictToSides: bool
    Creator: 'XlCreator'
    DataLabel: DataLabel
    Explosion: float
    Format: ChartFormat
    Has3DEffect: bool
    HasDataLabel: bool
    Height: float
    InvertIfNegative: bool
    IsTotal: bool
    Left: float
    MarkerBackgroundColor: float
    MarkerBackgroundColorIndex: 'XlColorIndex'
    MarkerForegroundColor: float
    MarkerForegroundColorIndex: 'XlColorIndex'
    MarkerSize: float
    MarkerStyle: 'XlMarkerStyle'
    Name: str
    Parent = None
    PictureType: 'XlChartPictureType'
    PictureUnit2: float
    SecondaryPlot: bool
    Shadow: bool
    Top: float
    Width: float
    def ApplyDataLabels(self, Type: 'XlDataLabelsType' = None, LegendKey = None, AutoText = None, HasLeaderLines = None, ShowSeriesName = None, ShowCategoryName = None, ShowValue = None, ShowPercentage = None, ShowBubbleSize = None, Separator = None): ...
    def ClearFormats(self, ): ...
    def Copy(self, ): ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Paste(self, ): ...
    def PieSliceLocation(self, loc: 'XlPieSliceLocation', Index: 'XlPieSliceIndex' = None) -> float: ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class Points:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Point': ...
    def Item(self, Index: float) -> Point: ...

class ProtectedViewWindow:
    Caption: str
    EnableResize: bool
    Height: float
    Left: float
    SourceName: str
    SourcePath: str
    Top: float
    Visible: bool
    Width: float
    WindowState: 'XlProtectedViewWindowState'
    Workbook: 'Workbook'
    def Activate(self, ): ...
    def Close(self, ) -> bool: ...
    def Edit(self, WriteResPassword = None, UpdateLinks = None) -> 'Workbook': ...

class ProtectedViewWindows:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'ProtectedViewWindow': ...
    @property
    def Item(self, Index) -> ProtectedViewWindow: ...
    def Open(self, Filename: str, Password = None, AddToMru = None, RepairMode = None) -> ProtectedViewWindow: ...

class Protection:
    AllowDeletingColumns: bool
    AllowDeletingRows: bool
    AllowEditRanges: AllowEditRanges
    AllowFiltering: bool
    AllowFormattingCells: bool
    AllowFormattingColumns: bool
    AllowFormattingRows: bool
    AllowInsertingColumns: bool
    AllowInsertingHyperlinks: bool
    AllowInsertingRows: bool
    AllowSorting: bool
    AllowUsingPivotTables: bool

class PublishObject:
    Application: Application
    AutoRepublish: bool
    Creator: 'XlCreator'
    DivID: str
    Filename: str
    HtmlType: 'XlHtmlType'
    Parent = None
    Sheet: str
    Source: str
    SourceType: 'XlSourceType'
    Title: str
    def Delete(self, ): ...
    def Publish(self, Create = None): ...

class PublishObjects:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PublishObject': ...
    def Add(self, SourceType: 'XlSourceType', Filename: str, Sheet = None, Source = None, HtmlType = None, DivID = None, Title = None) -> PublishObject: ...
    def Delete(self, ): ...
    @property
    def Item(self, Index) -> PublishObject: ...
    def Publish(self, ): ...

class Queries:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    FastCombine: bool
    Parent = None
    def __call__(self, Index) -> 'WorkbookQuery': ...
    def Add(self, Name: str, Formula: str, Description = None) -> 'WorkbookQuery': ...
    def Item(self, NameOrIndex) -> 'WorkbookQuery': ...

class QueryTable:
    AdjustColumnWidth: bool
    Application: Application
    BackgroundQuery: bool
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    Destination: 'Range'
    EditWebPage = None
    EnableEditing: bool
    EnableRefresh: bool
    FetchedRowOverflow: bool
    FieldNames: bool
    FillAdjacentFormulas: bool
    ListObject: ListObject
    MaintainConnection: bool
    Name: str
    Parameters: Parameters
    Parent = None
    PostText: str
    PreserveColumnInfo: bool
    PreserveFormatting: bool
    QueryType: 'XlQueryType'
    Recordset = None
    Refreshing: bool
    RefreshOnFileOpen: bool
    RefreshPeriod: float
    RefreshStyle: 'XlCellInsertionMode'
    ResultRange: 'Range'
    RobustConnect: 'XlRobustConnect'
    RowNumbers: bool
    SaveData: bool
    SavePassword: bool
    Sort: 'Sort'
    SourceConnectionFile: str
    SourceDataFile: str
    TextFileColumnDataTypes = None
    TextFileCommaDelimiter: bool
    TextFileConsecutiveDelimiter: bool
    TextFileDecimalSeparator: str
    TextFileFixedColumnWidths = None
    TextFileOtherDelimiter: str
    TextFileParseType: 'XlTextParsingType'
    TextFilePlatform: float
    TextFilePromptOnRefresh: bool
    TextFileSemicolonDelimiter: bool
    TextFileSpaceDelimiter: bool
    TextFileStartRow: float
    TextFileTabDelimiter: bool
    TextFileTextQualifier: 'XlTextQualifier'
    TextFileThousandsSeparator: str
    TextFileTrailingMinusNumbers: bool
    TextFileVisualLayout: 'XlTextVisualLayoutType'
    WebConsecutiveDelimitersAsOne: bool
    WebDisableDateRecognition: bool
    WebDisableRedirections: bool
    WebFormatting: 'XlWebFormatting'
    WebPreFormattedTextToColumns: bool
    WebSelectionType: 'XlWebSelectionType'
    WebSingleBlockTextImport: bool
    WebTables: str
    WorkbookConnection: 'WorkbookConnection'
    def CancelRefresh(self, ): ...
    def Delete(self, ): ...
    def Refresh(self, BackgroundQuery = None) -> bool: ...
    def ResetTimer(self, ): ...
    def SaveAsODC(self, ODCFileName: str, Description = None, Keywords = None): ...

class QueryTables:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'QueryTable': ...
    def Add(self, Connection, Destination: 'Range', Sql = None) -> QueryTable: ...
    def Item(self, Index) -> QueryTable: ...

class QuickAnalysis:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def Hide(self, XlQuickAnalysisMode: 'XlQuickAnalysisMode' = None): ...
    def Show(self, XlQuickAnalysisMode: 'XlQuickAnalysisMode' = None): ...

class Range:
    AddIndent = None
    AllowEdit: bool
    Application: Application
    Areas: Areas
    Borders: Borders
    Cells: 'Range'
    Column: float
    Columns: 'Range'
    ColumnWidth = None
    Comment: Comment
    CommentThreaded: CommentThreaded
    Count: float
    CountLarge = None
    Creator: 'XlCreator'
    CurrentArray: 'Range'
    CurrentRegion: 'Range'
    Dependents: 'Range'
    DirectDependents: 'Range'
    DirectPrecedents: 'Range'
    DisplayFormat: DisplayFormat
    EntireColumn: 'Range'
    EntireRow: 'Range'
    Errors: Errors
    Font: Font
    FormatConditions: FormatConditions
    Formula = None
    Formula2 = None
    Formula2Local = None
    Formula2R1C1 = None
    Formula2R1C1Local = None
    FormulaArray = None
    FormulaHidden = None
    FormulaLocal = None
    FormulaR1C1 = None
    FormulaR1C1Local = None
    HasArray = None
    HasFormula = None
    HasRichDataType = None
    HasSpill = None
    Height = None
    Hidden = None
    HorizontalAlignment = None
    Hyperlinks: Hyperlinks
    ID: str
    IndentLevel = None
    Interior: Interior
    Left = None
    LinkedDataTypeState = None
    ListHeaderRows: float
    ListObject: ListObject
    LocationInTable: 'XlLocationInTable'
    Locked = None
    MDX: str
    MergeArea: 'Range'
    MergeCells = None
    Name = None
    Next: 'Range'
    NumberFormat = None
    NumberFormatLocal = None
    Orientation = None
    OutlineLevel = None
    PageBreak: float
    Parent = None
    Phonetic: Phonetic
    Phonetics: Phonetics
    PivotCell: PivotCell
    PivotField: PivotField
    PivotItem: PivotItem
    PivotTable: PivotTable
    Precedents: 'Range'
    PrefixCharacter = None
    Previous: 'Range'
    QueryTable: QueryTable
    ReadingOrder: float
    Row: float
    RowHeight = None
    Rows: 'Range'
    SavedAsArray = None
    ServerActions: Actions
    ShowDetail = None
    ShrinkToFit = None
    SoundNote: 'SoundNote'
    SparklineGroups: 'SparklineGroups'
    SpillingToRange: 'Range'
    SpillParent: 'Range'
    Style = None
    Summary = None
    Text = None
    Top = None
    UseStandardHeight = None
    UseStandardWidth = None
    Validation: 'Validation'
    Value2 = None
    VerticalAlignment = None
    Width = None
    Worksheet: 'Worksheet'
    WrapText = None
    XPath: 'XPath'
    def __call__(self, Cell1, Cell2 = None) -> 'Range': ...
    def Activate(self, ): ...
    def AddComment(self, Text = None) -> Comment: ...
    def AddCommentThreaded(self, Text: str) -> CommentThreaded: ...
    @property
    def Address(self, RowAbsolute = None, ColumnAbsolute = None, ReferenceStyle: 'XlReferenceStyle' = None, External = None, RelativeTo = None) -> str: ...
    @property
    def AddressLocal(self, RowAbsolute = None, ColumnAbsolute = None, ReferenceStyle: 'XlReferenceStyle' = None, External = None, RelativeTo = None) -> str: ...
    def AdvancedFilter(self, Action: 'XlFilterAction', CriteriaRange = None, CopyToRange = None, Unique = None): ...
    def AllocateChanges(self, ): ...
    def ApplyNames(self, Names = None, IgnoreRelativeAbsolute = None, UseRowColumnNames = None, OmitColumn = None, OmitRow = None, Order: 'XlApplyNamesOrder' = None, AppendLast = None): ...
    def ApplyOutlineStyles(self, ): ...
    def AutoComplete(self, String: str) -> str: ...
    def AutoFill(self, Destination: 'Range', Type: 'XlAutoFillType' = None): ...
    def AutoFilter(self, Field = None, Criteria1 = None, Operator: 'XlAutoFilterOperator' = None, Criteria2 = None, VisibleDropDown = None, SubField = None): ...
    def AutoFit(self, ): ...
    def AutoOutline(self, ): ...
    def BorderAround(self, LineStyle = None, Weight: 'XlBorderWeight' = None, ColorIndex: 'XlColorIndex' = None, Color = None, ThemeColor = None): ...
    def Calculate(self, ): ...
    def CalculateRowMajorOrder(self, ): ...
    @property
    def Characters(self, Start = None, Length = None) -> Characters: ...
    def CheckSpelling(self, CustomDictionary = None, IgnoreUppercase = None, AlwaysSuggest = None, SpellLang = None): ...
    def Clear(self, ): ...
    def ClearComments(self, ): ...
    def ClearContents(self, ): ...
    def ClearFormats(self, ): ...
    def ClearHyperlinks(self, ): ...
    def ClearNotes(self, ): ...
    def ClearOutline(self, ): ...
    def ColumnDifferences(self, Comparison) -> 'Range': ...
    def Consolidate(self, Sources = None, Function = None, TopRow = None, LeftColumn = None, CreateLinks = None): ...
    def ConvertToLinkedDataType(self, ServiceID: float, LanguageCulture: str): ...
    def Copy(self, Destination = None): ...
    def CopyFromRecordset(self, Data: 'Unknown', MaxRows = None, MaxColumns = None) -> float: ...
    def CopyPicture(self, Appearance: 'XlPictureAppearance' = None, Format: 'XlCopyPictureFormat' = None): ...
    def CreateNames(self, Top = None, Left = None, Bottom = None, Right = None): ...
    def Cut(self, Destination = None): ...
    def DataSeries(self, Rowcol = None, Type: 'XlDataSeriesType' = None, Date: 'XlDataSeriesDate' = None, Step = None, Stop = None, Trend = None): ...
    def DataTypeToText(self, ): ...
    def Delete(self, Shift = None): ...
    def DialogBox(self, ): ...
    def Dirty(self, ): ...
    def DiscardChanges(self, ): ...
    def EditionOptions(self, Type: 'XlEditionType', Option: 'XlEditionOptionsOption', Name = None, Reference = None, Appearance: 'XlPictureAppearance' = None, ChartSize: 'XlPictureAppearance' = None, Format = None): ...
    @property
    def End(self, Direction: 'XlDirection') -> 'Range': ...
    def ExportAsFixedFormat(self, Type: 'XlFixedFormatType', Filename = None, Quality = None, IncludeDocProperties = None, IgnorePrintAreas = None, From = None, To = None, OpenAfterPublish = None, FixedFormatExtClassPtr = None, WorkIdentity = None): ...
    def FillDown(self, ): ...
    def FillLeft(self, ): ...
    def FillRight(self, ): ...
    def FillUp(self, ): ...
    def Find(self, What, After = None, LookIn = None, LookAt = None, SearchOrder = None, SearchDirection: 'XlSearchDirection' = None, MatchCase = None, MatchByte = None, SearchFormat = None) -> 'Range': ...
    def FindNext(self, After = None) -> 'Range': ...
    def FindPrevious(self, After = None) -> 'Range': ...
    def FlashFill(self, ): ...
    def FunctionWizard(self, ): ...
    def Group(self, Start = None, End = None, By = None, Periods = None): ...
    def Insert(self, Shift = None, CopyOrigin = None): ...
    def InsertIndent(self, InsertAmount: float): ...
    @property
    def Item(self, RowIndex, ColumnIndex = None) -> 'Range': ...
    def Justify(self, ): ...
    def ListNames(self, ): ...
    def Merge(self, Across = None): ...
    def NavigateArrow(self, TowardPrecedent = None, ArrowNumber = None, LinkNumber = None): ...
    def NoteText(self, Text = None, Start = None, Length = None) -> str: ...
    @property
    def Offset(self, RowOffset = None, ColumnOffset = None) -> 'Range': ...
    def Parse(self, ParseLine = None, Destination = None): ...
    def PasteSpecial(self, Paste: 'XlPasteType' = None, Operation: 'XlPasteSpecialOperation' = None, SkipBlanks = None, Transpose = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    @property
    def Range(self, Cell1, Cell2 = None) -> 'Range': ...
    def RefreshLinkedDataType(self, DomainID = None): ...
    def RemoveDuplicates(self, Columns = None, Header: 'XlYesNoGuess' = None): ...
    def RemoveSubtotal(self, ): ...
    def Replace(self, What, Replacement, LookAt = None, SearchOrder = None, MatchCase = None, MatchByte = None, SearchFormat = None, ReplaceFormat = None, FormulaVersion = None) -> bool: ...
    @property
    def Resize(self, RowSize = None, ColumnSize = None) -> 'Range': ...
    def RowDifferences(self, Comparison) -> 'Range': ...
    def Run(self, Arg1 = None, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None): ...
    def Select(self, ): ...
    def SetCellDataTypeFromCell(self, SourceCell: 'Range'): ...
    def SetPhonetic(self, ): ...
    def Show(self, ): ...
    def ShowCard(self, ): ...
    def ShowDependents(self, Remove = None): ...
    def ShowErrors(self, ): ...
    def ShowPrecedents(self, Remove = None): ...
    def Sort(self, Key1 = None, Order1: 'XlSortOrder' = None, Key2 = None, Type = None, Order2: 'XlSortOrder' = None, Key3 = None, Order3: 'XlSortOrder' = None, Header: 'XlYesNoGuess' = None, OrderCustom = None, MatchCase = None, Orientation: 'XlSortOrientation' = None, SortMethod: 'XlSortMethod' = None, DataOption1: 'XlSortDataOption' = None, DataOption2: 'XlSortDataOption' = None, DataOption3: 'XlSortDataOption' = None, SubField1 = None): ...
    def SortSpecial(self, SortMethod: 'XlSortMethod' = None, Key1 = None, Order1: 'XlSortOrder' = None, Type = None, Key2 = None, Order2: 'XlSortOrder' = None, Key3 = None, Order3: 'XlSortOrder' = None, Header: 'XlYesNoGuess' = None, OrderCustom = None, MatchCase = None, Orientation: 'XlSortOrientation' = None, DataOption1: 'XlSortDataOption' = None, DataOption2: 'XlSortDataOption' = None, DataOption3: 'XlSortDataOption' = None): ...
    def Speak(self, SpeakDirection = None, SpeakFormulas = None): ...
    def SpecialCells(self, Type: 'XlCellType', Value = None) -> 'Range': ...
    def SubscribeTo(self, Edition: str, Format: 'XlSubscribeToFormat' = None): ...
    def Subtotal(self, GroupBy: float, Function: 'XlConsolidationFunction', TotalList, Replace = None, PageBreaks = None, SummaryBelowData: 'XlSummaryRow' = None): ...
    def Table(self, RowInput = None, ColumnInput = None): ...
    def TextToColumns(self, Destination = None, DataType: 'XlTextParsingType' = None, TextQualifier: 'XlTextQualifier' = None, ConsecutiveDelimiter = None, Tab = None, Semicolon = None, Comma = None, Space = None, Other = None, OtherChar = None, FieldInfo = None, DecimalSeparator = None, ThousandsSeparator = None, TrailingMinusNumbers = None): ...
    def Ungroup(self, ): ...
    def UnMerge(self, ): ...
    @property
    def Value(self, RangeValueDataType = None): ...

class Ranges:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Range': ...
    @property
    def Item(self, Index) -> Range: ...

class RecentFile:
    Application: Application
    Creator: 'XlCreator'
    Index: float
    Name: str
    Parent = None
    Path: str
    def Delete(self, ): ...
    def Open(self, ) -> 'Workbook': ...

class RecentFiles:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Maximum: float
    Parent = None
    def __call__(self, Index) -> 'RecentFile': ...
    def Add(self, Name: str) -> RecentFile: ...
    @property
    def Item(self, Index: float) -> RecentFile: ...

class RectangularGradient:
    Application: Application
    ColorStops: ColorStops
    Creator: 'XlCreator'
    Parent = None
    RectangleBottom: float
    RectangleLeft: float
    RectangleRight: float
    RectangleTop: float

class Research:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def IsResearchService(self, ServiceID: str) -> bool: ...
    def Query(self, ServiceID: str, QueryString = None, QueryLanguage = None, UseSelection = None, LaunchQuery = None): ...
    def SetLanguagePair(self, LanguageFrom: float, LanguageTo: float): ...

class RoutingSlip:
    Application: Application
    Creator: 'XlCreator'
    Delivery: 'XlRoutingSlipDelivery'
    Message = None
    Parent = None
    ReturnWhenDone: bool
    Status: 'XlRoutingSlipStatus'
    Subject = None
    TrackStatus: bool
    @property
    def Recipients(self, Index = None): ...
    def Reset(self, ): ...

class RTD:
    ThrottleInterval: float
    def RefreshData(self, ): ...
    def RestartServers(self, ): ...

class Scenario:
    Application: Application
    ChangingCells: Range
    Comment: str
    Creator: 'XlCreator'
    Hidden: bool
    Index: float
    Locked: bool
    Name: str
    Parent = None
    def ChangeScenario(self, ChangingCells, Values = None): ...
    def Delete(self, ): ...
    def Show(self, ): ...
    @property
    def Values(self, Index = None): ...

class Scenarios:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Scenario': ...
    def Add(self, Name: str, ChangingCells, Values = None, Comment = None, Locked = None, Hidden = None) -> Scenario: ...
    def CreateSummary(self, ReportType: 'XlSummaryReportType' = None, ResultCells = None): ...
    def Item(self, Index) -> Scenario: ...
    def Merge(self, Source): ...

class Series:
    Application: Application
    ApplyPictToEnd: bool
    ApplyPictToFront: bool
    ApplyPictToSides: bool
    AxisGroup: 'XlAxisGroup'
    BarShape: 'XlBarShape'
    BubbleSizes = None
    ChartType: 'XlChartType'
    Creator: 'XlCreator'
    ErrorBars: ErrorBars
    Explosion: float
    Format: ChartFormat
    Formula: str
    FormulaLocal: str
    FormulaR1C1: str
    FormulaR1C1Local: str
    GeoMappingLevel: 'XlGeoMappingLevel'
    GeoProjectionType: 'XlGeoProjectionType'
    Has3DEffect: bool
    HasDataLabels: bool
    HasErrorBars: bool
    HasLeaderLines: bool
    InvertColor: float
    InvertColorIndex: float
    InvertIfNegative: bool
    IsFiltered: bool
    LeaderLines: LeaderLines
    MarkerBackgroundColor: float
    MarkerBackgroundColorIndex: 'XlColorIndex'
    MarkerForegroundColor: float
    MarkerForegroundColorIndex: 'XlColorIndex'
    MarkerSize: float
    MarkerStyle: 'XlMarkerStyle'
    Name: str
    Parent = None
    ParentDataLabelOption: 'XlParentDataLabelOptions'
    PictureType: 'XlChartPictureType'
    PictureUnit2: float
    PlotColorIndex: float
    PlotOrder: float
    QuartileCalculationInclusiveMedian: bool
    RegionLabelOption: 'XlRegionLabelOptions'
    SeriesColorGradientStyle: 'XlSeriesColorGradientStyle'
    SeriesColorMaxGradientStop: ChartSeriesGradientStopData
    SeriesColorMidGradientStop: ChartSeriesGradientStopData
    SeriesColorMinGradientStop: ChartSeriesGradientStopData
    Shadow: bool
    Smooth: bool
    Type: float
    Values = None
    ValueSortOrder: 'XlValueSortOrder'
    XValues = None
    def ApplyDataLabels(self, Type: 'XlDataLabelsType' = None, LegendKey = None, AutoText = None, HasLeaderLines = None, ShowSeriesName = None, ShowCategoryName = None, ShowValue = None, ShowPercentage = None, ShowBubbleSize = None, Separator = None): ...
    def ClearFormats(self, ): ...
    def Copy(self, ): ...
    def DataLabels(self, Index = None): ...
    def Delete(self, ): ...
    def ErrorBar(self, Direction: 'XlErrorBarDirection', Include: 'XlErrorBarInclude', Type: 'XlErrorBarType', Amount = None, MinusValues = None): ...
    def GetProperty(self, ID: str): ...
    def Paste(self, ): ...
    def Points(self, Index = None): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...
    def Trendlines(self, Index = None): ...

class SeriesCollection:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Series': ...
    def Add(self, Source, Rowcol: 'XlRowCol' = None, SeriesLabels = None, CategoryLabels = None, Replace = None) -> Series: ...
    def Extend(self, Source, Rowcol = None, CategoryLabels = None): ...
    def Item(self, Index) -> Series: ...
    def NewSeries(self, ) -> Series: ...
    def Paste(self, Rowcol: 'XlRowCol' = None, SeriesLabels = None, CategoryLabels = None, Replace = None, NewSeries = None): ...

class SeriesGradientStopColorFormat:
    Application: Application
    Creator: 'XlCreator'
    ObjectThemeColor: 'MsoThemeColorIndex'
    Parent = None
    RGB: float
    TintAndShade: 'Single'
    Transparency: 'Single'
    Type: 'MsoColorType'

class SeriesLines:
    Application: Application
    Border: Border
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class ServerViewableItems:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'None': ...
    def Add(self, Obj): ...
    def Delete(self, Index): ...
    def DeleteAll(self, ): ...
    def Item(self, Index): ...

class ShadowFormat:
    Application = None
    Blur: 'Single'
    Creator: float
    ForeColor: ColorFormat
    Obscured: 'MsoTriState'
    OffsetX: 'Single'
    OffsetY: 'Single'
    Parent = None
    RotateWithShape: 'MsoTriState'
    Size: 'Single'
    Style: 'MsoShadowStyle'
    Transparency: 'Single'
    Type: 'MsoShadowType'
    Visible: 'MsoTriState'
    def IncrementOffsetX(self, Increment: 'Single'): ...
    def IncrementOffsetY(self, Increment: 'Single'): ...

class Shape:
    Adjustments: Adjustments
    AlternativeText: str
    Application: Application
    AutoShapeType: 'MsoAutoShapeType'
    BackgroundStyle: 'MsoBackgroundStyleIndex'
    BlackWhiteMode: 'MsoBlackWhiteMode'
    BottomRightCell: Range
    Callout: CalloutFormat
    Chart: Chart
    Child: 'MsoTriState'
    ConnectionSiteCount: float
    Connector: 'MsoTriState'
    ConnectorFormat: ConnectorFormat
    ControlFormat: ControlFormat
    Creator: 'XlCreator'
    Decorative: 'MsoTriState'
    Fill: FillFormat
    FormControlType: 'XlFormControl'
    Glow: 'GlowFormat'
    GraphicStyle: 'MsoGraphicStyleIndex'
    GroupItems: GroupShapes
    HasChart: 'MsoTriState'
    HasSmartArt: 'MsoTriState'
    Height: 'Single'
    HorizontalFlip: 'MsoTriState'
    Hyperlink: Hyperlink
    ID: float
    Left: 'Single'
    Line: LineFormat
    LinkFormat: LinkFormat
    LockAspectRatio: 'MsoTriState'
    Locked: bool
    Model3D: Model3DFormat
    Name: str
    Nodes: 'ShapeNodes'
    OLEFormat: OLEFormat
    OnAction: str
    Parent = None
    ParentGroup: 'Shape'
    PictureFormat: PictureFormat
    Placement: 'XlPlacement'
    Reflection: 'ReflectionFormat'
    Rotation: 'Single'
    Shadow: ShadowFormat
    ShapeStyle: 'MsoShapeStyleIndex'
    SmartArt: 'SmartArt'
    SoftEdge: 'SoftEdgeFormat'
    TextEffect: 'TextEffectFormat'
    TextFrame: 'TextFrame'
    TextFrame2: 'TextFrame2'
    ThreeD: 'ThreeDFormat'
    Title: str
    Top: 'Single'
    TopLeftCell: Range
    Type: 'MsoShapeType'
    VerticalFlip: 'MsoTriState'
    Vertices = None
    Visible: 'MsoTriState'
    Width: 'Single'
    ZOrderPosition: float
    def Apply(self, ): ...
    def Copy(self, ): ...
    def CopyPicture(self, Appearance = None, Format = None): ...
    def Cut(self, ): ...
    def Delete(self, ): ...
    def Duplicate(self, ) -> 'Shape': ...
    def Flip(self, FlipCmd: 'MsoFlipCmd'): ...
    def IncrementLeft(self, Increment: 'Single'): ...
    def IncrementRotation(self, Increment: 'Single'): ...
    def IncrementTop(self, Increment: 'Single'): ...
    def PickUp(self, ): ...
    def RerouteConnections(self, ): ...
    def ScaleHeight(self, Factor: 'Single', RelativeToOriginalSize: 'MsoTriState', Scale = None): ...
    def ScaleWidth(self, Factor: 'Single', RelativeToOriginalSize: 'MsoTriState', Scale = None): ...
    def Select(self, Replace = None): ...
    def SetShapesDefaultProperties(self, ): ...
    def Ungroup(self, ) -> 'ShapeRange': ...
    def ZOrder(self, ZOrderCmd: 'MsoZOrderCmd'): ...

class ShapeNode:
    Application = None
    Creator: float
    EditingType: 'MsoEditingType'
    Parent = None
    Points = None
    SegmentType: 'MsoSegmentType'

class ShapeNodes:
    Application = None
    Count: float
    Creator: float
    Parent = None
    def __call__(self, Index) -> 'ShapeNode': ...
    def Delete(self, Index: float): ...
    def Insert(self, Index: float, SegmentType: 'MsoSegmentType', EditingType: 'MsoEditingType', X1: 'Single', Y1: 'Single', X2: 'Single' = None, Y2: 'Single' = None, X3: 'Single' = None, Y3: 'Single' = None): ...
    def Item(self, Index) -> ShapeNode: ...
    def SetEditingType(self, Index: float, EditingType: 'MsoEditingType'): ...
    def SetPosition(self, Index: float, X1: 'Single', Y1: 'Single'): ...
    def SetSegmentType(self, Index: float, SegmentType: 'MsoSegmentType'): ...

class ShapeRange:
    Adjustments: Adjustments
    AlternativeText: str
    Application: Application
    AutoShapeType: 'MsoAutoShapeType'
    BackgroundStyle: 'MsoBackgroundStyleIndex'
    BlackWhiteMode: 'MsoBlackWhiteMode'
    Callout: CalloutFormat
    Chart: Chart
    Child: 'MsoTriState'
    ConnectionSiteCount: float
    Connector: 'MsoTriState'
    ConnectorFormat: ConnectorFormat
    Count: float
    Creator: 'XlCreator'
    Decorative: 'MsoTriState'
    Fill: FillFormat
    Glow: 'GlowFormat'
    GraphicStyle: 'MsoGraphicStyleIndex'
    GroupItems: GroupShapes
    HasChart: 'MsoTriState'
    Height: 'Single'
    HorizontalFlip: 'MsoTriState'
    ID: float
    Left: 'Single'
    Line: LineFormat
    LockAspectRatio: 'MsoTriState'
    Model3D: Model3DFormat
    Name: str
    Nodes: ShapeNodes
    Parent = None
    ParentGroup: Shape
    PictureFormat: PictureFormat
    Reflection: 'ReflectionFormat'
    Rotation: 'Single'
    Shadow: ShadowFormat
    ShapeStyle: 'MsoShapeStyleIndex'
    SoftEdge: 'SoftEdgeFormat'
    TextEffect: 'TextEffectFormat'
    TextFrame: 'TextFrame'
    TextFrame2: 'TextFrame2'
    ThreeD: 'ThreeDFormat'
    Title: str
    Top: 'Single'
    Type: 'MsoShapeType'
    VerticalFlip: 'MsoTriState'
    Vertices = None
    Visible: 'MsoTriState'
    Width: 'Single'
    ZOrderPosition: float
    def __call__(self, Index) -> 'Shape': ...
    def Align(self, AlignCmd: 'MsoAlignCmd', RelativeTo: 'MsoTriState'): ...
    def Apply(self, ): ...
    def Delete(self, ): ...
    def Distribute(self, DistributeCmd: 'MsoDistributeCmd', RelativeTo: 'MsoTriState'): ...
    def Duplicate(self, ) -> 'ShapeRange': ...
    def Flip(self, FlipCmd: 'MsoFlipCmd'): ...
    def Group(self, ) -> Shape: ...
    def IncrementLeft(self, Increment: 'Single'): ...
    def IncrementRotation(self, Increment: 'Single'): ...
    def IncrementTop(self, Increment: 'Single'): ...
    def Item(self, Index) -> Shape: ...
    def PickUp(self, ): ...
    def Regroup(self, ) -> Shape: ...
    def RerouteConnections(self, ): ...
    def ScaleHeight(self, Factor: 'Single', RelativeToOriginalSize: 'MsoTriState', Scale = None): ...
    def ScaleWidth(self, Factor: 'Single', RelativeToOriginalSize: 'MsoTriState', Scale = None): ...
    def Select(self, Replace = None): ...
    def SetShapesDefaultProperties(self, ): ...
    def Ungroup(self, ) -> 'ShapeRange': ...
    def ZOrder(self, ZOrderCmd: 'MsoZOrderCmd'): ...

class Shapes:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Shape': ...
    def Add3DModel(self, Filename: str, LinkToFile = None, SaveWithDocument = None, Left = None, Top = None, Width = None, Height = None) -> Shape: ...
    def AddCallout(self, Type: 'MsoCalloutType', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single') -> Shape: ...
    def AddChart2(self, Style = None, XlChartType = None, Left = None, Top = None, Width = None, Height = None, NewLayout = None) -> Shape: ...
    def AddConnector(self, Type: 'MsoConnectorType', BeginX: 'Single', BeginY: 'Single', EndX: 'Single', EndY: 'Single') -> Shape: ...
    def AddCurve(self, SafeArrayOfPoints) -> Shape: ...
    def AddFormControl(self, Type: 'XlFormControl', Left: float, Top: float, Width: float, Height: float) -> Shape: ...
    def AddLabel(self, Orientation: 'MsoTextOrientation', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single') -> Shape: ...
    def AddLine(self, BeginX: 'Single', BeginY: 'Single', EndX: 'Single', EndY: 'Single') -> Shape: ...
    def AddOLEObject(self, ClassType = None, Filename = None, Link = None, DisplayAsIcon = None, IconFileName = None, IconIndex = None, IconLabel = None, Left = None, Top = None, Width = None, Height = None) -> Shape: ...
    def AddPicture(self, Filename: str, LinkToFile: 'MsoTriState', SaveWithDocument: 'MsoTriState', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single') -> Shape: ...
    def AddPicture2(self, Filename: str, LinkToFile: 'MsoTriState', SaveWithDocument: 'MsoTriState', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single', Compress: 'MsoPictureCompress') -> Shape: ...
    def AddPolyline(self, SafeArrayOfPoints) -> Shape: ...
    def AddShape(self, Type: 'MsoAutoShapeType', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single') -> Shape: ...
    def AddSmartArt(self, Layout: 'SmartArtLayout', Left = None, Top = None, Width = None, Height = None) -> Shape: ...
    def AddTextbox(self, Orientation: 'MsoTextOrientation', Left: 'Single', Top: 'Single', Width: 'Single', Height: 'Single') -> Shape: ...
    def AddTextEffect(self, PresetTextEffect: 'MsoPresetTextEffect', Text: str, FontName: str, FontSize: 'Single', FontBold: 'MsoTriState', FontItalic: 'MsoTriState', Left: 'Single', Top: 'Single') -> Shape: ...
    def BuildFreeform(self, EditingType: 'MsoEditingType', X1: 'Single', Y1: 'Single') -> FreeformBuilder: ...
    def Item(self, Index) -> Shape: ...
    @property
    def Range(self, Index) -> ShapeRange: ...
    def SelectAll(self, ): ...

class Sheets:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    HPageBreaks: HPageBreaks
    Parent = None
    Visible = None
    VPageBreaks: 'VPageBreaks'
    def __call__(self, Index) -> 'Worksheet': ...
    def Add(self, Before = None, After = None, Count = None, Type = None): ...
    def Add2(self, Before = None, After = None, Count = None, NewLayout = None): ...
    def Copy(self, Before = None, After = None): ...
    def Delete(self, ): ...
    def FillAcrossSheets(self, Range: Range, Type: 'XlFillWith' = None): ...
    @property
    def Item(self, Index) -> 'Worksheet': ...
    def Move(self, Before = None, After = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None, IgnorePrintAreas = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Select(self, Replace = None): ...

class SheetViews:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'None': ...
    @property
    def Item(self, Index): ...

class Slicer:
    ActiveItem: 'SlicerItem'
    Application: Application
    Caption: str
    ColumnWidth: float
    Creator: 'XlCreator'
    DisableMoveResizeUI: bool
    DisplayHeader: bool
    Height: float
    Left: float
    Locked: bool
    Name: str
    NumberOfColumns: float
    Parent = None
    RowHeight: float
    Shape: Shape
    SlicerCache: 'SlicerCache'
    SlicerCacheLevel: 'SlicerCacheLevel'
    SlicerCacheType: 'XlSlicerCacheType'
    Style = None
    TimelineViewState: 'TimelineViewState'
    Top: float
    Width: float
    def Copy(self, ): ...
    def Cut(self, ): ...
    def Delete(self, ): ...

class SlicerCache:
    Application: Application
    Creator: 'XlCreator'
    CrossFilterType: 'XlSlicerCrossFilterType'
    FilterCleared: bool
    Index: float
    List: bool
    ListObject: ListObject
    Name: str
    OLAP: bool
    Parent = None
    PivotTables: 'SlicerPivotTables'
    RequireManualUpdate: bool
    ShowAllItems: bool
    SlicerCacheLevels: 'SlicerCacheLevels'
    SlicerCacheType: 'XlSlicerCacheType'
    SlicerItems: 'SlicerItems'
    Slicers: 'Slicers'
    SortItems: 'XlSlicerSort'
    SortUsingCustomLists: bool
    SourceName: str
    SourceType: 'XlPivotTableSourceType'
    TimelineState: 'TimelineState'
    VisibleSlicerItems: 'SlicerItems'
    VisibleSlicerItemsList = None
    WorkbookConnection: 'WorkbookConnection'
    def ClearAllFilters(self, ): ...
    def ClearDateFilter(self, ): ...
    def ClearManualFilter(self, ): ...
    def Delete(self, ): ...

class SlicerCacheLevel:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    CrossFilterType: 'XlSlicerCrossFilterType'
    Name: str
    Ordinal: float
    Parent = None
    SlicerItems: 'SlicerItems'
    SortItems: 'XlSlicerSort'
    VisibleSlicerItemsList = None

class SlicerCacheLevels:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'SlicerCacheLevel': ...
    @property
    def Item(self, Level = None) -> SlicerCacheLevel: ...

class SlicerCaches:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'SlicerCache': ...
    def Add2(self, Source, SourceField, Name = None, SlicerCacheType = None) -> SlicerCache: ...
    @property
    def Item(self, Index) -> SlicerCache: ...

class SlicerItem:
    Application: Application
    Caption: str
    Creator: 'XlCreator'
    HasData: bool
    Name: str
    Parent: SlicerCache
    Selected: bool
    SourceName = None
    SourceNameStandard: str
    Value: str

class SlicerItems:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'SlicerItem': ...
    @property
    def Item(self, Index) -> SlicerItem: ...

class SlicerPivotTables:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'PivotTable': ...
    def AddPivotTable(self, PivotTable: PivotTable): ...
    @property
    def Item(self, Index) -> PivotTable: ...
    def RemovePivotTable(self, PivotTable): ...

class Slicers:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Slicer': ...
    def Add(self, SlicerDestination, Level = None, Name = None, Caption = None, Top = None, Left = None, Width = None, Height = None) -> Slicer: ...
    @property
    def Item(self, Index) -> Slicer: ...

class Sort:
    Application: Application
    Creator: 'XlCreator'
    Header: 'XlYesNoGuess'
    MatchCase: bool
    Orientation: 'XlSortOrientation'
    Parent = None
    Rng: Range
    SortFields: 'SortFields'
    SortMethod: 'XlSortMethod'
    def Apply(self, ): ...
    def SetRange(self, Rng: Range): ...

class SortField:
    Application: Application
    Creator: 'XlCreator'
    CustomOrder = None
    DataOption: 'XlSortDataOption'
    Key: Range
    Order: 'XlSortOrder'
    Parent = None
    Priority: float
    SortOn: 'XlSortOn'
    SortOnValue = None
    SubField = None
    def Delete(self, ): ...
    def ModifyKey(self, Key: Range): ...
    def SetIcon(self, Icon: Icon): ...

class SortFields:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'SortField': ...
    def Add(self, Key: Range, SortOn = None, Order = None, CustomOrder = None, DataOption = None) -> SortField: ...
    def Add2(self, Key: Range, SortOn = None, Order = None, CustomOrder = None, DataOption = None, SubField = None) -> SortField: ...
    def Clear(self, ): ...
    @property
    def Item(self, Index) -> SortField: ...

class SoundNote:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def Delete(self, ): ...
    def Import(self, Filename: str): ...
    def Play(self, ): ...
    def Record(self, ): ...

class SparkAxes:
    Application: Application
    Creator: 'XlCreator'
    Horizontal: 'SparkHorizontalAxis'
    Parent = None
    Vertical: 'SparkVerticalAxis'

class SparkColor:
    Application: Application
    Color: FormatColor
    Creator: 'XlCreator'
    Parent = None
    Visible: bool

class SparkHorizontalAxis:
    Application: Application
    Axis: SparkColor
    Creator: 'XlCreator'
    IsDateAxis: bool
    Parent = None
    RightToLeftPlotOrder: bool

class Sparkline:
    Application: Application
    Creator: 'XlCreator'
    Location: Range
    Parent = None
    SourceData: str
    def ModifyLocation(self, Range: Range): ...
    def ModifySourceData(self, Formula: str): ...

class SparklineGroup:
    Application: Application
    Axes: SparkAxes
    Count: float
    Creator: 'XlCreator'
    DateRange: str
    DisplayBlanksAs: 'XlDisplayBlanksAs'
    DisplayHidden: bool
    LineWeight = None
    Location: Range
    Parent = None
    PlotBy: 'XlSparklineRowCol'
    Points: 'SparkPoints'
    SeriesColor: FormatColor
    SourceData: str
    Type: 'XlSparkType'
    def __call__(self, Index) -> 'Sparkline': ...
    def Delete(self, ): ...
    @property
    def Item(self, Index) -> Sparkline: ...
    def Modify(self, Location: Range, SourceData: str): ...
    def ModifyDateRange(self, DateRange: str): ...
    def ModifyLocation(self, Location: Range): ...
    def ModifySourceData(self, SourceData: str): ...

class SparklineGroups:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'SparklineGroup': ...
    def Add(self, Type: 'XlSparkType', SourceData: str) -> SparklineGroup: ...
    def Clear(self, ): ...
    def ClearGroups(self, ): ...
    def Group(self, Location: Range): ...
    @property
    def Item(self, Index) -> SparklineGroup: ...
    def Ungroup(self, ): ...

class SparkPoints:
    Application: Application
    Creator: 'XlCreator'
    Firstpoint: SparkColor
    Highpoint: SparkColor
    Lastpoint: SparkColor
    Lowpoint: SparkColor
    Markers: SparkColor
    Negative: SparkColor
    Parent = None

class SparkVerticalAxis:
    Application: Application
    Creator: 'XlCreator'
    CustomMaxScaleValue = None
    CustomMinScaleValue = None
    MaxScaleType: 'XlSparkScale'
    MinScaleType: 'XlSparkScale'
    Parent = None

class Speech:
    Direction: 'XlSpeakDirection'
    SpeakCellOnEnter: bool
    def Speak(self, Text: str, SpeakAsync = None, SpeakXML = None, Purge = None): ...

class SpellingOptions:
    ArabicModes: 'XlArabicModes'
    ArabicStrictAlefHamza: bool
    ArabicStrictFinalYaa: bool
    ArabicStrictTaaMarboota: bool
    BrazilReform: 'XlPortugueseReform'
    DictLang: float
    GermanPostReform: bool
    HebrewModes: 'XlHebrewModes'
    IgnoreCaps: bool
    IgnoreFileNames: bool
    IgnoreMixedDigits: bool
    KoreanCombineAux: bool
    KoreanProcessCompound: bool
    KoreanUseAutoChangeList: bool
    PortugalReform: 'XlPortugueseReform'
    RussianStrictE: bool
    SpanishModes: 'XlSpanishModes'
    SuggestMainOnly: bool
    UserDict: str

class Style:
    AddIndent: bool
    Application: Application
    Borders: Borders
    BuiltIn: bool
    Creator: 'XlCreator'
    Font: Font
    FormulaHidden: bool
    HorizontalAlignment: 'XlHAlign'
    IncludeAlignment: bool
    IncludeBorder: bool
    IncludeFont: bool
    IncludeNumber: bool
    IncludePatterns: bool
    IncludeProtection: bool
    IndentLevel: float
    Interior: Interior
    Locked: bool
    MergeCells = None
    Name: str
    NameLocal: str
    NumberFormat: str
    NumberFormatLocal: str
    Orientation: 'XlOrientation'
    Parent = None
    ReadingOrder: float
    ShrinkToFit: bool
    Value: str
    VerticalAlignment: 'XlVAlign'
    WrapText: bool
    def Delete(self, ): ...

class Styles:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Style': ...
    def Add(self, Name: str, BasedOn = None) -> Style: ...
    @property
    def Item(self, Index) -> Style: ...
    def Merge(self, Workbook): ...

class Tab:
    Application: Application
    Color = None
    ColorIndex: 'XlColorIndex'
    Creator: 'XlCreator'
    Parent = None
    ThemeColor: 'XlThemeColor'
    TintAndShade = None

class TableObject:
    AdjustColumnWidth: bool
    Application: Application
    Creator: 'XlCreator'
    Destination: Range
    EnableEditing: bool
    EnableRefresh: bool
    FetchedRowOverflow: bool
    ListObject: ListObject
    Parent = None
    PreserveColumnInfo: bool
    PreserveFormatting: bool
    RefreshStyle: 'XlCellInsertionMode'
    ResultRange: Range
    RowNumbers: bool
    WorkbookConnection: 'WorkbookConnection'
    def Delete(self, ): ...
    def Refresh(self, ) -> bool: ...

class TableStyle:
    Application: Application
    BuiltIn: bool
    Creator: 'XlCreator'
    Name: str
    NameLocal: str
    Parent = None
    ShowAsAvailablePivotTableStyle: bool
    ShowAsAvailableSlicerStyle: bool
    ShowAsAvailableTableStyle: bool
    ShowAsAvailableTimelineStyle: bool
    TableStyleElements: 'TableStyleElements'
    def Delete(self, ): ...
    def Duplicate(self, NewTableStyleName = None) -> 'TableStyle': ...

class TableStyleElement:
    Application: Application
    Borders: Borders
    Creator: 'XlCreator'
    Font: Font
    HasFormat: bool
    Interior: Interior
    Parent = None
    StripeSize: float
    def Clear(self, ): ...

class TableStyleElements:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'TableStyleElement': ...
    def Item(self, Index: 'XlTableStyleElementType') -> TableStyleElement: ...

class TableStyles:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'TableStyle': ...
    def Add(self, TableStyleName: str) -> TableStyle: ...
    def Item(self, Index) -> TableStyle: ...

class TextConnection:
    Application: Application
    Connection = None
    Creator: 'XlCreator'
    Parent = None
    TextFileColumnDataTypes = None
    TextFileCommaDelimiter: bool
    TextFileConsecutiveDelimiter: bool
    TextFileDecimalSeparator: str
    TextFileFixedColumnWidths = None
    TextFileHeaderRow: bool
    TextFileOtherDelimiter: str
    TextFileParseType: 'XlTextParsingType'
    TextFilePlatform: 'XlPlatform'
    TextFilePromptOnRefresh: bool
    TextFileSemicolonDelimiter: bool
    TextFileSpaceDelimiter: bool
    TextFileStartRow: float
    TextFileTabDelimiter: bool
    TextFileTextQualifier: 'XlTextQualifier'
    TextFileThousandsSeparator: str
    TextFileTrailingMinusNumbers: bool
    TextFileVisualLayout: 'XlTextVisualLayoutType'

class TextEffectFormat:
    Alignment: 'MsoTextEffectAlignment'
    Application = None
    Creator: float
    FontBold: 'MsoTriState'
    FontItalic: 'MsoTriState'
    FontName: str
    FontSize: 'Single'
    KernedPairs: 'MsoTriState'
    NormalizedHeight: 'MsoTriState'
    Parent = None
    PresetShape: 'MsoPresetTextEffectShape'
    PresetTextEffect: 'MsoPresetTextEffect'
    RotatedChars: 'MsoTriState'
    Text: str
    Tracking: 'Single'
    def ToggleVerticalText(self, ): ...

class TextFrame:
    Application: Application
    AutoMargins: bool
    AutoSize: bool
    Creator: 'XlCreator'
    HorizontalAlignment: 'XlHAlign'
    HorizontalOverflow: 'XlOartHorizontalOverflow'
    MarginBottom: 'Single'
    MarginLeft: 'Single'
    MarginRight: 'Single'
    MarginTop: 'Single'
    Orientation: 'MsoTextOrientation'
    Parent = None
    ReadingOrder: float
    VerticalAlignment: 'XlVAlign'
    VerticalOverflow: 'XlOartVerticalOverflow'
    def Characters(self, Start = None, Length = None) -> Characters: ...

class TextFrame2:
    Application = None
    AutoSize: 'MsoAutoSize'
    Column: 'TextColumn2'
    Creator: float
    HasText: 'MsoTriState'
    HorizontalAnchor: 'MsoHorizontalAnchor'
    MarginBottom: 'Single'
    MarginLeft: 'Single'
    MarginRight: 'Single'
    MarginTop: 'Single'
    NoTextRotation: 'MsoTriState'
    Orientation: 'MsoTextOrientation'
    Parent = None
    PathFormat: 'MsoPathFormat'
    Ruler: 'Ruler2'
    TextRange: 'TextRange2'
    ThreeD: 'ThreeDFormat'
    VerticalAnchor: 'MsoVerticalAnchor'
    WarpFormat: 'MsoWarpFormat'
    WordArtformat: 'MsoPresetTextEffect'
    WordWrap: 'MsoTriState'
    def DeleteText(self, ): ...

class ThreeDFormat:
    Application = None
    BevelBottomDepth: 'Single'
    BevelBottomInset: 'Single'
    BevelBottomType: 'MsoBevelType'
    BevelTopDepth: 'Single'
    BevelTopInset: 'Single'
    BevelTopType: 'MsoBevelType'
    ContourColor: ColorFormat
    ContourWidth: 'Single'
    Creator: float
    Depth: 'Single'
    ExtrusionColor: ColorFormat
    ExtrusionColorType: 'MsoExtrusionColorType'
    FieldOfView: 'Single'
    LightAngle: 'Single'
    Parent = None
    Perspective: 'MsoTriState'
    PresetCamera: 'MsoPresetCamera'
    PresetExtrusionDirection: 'MsoPresetExtrusionDirection'
    PresetLighting: 'MsoLightRigType'
    PresetLightingDirection: 'MsoPresetLightingDirection'
    PresetLightingSoftness: 'MsoPresetLightingSoftness'
    PresetMaterial: 'MsoPresetMaterial'
    PresetThreeDFormat: 'MsoPresetThreeDFormat'
    ProjectText: 'MsoTriState'
    RotationX: 'Single'
    RotationY: 'Single'
    RotationZ: 'Single'
    Visible: 'MsoTriState'
    Z: 'Single'
    def IncrementRotationHorizontal(self, Increment: 'Single'): ...
    def IncrementRotationVertical(self, Increment: 'Single'): ...
    def IncrementRotationX(self, Increment: 'Single'): ...
    def IncrementRotationY(self, Increment: 'Single'): ...
    def IncrementRotationZ(self, Increment: 'Single'): ...
    def ResetRotation(self, ): ...
    def SetExtrusionDirection(self, PresetExtrusionDirection: 'MsoPresetExtrusionDirection'): ...
    def SetPresetCamera(self, PresetCamera: 'MsoPresetCamera'): ...
    def SetThreeDFormat(self, PresetThreeDFormat: 'MsoPresetThreeDFormat'): ...

class TickLabels:
    Alignment: float
    Application: Application
    Creator: 'XlCreator'
    Depth: float
    Font: Font
    Format: ChartFormat
    MultiLevel: bool
    Name: str
    NumberFormat: str
    NumberFormatLinked: bool
    NumberFormatLocal = None
    Offset: float
    Orientation: 'XlTickLabelOrientation'
    Parent = None
    ReadingOrder: float
    def Delete(self, ): ...
    def Select(self, ): ...

class TimelineState:
    Application: Application
    Creator: 'XlCreator'
    EndDate = None
    FilterType: 'XlPivotFilterType'
    FilterValue1 = None
    FilterValue2 = None
    Parent = None
    SingleRangeFilterState: bool
    StartDate = None
    def SetFilterDateRange(self, StartDate, EndDate) -> 'XlFilterStatus': ...

class TimelineViewState:
    Application: Application
    Creator: 'XlCreator'
    Level: 'XlTimelineLevel'
    Parent = None
    ShowHeader: bool
    ShowHorizontalScrollbar: bool
    ShowSelectionLabel: bool
    ShowTimeLevel: bool

class Top10:
    Application: Application
    AppliesTo: Range
    Borders: Borders
    CalcFor: 'XlCalcFor'
    Creator: 'XlCreator'
    Font: Font
    Interior: Interior
    NumberFormat = None
    Parent = None
    Percent: bool
    Priority: float
    PTCondition: bool
    Rank: float
    ScopeType: 'XlPivotConditionScope'
    StopIfTrue: bool
    TopBottom: 'XlTopBottom'
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: Range): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class TreeviewControl:
    Application: Application
    Creator: 'XlCreator'
    Drilled = None
    Hidden = None
    Parent = None

class Trendline:
    Application: Application
    Backward2: float
    Border: Border
    Creator: 'XlCreator'
    DataLabel: DataLabel
    DisplayEquation: bool
    DisplayRSquared: bool
    Format: ChartFormat
    Forward2: float
    Index: float
    Intercept: float
    InterceptIsAuto: bool
    Name: str
    NameIsAuto: bool
    Order: float
    Parent = None
    Period: float
    Type: 'XlTrendlineType'
    def ClearFormats(self, ): ...
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class Trendlines:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Trendline': ...
    def Add(self, Type: 'XlTrendlineType' = None, Order = None, Period = None, Forward = None, Backward = None, Intercept = None, DisplayEquation = None, DisplayRSquared = None, Name = None) -> Trendline: ...
    def Item(self, Index = None) -> Trendline: ...

class UniqueValues:
    Application: Application
    AppliesTo: Range
    Borders: Borders
    Creator: 'XlCreator'
    DupeUnique: 'XlDupeUnique'
    Font: Font
    Interior: Interior
    NumberFormat = None
    Parent = None
    Priority: float
    PTCondition: bool
    ScopeType: 'XlPivotConditionScope'
    StopIfTrue: bool
    Type: float
    def Delete(self, ): ...
    def ModifyAppliesToRange(self, Range: Range): ...
    def SetFirstPriority(self, ): ...
    def SetLastPriority(self, ): ...

class UpBars:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    def Delete(self, ): ...
    def GetProperty(self, ID: str): ...
    def Select(self, ): ...
    def SetProperty(self, ID: str, Value): ...

class UsedObjects:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'None': ...
    @property
    def Item(self, Index): ...

class UserAccess:
    AllowEdit: bool
    Name: str
    def Delete(self, ): ...

class UserAccessList:
    Count: float
    def __call__(self, Index) -> 'UserAccess': ...
    def Add(self, Name: str, AllowEdit: bool) -> UserAccess: ...
    def DeleteAll(self, ): ...
    @property
    def Item(self, Index) -> UserAccess: ...

class Validation:
    AlertStyle: float
    Application: Application
    Creator: 'XlCreator'
    ErrorMessage: str
    ErrorTitle: str
    Formula1: str
    Formula2: str
    IgnoreBlank: bool
    IMEMode: float
    InCellDropdown: bool
    InputMessage: str
    InputTitle: str
    Operator: float
    Parent = None
    ShowError: bool
    ShowInput: bool
    Type: float
    Value: bool
    def Add(self, Type: 'XlDVType', AlertStyle = None, Operator = None, Formula1 = None, Formula2 = None): ...
    def Delete(self, ): ...
    def Modify(self, Type = None, AlertStyle = None, Operator = None, Formula1 = None, Formula2 = None): ...

class ValueChange:
    AllocationMethod: 'XlAllocationMethod'
    AllocationValue: 'XlAllocationValue'
    AllocationWeightExpression: str
    Application: Application
    Creator: 'XlCreator'
    Order: float
    Parent = None
    PivotCell: PivotCell
    Tuple: str
    Value: float
    VisibleInPivotTable: bool
    def Delete(self, ): ...

class VPageBreak:
    Application: Application
    Creator: 'XlCreator'
    Extent: 'XlPageBreakExtent'
    Location: Range
    Parent: 'Worksheet'
    Type: 'XlPageBreak'
    def Delete(self, ): ...
    def DragOff(self, Direction: 'XlDirection', RegionIndex: float): ...

class VPageBreaks:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'VPageBreak': ...
    def Add(self, Before) -> VPageBreak: ...
    @property
    def Item(self, Index: float) -> VPageBreak: ...

class Walls:
    Application: Application
    Creator: 'XlCreator'
    Format: ChartFormat
    Name: str
    Parent = None
    PictureType = None
    PictureUnit = None
    Thickness: float
    def ClearFormats(self, ): ...
    def Paste(self, ): ...
    def Select(self, ): ...

class Watch:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    Source = None
    def Delete(self, ): ...

class Watches:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Watch': ...
    def Add(self, Source) -> Watch: ...
    def Delete(self, ): ...
    @property
    def Item(self, Index) -> Watch: ...

class WebOptions:
    AllowPNG: bool
    Application: Application
    Creator: 'XlCreator'
    DownloadComponents: bool
    Encoding: 'MsoEncoding'
    FolderSuffix: str
    LocationOfComponents: str
    OrganizeInFolder: bool
    Parent = None
    PixelsPerInch: float
    RelyOnCSS: bool
    RelyOnVML: bool
    ScreenSize: 'MsoScreenSize'
    TargetBrowser: 'MsoTargetBrowser'
    UseLongFileNames: bool
    def UseDefaultFolderSuffix(self, ): ...

class Window:
    ActiveCell: Range
    ActiveChart: Chart
    ActivePane: Pane
    ActiveSheet: 'Worksheet'
    ActiveSheetView = None
    Application: Application
    AutoFilterDateGrouping: bool
    Caption = None
    Creator: 'XlCreator'
    DisplayFormulas: bool
    DisplayGridlines: bool
    DisplayHeadings: bool
    DisplayHorizontalScrollBar: bool
    DisplayOutline: bool
    DisplayRightToLeft: bool
    DisplayRuler: bool
    DisplayVerticalScrollBar: bool
    DisplayWhitespace: bool
    DisplayWorkbookTabs: bool
    DisplayZeros: bool
    EnableResize: bool
    FreezePanes: bool
    GridlineColor: float
    GridlineColorIndex: 'XlColorIndex'
    Height: float
    Hwnd: float
    Index: float
    Left: float
    OnWindow: str
    Panes: Panes
    Parent = None
    RangeSelection: Range
    ScrollColumn: float
    ScrollRow: float
    SelectedSheets: Sheets
    Selection = None
    SheetViews: SheetViews
    Split: bool
    SplitColumn: float
    SplitHorizontal: float
    SplitRow: float
    SplitVertical: float
    TabRatio: float
    Top: float
    Type: 'XlWindowType'
    UsableHeight: float
    UsableWidth: float
    View: 'XlWindowView'
    Visible: bool
    VisibleRange: Range
    Width: float
    WindowNumber: float
    WindowState: 'XlWindowState'
    Zoom = None
    def Activate(self, ): ...
    def ActivateNext(self, ): ...
    def ActivatePrevious(self, ): ...
    def Close(self, SaveChanges = None, Filename = None, RouteWorkbook = None) -> bool: ...
    def LargeScroll(self, Down = None, Up = None, ToRight = None, ToLeft = None): ...
    def NewWindow(self, ) -> 'Window': ...
    def PointsToScreenPixelsX(self, Points: float) -> float: ...
    def PointsToScreenPixelsY(self, Points: float) -> float: ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def RangeFromPoint(self, x: float, y: float): ...
    def ScrollIntoView(self, Left: float, Top: float, Width: float, Height: float, Start = None): ...
    def ScrollWorkbookTabs(self, Sheets = None, Position = None): ...
    def SmallScroll(self, Down = None, Up = None, ToRight = None, ToLeft = None): ...

class Windows:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    SyncScrollingSideBySide: bool
    def __call__(self, Index) -> 'Window': ...
    def Arrange(self, ArrangeStyle: 'XlArrangeStyle' = None, ActiveWorkbook = None, SyncHorizontal = None, SyncVertical = None): ...
    def BreakSideBySide(self, ) -> bool: ...
    def CompareSideBySideWith(self, WindowName) -> bool: ...
    @property
    def Item(self, Index) -> Window: ...
    def ResetPositionsSideBySide(self, ): ...

class Workbook:
    AccuracyVersion: float
    ActiveChart: Chart
    ActiveSheet: 'Worksheet'
    ActiveSlicer: Slicer
    Application: Application
    AutoSaveOn: bool
    AutoUpdateFrequency: float
    AutoUpdateSaveChanges: bool
    BuiltinDocumentProperties = None
    CalculationVersion: float
    CaseSensitive: bool
    ChangeHistoryDuration: float
    ChartDataPointTrack: bool
    Charts: Sheets
    CheckCompatibility: bool
    CodeName: str
    CommandBars: 'CommandBars'
    ConflictResolution: 'XlSaveConflictResolution'
    Connections: Connections
    ConnectionsDisabled: bool
    Container = None
    ContentTypeProperties: 'MetaProperties'
    CreateBackup: bool
    Creator: 'XlCreator'
    CustomDocumentProperties = None
    CustomViews: CustomViews
    CustomXMLParts: 'CustomXMLParts'
    Date1904: bool
    DefaultPivotTableStyle = None
    DefaultSlicerStyle = None
    DefaultTableStyle = None
    DefaultTimelineStyle = None
    DisplayDrawingObjects: 'XlDisplayDrawingObjects'
    DisplayInkComments: bool
    DocumentInspectors: 'DocumentInspectors'
    DocumentLibraryVersions: 'DocumentLibraryVersions'
    DoNotPromptForConvert: bool
    EnableAutoRecover: bool
    EncryptionProvider: str
    EnvelopeVisible: bool
    Excel4IntlMacroSheets: Sheets
    Excel4MacroSheets: Sheets
    Excel8CompatibilityMode: bool
    FileFormat: 'XlFileFormat'
    Final: bool
    ForceFullCalculation: bool
    FullName: str
    FullNameURLEncoded: str
    HasPassword: bool
    HasVBProject: bool
    HighlightChangesOnScreen: bool
    IconSets: IconSets
    InactiveListBorderVisible: bool
    IsAddin: bool
    IsInplace: bool
    KeepChangeHistory: bool
    ListChangesOnNewSheet: bool
    Mailer: Mailer
    Model: Model
    MultiUserEditing: bool
    Name: str
    Names: Names
    Parent = None
    Password: str
    PasswordEncryptionAlgorithm: str
    PasswordEncryptionFileProperties: bool
    PasswordEncryptionKeyLength: float
    PasswordEncryptionProvider: str
    Path: str
    Permission: 'Permission'
    PersonalViewListSettings: bool
    PersonalViewPrintSettings: bool
    PivotTables = None
    PrecisionAsDisplayed: bool
    ProtectStructure: bool
    ProtectWindows: bool
    PublishObjects: PublishObjects
    Queries: Queries
    ReadOnly: bool
    ReadOnlyRecommended: bool
    RemovePersonalInformation: bool
    Research: Research
    RevisionNumber: float
    Saved: bool
    SaveLinkValues: bool
    SensitivityLabel: 'SensitivityLabel'
    ServerPolicy: 'ServerPolicy'
    ServerViewableItems: ServerViewableItems
    SharedWorkspace: 'SharedWorkspace'
    Sheets: Sheets
    ShowConflictHistory: bool
    ShowPivotChartActiveFields: bool
    ShowPivotTableFieldList: bool
    Signatures: 'SignatureSet'
    SlicerCaches: SlicerCaches
    SmartDocument: 'SmartDocument'
    Styles: Styles
    Sync: 'Sync'
    TableStyles: TableStyles
    TemplateRemoveExtData: bool
    Theme: 'OfficeTheme'
    UpdateLinks: 'XlUpdateLinks'
    UpdateRemoteReferences: bool
    UserStatus = None
    UseWholeCellCriteria: bool
    UseWildcards: bool
    VBASigned: bool
    VBProject: 'VBProject'
    WebOptions: WebOptions
    Windows: Windows
    WorkIdentity: str
    Worksheets: Sheets
    WritePassword: str
    WriteReserved: bool
    WriteReservedBy: str
    XmlMaps: 'XmlMaps'
    XmlNamespaces: 'XmlNamespaces'
    def AcceptAllChanges(self, When = None, Who = None, Where = None): ...
    def AddToFavorites(self, ): ...
    def ApplyTheme(self, Filename: str): ...
    def BreakLink(self, Name: str, Type: 'XlLinkType'): ...
    def CanCheckIn(self, ) -> bool: ...
    def ChangeFileAccess(self, Mode: 'XlFileAccess', WritePassword = None, Notify = None): ...
    def ChangeLink(self, Name: str, NewName: str, Type: 'XlLinkType' = None): ...
    def CheckIn(self, SaveChanges = None, Comments = None, MakePublic = None): ...
    def CheckInWithVersion(self, SaveChanges = None, Comments = None, MakePublic = None, VersionType = None): ...
    def Close(self, SaveChanges = None, Filename = None, RouteWorkbook = None): ...
    @property
    def Colors(self, Index = None): ...
    def ConvertComments(self, ): ...
    def CreateForecastSheet(self, Timeline: Range, Values: Range, ForecastStart = None, ForecastEnd = None, ConfInt = None, Seasonality = None, DataCompletion = None, Aggregation = None, ChartType = None, ShowStatsTable = None): ...
    def DeleteNumberFormat(self, NumberFormat: str): ...
    def EnableConnections(self, ): ...
    def EndReview(self, ): ...
    def ExclusiveAccess(self, ) -> bool: ...
    def ExportAsFixedFormat(self, Type: 'XlFixedFormatType', Filename = None, Quality = None, IncludeDocProperties = None, IgnorePrintAreas = None, From = None, To = None, OpenAfterPublish = None, FixedFormatExtClassPtr = None, WorkIdentity = None): ...
    def FollowHyperlink(self, Address: str, SubAddress = None, NewWindow = None, AddHistory = None, ExtraInfo = None, Method = None, HeaderInfo = None): ...
    def ForwardMailer(self, ): ...
    def GetWorkflowTasks(self, ) -> 'WorkflowTasks': ...
    def GetWorkflowTemplates(self, ) -> 'WorkflowTemplates': ...
    def HighlightChangesOptions(self, When = None, Who = None, Where = None): ...
    def LinkInfo(self, Name: str, LinkInfo: 'XlLinkInfo', Type = None, EditionRef = None): ...
    def LinkSources(self, Type = None): ...
    def LockServerFile(self, ): ...
    def MergeWorkbook(self, Filename): ...
    def NewWindow(self, ) -> Window: ...
    def OpenLinks(self, Name: str, ReadOnly = None, Type = None): ...
    def PivotCaches(self, ) -> PivotCaches: ...
    def Post(self, DestName = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None, IgnorePrintAreas = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Protect(self, Password = None, Structure = None, Windows = None): ...
    def ProtectSharing(self, Filename = None, Password = None, WriteResPassword = None, ReadOnlyRecommended = None, CreateBackup = None, SharingPassword = None, FileFormat = None): ...
    def PublishToPBI(self, PublishType = None, nameConflict = None, bstrGroupName = None) -> str: ...
    def PurgeChangeHistoryNow(self, Days: float, SharingPassword = None): ...
    def RefreshAll(self, ): ...
    def RejectAllChanges(self, When = None, Who = None, Where = None): ...
    def ReloadAs(self, Encoding: 'MsoEncoding'): ...
    def RemoveDocumentInformation(self, RemoveDocInfoType: 'XlRemoveDocInfoType'): ...
    def RemoveUser(self, Index: float): ...
    def Reply(self, ): ...
    def ReplyAll(self, ): ...
    def ReplyWithChanges(self, ShowMessage = None): ...
    def ResetColors(self, ): ...
    def RunAutoMacros(self, Which: 'XlRunAutoMacro'): ...
    def Save(self, ): ...
    def SaveAs(self, Filename = None, FileFormat = None, Password = None, WriteResPassword = None, ReadOnlyRecommended = None, CreateBackup = None, AccessMode: 'XlSaveAsAccessMode' = None, ConflictResolution = None, AddToMru = None, TextCodepage = None, TextVisualLayout = None, Local = None, WorkIdentity = None): ...
    def SaveAsXMLData(self, Filename: str, Map: 'XmlMap'): ...
    def SaveCopyAs(self, Filename = None): ...
    def SendFaxOverInternet(self, Recipients = None, Subject = None, ShowMessage = None): ...
    def SendForReview(self, Recipients = None, Subject = None, ShowMessage = None, IncludeAttachment = None): ...
    def SendMail(self, Recipients, Subject = None, ReturnReceipt = None): ...
    def SendMailer(self, FileFormat = None, Priority: 'XlPriority' = None): ...
    def SetLinkOnData(self, Name: str, Procedure = None): ...
    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider = None, PasswordEncryptionAlgorithm = None, PasswordEncryptionKeyLength = None, PasswordEncryptionFileProperties = None): ...
    def ToggleFormsDesign(self, ): ...
    def Unprotect(self, Password = None): ...
    def UnprotectSharing(self, SharingPassword = None): ...
    def UpdateFromFile(self, ): ...
    def UpdateLink(self, Name = None, Type = None): ...
    def WebPagePreview(self, ): ...
    def XmlImport(self, Url: str, ImportMap: 'XmlMap', Overwrite = None, Destination = None) -> 'XlXmlImportResult': ...
    def XmlImportXml(self, Data: str, ImportMap: 'XmlMap', Overwrite = None, Destination = None) -> 'XlXmlImportResult': ...

class WorkbookConnection:
    Application: Application
    Creator: 'XlCreator'
    DataFeedConnection: DataFeedConnection
    Description: str
    InModel: bool
    ModelConnection: ModelConnection
    ModelTables: ModelTables
    Name: str
    ODBCConnection: ODBCConnection
    OLEDBConnection: OLEDBConnection
    Parent = None
    Ranges: Ranges
    RefreshWithRefreshAll: bool
    TextConnection: TextConnection
    Type: 'XlConnectionType'
    WorksheetDataConnection: 'WorksheetDataConnection'
    def Delete(self, ): ...
    def Refresh(self, ): ...

class WorkbookQuery:
    Application: Application
    Creator: 'XlCreator'
    Description: str
    Formula: str
    Name: str
    Parent = None
    def Delete(self, DeleteConnection = None): ...
    def Refresh(self, ): ...

class Workbooks:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    Parent = None
    def __call__(self, Index) -> 'Workbook': ...
    def Add(self, Template = None) -> Workbook: ...
    def CanCheckOut(self, Filename: str) -> bool: ...
    def CheckOut(self, Filename: str): ...
    def Close(self, ): ...
    @property
    def Item(self, Index) -> Workbook: ...
    def Open(self, Filename: str, UpdateLinks = None, ReadOnly = None, Format = None, Password = None, WriteResPassword = None, IgnoreReadOnlyRecommended = None, Origin = None, Delimiter = None, Editable = None, Notify = None, Converter = None, AddToMru = None, Local = None, CorruptLoad = None) -> Workbook: ...
    def OpenDatabase(self, Filename: str, CommandText = None, CommandType = None, BackgroundQuery = None, ImportDataAs = None) -> Workbook: ...
    def OpenText(self, Filename: str, Origin = None, StartRow = None, DataType = None, TextQualifier: 'XlTextQualifier' = None, ConsecutiveDelimiter = None, Tab = None, Semicolon = None, Comma = None, Space = None, Other = None, OtherChar = None, FieldInfo = None, TextVisualLayout = None, DecimalSeparator = None, ThousandsSeparator = None, TrailingMinusNumbers = None, Local = None): ...
    def OpenXML(self, Filename: str, Stylesheets = None, LoadOption = None) -> Workbook: ...

class Worksheet:
    Application: Application
    AutoFilter: AutoFilter
    AutoFilterMode: bool
    Cells: Range
    CircularReference: Range
    CodeName: str
    Columns: Range
    Comments: Comments
    CommentsThreaded: CommentsThreaded
    ConsolidationFunction: 'XlConsolidationFunction'
    ConsolidationOptions = None
    ConsolidationSources = None
    Creator: 'XlCreator'
    CustomProperties: CustomProperties
    DisplayPageBreaks: bool
    DisplayRightToLeft: bool
    EnableAutoFilter: bool
    EnableCalculation: bool
    EnableFormatConditionsCalculation: bool
    EnableOutlining: bool
    EnablePivotTable: bool
    EnableSelection: 'XlEnableSelection'
    FilterMode: bool
    HPageBreaks: HPageBreaks
    Hyperlinks: Hyperlinks
    Index: float
    ListObjects: ListObjects
    MailEnvelope: 'MsoEnvelope'
    Name: str
    NamedSheetViews: NamedSheetViewCollection
    Names: Names
    Next = None
    Outline: Outline
    PageSetup: PageSetup
    Parent = None
    Previous = None
    PrintedCommentPages: float
    ProtectContents: bool
    ProtectDrawingObjects: bool
    Protection: Protection
    ProtectionMode: bool
    ProtectScenarios: bool
    QueryTables: QueryTables
    Rows: Range
    ScrollArea: str
    Shapes: Shapes
    Sort: Sort
    StandardHeight: float
    StandardWidth: float
    Tab: Tab
    TransitionExpEval: bool
    TransitionFormEntry: bool
    Type: 'XlSheetType'
    UsedRange: Range
    Visible: 'XlSheetVisibility'
    VPageBreaks: VPageBreaks
    def Activate(self, ): ...
    def Calculate(self, ): ...
    def ChartObjects(self, Index = None): ...
    def CheckSpelling(self, CustomDictionary = None, IgnoreUppercase = None, AlwaysSuggest = None, SpellLang = None): ...
    def CircleInvalid(self, ): ...
    def ClearArrows(self, ): ...
    def ClearCircles(self, ): ...
    def Copy(self, Before = None, After = None): ...
    def Delete(self, ): ...
    def Evaluate(self, Name): ...
    def ExportAsFixedFormat(self, Type: 'XlFixedFormatType', Filename = None, Quality = None, IncludeDocProperties = None, IgnorePrintAreas = None, From = None, To = None, OpenAfterPublish = None, FixedFormatExtClassPtr = None, WorkIdentity = None): ...
    def Move(self, Before = None, After = None): ...
    def OLEObjects(self, Index = None): ...
    def Paste(self, Destination = None, Link = None): ...
    def PasteSpecial(self, Format = None, Link = None, DisplayAsIcon = None, IconFileName = None, IconIndex = None, IconLabel = None, NoHTMLFormatting = None): ...
    def PivotTables(self, Index = None): ...
    def PivotTableWizard(self, SourceType = None, SourceData = None, TableDestination = None, TableName = None, RowGrand = None, ColumnGrand = None, SaveData = None, HasAutoFormat = None, AutoPage = None, Reserved = None, BackgroundQuery = None, OptimizeCache = None, PageFieldOrder = None, PageFieldWrapCount = None, ReadData = None, Connection = None) -> PivotTable: ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None, IgnorePrintAreas = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Protect(self, Password = None, DrawingObjects = None, Contents = None, Scenarios = None, UserInterfaceOnly = None, AllowFormattingCells = None, AllowFormattingColumns = None, AllowFormattingRows = None, AllowInsertingColumns = None, AllowInsertingRows = None, AllowInsertingHyperlinks = None, AllowDeletingColumns = None, AllowDeletingRows = None, AllowSorting = None, AllowFiltering = None, AllowUsingPivotTables = None): ...
    @property
    def Range(self, Cell1, Cell2 = None) -> Range: ...
    def ResetAllPageBreaks(self, ): ...
    def SaveAs(self, Filename: str, FileFormat = None, Password = None, WriteResPassword = None, ReadOnlyRecommended = None, CreateBackup = None, AddToMru = None, TextCodepage = None, TextVisualLayout = None, Local = None): ...
    def Scenarios(self, Index = None): ...
    def Select(self, Replace = None): ...
    def SetBackgroundPicture(self, Filename: str): ...
    def ShowAllData(self, ): ...
    def ShowDataForm(self, ): ...
    def Unprotect(self, Password = None): ...
    def XmlDataQuery(self, XPath: str, SelectionNamespaces = None, Map = None) -> Range: ...
    def XmlMapQuery(self, XPath: str, SelectionNamespaces = None, Map = None) -> Range: ...

class WorksheetDataConnection:
    Application: Application
    CommandText = None
    CommandType: 'XlCmdType'
    Connection = None
    Creator: 'XlCreator'
    Parent = None

class WorksheetFunction:
    Application: Application
    Creator: 'XlCreator'
    Parent = None
    def AccrInt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7 = None) -> float: ...
    def AccrIntM(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def Acos(self, Arg1: float) -> float: ...
    def Acosh(self, Arg1: float) -> float: ...
    def Acot(self, Arg1: float) -> float: ...
    def Acoth(self, Arg1: float) -> float: ...
    def Aggregate(self, Arg1: float, Arg2: float, Arg3: Range, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def AmorDegrc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7 = None) -> float: ...
    def AmorLinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7 = None) -> float: ...
    def And(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> bool: ...
    def Arabic(self, Arg1: str) -> float: ...
    def ArrayToText(self, Arg1, Arg2 = None) -> str: ...
    def Asc(self, Arg1: str) -> str: ...
    def Asin(self, Arg1: float) -> float: ...
    def Asinh(self, Arg1: float) -> float: ...
    def Atan2(self, Arg1: float, Arg2: float) -> float: ...
    def Atanh(self, Arg1: float) -> float: ...
    def AveDev(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Average(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def AverageIf(self, Arg1: Range, Arg2, Arg3 = None) -> float: ...
    def AverageIfs(self, Arg1: Range, Arg2: Range, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> float: ...
    def BahtText(self, Arg1: float) -> str: ...
    def Base(self, Arg1: float, Arg2: float, Arg3 = None) -> str: ...
    def BesselI(self, Arg1, Arg2) -> float: ...
    def BesselJ(self, Arg1, Arg2) -> float: ...
    def BesselK(self, Arg1, Arg2) -> float: ...
    def BesselY(self, Arg1, Arg2) -> float: ...
    def Beta_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool, Arg5 = None, Arg6 = None) -> float: ...
    def Beta_Inv(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def BetaDist(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def BetaInv(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def Bin2Dec(self, Arg1) -> str: ...
    def Bin2Hex(self, Arg1, Arg2 = None) -> str: ...
    def Bin2Oct(self, Arg1, Arg2 = None) -> str: ...
    def Binom_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def Binom_Dist_Range(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None) -> float: ...
    def Binom_Inv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def BinomDist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def Bitand(self, Arg1: float, Arg2: float) -> float: ...
    def Bitlshift(self, Arg1: float, Arg2: float) -> float: ...
    def Bitor(self, Arg1: float, Arg2: float) -> float: ...
    def Bitrshift(self, Arg1: float, Arg2: float) -> float: ...
    def Bitxor(self, Arg1: float, Arg2: float) -> float: ...
    def Ceiling(self, Arg1: float, Arg2: float) -> float: ...
    def Ceiling_Math(self, Arg1: float, Arg2 = None, Arg3 = None) -> float: ...
    def Ceiling_Precise(self, Arg1: float, Arg2 = None) -> float: ...
    def ChiDist(self, Arg1: float, Arg2: float) -> float: ...
    def ChiInv(self, Arg1: float, Arg2: float) -> float: ...
    def ChiSq_Dist(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def ChiSq_Dist_RT(self, Arg1: float, Arg2: float) -> float: ...
    def ChiSq_Inv(self, Arg1: float, Arg2: float) -> float: ...
    def ChiSq_Inv_RT(self, Arg1: float, Arg2: float) -> float: ...
    def ChiSq_Test(self, Arg1, Arg2) -> float: ...
    def ChiTest(self, Arg1, Arg2) -> float: ...
    def Choose(self, Arg1, Arg2, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None): ...
    def Clean(self, Arg1: str) -> str: ...
    def Combin(self, Arg1: float, Arg2: float) -> float: ...
    def Combina(self, Arg1: float, Arg2: float) -> float: ...
    def Complex(self, Arg1, Arg2, Arg3 = None) -> str: ...
    def Concat(self, Arg1: str, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> str: ...
    def Confidence(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Confidence_Norm(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Confidence_T(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Convert(self, Arg1, Arg2, Arg3) -> float: ...
    def Correl(self, Arg1, Arg2) -> float: ...
    def Cosh(self, Arg1: float) -> float: ...
    def Cot(self, Arg1: float) -> float: ...
    def Coth(self, Arg1: float) -> float: ...
    def Count(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def CountA(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def CountBlank(self, Arg1: Range) -> float: ...
    def CountIf(self, Arg1: Range, Arg2) -> float: ...
    def CountIfs(self, Arg1: Range, Arg2, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def CoupDayBs(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def CoupDays(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def CoupDaysNc(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def CoupNcd(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def CoupNum(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def CoupPcd(self, Arg1, Arg2, Arg3, Arg4 = None) -> float: ...
    def Covar(self, Arg1, Arg2) -> float: ...
    def Covariance_P(self, Arg1, Arg2) -> float: ...
    def Covariance_S(self, Arg1, Arg2) -> float: ...
    def CritBinom(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Csc(self, Arg1: float) -> float: ...
    def Csch(self, Arg1: float) -> float: ...
    def CumIPmt(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float: ...
    def CumPrinc(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6) -> float: ...
    def DAverage(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def Days(self, Arg1, Arg2) -> float: ...
    def Days360(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def Db(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5 = None) -> float: ...
    def Dbcs(self, Arg1: str) -> str: ...
    def DCount(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DCountA(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def Ddb(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5 = None) -> float: ...
    def Dec2Bin(self, Arg1, Arg2 = None) -> str: ...
    def Dec2Hex(self, Arg1, Arg2 = None) -> str: ...
    def Dec2Oct(self, Arg1, Arg2 = None) -> str: ...
    def Decimal(self, Arg1: str, Arg2: float) -> float: ...
    def Degrees(self, Arg1: float) -> float: ...
    def Delta(self, Arg1, Arg2 = None) -> float: ...
    def DevSq(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def DGet(self, Arg1: Range, Arg2, Arg3): ...
    def Disc(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def DMax(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DMin(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def Dollar(self, Arg1: float, Arg2 = None) -> str: ...
    def DollarDe(self, Arg1, Arg2) -> float: ...
    def DollarFr(self, Arg1, Arg2) -> float: ...
    def DProduct(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DStDev(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DStDevP(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DSum(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def Duration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6 = None) -> float: ...
    def DVar(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def DVarP(self, Arg1: Range, Arg2, Arg3) -> float: ...
    def EDate(self, Arg1, Arg2) -> float: ...
    def Effect(self, Arg1, Arg2) -> float: ...
    def EncodeURL(self, Arg1: str): ...
    def EoMonth(self, Arg1, Arg2) -> float: ...
    def Erf(self, Arg1, Arg2 = None) -> float: ...
    def Erf_Precise(self, Arg1) -> float: ...
    def ErfC(self, Arg1) -> float: ...
    def ErfC_Precise(self, Arg1) -> float: ...
    def Even(self, Arg1: float) -> float: ...
    def Expon_Dist(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def ExponDist(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def F_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def F_Dist_RT(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def F_Inv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def F_Inv_RT(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def F_Test(self, Arg1, Arg2) -> float: ...
    def Fact(self, Arg1: float) -> float: ...
    def FactDouble(self, Arg1) -> float: ...
    def FDist(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def FieldValue(self, Arg1, Arg2: str): ...
    def Filter(self, Arg1, Arg2, Arg3 = None): ...
    def FilterXML(self, Arg1: str, Arg2: str): ...
    def Find(self, Arg1: str, Arg2: str, Arg3 = None) -> float: ...
    def FindB(self, Arg1: str, Arg2: str, Arg3 = None) -> float: ...
    def FInv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Fisher(self, Arg1: float) -> float: ...
    def FisherInv(self, Arg1: float) -> float: ...
    def Fixed(self, Arg1: float, Arg2 = None, Arg3 = None) -> str: ...
    def Floor(self, Arg1: float, Arg2: float) -> float: ...
    def Floor_Math(self, Arg1: float, Arg2 = None, Arg3 = None) -> float: ...
    def Floor_Precise(self, Arg1: float, Arg2 = None) -> float: ...
    def Forecast_ETS(self, Arg1: float, Arg2, Arg3, Arg4 = None, Arg5 = None, Arg6 = None) -> float: ...
    def Forecast_ETS_ConfInt(self, Arg1: float, Arg2, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None) -> float: ...
    def Forecast_ETS_Seasonality(self, Arg1, Arg2, Arg3 = None, Arg4 = None) -> float: ...
    def Forecast_ETS_STAT(self, Arg1, Arg2, Arg3: float, Arg4 = None, Arg5 = None, Arg6 = None) -> float: ...
    def Forecast_Linear(self, Arg1: float, Arg2, Arg3) -> float: ...
    def Frequency(self, Arg1, Arg2): ...
    def FTest(self, Arg1, Arg2) -> float: ...
    def Fv(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def FVSchedule(self, Arg1, Arg2) -> float: ...
    def Gamma(self, Arg1: float) -> float: ...
    def Gamma_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def Gamma_Inv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def GammaDist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def GammaInv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def GammaLn(self, Arg1: float) -> float: ...
    def GammaLn_Precise(self, Arg1: float) -> float: ...
    def Gauss(self, Arg1: float) -> float: ...
    def Gcd(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def GeoMean(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def GeStep(self, Arg1, Arg2 = None) -> float: ...
    def Growth(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def HarMean(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Hex2Bin(self, Arg1, Arg2 = None) -> str: ...
    def Hex2Dec(self, Arg1) -> str: ...
    def Hex2Oct(self, Arg1, Arg2 = None) -> str: ...
    def HLookup(self, Arg1, Arg2, Arg3, Arg4 = None): ...
    def HypGeom_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5: bool) -> float: ...
    def HypGeomDist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float) -> float: ...
    def IfError(self, Arg1, Arg2): ...
    def IfNa(self, Arg1, Arg2): ...
    def ImAbs(self, Arg1) -> str: ...
    def Imaginary(self, Arg1) -> float: ...
    def ImArgument(self, Arg1) -> str: ...
    def ImConjugate(self, Arg1) -> str: ...
    def ImCos(self, Arg1) -> str: ...
    def ImCosh(self, Arg1) -> str: ...
    def ImCot(self, Arg1) -> str: ...
    def ImCsc(self, Arg1) -> str: ...
    def ImCsch(self, Arg1) -> str: ...
    def ImDiv(self, Arg1, Arg2) -> str: ...
    def ImExp(self, Arg1) -> str: ...
    def ImLn(self, Arg1) -> str: ...
    def ImLog10(self, Arg1) -> str: ...
    def ImLog2(self, Arg1) -> str: ...
    def ImPower(self, Arg1, Arg2) -> str: ...
    def ImProduct(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> str: ...
    def ImReal(self, Arg1) -> float: ...
    def ImSec(self, Arg1) -> str: ...
    def ImSech(self, Arg1) -> str: ...
    def ImSin(self, Arg1) -> str: ...
    def ImSinh(self, Arg1) -> str: ...
    def ImSqrt(self, Arg1) -> str: ...
    def ImSub(self, Arg1, Arg2) -> str: ...
    def ImSum(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> str: ...
    def ImTan(self, Arg1) -> str: ...
    def Index(self, Arg1, Arg2: float, Arg3 = None, Arg4 = None): ...
    def Intercept(self, Arg1, Arg2) -> float: ...
    def IntRate(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def Ipmt(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5 = None, Arg6 = None) -> float: ...
    def Irr(self, Arg1, Arg2 = None) -> float: ...
    def IsErr(self, Arg1) -> bool: ...
    def IsError(self, Arg1) -> bool: ...
    def IsEven(self, Arg1) -> bool: ...
    def IsFormula(self, Arg1: Range) -> bool: ...
    def IsLogical(self, Arg1) -> bool: ...
    def IsNA(self, Arg1) -> bool: ...
    def IsNonText(self, Arg1) -> bool: ...
    def IsNumber(self, Arg1) -> bool: ...
    def ISO_Ceiling(self, Arg1: float, Arg2 = None) -> float: ...
    def IsOdd(self, Arg1) -> bool: ...
    def IsoWeekNum(self, Arg1: float, Arg2 = None) -> float: ...
    def Ispmt(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float) -> float: ...
    def IsText(self, Arg1) -> bool: ...
    def Kurt(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Large(self, Arg1, Arg2: float) -> float: ...
    def Lcm(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def LinEst(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def Ln(self, Arg1: float) -> float: ...
    def Log(self, Arg1: float, Arg2 = None) -> float: ...
    def Log10(self, Arg1: float) -> float: ...
    def LogEst(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def LogInv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def LogNorm_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def LogNorm_Inv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def LogNormDist(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Lookup(self, Arg1, Arg2, Arg3 = None): ...
    def Match(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def Max(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def MaxIfs(self, Arg1: Range, Arg2: Range, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> float: ...
    def MDeterm(self, Arg1) -> float: ...
    def MDuration(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6 = None) -> float: ...
    def Median(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Min(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def MinIfs(self, Arg1: Range, Arg2: Range, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> float: ...
    def MInverse(self, Arg1): ...
    def MIrr(self, Arg1, Arg2: float, Arg3: float) -> float: ...
    def MMult(self, Arg1, Arg2): ...
    def Mode(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Mode_Mult(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None): ...
    def Mode_Sngl(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def MRound(self, Arg1, Arg2) -> float: ...
    def MultiNomial(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Munit(self, Arg1: float): ...
    def NegBinom_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def NegBinomDist(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def NetworkDays(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def NetworkDays_Intl(self, Arg1, Arg2, Arg3 = None, Arg4 = None) -> float: ...
    def Nominal(self, Arg1, Arg2) -> float: ...
    def Norm_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def Norm_Inv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Norm_S_Dist(self, Arg1: float, Arg2: bool) -> float: ...
    def Norm_S_Inv(self, Arg1: float) -> float: ...
    def NormDist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def NormInv(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def NormSDist(self, Arg1: float) -> float: ...
    def NormSInv(self, Arg1: float) -> float: ...
    def NPer(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def Npv(self, Arg1: float, Arg2, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def NumberValue(self, Arg1: str, Arg2: str, Arg3: str) -> float: ...
    def Oct2Bin(self, Arg1, Arg2 = None) -> str: ...
    def Oct2Dec(self, Arg1) -> str: ...
    def Oct2Hex(self, Arg1, Arg2 = None) -> str: ...
    def Odd(self, Arg1: float) -> float: ...
    def OddFPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9 = None) -> float: ...
    def OddFYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9 = None) -> float: ...
    def OddLPrice(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8 = None) -> float: ...
    def OddLYield(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8 = None) -> float: ...
    def Or(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> bool: ...
    def PDuration(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Pearson(self, Arg1, Arg2) -> float: ...
    def Percentile(self, Arg1, Arg2: float) -> float: ...
    def Percentile_Exc(self, Arg1, Arg2: float) -> float: ...
    def Percentile_Inc(self, Arg1, Arg2: float) -> float: ...
    def PercentRank(self, Arg1, Arg2: float, Arg3 = None) -> float: ...
    def PercentRank_Exc(self, Arg1, Arg2: float, Arg3 = None) -> float: ...
    def PercentRank_Inc(self, Arg1, Arg2: float, Arg3 = None) -> float: ...
    def Permut(self, Arg1: float, Arg2: float) -> float: ...
    def Permutationa(self, Arg1: float, Arg2: float) -> float: ...
    def Phi(self, Arg1: float) -> float: ...
    def Phonetic(self, Arg1: Range) -> str: ...
    def Pi(self, ) -> float: ...
    def Pmt(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def Poisson(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def Poisson_Dist(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def Power(self, Arg1: float, Arg2: float) -> float: ...
    def Ppmt(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5 = None, Arg6 = None) -> float: ...
    def Price(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7 = None) -> float: ...
    def PriceDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def PriceMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6 = None) -> float: ...
    def Prob(self, Arg1, Arg2, Arg3: float, Arg4 = None) -> float: ...
    def Product(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Proper(self, Arg1: str) -> str: ...
    def Pv(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None) -> float: ...
    def Quartile(self, Arg1, Arg2: float) -> float: ...
    def Quartile_Exc(self, Arg1, Arg2: float) -> float: ...
    def Quartile_Inc(self, Arg1, Arg2: float) -> float: ...
    def Quotient(self, Arg1, Arg2) -> float: ...
    def Radians(self, Arg1: float) -> float: ...
    def RandArray(self, Arg1 = None, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None): ...
    def RandBetween(self, Arg1, Arg2) -> float: ...
    def Rank(self, Arg1: float, Arg2: Range, Arg3 = None) -> float: ...
    def Rank_Avg(self, Arg1: float, Arg2: Range, Arg3 = None) -> float: ...
    def Rank_Eq(self, Arg1: float, Arg2: Range, Arg3 = None) -> float: ...
    def Rate(self, Arg1: float, Arg2: float, Arg3: float, Arg4 = None, Arg5 = None, Arg6 = None) -> float: ...
    def Received(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def Replace(self, Arg1: str, Arg2: float, Arg3: float, Arg4: str) -> str: ...
    def ReplaceB(self, Arg1: str, Arg2: float, Arg3: float, Arg4: str) -> str: ...
    def Rept(self, Arg1: str, Arg2: float) -> str: ...
    def Roman(self, Arg1: float, Arg2 = None) -> str: ...
    def Round(self, Arg1: float, Arg2: float) -> float: ...
    def RoundDown(self, Arg1: float, Arg2: float) -> float: ...
    def RoundUp(self, Arg1: float, Arg2: float) -> float: ...
    def Rri(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def RSq(self, Arg1, Arg2) -> float: ...
    def RTD(self, progID, server, topic1, topic2 = None, topic3 = None, topic4 = None, topic5 = None, topic6 = None, topic7 = None, topic8 = None, topic9 = None, topic10 = None, topic11 = None, topic12 = None, topic13 = None, topic14 = None, topic15 = None, topic16 = None, topic17 = None, topic18 = None, topic19 = None, topic20 = None, topic21 = None, topic22 = None, topic23 = None, topic24 = None, topic25 = None, topic26 = None, topic27 = None, topic28 = None): ...
    def Search(self, Arg1: str, Arg2: str, Arg3 = None) -> float: ...
    def SearchB(self, Arg1: str, Arg2: str, Arg3 = None) -> float: ...
    def Sec(self, Arg1: float) -> float: ...
    def Sech(self, Arg1: float) -> float: ...
    def Sequence(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def SeriesSum(self, Arg1, Arg2, Arg3, Arg4) -> float: ...
    def Single(self, Arg1): ...
    def Sinh(self, Arg1: float) -> float: ...
    def Skew(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Skew_p(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Sln(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Slope(self, Arg1, Arg2) -> float: ...
    def Small(self, Arg1, Arg2: float) -> float: ...
    def Sort(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def SortBy(self, Arg1, Arg2, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None): ...
    def SqrtPi(self, Arg1) -> float: ...
    def Standardize(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def StDev(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def StDev_P(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def StDev_S(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def StDevP(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def StEyx(self, Arg1, Arg2) -> float: ...
    def StockHistory(self, Arg1, Arg2, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None): ...
    def Substitute(self, Arg1: str, Arg2: str, Arg3: str, Arg4 = None) -> str: ...
    def Subtotal(self, Arg1: float, Arg2: Range, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Sum(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def SumIf(self, Arg1: Range, Arg2, Arg3 = None) -> float: ...
    def SumIfs(self, Arg1: Range, Arg2: Range, Arg3, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> float: ...
    def SumProduct(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def SumSq(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def SumX2MY2(self, Arg1, Arg2) -> float: ...
    def SumX2PY2(self, Arg1, Arg2) -> float: ...
    def SumXMY2(self, Arg1, Arg2) -> float: ...
    def Syd(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float) -> float: ...
    def T_Dist(self, Arg1: float, Arg2: float, Arg3: bool) -> float: ...
    def T_Dist_2T(self, Arg1: float, Arg2: float) -> float: ...
    def T_Dist_RT(self, Arg1: float, Arg2: float) -> float: ...
    def T_Inv(self, Arg1: float, Arg2: float) -> float: ...
    def T_Inv_2T(self, Arg1: float, Arg2: float) -> float: ...
    def T_Test(self, Arg1, Arg2, Arg3: float, Arg4: float) -> float: ...
    def Tanh(self, Arg1: float) -> float: ...
    def TBillEq(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def TBillPrice(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def TBillYield(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def TDist(self, Arg1: float, Arg2: float, Arg3: float) -> float: ...
    def Text(self, Arg1, Arg2: str) -> str: ...
    def TextJoin(self, Arg1: str, Arg2: bool, Arg3: str, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None) -> str: ...
    def TInv(self, Arg1: float, Arg2: float) -> float: ...
    def Transpose(self, Arg1): ...
    def Trend(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None): ...
    def Trim(self, Arg1: str) -> str: ...
    def TrimMean(self, Arg1, Arg2: float) -> float: ...
    def TTest(self, Arg1, Arg2, Arg3: float, Arg4: float) -> float: ...
    def Unichar(self, Arg1: float) -> str: ...
    def Unicode(self, Arg1: str) -> float: ...
    def Unique(self, Arg1, Arg2 = None, Arg3 = None): ...
    def USDollar(self, Arg1: float, Arg2: float) -> str: ...
    def ValueToText(self, Arg1, Arg2 = None) -> str: ...
    def Var(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Var_P(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Var_S(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def VarP(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> float: ...
    def Vdb(self, Arg1: float, Arg2: float, Arg3: float, Arg4: float, Arg5: float, Arg6 = None, Arg7 = None) -> float: ...
    def VLookup(self, Arg1, Arg2, Arg3, Arg4 = None): ...
    def WebService(self, Arg1: str): ...
    def Weekday(self, Arg1, Arg2 = None) -> float: ...
    def WeekNum(self, Arg1, Arg2 = None) -> float: ...
    def Weibull(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def Weibull_Dist(self, Arg1: float, Arg2: float, Arg3: float, Arg4: bool) -> float: ...
    def WorkDay(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def WorkDay_Intl(self, Arg1, Arg2, Arg3 = None, Arg4 = None) -> float: ...
    def Xirr(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def XLookup(self, Arg1, Arg2, Arg3, Arg4 = None, Arg5 = None, Arg6 = None): ...
    def XMatch(self, Arg1, Arg2, Arg3 = None, Arg4 = None) -> float: ...
    def Xnpv(self, Arg1, Arg2) -> float: ...
    def Xor(self, Arg1, Arg2 = None, Arg3 = None, Arg4 = None, Arg5 = None, Arg6 = None, Arg7 = None, Arg8 = None, Arg9 = None, Arg10 = None, Arg11 = None, Arg12 = None, Arg13 = None, Arg14 = None, Arg15 = None, Arg16 = None, Arg17 = None, Arg18 = None, Arg19 = None, Arg20 = None, Arg21 = None, Arg22 = None, Arg23 = None, Arg24 = None, Arg25 = None, Arg26 = None, Arg27 = None, Arg28 = None, Arg29 = None, Arg30 = None) -> bool: ...
    def YearFrac(self, Arg1, Arg2, Arg3 = None) -> float: ...
    def YieldDisc(self, Arg1, Arg2, Arg3, Arg4, Arg5 = None) -> float: ...
    def YieldMat(self, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6 = None) -> float: ...
    def Z_Test(self, Arg1, Arg2: float, Arg3 = None) -> float: ...
    def ZTest(self, Arg1, Arg2: float, Arg3 = None) -> float: ...

class Worksheets:
    Application: Application
    Count: float
    Creator: 'XlCreator'
    HPageBreaks: HPageBreaks
    Parent = None
    Visible = None
    VPageBreaks: VPageBreaks
    def __call__(self, Index) -> 'Worksheet': ...
    def Add(self, Before = None, After = None, Count = None, Type = None): ...
    def Add2(self, Before = None, After = None, Count = None, NewLayout = None): ...
    def Copy(self, Before = None, After = None): ...
    def Delete(self, ): ...
    def FillAcrossSheets(self, Range: Range, Type: 'XlFillWith' = None): ...
    @property
    def Item(self, Index) -> Worksheet: ...
    def Move(self, Before = None, After = None): ...
    def PrintOut(self, From = None, To = None, Copies = None, Preview = None, ActivePrinter = None, PrintToFile = None, Collate = None, PrToFileName = None, IgnorePrintAreas = None): ...
    def PrintPreview(self, EnableChanges = None): ...
    def Select(self, Replace = None): ...

class WorksheetView:
    Application: Application
    Creator: 'XlCreator'
    DisplayFormulas: bool
    DisplayGridlines: bool
    DisplayHeadings: bool
    DisplayOutline: bool
    DisplayZeros: bool
    Parent = None
    Sheet = None

class XlAboveBelow(IntEnum):
    xlAboveAverage = 0
    xlAboveStdDev = 4
    xlBelowAverage = 1
    xlBelowStdDev = 5
    xlEqualAboveAverage = 2
    xlEqualBelowAverage = 3

class XlActionType(IntEnum):
    xlActionTypeDrillthrough = 256 
    xlActionTypeReport = 128 
    xlActionTypeRowset = 16 
    xlActionTypeUrl = 1

class XlAllocation(IntEnum):
    xlAutomaticAllocation = 2
    xlManualAllocation = 1

class XlAllocationMethod(IntEnum):
    xlEqualAllocation = 1
    xlWeightedAllocation = 2

class XlAllocationValue(IntEnum):
    xlAllocateIncrement = 2
    xlAllocateValue = 1

class XlApplicationInternational(IntEnum):
    xl24HourClock = 33 
    xl4DigitYears = 43 
    xlAlternateArraySeparator = 16 
    xlColumnSeparator = 14
    xlCountryCode = 1
    xlCountrySetting = 2
    xlCurrencyBefore = 37 
    xlCurrencyCode = 25 
    xlCurrencyDigits = 27 
    xlCurrencyLeadingZeros = 40 
    xlCurrencyMinusSign = 38 
    xlCurrencyNegative = 28 
    xlCurrencySpaceBefore = 36 
    xlCurrencyTrailingZeros = 39 
    xlDateOrder = 32 
    xlDateSeparator = 17 
    xlDayCode = 21 
    xlDayLeadingZero = 42 
    xlDecimalSeparator = 3
    xlGeneralFormatName = 26 
    xlHourCode = 22 
    xlLeftBrace = 12
    xlLeftBracket = 10
    xlListSeparator = 5
    xlLowerCaseColumnLetter = 9
    xlLowerCaseRowLetter = 8
    xlMDY = 44 
    xlMetric = 35 
    xlMinuteCode = 23 
    xlMonthCode = 20 
    xlMonthLeadingZero = 41 
    xlMonthNameChars = 30 
    xlNoncurrencyDigits = 29 
    xlNonEnglishFunctions = 34 
    xlRightBrace = 13
    xlRightBracket = 11
    xlRowSeparator = 15
    xlSecondCode = 24 
    xlThousandsSeparator = 4
    xlTimeLeadingZero = 45 
    xlTimeSeparator = 18 
    xlUICultureTag = 46 
    xlUpperCaseColumnLetter = 7
    xlUpperCaseRowLetter = 6
    xlWeekdayNameChars = 31 
    xlYearCode = 19 

class XlApplyNamesOrder(IntEnum):
    xlColumnThenRow = 2
    xlRowThenColumn = 1

class XlArabicModes(IntEnum):
    xlArabicBothStrict = 3
    xlArabicNone = 0
    xlArabicStrictAlefHamza = 1
    xlArabicStrictFinalYaa = 2

class XlArrangeStyle(IntEnum):
    xlArrangeStyleCascade = 7
    xlArrangeStyleHorizontal = -4128 
    xlArrangeStyleTiled = 1
    xlArrangeStyleVertical = -4166 

class XlArrowHeadLength(IntEnum):
    xlArrowHeadLengthLong = 3
    xlArrowHeadLengthMedium = -4138 
    xlArrowHeadLengthShort = 1

class XlArrowHeadStyle(IntEnum):
    xlArrowHeadStyleClosed = 3
    xlArrowHeadStyleDoubleClosed = 5
    xlArrowHeadStyleDoubleOpen = 4
    xlArrowHeadStyleNone = -4142 
    xlArrowHeadStyleOpen = 2

class XlArrowHeadWidth(IntEnum):
    xlArrowHeadWidthMedium = -4138 
    xlArrowHeadWidthNarrow = 1
    xlArrowHeadWidthWide = 3

class XlAutoFillType(IntEnum):
    xlFillCopy = 1
    xlFillDays = 5
    xlFillDefault = 0
    xlFillFormats = 3
    xlFillMonths = 7
    xlFillSeries = 2
    xlFillValues = 4
    xlFillWeekdays = 6
    xlFillYears = 8
    xlFlashFill = 11
    xlGrowthTrend = 10
    xlLinearTrend = 9

class XlAutoFilterOperator(IntEnum):
    xlAnd = 1
    xlBottom10Items = 4
    xlBottom10Percent = 6
    xlFilterAutomaticFontColor = 13
    xlFilterCellColor = 8
    xlFilterDynamic = 11
    xlFilterFontColor = 9
    xlFilterIcon = 10
    xlFilterNoFill = 12
    xlFilterNoIcon = 14
    xlFilterValues = 7
    xlOr = 2
    xlTop10Items = 3
    xlTop10Percent = 5

class XlAxisCrosses(IntEnum):
    xlAxisCrossesAutomatic = -4105 
    xlAxisCrossesCustom = -4114 
    xlAxisCrossesMaximum = 2
    xlAxisCrossesMinimum = 4

class XlAxisGroup(IntEnum):
    xlPrimary = 1
    xlSecondary = 2

class XlAxisType(IntEnum):
    xlCategory = 1
    xlSeriesAxis = 3
    xlValue = 2

class XlBackground(IntEnum):
    xlBackgroundAutomatic = -4105 
    xlBackgroundOpaque = 3
    xlBackgroundTransparent = 2

class XlBarShape(IntEnum):
    xlBox = 0
    xlConeToMax = 5
    xlConeToPoint = 4
    xlCylinder = 3
    xlPyramidToMax = 2
    xlPyramidToPoint = 1

class XlBinsType(IntEnum):
    xlBinsTypeAutomatic = 0
    xlBinsTypeBinCount = 4
    xlBinsTypeBinSize = 3
    xlBinsTypeCategorical = 1
    xlBinsTypeManual = 2

class XlBordersIndex(IntEnum):
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeBottom = 9
    xlEdgeLeft = 7
    xlEdgeRight = 10
    xlEdgeTop = 8
    xlInsideHorizontal = 12
    xlInsideVertical = 11

class XlBorderWeight(IntEnum):
    xlHairline = 1
    xlMedium = -4138 
    xlThick = 4
    xlThin = 2

class XlBuiltInDialog(IntEnum):
    xlDialogActivate = 103 
    xlDialogActiveCellFont = 476 
    xlDialogAddChartAutoformat = 390 
    xlDialogAddinManager = 321 
    xlDialogAlignment = 43 
    xlDialogApplyNames = 133 
    xlDialogApplyStyle = 212 
    xlDialogAppMove = 170 
    xlDialogAppSize = 171 
    xlDialogArrangeAll = 12
    xlDialogAssignToObject = 213 
    xlDialogAssignToTool = 293 
    xlDialogAttachText = 80 
    xlDialogAttachToolbars = 323 
    xlDialogAutoCorrect = 485 
    xlDialogAxes = 78 
    xlDialogBorder = 45 
    xlDialogCalculation = 32 
    xlDialogCellProtection = 46 
    xlDialogChangeLink = 166 
    xlDialogChartAddData = 392 
    xlDialogChartLocation = 527 
    xlDialogChartOptionsDataLabelMultiple = 724 
    xlDialogChartOptionsDataLabels = 505 
    xlDialogChartOptionsDataTable = 506 
    xlDialogChartSourceData = 540 
    xlDialogChartTrend = 350 
    xlDialogChartType = 526 
    xlDialogChartWizard = 288 
    xlDialogCheckboxProperties = 435 
    xlDialogClear = 52 
    xlDialogColorPalette = 161 
    xlDialogColumnWidth = 47 
    xlDialogCombination = 73 
    xlDialogConditionalFormatting = 583 
    xlDialogConsolidate = 191 
    xlDialogCopyChart = 147 
    xlDialogCopyPicture = 108 
    xlDialogCreateList = 796 
    xlDialogCreateNames = 62 
    xlDialogCreatePublisher = 217 
    xlDialogCreateRelationship = 1272 
    xlDialogCustomizeToolbar = 276 
    xlDialogCustomViews = 493 
    xlDialogDataDelete = 36 
    xlDialogDataLabel = 379 
    xlDialogDataLabelMultiple = 723 
    xlDialogDataSeries = 40 
    xlDialogDataValidation = 525 
    xlDialogDefineName = 61 
    xlDialogDefineStyle = 229 
    xlDialogDeleteFormat = 111 
    xlDialogDeleteName = 110 
    xlDialogDemote = 203 
    xlDialogDisplay = 27 
    xlDialogDocumentInspector = 862 
    xlDialogEditboxProperties = 438 
    xlDialogEditColor = 223 
    xlDialogEditDelete = 54 
    xlDialogEditionOptions = 251 
    xlDialogEditSeries = 228 
    xlDialogErrorbarX = 463 
    xlDialogErrorbarY = 464 
    xlDialogErrorChecking = 732 
    xlDialogEvaluateFormula = 709 
    xlDialogExternalDataProperties = 530 
    xlDialogExtract = 35 
    xlDialogFileDelete = 6
    xlDialogFileSharing = 481 
    xlDialogFillGroup = 200 
    xlDialogFillWorkgroup = 301 
    xlDialogFilter = 447 
    xlDialogFilterAdvanced = 370 
    xlDialogFindFile = 475 
    xlDialogFont = 26 
    xlDialogFontProperties = 381 
    xlDialogForecastETS = 1300 
    xlDialogFormatAuto = 269 
    xlDialogFormatChart = 465 
    xlDialogFormatCharttype = 423 
    xlDialogFormatFont = 150 
    xlDialogFormatLegend = 88 
    xlDialogFormatMain = 225 
    xlDialogFormatMove = 128 
    xlDialogFormatNumber = 42 
    xlDialogFormatOverlay = 226 
    xlDialogFormatSize = 129 
    xlDialogFormatText = 89 
    xlDialogFormulaFind = 64 
    xlDialogFormulaGoto = 63 
    xlDialogFormulaReplace = 130 
    xlDialogFunctionWizard = 450 
    xlDialogGallery3dArea = 193 
    xlDialogGallery3dBar = 272 
    xlDialogGallery3dColumn = 194 
    xlDialogGallery3dLine = 195 
    xlDialogGallery3dPie = 196 
    xlDialogGallery3dSurface = 273 
    xlDialogGalleryArea = 67 
    xlDialogGalleryBar = 68 
    xlDialogGalleryColumn = 69 
    xlDialogGalleryCustom = 388 
    xlDialogGalleryDoughnut = 344 
    xlDialogGalleryLine = 70 
    xlDialogGalleryPie = 71 
    xlDialogGalleryRadar = 249 
    xlDialogGalleryScatter = 72 
    xlDialogGoalSeek = 198 
    xlDialogGridlines = 76 
    xlDialogImportTextFile = 666 
    xlDialogInsert = 55 
    xlDialogInsertHyperlink = 596 
    xlDialogInsertNameLabel = 496 
    xlDialogInsertObject = 259 
    xlDialogInsertPicture = 342 
    xlDialogInsertTitle = 380 
    xlDialogLabelProperties = 436 
    xlDialogListboxProperties = 437 
    xlDialogMacroOptions = 382 
    xlDialogMailEditMailer = 470 
    xlDialogMailLogon = 339 
    xlDialogMailNextLetter = 378 
    xlDialogMainChart = 85 
    xlDialogMainChartType = 185 
    xlDialogManageRelationships = 1271 
    xlDialogMenuEditor = 322 
    xlDialogMove = 262 
    xlDialogMyPermission = 834 
    xlDialogNameManager = 977 
    xlDialogNew = 119 
    xlDialogNewName = 978 
    xlDialogNewWebQuery = 667 
    xlDialogNote = 154 
    xlDialogObjectProperties = 207 
    xlDialogObjectProtection = 214 
    xlDialogOpen = 1
    xlDialogOpenLinks = 2
    xlDialogOpenMail = 188 
    xlDialogOpenText = 441 
    xlDialogOptionsCalculation = 318 
    xlDialogOptionsChart = 325 
    xlDialogOptionsEdit = 319 
    xlDialogOptionsGeneral = 356 
    xlDialogOptionsListsAdd = 458 
    xlDialogOptionsME = 647 
    xlDialogOptionsTransition = 355 
    xlDialogOptionsView = 320 
    xlDialogOutline = 142 
    xlDialogOverlay = 86 
    xlDialogOverlayChartType = 186 
    xlDialogPageSetup = 7
    xlDialogParse = 91 
    xlDialogPasteNames = 58 
    xlDialogPasteSpecial = 53 
    xlDialogPatterns = 84 
    xlDialogPermission = 832 
    xlDialogPhonetic = 656 
    xlDialogPivotCalculatedField = 570 
    xlDialogPivotCalculatedItem = 572 
    xlDialogPivotClientServerSet = 689 
    xlDialogPivotDefaultLayout = 1360 
    xlDialogPivotFieldGroup = 433 
    xlDialogPivotFieldProperties = 313 
    xlDialogPivotFieldUngroup = 434 
    xlDialogPivotShowPages = 421 
    xlDialogPivotSolveOrder = 568 
    xlDialogPivotTableOptions = 567 
    xlDialogPivotTableSlicerConnections = 1183 
    xlDialogPivotTableWhatIfAnalysisSettings = 1153 
    xlDialogPivotTableWizard = 312 
    xlDialogPlacement = 300 
    xlDialogPrint = 8
    xlDialogPrinterSetup = 9
    xlDialogPrintPreview = 222 
    xlDialogPromote = 202 
    xlDialogProperties = 474 
    xlDialogPropertyFields = 754 
    xlDialogProtectDocument = 28 
    xlDialogProtectSharing = 620 
    xlDialogPublishAsWebPage = 653 
    xlDialogPushbuttonProperties = 445 
    xlDialogRecommendedPivotTables = 1258 
    xlDialogReplaceFont = 134 
    xlDialogRoutingSlip = 336 
    xlDialogRowHeight = 127 
    xlDialogRun = 17 
    xlDialogSaveAs = 5
    xlDialogSaveCopyAs = 456 
    xlDialogSaveNewObject = 208 
    xlDialogSaveWorkbook = 145 
    xlDialogSaveWorkspace = 285 
    xlDialogScale = 87 
    xlDialogScenarioAdd = 307 
    xlDialogScenarioCells = 305 
    xlDialogScenarioEdit = 308 
    xlDialogScenarioMerge = 473 
    xlDialogScenarioSummary = 311 
    xlDialogScrollbarProperties = 420 
    xlDialogSearch = 731 
    xlDialogSelectSpecial = 132 
    xlDialogSendMail = 189 
    xlDialogSeriesAxes = 460 
    xlDialogSeriesOptions = 557 
    xlDialogSeriesOrder = 466 
    xlDialogSeriesShape = 504 
    xlDialogSeriesX = 461 
    xlDialogSeriesY = 462 
    xlDialogSetBackgroundPicture = 509 
    xlDialogSetManager = 1109 
    xlDialogSetMDXEditor = 1208 
    xlDialogSetPrintTitles = 23 
    xlDialogSetTupleEditorOnColumns = 1108 
    xlDialogSetTupleEditorOnRows = 1107 
    xlDialogSetUpdateStatus = 159 
    xlDialogShowDetail = 204 
    xlDialogShowToolbar = 220 
    xlDialogSize = 261 
    xlDialogSlicerCreation = 1182 
    xlDialogSlicerPivotTableConnections = 1184 
    xlDialogSlicerSettings = 1179 
    xlDialogSort = 39 
    xlDialogSortSpecial = 192 
    xlDialogSparklineInsertColumn = 1134 
    xlDialogSparklineInsertLine = 1133 
    xlDialogSparklineInsertWinLoss = 1135 
    xlDialogSplit = 137 
    xlDialogStandardFont = 190 
    xlDialogStandardWidth = 472 
    xlDialogStyle = 44 
    xlDialogSubscribeTo = 218 
    xlDialogSubtotalCreate = 398 
    xlDialogSummaryInfo = 474 
    xlDialogTable = 41 
    xlDialogTabOrder = 394 
    xlDialogTextToColumns = 422 
    xlDialogUnhide = 94 
    xlDialogUpdateLink = 201 
    xlDialogVbaInsertFile = 328 
    xlDialogVbaMakeAddin = 478 
    xlDialogVbaProcedureDefinition = 330 
    xlDialogView3d = 197 
    xlDialogWebOptionsBrowsers = 773 
    xlDialogWebOptionsEncoding = 686 
    xlDialogWebOptionsFiles = 684 
    xlDialogWebOptionsFonts = 687 
    xlDialogWebOptionsGeneral = 683 
    xlDialogWebOptionsPictures = 685 
    xlDialogWindowMove = 14
    xlDialogWindowSize = 13
    xlDialogWorkbookAdd = 281 
    xlDialogWorkbookCopy = 283 
    xlDialogWorkbookInsert = 354 
    xlDialogWorkbookMove = 282 
    xlDialogWorkbookName = 386 
    xlDialogWorkbookNew = 302 
    xlDialogWorkbookOptions = 284 
    xlDialogWorkbookProtect = 417 
    xlDialogWorkbookTabSplit = 415 
    xlDialogWorkbookUnhide = 384 
    xlDialogWorkgroup = 199 
    xlDialogWorkspace = 95 
    xlDialogZoom = 256 

class XlCalcFor(IntEnum):
    xlAllValues = 0
    xlColGroups = 2
    xlRowGroups = 1

class XlCalcMemNumberFormatType(IntEnum):
    xlNumberFormatTypeDefault = 0
    xlNumberFormatTypeNumber = 1
    xlNumberFormatTypePercent = 2

class XlCalculatedMemberType(IntEnum):
    xlCalculatedMeasure = 2
    xlCalculatedMember = 0
    xlCalculatedSet = 1

class XlCalculation(IntEnum):
    xlCalculationAutomatic = -4105 
    xlCalculationManual = -4135 
    xlCalculationSemiautomatic = 2

class XlCalculationInterruptKey(IntEnum):
    xlAnyKey = 2
    xlEscKey = 1
    xlNoKey = 0

class XlCalculationState(IntEnum):
    xlCalculating = 1
    xlDone = 0
    xlPending = 2

class XlCategoryLabelLevel(IntEnum):
    xlCategoryLabelLevelAll = -1 
    xlCategoryLabelLevelCustom = -2 
    xlCategoryLabelLevelNone = -3 

class XlCategorySortOrder(IntEnum):
    xlCategoryAscending = 2
    xlCategoryDescending = 3
    xlIndexAscending = 0
    xlIndexDescending = 1

class XlCategoryType(IntEnum):
    xlAutomaticScale = -4105 
    xlCategoryScale = 2
    xlTimeScale = 3

class XlCellChangedState(IntEnum):
    xlCellChangeApplied = 3
    xlCellChanged = 2
    xlCellNotChanged = 1

class XlCellInsertionMode(IntEnum):
    xlInsertDeleteCells = 1
    xlInsertEntireRows = 2
    xlOverwriteCells = 0

class XlCellType(IntEnum):
    xlCellTypeAllFormatConditions = -4172 
    xlCellTypeAllValidation = -4174 
    xlCellTypeBlanks = 4
    xlCellTypeComments = -4144 
    xlCellTypeConstants = 2
    xlCellTypeFormulas = -4123 
    xlCellTypeLastCell = 11
    xlCellTypeSameFormatConditions = -4173 
    xlCellTypeSameValidation = -4175 
    xlCellTypeVisible = 12

class XlChartElementPosition(IntEnum):
    xlChartElementPositionAutomatic = -4105 
    xlChartElementPositionCustom = -4114 

class XlChartGallery(IntEnum):
    xlAnyGallery = 23 
    xlBuiltIn = 21 
    xlUserDefined = 22 

class XlChartItem(IntEnum):
    xlAxis = 21 
    xlAxisTitle = 17 
    xlChartArea = 2
    xlChartTitle = 4
    xlCorners = 6
    xlDataLabel = 0
    xlDataTable = 7
    xlDisplayUnitLabel = 30 
    xlDownBars = 20 
    xlDropLines = 26 
    xlErrorBars = 9
    xlFloor = 23 
    xlHiLoLines = 25 
    xlLeaderLines = 29 
    xlLegend = 24 
    xlLegendEntry = 12
    xlLegendKey = 13
    xlMajorGridlines = 15
    xlMinorGridlines = 16 
    xlNothing = 28 
    xlPivotChartCollapseEntireFieldButton = 34 
    xlPivotChartDropZone = 32 
    xlPivotChartExpandEntireFieldButton = 33 
    xlPivotChartFieldButton = 31 
    xlPlotArea = 19 
    xlRadarAxisLabels = 27 
    xlSeries = 3
    xlSeriesLines = 22 
    xlShape = 14
    xlTrendline = 8
    xlUpBars = 18 
    xlWalls = 5
    xlXErrorBars = 10
    xlYErrorBars = 11

class XlChartLocation(IntEnum):
    xlLocationAsNewSheet = 1
    xlLocationAsObject = 2
    xlLocationAutomatic = 3

class XlChartPicturePlacement(IntEnum):
    xlAllFaces = 7
    xlEnd = 2
    xlEndSides = 3
    xlFront = 4
    xlFrontEnd = 6
    xlFrontSides = 5
    xlSides = 1

class XlChartPictureType(IntEnum):
    xlStack = 2
    xlStackScale = 3
    xlStretch = 1

class XlChartSplitType(IntEnum):
    xlSplitByCustomSplit = 4
    xlSplitByPercentValue = 3
    xlSplitByPosition = 1
    xlSplitByValue = 2

class XlChartType(IntEnum):
    xl3DArea = -4098 
    xl3DAreaStacked = 78 
    xl3DAreaStacked100 = 79 
    xl3DBarClustered = 60 
    xl3DBarStacked = 61 
    xl3DBarStacked100 = 62 
    xl3DColumn = -4100 
    xl3DColumnClustered = 54 
    xl3DColumnStacked = 55 
    xl3DColumnStacked100 = 56 
    xl3DLine = -4101 
    xl3DPie = -4102 
    xl3DPieExploded = 70 
    xlArea = 1
    xlAreaStacked = 76 
    xlAreaStacked100 = 77 
    xlBarClustered = 57 
    xlBarOfPie = 71 
    xlBarStacked = 58 
    xlBarStacked100 = 59 
    xlBoxwhisker = 121 
    xlBubble = 15
    xlBubble3DEffect = 87 
    xlColumnClustered = 51 
    xlColumnStacked = 52 
    xlColumnStacked100 = 53 
    xlConeBarClustered = 102 
    xlConeBarStacked = 103 
    xlConeBarStacked100 = 104 
    xlConeCol = 105 
    xlConeColClustered = 99 
    xlConeColStacked = 100 
    xlConeColStacked100 = 101 
    xlCylinderBarClustered = 95 
    xlCylinderBarStacked = 96 
    xlCylinderBarStacked100 = 97 
    xlCylinderCol = 98 
    xlCylinderColClustered = 92 
    xlCylinderColStacked = 93 
    xlCylinderColStacked100 = 94 
    xlDoughnut = -4120 
    xlDoughnutExploded = 80 
    xlFunnel = 123 
    xlHistogram = 118 
    xlLine = 4
    xlLineMarkers = 65 
    xlLineMarkersStacked = 66 
    xlLineMarkersStacked100 = 67 
    xlLineStacked = 63 
    xlLineStacked100 = 64 
    xlPareto = 122 
    xlPie = 5
    xlPieExploded = 69 
    xlPieOfPie = 68 
    xlPyramidBarClustered = 109 
    xlPyramidBarStacked = 110 
    xlPyramidBarStacked100 = 111 
    xlPyramidCol = 112 
    xlPyramidColClustered = 106 
    xlPyramidColStacked = 107 
    xlPyramidColStacked100 = 108 
    xlRadar = -4151 
    xlRadarFilled = 82 
    xlRadarMarkers = 81 
    xlRegionMap = 140 
    xlStockHLC = 88 
    xlStockOHLC = 89 
    xlStockVHLC = 90 
    xlStockVOHLC = 91 
    xlSunburst = 120 
    xlSurface = 83 
    xlSurfaceTopView = 85 
    xlSurfaceTopViewWireframe = 86 
    xlSurfaceWireframe = 84 
    xlTreemap = 117 
    xlWaterfall = 119 
    xlXYScatter = -4169 
    xlXYScatterLines = 74 
    xlXYScatterLinesNoMarkers = 75 
    xlXYScatterSmooth = 72 
    xlXYScatterSmoothNoMarkers = 73 

class XlCheckInVersionType(IntEnum):
    xlCheckInMajorVersion = 1
    xlCheckInMinorVersion = 0
    xlCheckInOverwriteVersion = 2

class XlClipboardFormat(IntEnum):
    xlClipboardFormatBIFF = 8
    xlClipboardFormatBIFF12 = 63 
    xlClipboardFormatBIFF2 = 18 
    xlClipboardFormatBIFF3 = 20 
    xlClipboardFormatBIFF4 = 30 
    xlClipboardFormatBinary = 15
    xlClipboardFormatBitmap = 9
    xlClipboardFormatCGM = 13
    xlClipboardFormatCSV = 5
    xlClipboardFormatDIF = 4
    xlClipboardFormatDspText = 12
    xlClipboardFormatEmbeddedObject = 21 
    xlClipboardFormatEmbedSource = 22 
    xlClipboardFormatLink = 11
    xlClipboardFormatLinkSource = 23 
    xlClipboardFormatLinkSourceDesc = 32 
    xlClipboardFormatMovie = 24 
    xlClipboardFormatNative = 14
    xlClipboardFormatObjectDesc = 31 
    xlClipboardFormatObjectLink = 19 
    xlClipboardFormatOwnerLink = 17 
    xlClipboardFormatPICT = 2
    xlClipboardFormatPrintPICT = 3
    xlClipboardFormatRTF = 7
    xlClipboardFormatScreenPICT = 29 
    xlClipboardFormatStandardFont = 28 
    xlClipboardFormatStandardScale = 27 
    xlClipboardFormatSYLK = 6
    xlClipboardFormatTable = 16 
    xlClipboardFormatText = 0
    xlClipboardFormatToolFace = 25 
    xlClipboardFormatToolFacePICT = 26 
    xlClipboardFormatVALU = 1
    xlClipboardFormatWK1 = 10

class XlCmdType(IntEnum):
    xlCmdCube = 1
    xlCmdDAX = 8
    xlCmdDefault = 4
    xlCmdExcel = 7
    xlCmdList = 5
    xlCmdSql = 2
    xlCmdTable = 3
    xlCmdTableCollection = 6

class XlColorIndex(IntEnum):
    xlColorIndexAutomatic = -4105 
    xlColorIndexNone = -4142 

class XlColumnDataType(IntEnum):
    xlDMYFormat = 4
    xlDYMFormat = 7
    xlEMDFormat = 10
    xlGeneralFormat = 1
    xlMDYFormat = 3
    xlMYDFormat = 6
    xlSkipColumn = 9
    xlTextFormat = 2
    xlYDMFormat = 8
    xlYMDFormat = 5

class XlCommandUnderlines(IntEnum):
    xlCommandUnderlinesAutomatic = -4105 
    xlCommandUnderlinesOff = -4146 
    xlCommandUnderlinesOn = 1

class XlCommentDisplayMode(IntEnum):
    xlCommentAndIndicator = 1
    xlCommentIndicatorOnly = -1 
    xlNoIndicator = 0

class XlConditionValueTypes(IntEnum):
    xlConditionValueAutomaticMax = 7
    xlConditionValueAutomaticMin = 6
    xlConditionValueFormula = 4
    xlConditionValueHighestValue = 2
    xlConditionValueLowestValue = 1
    xlConditionValueNone = -1 
    xlConditionValueNumber = 0
    xlConditionValuePercent = 3
    xlConditionValuePercentile = 5

class XlConnectionType(IntEnum):
    xlConnectionTypeDATAFEED = 6
    xlConnectionTypeMODEL = 7
    xlConnectionTypeNOSOURCE = 9
    xlConnectionTypeODBC = 2
    xlConnectionTypeOLEDB = 1
    xlConnectionTypeTEXT = 4
    xlConnectionTypeWEB = 5
    xlConnectionTypeWORKSHEET = 8
    xlConnectionTypeXMLMAP = 3

class XlConsolidationFunction(IntEnum):
    xlAverage = -4106 
    xlCount = -4112 
    xlCountNums = -4113 
    xlDistinctCount = 11
    xlMax = -4136 
    xlMin = -4139 
    xlProduct = -4149 
    xlStDev = -4155 
    xlStDevP = -4156 
    xlSum = -4157 
    xlUnknown = 1000 
    xlVar = -4164 
    xlVarP = -4165 

class XlContainsOperator(IntEnum):
    xlBeginsWith = 2
    xlContains = 0
    xlDoesNotContain = 1
    xlEndsWith = 3

class XlCopyPictureFormat(IntEnum):
    xlBitmap = 2
    xlPicture = -4147 

class XlCorruptLoad(IntEnum):
    xlExtractData = 2
    xlNormalLoad = 0
    xlRepairFile = 1

class XlCreator(IntEnum):
    xlCreatorCode = 1480803660 

class XlCredentialsMethod(IntEnum):
    xlCredentialsMethodIntegrated = 0
    xlCredentialsMethodNone = 1
    xlCredentialsMethodStored = 2

class XlCubeFieldSubType(IntEnum):
    xlCubeAttribute = 4
    xlCubeCalculatedMeasure = 5
    xlCubeHierarchy = 1
    xlCubeImplicitMeasure = 11
    xlCubeKPIGoal = 7
    xlCubeKPIStatus = 8
    xlCubeKPITrend = 9
    xlCubeKPIValue = 6
    xlCubeKPIWeight = 10
    xlCubeMeasure = 2
    xlCubeSet = 3

class XlCubeFieldType(IntEnum):
    xlHierarchy = 1
    xlMeasure = 2
    xlSet = 3

class XlCutCopyMode(IntEnum):
    xlCopy = 1
    xlCut = 2

class XlCVError(IntEnum):
    xlErrBlocked = 2047 
    xlErrCalc = 2050 
    xlErrConnect = 2046 
    xlErrDiv0 = 2007 
    xlErrField = 2049 
    xlErrGettingData = 2043 
    xlErrNA = 2042 
    xlErrName = 2029 
    xlErrNull = 2000 
    xlErrNum = 2036 
    xlErrRef = 2023 
    xlErrSpill = 2045 
    xlErrUnknown = 2048 
    xlErrValue = 2015 

class XlDataBarAxisPosition(IntEnum):
    xlDataBarAxisAutomatic = 0
    xlDataBarAxisMidpoint = 1
    xlDataBarAxisNone = 2

class XlDataBarBorderType(IntEnum):
    xlDataBarBorderNone = 0
    xlDataBarBorderSolid = 1

class XlDataBarFillType(IntEnum):
    xlDataBarFillGradient = 1
    xlDataBarFillSolid = 0

class XlDataBarNegativeColorType(IntEnum):
    xlDataBarColor = 0
    xlDataBarSameAsPositive = 1

class XlDataLabelPosition(IntEnum):
    xlLabelPositionAbove = 0
    xlLabelPositionBelow = 1
    xlLabelPositionBestFit = 5
    xlLabelPositionCenter = -4108 
    xlLabelPositionCustom = 7
    xlLabelPositionInsideBase = 4
    xlLabelPositionInsideEnd = 3
    xlLabelPositionLeft = -4131 
    xlLabelPositionMixed = 6
    xlLabelPositionOutsideEnd = 2
    xlLabelPositionRight = -4152 

class XlDataLabelSeparator(IntEnum):
    xlDataLabelSeparatorDefault = 1

class XlDataLabelsType(IntEnum):
    xlDataLabelsShowBubbleSizes = 6
    xlDataLabelsShowLabel = 4
    xlDataLabelsShowLabelAndPercent = 5
    xlDataLabelsShowNone = -4142 
    xlDataLabelsShowPercent = 3
    xlDataLabelsShowValue = 2

class XlDataSeriesDate(IntEnum):
    xlDay = 1
    xlMonth = 3
    xlWeekday = 2
    xlYear = 4

class XlDataSeriesType(IntEnum):
    xlAutoFill = 4
    xlChronological = 3
    xlDataSeriesLinear = -4132 
    xlGrowth = 2

class XlDeleteShiftDirection(IntEnum):
    xlShiftToLeft = -4159 
    xlShiftUp = -4162 

class XlDirection(IntEnum):
    xlDown = -4121 
    xlToLeft = -4159 
    xlToRight = -4161 
    xlUp = -4162 

class XlDisplayBlanksAs(IntEnum):
    xlInterpolated = 3
    xlNotPlotted = 1
    xlZero = 2

class XlDisplayDrawingObjects(IntEnum):
    xlDisplayShapes = -4104 
    xlHide = 3
    xlPlaceholders = 2

class XlDisplayUnit(IntEnum):
    xlHundredMillions = -8 
    xlHundreds = -2 
    xlHundredThousands = -5 
    xlMillionMillions = -10 
    xlMillions = -6 
    xlTenMillions = -7 
    xlTenThousands = -4 
    xlThousandMillions = -9 
    xlThousands = -3 

class XlDupeUnique(IntEnum):
    xlDuplicate = 1
    xlUnique = 0

class XlDVAlertStyle(IntEnum):
    xlValidAlertInformation = 3
    xlValidAlertStop = 1
    xlValidAlertWarning = 2

class XlDVType(IntEnum):
    xlValidateCustom = 7
    xlValidateDate = 4
    xlValidateDecimal = 2
    xlValidateInputOnly = 0
    xlValidateList = 3
    xlValidateTextLength = 6
    xlValidateTime = 5
    xlValidateWholeNumber = 1

class XlDynamicFilterCriteria(IntEnum):
    xlFilterAboveAverage = 33 
    xlFilterAllDatesInPeriodApril = 24 
    xlFilterAllDatesInPeriodAugust = 28 
    xlFilterAllDatesInPeriodDecember = 32 
    xlFilterAllDatesInPeriodFebruray = 22 
    xlFilterAllDatesInPeriodJanuary = 21 
    xlFilterAllDatesInPeriodJuly = 27 
    xlFilterAllDatesInPeriodJune = 26 
    xlFilterAllDatesInPeriodMarch = 23 
    xlFilterAllDatesInPeriodMay = 25 
    xlFilterAllDatesInPeriodNovember = 31 
    xlFilterAllDatesInPeriodOctober = 30 
    xlFilterAllDatesInPeriodQuarter1 = 17 
    xlFilterAllDatesInPeriodQuarter2 = 18 
    xlFilterAllDatesInPeriodQuarter3 = 19 
    xlFilterAllDatesInPeriodQuarter4 = 20 
    xlFilterAllDatesInPeriodSeptember = 29 
    xlFilterBelowAverage = 34 
    xlFilterLastMonth = 8
    xlFilterLastQuarter = 11
    xlFilterLastWeek = 5
    xlFilterLastYear = 14
    xlFilterNextMonth = 9
    xlFilterNextQuarter = 12
    xlFilterNextWeek = 6
    xlFilterNextYear = 15
    xlFilterThisMonth = 7
    xlFilterThisQuarter = 10
    xlFilterThisWeek = 4
    xlFilterThisYear = 13
    xlFilterToday = 1
    xlFilterTomorrow = 3
    xlFilterYearToDate = 16 
    xlFilterYesterday = 2

class XlEditionFormat(IntEnum):
    xlBIFF = 2
    xlPICT = 1
    xlRTF = 4
    xlVALU = 8

class XlEditionOptionsOption(IntEnum):
    xlAutomaticUpdate = 4
    xlCancel = 1
    xlChangeAttributes = 6
    xlManualUpdate = 5
    xlOpenSource = 3
    xlSelect = 3
    xlSendPublisher = 2
    xlUpdateSubscriber = 2

class XlEditionType(IntEnum):
    xlPublisher = 1
    xlSubscriber = 2

class XlEnableCancelKey(IntEnum):
    xlDisabled = 0
    xlErrorHandler = 2
    xlInterrupt = 1

class XlEnableSelection(IntEnum):
    xlNoRestrictions = 0
    xlNoSelection = -4142 
    xlUnlockedCells = 1

class XlEndStyleCap(IntEnum):
    xlCap = 1
    xlNoCap = 2

class XlErrorBarDirection(IntEnum):
    xlX = -4168 
    xlY = 1

class XlErrorBarInclude(IntEnum):
    xlErrorBarIncludeBoth = 1
    xlErrorBarIncludeMinusValues = 3
    xlErrorBarIncludeNone = -4142 
    xlErrorBarIncludePlusValues = 2

class XlErrorBarType(IntEnum):
    xlErrorBarTypeCustom = -4114 
    xlErrorBarTypeFixedValue = 1
    xlErrorBarTypePercent = 2
    xlErrorBarTypeStDev = -4155 
    xlErrorBarTypeStError = 4

class XlErrorChecks(IntEnum):
    xlEmptyCellReferences = 7
    xlEvaluateToError = 1
    xlInconsistentFormula = 4
    xlInconsistentListFormula = 9
    xlListDataValidation = 8
    xlMisleadingFormat = 10
    xlNumberAsText = 3
    xlOmittedCells = 5
    xlOutdatedLinkedDataType = 11
    xlTextDate = 2
    xlUnlockedFormulaCells = 6

class XlFileAccess(IntEnum):
    xlReadOnly = 3
    xlReadWrite = 2

class XlFileFormat(IntEnum):
    xlAddIn = 18 
    xlAddIn8 = 18 
    xlCSV = 6
    xlCSVMac = 22 
    xlCSVMSDOS = 24 
    xlCSVUTF8 = 62 
    xlCSVWindows = 23 
    xlCurrentPlatformText = -4158 
    xlDBF2 = 7
    xlDBF3 = 8
    xlDBF4 = 11
    xlDIF = 9
    xlExcel12 = 50 
    xlExcel2 = 16 
    xlExcel2FarEast = 27 
    xlExcel3 = 29 
    xlExcel4 = 33 
    xlExcel4Workbook = 35 
    xlExcel5 = 39 
    xlExcel7 = 39 
    xlExcel8 = 56 
    xlExcel9795 = 43 
    xlHtml = 44 
    xlIntlAddIn = 26 
    xlIntlMacro = 25 
    xlOpenDocumentSpreadsheet = 60 
    xlOpenXMLAddIn = 55 
    xlOpenXMLStrictWorkbook = 61 
    xlOpenXMLTemplate = 54 
    xlOpenXMLTemplateMacroEnabled = 53 
    xlOpenXMLWorkbook = 51 
    xlOpenXMLWorkbookMacroEnabled = 52 
    xlSYLK = 2
    xlTemplate = 17 
    xlTemplate8 = 17 
    xlTextMac = 19 
    xlTextMSDOS = 21 
    xlTextPrinter = 36 
    xlTextWindows = 20 
    xlUnicodeText = 42 
    xlWebArchive = 45 
    xlWJ2WD1 = 14
    xlWJ3 = 40 
    xlWJ3FJ3 = 41 
    xlWK1 = 5
    xlWK1ALL = 31 
    xlWK1FMT = 30 
    xlWK3 = 15
    xlWK3FM3 = 32 
    xlWK4 = 38 
    xlWKS = 4
    xlWorkbookDefault = 51 
    xlWorkbookNormal = -4143 
    xlWorks2FarEast = 28 
    xlWQ1 = 34 
    xlXMLSpreadsheet = 46 

class XlFileValidationPivotMode(IntEnum):
    xlFileValidationPivotDefault = 0
    xlFileValidationPivotRun = 1
    xlFileValidationPivotSkip = 2

class XlFillWith(IntEnum):
    xlFillWithAll = -4104 
    xlFillWithContents = 2
    xlFillWithFormats = -4122 

class XlFilterAction(IntEnum):
    xlFilterCopy = 2
    xlFilterInPlace = 1

class XlFilterAllDatesInPeriod(IntEnum):
    xlFilterAllDatesInPeriodDay = 2
    xlFilterAllDatesInPeriodHour = 3
    xlFilterAllDatesInPeriodMinute = 4
    xlFilterAllDatesInPeriodMonth = 1
    xlFilterAllDatesInPeriodSecond = 5
    xlFilterAllDatesInPeriodYear = 0

class XlFilterStatus(IntEnum):
    xlFilterStatusDateHasTime = 2
    xlFilterStatusDateWrongOrder = 1
    xlFilterStatusInvalidDate = 3
    xlFilterStatusOK = 0

class XlFindLookIn(IntEnum):
    xlComments = -4144 
    xlCommentsThreaded = -4184 
    xlFormulas = -4123 
    xlFormulas2 = -4185 
    xlValues = -4163 

class XlFixedFormatQuality(IntEnum):
    xlQualityMinimum = 1
    xlQualityStandard = 0

class XlFixedFormatType(IntEnum):
    xlTypePDF = 0
    xlTypeXPS = 1

class XlForecastAggregation(IntEnum):
    xlForecastAggregationAverage = 1
    xlForecastAggregationCount = 2
    xlForecastAggregationCountA = 3
    xlForecastAggregationMax = 4
    xlForecastAggregationMedian = 5
    xlForecastAggregationMin = 6
    xlForecastAggregationSum = 7

class XlForecastChartType(IntEnum):
    xlForecastChartTypeColumn = 1
    xlForecastChartTypeLine = 0

class XlForecastDataCompletion(IntEnum):
    xlForecastDataCompletionInterpolate = 1
    xlForecastDataCompletionZeros = 0

class XlFormatConditionOperator(IntEnum):
    xlBetween = 1
    xlEqual = 3
    xlGreater = 5
    xlGreaterEqual = 7
    xlLess = 6
    xlLessEqual = 8
    xlNotBetween = 2
    xlNotEqual = 4

class XlFormatConditionType(IntEnum):
    xlAboveAverageCondition = 12
    xlBlanksCondition = 10
    xlCellValue = 1
    xlColorScale = 3
    xlDatabar = 4
    xlErrorsCondition = 16 
    xlExpression = 2
    xlIconSets = 6
    xlNoBlanksCondition = 13
    xlNoErrorsCondition = 17 
    xlTextString = 9
    xlTimePeriod = 11
    xlTop10 = 5
    xlUniqueValues = 8

class XlFormatFilterTypes(IntEnum):
    xlFilterBottom = 0
    xlFilterBottomPercent = 2
    xlFilterTop = 1
    xlFilterTopPercent = 3

class XlFormControl(IntEnum):
    xlButtonControl = 0
    xlCheckBox = 1
    xlDropDown = 2
    xlEditBox = 3
    xlGroupBox = 4
    xlLabel = 5
    xlListBox = 6
    xlOptionButton = 7
    xlScrollBar = 8
    xlSpinner = 9

class XlFormulaLabel(IntEnum):
    xlColumnLabels = 2
    xlMixedLabels = 3
    xlNoLabels = -4142 
    xlRowLabels = 1

class XlFormulaVersion(IntEnum):
    xlReplaceFormula = 0
    xlReplaceFormula2 = 1

class XlGenerateTableRefs(IntEnum):
    xlGenerateTableRefA1 = 0
    xlGenerateTableRefStruct = 1

class XlGeoMappingLevel(IntEnum):
    xlGeoMappingLevelAutomatic = 0
    xlGeoMappingLevelCountryRegion = 5
    xlGeoMappingLevelCountryRegionList = 6
    xlGeoMappingLevelCounty = 3
    xlGeoMappingLevelDataOnly = 1
    xlGeoMappingLevelPostalCode = 2
    xlGeoMappingLevelState = 4
    xlGeoMappingLevelWorld = 7

class XlGeoProjectionType(IntEnum):
    xlGeoProjectionTypeAlbers = 3
    xlGeoProjectionTypeAutomatic = 0
    xlGeoProjectionTypeMercator = 1
    xlGeoProjectionTypeMiller = 2
    xlGeoProjectionTypeRobinson = 4

class XlGradientFillType(IntEnum):
    xlGradientFillLinear = 0
    xlGradientFillPath = 1

class XlGradientStopPositionType(IntEnum):
    xlGradientStopPositionTypeExtremeValue = 0
    xlGradientStopPositionTypeNumber = 1
    xlGradientStopPositionTypePercent = 2

class XlHAlign(IntEnum):
    xlHAlignCenter = -4108 
    xlHAlignCenterAcrossSelection = 7
    xlHAlignDistributed = -4117 
    xlHAlignFill = 5
    xlHAlignGeneral = 1
    xlHAlignJustify = -4130 
    xlHAlignLeft = -4131 
    xlHAlignRight = -4152 

class XlHebrewModes(IntEnum):
    xlHebrewFullScript = 0
    xlHebrewMixedAuthorizedScript = 3
    xlHebrewMixedScript = 2
    xlHebrewPartialScript = 1

class XlHighlightChangesTime(IntEnum):
    xlAllChanges = 2
    xlNotYetReviewed = 3
    xlSinceMyLastSave = 1

class XlHtmlType(IntEnum):
    xlHtmlCalc = 1
    xlHtmlChart = 3
    xlHtmlList = 2
    xlHtmlStatic = 0

class XlIcon(IntEnum):
    xlIcon0Bars = 37 
    xlIcon0FilledBoxes = 52 
    xlIcon1Bar = 38 
    xlIcon1FilledBox = 51 
    xlIcon2Bars = 39 
    xlIcon2FilledBoxes = 50 
    xlIcon3Bars = 40 
    xlIcon3FilledBoxes = 49 
    xlIcon4Bars = 41 
    xlIcon4FilledBoxes = 48 
    xlIconBlackCircle = 32 
    xlIconBlackCircleWithBorder = 13
    xlIconCircleWithOneWhiteQuarter = 33 
    xlIconCircleWithThreeWhiteQuarters = 35 
    xlIconCircleWithTwoWhiteQuarters = 34 
    xlIconGoldStar = 42 
    xlIconGrayCircle = 31 
    xlIconGrayDownArrow = 6
    xlIconGrayDownInclineArrow = 28 
    xlIconGraySideArrow = 5
    xlIconGrayUpArrow = 4
    xlIconGrayUpInclineArrow = 27 
    xlIconGreenCheck = 22 
    xlIconGreenCheckSymbol = 19 
    xlIconGreenCircle = 10
    xlIconGreenFlag = 7
    xlIconGreenTrafficLight = 14
    xlIconGreenUpArrow = 1
    xlIconGreenUpTriangle = 45 
    xlIconHalfGoldStar = 43 
    xlIconNoCellIcon = -1 
    xlIconPinkCircle = 30 
    xlIconRedCircle = 29 
    xlIconRedCircleWithBorder = 12
    xlIconRedCross = 24 
    xlIconRedCrossSymbol = 21 
    xlIconRedDiamond = 18 
    xlIconRedDownArrow = 3
    xlIconRedDownTriangle = 47 
    xlIconRedFlag = 9
    xlIconRedTrafficLight = 16 
    xlIconSilverStar = 44 
    xlIconWhiteCircleAllWhiteQuarters = 36 
    xlIconYellowCircle = 11
    xlIconYellowDash = 46 
    xlIconYellowDownInclineArrow = 26 
    xlIconYellowExclamation = 23 
    xlIconYellowExclamationSymbol = 20 
    xlIconYellowFlag = 8
    xlIconYellowSideArrow = 2
    xlIconYellowTrafficLight = 15
    xlIconYellowTriangle = 17 
    xlIconYellowUpInclineArrow = 25 

class XlIconSet(IntEnum):
    xl3Arrows = 1
    xl3ArrowsGray = 2
    xl3Flags = 3
    xl3Signs = 6
    xl3Stars = 18 
    xl3Symbols = 7
    xl3Symbols2 = 8
    xl3TrafficLights1 = 4
    xl3TrafficLights2 = 5
    xl3Triangles = 19 
    xl4Arrows = 9
    xl4ArrowsGray = 10
    xl4CRV = 12
    xl4RedToBlack = 11
    xl4TrafficLights = 13
    xl5Arrows = 14
    xl5ArrowsGray = 15
    xl5Boxes = 20 
    xl5CRV = 16 
    xl5Quarters = 17 
    xlCustomSet = -1 

class XlIMEMode(IntEnum):
    xlIMEModeAlpha = 8
    xlIMEModeAlphaFull = 7
    xlIMEModeDisable = 3
    xlIMEModeHangul = 10
    xlIMEModeHangulFull = 9
    xlIMEModeHiragana = 4
    xlIMEModeKatakana = 5
    xlIMEModeKatakanaHalf = 6
    xlIMEModeNoControl = 0
    xlIMEModeOff = 2
    xlIMEModeOn = 1

class XlImportDataAs(IntEnum):
    xlPivotTableReport = 1
    xlQueryTable = 0
    xlTable = 2

class XlInsertFormatOrigin(IntEnum):
    xlFormatFromLeftOrAbove = 0
    xlFormatFromRightOrBelow = 1

class XlInsertShiftDirection(IntEnum):
    xlShiftDown = -4121 
    xlShiftToRight = -4161 

class XlLayoutFormType(IntEnum):
    xlOutline = 1
    xlTabular = 0

class XlLayoutRowType(IntEnum):
    xlCompactRow = 0
    xlOutlineRow = 2
    xlTabularRow = 1

class XlLegendPosition(IntEnum):
    xlLegendPositionBottom = -4107 
    xlLegendPositionCorner = 2
    xlLegendPositionCustom = -4161 
    xlLegendPositionLeft = -4131 
    xlLegendPositionRight = -4152 
    xlLegendPositionTop = -4160 

class XlLineStyle(IntEnum):
    xlContinuous = 1
    xlDash = -4115 
    xlDashDot = 4
    xlDashDotDot = 5
    xlDot = -4118 
    xlDouble = -4119 
    xlLineStyleNone = -4142 
    xlSlantDashDot = 13

class XlLink(IntEnum):
    xlExcelLinks = 1
    xlOLELinks = 2
    xlPublishers = 5
    xlSubscribers = 6

class XlLinkedDataTypeState(IntEnum):
    xlLinkedDataTypeStateBrokenLinkedData = 3
    xlLinkedDataTypeStateDisambiguationNeeded = 2
    xlLinkedDataTypeStateFetchingData = 4
    xlLinkedDataTypeStateNone = 0
    xlLinkedDataTypeStateValidLinkedData = 1

class XlLinkInfo(IntEnum):
    xlEditionDate = 2
    xlLinkInfoStatus = 3
    xlUpdateState = 1

class XlLinkInfoType(IntEnum):
    xlLinkInfoOLELinks = 2
    xlLinkInfoPublishers = 5
    xlLinkInfoSubscribers = 6

class XlLinkStatus(IntEnum):
    xlLinkStatusCopiedValues = 10
    xlLinkStatusIndeterminate = 5
    xlLinkStatusInvalidName = 7
    xlLinkStatusMissingFile = 1
    xlLinkStatusMissingSheet = 2
    xlLinkStatusNotStarted = 6
    xlLinkStatusOK = 0
    xlLinkStatusOld = 3
    xlLinkStatusSourceNotCalculated = 4
    xlLinkStatusSourceNotOpen = 8
    xlLinkStatusSourceOpen = 9

class XlLinkType(IntEnum):
    xlLinkTypeExcelLinks = 1
    xlLinkTypeOLELinks = 2

class XlListConflict(IntEnum):
    xlListConflictDialog = 0
    xlListConflictDiscardAllConflicts = 2
    xlListConflictError = 3
    xlListConflictRetryAllConflicts = 1

class XlListDataType(IntEnum):
    xlListDataTypeCheckbox = 9
    xlListDataTypeChoice = 6
    xlListDataTypeChoiceMulti = 7
    xlListDataTypeCounter = 11
    xlListDataTypeCurrency = 4
    xlListDataTypeDateTime = 5
    xlListDataTypeHyperLink = 10
    xlListDataTypeListLookup = 8
    xlListDataTypeMultiLineRichText = 12
    xlListDataTypeMultiLineText = 2
    xlListDataTypeNone = 0
    xlListDataTypeNumber = 3
    xlListDataTypeText = 1

class XlListObjectSourceType(IntEnum):
    xlSrcExternal = 0
    xlSrcModel = 4
    xlSrcQuery = 3
    xlSrcRange = 1
    xlSrcXml = 2

class XlLocationInTable(IntEnum):
    xlColumnHeader = -4110 
    xlColumnItem = 5
    xlDataHeader = 3
    xlDataItem = 7
    xlPageHeader = 2
    xlPageItem = 6
    xlRowHeader = -4153 
    xlRowItem = 4
    xlTableBody = 8

class XlLookAt(IntEnum):
    xlPart = 2
    xlWhole = 1

class XlLookFor(IntEnum):
    xlLookForBlanks = 0
    xlLookForErrors = 1
    xlLookForFormulas = 2

class XlMailSystem(IntEnum):
    xlMAPI = 1
    xlNoMailSystem = 0
    xlPowerTalk = 2

class XlMarkerStyle(IntEnum):
    xlMarkerStyleAutomatic = -4105 
    xlMarkerStyleCircle = 8
    xlMarkerStyleDash = -4115 
    xlMarkerStyleDiamond = 2
    xlMarkerStyleDot = -4118 
    xlMarkerStyleNone = -4142 
    xlMarkerStylePicture = -4147 
    xlMarkerStylePlus = 9
    xlMarkerStyleSquare = 1
    xlMarkerStyleStar = 5
    xlMarkerStyleTriangle = 3
    xlMarkerStyleX = -4168 

class XlMeasurementUnits(IntEnum):
    xlCentimeters = 1
    xlInches = 0
    xlMillimeters = 2

class XlModelChangeSource(IntEnum):
    xlChangeByExcel = 0
    xlChangeByPowerPivotAddIn = 1

class XlMouseButton(IntEnum):
    xlNoButton = 0
    xlPrimaryButton = 1
    xlSecondaryButton = 2

class XlMousePointer(IntEnum):
    xlDefault = -4143 
    xlIBeam = 3
    xlNorthwestArrow = 1
    xlWait = 2

class XlMSApplication(IntEnum):
    xlMicrosoftAccess = 4
    xlMicrosoftFoxPro = 5
    xlMicrosoftMail = 3
    xlMicrosoftPowerPoint = 2
    xlMicrosoftProject = 6
    xlMicrosoftSchedulePlus = 7
    xlMicrosoftWord = 1

class XlOartHorizontalOverflow(IntEnum):
    xlOartHorizontalOverflowClip = 1
    xlOartHorizontalOverflowOverflow = 0

class XlOartVerticalOverflow(IntEnum):
    xlOartVerticalOverflowClip = 1
    xlOartVerticalOverflowEllipsis = 2
    xlOartVerticalOverflowOverflow = 0

class XlObjectSize(IntEnum):
    xlFitToPage = 2
    xlFullPage = 3
    xlScreenSize = 1

class XlOLEType(IntEnum):
    xlOLEControl = 2
    xlOLEEmbed = 1
    xlOLELink = 0

class XlOLEVerb(IntEnum):
    xlVerbOpen = 2
    xlVerbPrimary = 1

class XlOrder(IntEnum):
    xlDownThenOver = 1
    xlOverThenDown = 2

class XlOrientation(IntEnum):
    xlDownward = -4170 
    xlHorizontal = -4128 
    xlUpward = -4171 
    xlVertical = -4166 

class XlPageBreak(IntEnum):
    xlPageBreakAutomatic = -4105 
    xlPageBreakManual = -4135 
    xlPageBreakNone = -4142 

class XlPageBreakExtent(IntEnum):
    xlPageBreakFull = 1
    xlPageBreakPartial = 2

class XlPageOrientation(IntEnum):
    xlLandscape = 2
    xlPortrait = 1

class XlPaperSize(IntEnum):
    xlPaper10x14 = 16 
    xlPaper11x17 = 17 
    xlPaperA3 = 8
    xlPaperA4 = 9
    xlPaperA4Small = 10
    xlPaperA5 = 11
    xlPaperB4 = 12
    xlPaperB5 = 13
    xlPaperCsheet = 24 
    xlPaperDsheet = 25 
    xlPaperEnvelope10 = 20 
    xlPaperEnvelope11 = 21 
    xlPaperEnvelope12 = 22 
    xlPaperEnvelope14 = 23 
    xlPaperEnvelope9 = 19 
    xlPaperEnvelopeB4 = 33 
    xlPaperEnvelopeB5 = 34 
    xlPaperEnvelopeB6 = 35 
    xlPaperEnvelopeC3 = 29 
    xlPaperEnvelopeC4 = 30 
    xlPaperEnvelopeC5 = 28 
    xlPaperEnvelopeC6 = 31 
    xlPaperEnvelopeC65 = 32 
    xlPaperEnvelopeDL = 27 
    xlPaperEnvelopeItaly = 36 
    xlPaperEnvelopeMonarch = 37 
    xlPaperEnvelopePersonal = 38 
    xlPaperEsheet = 26 
    xlPaperExecutive = 7
    xlPaperFanfoldLegalGerman = 41 
    xlPaperFanfoldStdGerman = 40 
    xlPaperFanfoldUS = 39 
    xlPaperFolio = 14
    xlPaperLedger = 4
    xlPaperLegal = 5
    xlPaperLetter = 1
    xlPaperLetterSmall = 2
    xlPaperNote = 18 
    xlPaperQuarto = 15
    xlPaperStatement = 6
    xlPaperTabloid = 3
    xlPaperUser = 256 

class XlParameterDataType(IntEnum):
    xlParamTypeBigInt = -5 
    xlParamTypeBinary = -2 
    xlParamTypeBit = -7 
    xlParamTypeChar = 1
    xlParamTypeDate = 9
    xlParamTypeDecimal = 3
    xlParamTypeDouble = 8
    xlParamTypeFloat = 6
    xlParamTypeInteger = 4
    xlParamTypeLongVarBinary = -4 
    xlParamTypeLongVarChar = -1 
    xlParamTypeNumeric = 2
    xlParamTypeReal = 7
    xlParamTypeSmallInt = 5
    xlParamTypeTime = 10
    xlParamTypeTimestamp = 11
    xlParamTypeTinyInt = -6 
    xlParamTypeUnknown = 0
    xlParamTypeVarBinary = -3 
    xlParamTypeVarChar = 12
    xlParamTypeWChar = -8 

class XlParameterType(IntEnum):
    xlConstant = 1
    xlPrompt = 0
    xlRange = 2

class XlParentDataLabelOptions(IntEnum):
    xlParentDataLabelOptionsBanner = 1
    xlParentDataLabelOptionsNone = 0
    xlParentDataLabelOptionsOverlapping = 2

class XlPasteSpecialOperation(IntEnum):
    xlPasteSpecialOperationAdd = 2
    xlPasteSpecialOperationDivide = 5
    xlPasteSpecialOperationMultiply = 4
    xlPasteSpecialOperationNone = -4142 
    xlPasteSpecialOperationSubtract = 3

class XlPasteType(IntEnum):
    xlPasteAll = -4104 
    xlPasteAllExceptBorders = 7
    xlPasteAllMergingConditionalFormats = 14
    xlPasteAllUsingSourceTheme = 13
    xlPasteColumnWidths = 8
    xlPasteComments = -4144 
    xlPasteFormats = -4122 
    xlPasteFormulas = -4123 
    xlPasteFormulasAndNumberFormats = 11
    xlPasteValidation = 6
    xlPasteValues = -4163 
    xlPasteValuesAndNumberFormats = 12

class XlPattern(IntEnum):
    xlPatternAutomatic = -4105 
    xlPatternChecker = 9
    xlPatternCrissCross = 16 
    xlPatternDown = -4121 
    xlPatternGray16 = 17 
    xlPatternGray25 = -4124 
    xlPatternGray50 = -4125 
    xlPatternGray75 = -4126 
    xlPatternGray8 = 18 
    xlPatternGrid = 15
    xlPatternHorizontal = -4128 
    xlPatternLightDown = 13
    xlPatternLightHorizontal = 11
    xlPatternLightUp = 14
    xlPatternLightVertical = 12
    xlPatternLinearGradient = 4000 
    xlPatternNone = -4142 
    xlPatternRectangularGradient = 4001 
    xlPatternSemiGray75 = 10
    xlPatternSolid = 1
    xlPatternUp = -4162 
    xlPatternVertical = -4166 

class XlPhoneticAlignment(IntEnum):
    xlPhoneticAlignCenter = 2
    xlPhoneticAlignDistributed = 3
    xlPhoneticAlignLeft = 1
    xlPhoneticAlignNoControl = 0

class XlPhoneticCharacterType(IntEnum):
    xlHiragana = 2
    xlKatakana = 1
    xlKatakanaHalf = 0
    xlNoConversion = 3

class XlPictureAppearance(IntEnum):
    xlPrinter = 2
    xlScreen = 1

class XlPictureConvertorType(IntEnum):
    xlBMP = 1
    xlCGM = 7
    xlDRW = 4
    xlDXF = 5
    xlEPS = 8
    xlHGL = 6
    xlPCT = 13
    xlPCX = 10
    xlPIC = 11
    xlPLT = 12
    xlTIF = 9
    xlWMF = 2
    xlWPG = 3

class XlPieSliceIndex(IntEnum):
    xlCenterPoint = 5
    xlInnerCenterPoint = 8
    xlInnerClockwisePoint = 7
    xlInnerCounterClockwisePoint = 9
    xlMidClockwiseRadiusPoint = 4
    xlMidCounterClockwiseRadiusPoint = 6
    xlOuterCenterPoint = 2
    xlOuterClockwisePoint = 3
    xlOuterCounterClockwisePoint = 1

class XlPieSliceLocation(IntEnum):
    xlHorizontalCoordinate = 1
    xlVerticalCoordinate = 2

class XlPivotCellType(IntEnum):
    xlPivotCellBlankCell = 9
    xlPivotCellCustomSubtotal = 7
    xlPivotCellDataField = 4
    xlPivotCellDataPivotField = 8
    xlPivotCellGrandTotal = 3
    xlPivotCellPageFieldItem = 6
    xlPivotCellPivotField = 5
    xlPivotCellPivotItem = 1
    xlPivotCellSubtotal = 2
    xlPivotCellValue = 0

class XlPivotConditionScope(IntEnum):
    xlDataFieldScope = 2
    xlFieldsScope = 1
    xlSelectionScope = 0

class XlPivotFieldCalculation(IntEnum):
    xlDifferenceFrom = 2
    xlIndex = 9
    xlNoAdditionalCalculation = -4143 
    xlPercentDifferenceFrom = 4
    xlPercentOf = 3
    xlPercentOfColumn = 7
    xlPercentOfParent = 12
    xlPercentOfParentColumn = 11
    xlPercentOfParentRow = 10
    xlPercentOfRow = 6
    xlPercentOfTotal = 8
    xlPercentRunningTotal = 13
    xlRankAscending = 14
    xlRankDecending = 15
    xlRunningTotal = 5

class XlPivotFieldDataType(IntEnum):
    xlDate = 2
    xlNumber = -4145 
    xlText = -4158 

class XlPivotFieldOrientation(IntEnum):
    xlColumnField = 2
    xlDataField = 4
    xlHidden = 0
    xlPageField = 3
    xlRowField = 1

class XlPivotFieldRepeatLabels(IntEnum):
    xlDoNotRepeatLabels = 1
    xlRepeatLabels = 2

class XlPivotFilterType(IntEnum):
    xlAfter = 33 
    xlAfterOrEqualTo = 34 
    xlAllDatesInPeriodApril = 60 
    xlAllDatesInPeriodAugust = 64 
    xlAllDatesInPeriodDecember = 68 
    xlAllDatesInPeriodFebruary = 58 
    xlAllDatesInPeriodJanuary = 57 
    xlAllDatesInPeriodJuly = 63 
    xlAllDatesInPeriodJune = 62 
    xlAllDatesInPeriodMarch = 59 
    xlAllDatesInPeriodMay = 61 
    xlAllDatesInPeriodNovember = 67 
    xlAllDatesInPeriodOctober = 66 
    xlAllDatesInPeriodQuarter1 = 53 
    xlAllDatesInPeriodQuarter2 = 54 
    xlAllDatesInPeriodQuarter3 = 55 
    xlAllDatesInPeriodQuarter4 = 56 
    xlAllDatesInPeriodSeptember = 65 
    xlBefore = 31 
    xlBeforeOrEqualTo = 32 
    xlBottomCount = 2
    xlBottomPercent = 4
    xlBottomSum = 6
    xlCaptionBeginsWith = 17 
    xlCaptionContains = 21 
    xlCaptionDoesNotBeginWith = 18 
    xlCaptionDoesNotContain = 22 
    xlCaptionDoesNotEndWith = 20 
    xlCaptionDoesNotEqual = 16 
    xlCaptionEndsWith = 19 
    xlCaptionEquals = 15
    xlCaptionIsBetween = 27 
    xlCaptionIsGreaterThan = 23 
    xlCaptionIsGreaterThanOrEqualTo = 24 
    xlCaptionIsLessThan = 25 
    xlCaptionIsLessThanOrEqualTo = 26 
    xlCaptionIsNotBetween = 28 
    xlDateBetween = 35 
    xlDateLastMonth = 45 
    xlDateLastQuarter = 48 
    xlDateLastWeek = 42 
    xlDateLastYear = 51 
    xlDateNextMonth = 43 
    xlDateNextQuarter = 46 
    xlDateNextWeek = 40 
    xlDateNextYear = 49 
    xlDateNotBetween = 36 
    xlDateThisMonth = 44 
    xlDateThisQuarter = 47 
    xlDateThisWeek = 41 
    xlDateThisYear = 50 
    xlDateToday = 38 
    xlDateTomorrow = 37 
    xlDateYesterday = 39 
    xlNotSpecificDate = 30 
    xlSpecificDate = 29 
    xlTopCount = 1
    xlTopPercent = 3
    xlTopSum = 5
    xlValueDoesNotEqual = 8
    xlValueEquals = 7
    xlValueIsBetween = 13
    xlValueIsGreaterThan = 9
    xlValueIsGreaterThanOrEqualTo = 10
    xlValueIsLessThan = 11
    xlValueIsLessThanOrEqualTo = 12
    xlValueIsNotBetween = 14
    xlYearToDate = 52 

class XlPivotFormatType(IntEnum):
    xlPTClassic = 20 
    xlPTNone = 21 
    xlReport1 = 0
    xlReport10 = 9
    xlReport2 = 1
    xlReport3 = 2
    xlReport4 = 3
    xlReport5 = 4
    xlReport6 = 5
    xlReport7 = 6
    xlReport8 = 7
    xlReport9 = 8
    xlTable1 = 10
    xlTable10 = 19 
    xlTable2 = 11
    xlTable3 = 12
    xlTable4 = 13
    xlTable5 = 14
    xlTable6 = 15
    xlTable7 = 16 
    xlTable8 = 17 
    xlTable9 = 18 

class XlPivotLineType(IntEnum):
    xlPivotLineBlank = 3
    xlPivotLineGrandTotal = 2
    xlPivotLineRegular = 0
    xlPivotLineSubtotal = 1

class XlPivotTableMissingItems(IntEnum):
    xlMissingItemsDefault = -1 
    xlMissingItemsMax = 32500 
    xlMissingItemsMax2 = 1048576 
    xlMissingItemsNone = 0

class XlPivotTableSourceType(IntEnum):
    xlConsolidation = 3
    xlDatabase = 1
    xlExternal = 2
    xlPivotTable = -4148 
    xlScenario = 4

class XlPivotTableVersionList(IntEnum):
    xlPivotTableVersion10 = 1
    xlPivotTableVersion11 = 2
    xlPivotTableVersion12 = 3
    xlPivotTableVersion14 = 4
    xlPivotTableVersion15 = 5
    xlPivotTableVersion2000 = 0
    xlPivotTableVersionCurrent = -1 

class XlPlacement(IntEnum):
    xlFreeFloating = 3
    xlMove = 2
    xlMoveAndSize = 1

class XlPlatform(IntEnum):
    xlMacintosh = 1
    xlMSDOS = 3
    xlWindows = 2

class XlPortugueseReform(IntEnum):
    xlPortugueseBoth = 3
    xlPortuguesePostReform = 2
    xlPortuguesePreReform = 1

class XlPrintErrors(IntEnum):
    xlPrintErrorsBlank = 1
    xlPrintErrorsDash = 2
    xlPrintErrorsDisplayed = 0
    xlPrintErrorsNA = 3

class XlPrintLocation(IntEnum):
    xlPrintInPlace = 16 
    xlPrintNoComments = -4142 
    xlPrintSheetEnd = 1

class XlPriority(IntEnum):
    xlPriorityHigh = -4127 
    xlPriorityLow = -4134 
    xlPriorityNormal = -4143 

class XlPropertyDisplayedIn(IntEnum):
    xlDisplayPropertyInPivotTable = 1
    xlDisplayPropertyInPivotTableAndTooltip = 3
    xlDisplayPropertyInTooltip = 2

class XlProtectedViewCloseReason(IntEnum):
    xlProtectedViewCloseEdit = 1
    xlProtectedViewCloseForced = 2
    xlProtectedViewCloseNormal = 0

class XlProtectedViewWindowState(IntEnum):
    xlProtectedViewWindowMaximized = 2
    xlProtectedViewWindowMinimized = 1
    xlProtectedViewWindowNormal = 0

class XlPTSelectionMode(IntEnum):
    xlBlanks = 4
    xlButton = 15
    xlDataAndLabel = 0
    xlDataOnly = 2
    xlFirstRow = 256 
    xlLabelOnly = 1
    xlOrigin = 3

class XlPublishToDocsDisclosureScope(IntEnum):
    msoLimited = 1
    msoNoOverwrite = 3
    msoOrganization = 2
    msoPublic = 0

class XlPublishToPBINameConflictAction(IntEnum):
    msoPBIAbort = 1
    msoPBIIgnore = 0
    msoPBIOverwrite = 2

class XlPublishToPBIPublishType(IntEnum):
    msoPBIExport = 0
    msoPBIUpload = 1

class XlQueryType(IntEnum):
    xlADORecordset = 7
    xlDAORecordset = 2
    xlODBCQuery = 1
    xlOLEDBQuery = 5
    xlTextImport = 6
    xlWebQuery = 4

class XlQuickAnalysisMode(IntEnum):
    xlFormatConditions = 1
    xlLensOnly = 0
    xlRecommendedCharts = 2
    xlSparklines = 5
    xlTables = 4
    xlTotals = 3

class XlRangeAutoFormat(IntEnum):
    xlRangeAutoFormat3DEffects1 = 13
    xlRangeAutoFormat3DEffects2 = 14
    xlRangeAutoFormatAccounting1 = 4
    xlRangeAutoFormatAccounting2 = 5
    xlRangeAutoFormatAccounting3 = 6
    xlRangeAutoFormatAccounting4 = 17 
    xlRangeAutoFormatClassic1 = 1
    xlRangeAutoFormatClassic2 = 2
    xlRangeAutoFormatClassic3 = 3
    xlRangeAutoFormatClassicPivotTable = 31 
    xlRangeAutoFormatColor1 = 7
    xlRangeAutoFormatColor2 = 8
    xlRangeAutoFormatColor3 = 9
    xlRangeAutoFormatList1 = 10
    xlRangeAutoFormatList2 = 11
    xlRangeAutoFormatList3 = 12
    xlRangeAutoFormatLocalFormat1 = 15
    xlRangeAutoFormatLocalFormat2 = 16 
    xlRangeAutoFormatLocalFormat3 = 19 
    xlRangeAutoFormatLocalFormat4 = 20 
    xlRangeAutoFormatNone = -4142 
    xlRangeAutoFormatPTNone = 42 
    xlRangeAutoFormatReport1 = 21 
    xlRangeAutoFormatReport10 = 30 
    xlRangeAutoFormatReport2 = 22 
    xlRangeAutoFormatReport3 = 23 
    xlRangeAutoFormatReport4 = 24 
    xlRangeAutoFormatReport5 = 25 
    xlRangeAutoFormatReport6 = 26 
    xlRangeAutoFormatReport7 = 27 
    xlRangeAutoFormatReport8 = 28 
    xlRangeAutoFormatReport9 = 29 
    xlRangeAutoFormatSimple = -4154 
    xlRangeAutoFormatTable1 = 32 
    xlRangeAutoFormatTable10 = 41 
    xlRangeAutoFormatTable2 = 33 
    xlRangeAutoFormatTable3 = 34 
    xlRangeAutoFormatTable4 = 35 
    xlRangeAutoFormatTable5 = 36 
    xlRangeAutoFormatTable6 = 37 
    xlRangeAutoFormatTable7 = 38 
    xlRangeAutoFormatTable8 = 39 
    xlRangeAutoFormatTable9 = 40 

class XlRangeValueDataType(IntEnum):
    xlRangeValueDefault = 10
    xlRangeValueMSPersistXML = 12
    xlRangeValueXMLSpreadsheet = 11

class XlReferenceStyle(IntEnum):
    xlA1 = 1
    xlR1C1 = -4150 

class XlReferenceType(IntEnum):
    xlAbsolute = 1
    xlAbsRowRelColumn = 2
    xlRelative = 4
    xlRelRowAbsColumn = 3

class XlRegionLabelOptions(IntEnum):
    xlRegionLabelOptionsBestFitOnly = 1
    xlRegionLabelOptionsNone = 0
    xlRegionLabelOptionsShowAll = 2

class XlRemoveDocInfoType(IntEnum):
    xlRDIAll = 99 
    xlRDIComments = 1
    xlRDIContentType = 16 
    xlRDIDefinedNameComments = 18 
    xlRDIDocumentManagementPolicy = 15
    xlRDIDocumentProperties = 8
    xlRDIDocumentServerProperties = 14
    xlRDIDocumentWorkspace = 10
    xlRDIEmailHeader = 5
    xlRDIExcelDataModel = 23 
    xlRDIInactiveDataConnections = 19 
    xlRDIInkAnnotations = 11
    xlRDIInlineWebExtensions = 21 
    xlRDIPrinterPath = 20 
    xlRDIPublishInfo = 13
    xlRDIRemovePersonalInformation = 4
    xlRDIRoutingSlip = 6
    xlRDIScenarioComments = 12
    xlRDISendForReview = 7
    xlRDITaskpaneWebExtensions = 22 

class XlRgbColor(IntEnum):
    rgbAliceBlue = 16775408 
    rgbAntiqueWhite = 14150650 
    rgbAqua = 16776960 
    rgbAquamarine = 13959039 
    rgbAzure = 16777200 
    rgbBeige = 14480885 
    rgbBisque = 12903679 
    rgbBlack = 0
    rgbBlanchedAlmond = 13495295 
    rgbBlue = 16711680 
    rgbBlueViolet = 14822282 
    rgbBrown = 2763429 
    rgbBurlyWood = 8894686 
    rgbCadetBlue = 10526303 
    rgbChartreuse = 65407 
    rgbCoral = 5275647 
    rgbCornflowerBlue = 15570276 
    rgbCornsilk = 14481663 
    rgbCrimson = 3937500 
    rgbDarkBlue = 9109504 
    rgbDarkCyan = 9145088 
    rgbDarkGoldenrod = 755384 
    rgbDarkGray = 11119017 
    rgbDarkGreen = 25600 
    rgbDarkGrey = 11119017 
    rgbDarkKhaki = 7059389 
    rgbDarkMagenta = 9109643 
    rgbDarkOliveGreen = 3107669 
    rgbDarkOrange = 36095 
    rgbDarkOrchid = 13382297 
    rgbDarkRed = 139 
    rgbDarkSalmon = 8034025 
    rgbDarkSeaGreen = 9419919 
    rgbDarkSlateBlue = 9125192 
    rgbDarkSlateGray = 5197615 
    rgbDarkSlateGrey = 5197615 
    rgbDarkTurquoise = 13749760 
    rgbDarkViolet = 13828244 
    rgbDeepPink = 9639167 
    rgbDeepSkyBlue = 16760576 
    rgbDimGray = 6908265 
    rgbDimGrey = 6908265 
    rgbDodgerBlue = 16748574 
    rgbFireBrick = 2237106 
    rgbFloralWhite = 15792895 
    rgbForestGreen = 2263842 
    rgbFuchsia = 16711935 
    rgbGainsboro = 14474460 
    rgbGhostWhite = 16775416 
    rgbGold = 55295 
    rgbGoldenrod = 2139610 
    rgbGray = 8421504 
    rgbGreen = 32768 
    rgbGreenYellow = 3145645 
    rgbGrey = 8421504 
    rgbHoneydew = 15794160 
    rgbHotPink = 11823615 
    rgbIndianRed = 6053069 
    rgbIndigo = 8519755 
    rgbIvory = 15794175 
    rgbKhaki = 9234160 
    rgbLavender = 16443110 
    rgbLavenderBlush = 16118015 
    rgbLawnGreen = 64636 
    rgbLemonChiffon = 13499135 
    rgbLightBlue = 15128749 
    rgbLightCoral = 8421616 
    rgbLightCyan = 9145088 
    rgbLightGoldenrodYellow = 13826810 
    rgbLightGray = 13882323 
    rgbLightGreen = 9498256 
    rgbLightGrey = 13882323 
    rgbLightPink = 12695295 
    rgbLightSalmon = 8036607 
    rgbLightSeaGreen = 11186720 
    rgbLightSkyBlue = 16436871 
    rgbLightSlateGray = 10061943 
    rgbLightSlateGrey = 10061943 
    rgbLightSteelBlue = 14599344 
    rgbLightYellow = 14745599 
    rgbLime = 65280 
    rgbLimeGreen = 3329330 
    rgbLinen = 15134970 
    rgbMaroon = 128 
    rgbMediumAquamarine = 11206502 
    rgbMediumBlue = 13434880 
    rgbMediumOrchid = 13850042 
    rgbMediumPurple = 14381203 
    rgbMediumSeaGreen = 7451452 
    rgbMediumSlateBlue = 15624315 
    rgbMediumSpringGreen = 10156544 
    rgbMediumTurquoise = 13422920 
    rgbMediumVioletRed = 8721863 
    rgbMidnightBlue = 7346457 
    rgbMintCream = 16449525 
    rgbMistyRose = 14804223 
    rgbMoccasin = 11920639 
    rgbNavajoWhite = 11394815 
    rgbNavy = 8388608 
    rgbNavyBlue = 8388608 
    rgbOldLace = 15136253 
    rgbOlive = 32896 
    rgbOliveDrab = 2330219 
    rgbOrange = 42495 
    rgbOrangeRed = 17919 
    rgbOrchid = 14053594 
    rgbPaleGoldenrod = 7071982 
    rgbPaleGreen = 10025880 
    rgbPaleTurquoise = 15658671 
    rgbPaleVioletRed = 9662683 
    rgbPapayaWhip = 14020607 
    rgbPeachPuff = 12180223 
    rgbPeru = 4163021 
    rgbPink = 13353215 
    rgbPlum = 14524637 
    rgbPowderBlue = 15130800 
    rgbPurple = 8388736 
    rgbRed = 255 
    rgbRosyBrown = 9408444 
    rgbRoyalBlue = 14772545 
    rgbSalmon = 7504122 
    rgbSandyBrown = 6333684 
    rgbSeaGreen = 5737262 
    rgbSeashell = 15660543 
    rgbSienna = 2970272 
    rgbSilver = 12632256 
    rgbSkyBlue = 15453831 
    rgbSlateBlue = 13458026 
    rgbSlateGray = 9470064 
    rgbSlateGrey = 9470064 
    rgbSnow = 16448255 
    rgbSpringGreen = 8388352 
    rgbSteelBlue = 11829830 
    rgbTan = 9221330 
    rgbTeal = 8421376 
    rgbThistle = 14204888 
    rgbTomato = 4678655 
    rgbTurquoise = 13688896 
    rgbViolet = 15631086 
    rgbWheat = 11788021 
    rgbWhite = 16777215 
    rgbWhiteSmoke = 16119285 
    rgbYellow = 65535 
    rgbYellowGreen = 3329434 

class XlRobustConnect(IntEnum):
    xlAlways = 1
    xlAsRequired = 0
    xlNever = 2

class XlRoutingSlipDelivery(IntEnum):
    xlAllAtOnce = 2
    xlOneAfterAnother = 1

class XlRoutingSlipStatus(IntEnum):
    xlNotYetRouted = 0
    xlRoutingComplete = 2
    xlRoutingInProgress = 1

class XlRowCol(IntEnum):
    xlColumns = 2
    xlRows = 1

class XlRunAutoMacro(IntEnum):
    xlAutoActivate = 3
    xlAutoClose = 2
    xlAutoDeactivate = 4
    xlAutoOpen = 1

class XlSaveAction(IntEnum):
    xlDoNotSaveChanges = 2
    xlSaveChanges = 1

class XlSaveAsAccessMode(IntEnum):
    xlExclusive = 3
    xlNoChange = 1
    xlShared = 2

class XlSaveConflictResolution(IntEnum):
    xlLocalSessionChanges = 2
    xlOtherSessionChanges = 3
    xlUserResolution = 1

class XlScaleType(IntEnum):
    xlScaleLinear = -4132 
    xlScaleLogarithmic = -4133 

class XlSearchDirection(IntEnum):
    xlNext = 1
    xlPrevious = 2

class XlSearchOrder(IntEnum):
    xlByColumns = 2
    xlByRows = 1

class XlSearchWithin(IntEnum):
    xlWithinSheet = 1
    xlWithinWorkbook = 2

class XlSeriesColorGradientStyle(IntEnum):
    xlSeriesColorGradientStyleDiverging = 1
    xlSeriesColorGradientStyleSequential = 0

class XlSeriesNameLevel(IntEnum):
    xlSeriesNameLevelAll = -1 
    xlSeriesNameLevelCustom = -2 
    xlSeriesNameLevelNone = -3 

class XlSheetType(IntEnum):
    xlChart = -4109 
    xlDialogSheet = -4116 
    xlExcel4IntlMacroSheet = 4
    xlExcel4MacroSheet = 3
    xlWorksheet = -4167 

class XlSheetVisibility(IntEnum):
    xlSheetHidden = 0
    xlSheetVeryHidden = 2
    xlSheetVisible = -1 

class XlSizeRepresents(IntEnum):
    xlSizeIsArea = 1
    xlSizeIsWidth = 2

class XlSlicerCacheType(IntEnum):
    xlSlicer = 1
    xlTimeline = 2

class XlSlicerCrossFilterType(IntEnum):
    xlSlicerCrossFilterHideButtonsWithNoData = 4
    xlSlicerCrossFilterShowItemsWithDataAtTop = 2
    xlSlicerCrossFilterShowItemsWithNoData = 3
    xlSlicerNoCrossFilter = 1

class XlSlicerSort(IntEnum):
    xlSlicerSortAscending = 2
    xlSlicerSortDataSourceOrder = 1
    xlSlicerSortDescending = 3

class XlSmartTagControlType(IntEnum):
    xlSmartTagControlActiveX = 13
    xlSmartTagControlButton = 6
    xlSmartTagControlCheckbox = 9
    xlSmartTagControlCombo = 12
    xlSmartTagControlHelp = 3
    xlSmartTagControlHelpURL = 4
    xlSmartTagControlImage = 8
    xlSmartTagControlLabel = 7
    xlSmartTagControlLink = 2
    xlSmartTagControlListbox = 11
    xlSmartTagControlRadioGroup = 14
    xlSmartTagControlSeparator = 5
    xlSmartTagControlSmartTag = 1
    xlSmartTagControlTextbox = 10

class XlSmartTagDisplayMode(IntEnum):
    xlButtonOnly = 2
    xlDisplayNone = 1
    xlIndicatorAndButton = 0

class XlSortDataOption(IntEnum):
    xlSortNormal = 0
    xlSortTextAsNumbers = 1

class XlSortMethod(IntEnum):
    xlPinYin = 1
    xlStroke = 2

class XlSortMethodOld(IntEnum):
    xlCodePage = 2
    xlSyllabary = 1

class XlSortOn(IntEnum):
    xlSortOnCellColor = 1
    xlSortOnFontColor = 2
    xlSortOnIcon = 3
    xlSortOnValues = 0

class XlSortOrder(IntEnum):
    xlAscending = 1
    xlDescending = 2

class XlSortOrientation(IntEnum):
    xlSortColumns = 1
    xlSortRows = 2

class XlSortType(IntEnum):
    xlSortLabels = 2
    xlSortValues = 1

class XlSourceType(IntEnum):
    xlSourceAutoFilter = 3
    xlSourceChart = 5
    xlSourcePivotTable = 6
    xlSourcePrintArea = 2
    xlSourceQuery = 7
    xlSourceRange = 4
    xlSourceSheet = 1
    xlSourceWorkbook = 0

class XlSpanishModes(IntEnum):
    xlSpanishTuteoAndVoseo = 1
    xlSpanishTuteoOnly = 0
    xlSpanishVoseoOnly = 2

class XlSparklineRowCol(IntEnum):
    xlSparklineColumnsSquare = 2
    xlSparklineNonSquare = 0
    xlSparklineRowsSquare = 1

class XlSparkScale(IntEnum):
    xlSparkScaleCustom = 3
    xlSparkScaleGroup = 1
    xlSparkScaleSingle = 2

class XlSparkType(IntEnum):
    xlSparkColumn = 2
    xlSparkColumnStacked100 = 3
    xlSparkLine = 1

class XlSpeakDirection(IntEnum):
    xlSpeakByColumns = 1
    xlSpeakByRows = 0

class XlSpecialCellsValue(IntEnum):
    xlErrors = 16 
    xlLogical = 4
    xlNumbers = 1
    xlTextValues = 2

class XlStdColorScale(IntEnum):
    xlColorScaleBlackWhite = 3
    xlColorScaleGYR = 2
    xlColorScaleRYG = 1
    xlColorScaleWhiteBlack = 4

class XlSubscribeToFormat(IntEnum):
    xlSubscribeToPicture = -4147 
    xlSubscribeToText = -4158 

class XlSubtototalLocationType(IntEnum):
    xlAtBottom = 2
    xlAtTop = 1

class XlSummaryColumn(IntEnum):
    xlSummaryOnLeft = -4131 
    xlSummaryOnRight = -4152 

class XlSummaryReportType(IntEnum):
    xlStandardSummary = 1
    xlSummaryPivotTable = -4148 

class XlSummaryRow(IntEnum):
    xlSummaryAbove = 0
    xlSummaryBelow = 1

class XlTableStyleElementType(IntEnum):
    xlBlankRow = 19 
    xlColumnStripe1 = 7
    xlColumnStripe2 = 8
    xlColumnSubheading1 = 20 
    xlColumnSubheading2 = 21 
    xlColumnSubheading3 = 22 
    xlFirstColumn = 3
    xlFirstHeaderCell = 9
    xlFirstTotalCell = 11
    xlGrandTotalColumn = 4
    xlGrandTotalRow = 2
    xlHeaderRow = 1
    xlLastColumn = 4
    xlLastHeaderCell = 10
    xlLastTotalCell = 12
    xlPageFieldLabels = 26 
    xlPageFieldValues = 27 
    xlRowStripe1 = 5
    xlRowStripe2 = 6
    xlRowSubheading1 = 23 
    xlRowSubheading2 = 24 
    xlRowSubheading3 = 25 
    xlSlicerHoveredSelectedItemWithData = 33 
    xlSlicerHoveredSelectedItemWithNoData = 35 
    xlSlicerHoveredUnselectedItemWithData = 32 
    xlSlicerHoveredUnselectedItemWithNoData = 34 
    xlSlicerSelectedItemWithData = 30 
    xlSlicerSelectedItemWithNoData = 31 
    xlSlicerUnselectedItemWithData = 28 
    xlSlicerUnselectedItemWithNoData = 29 
    xlSubtotalColumn1 = 13
    xlSubtotalColumn2 = 14
    xlSubtotalColumn3 = 15
    xlSubtotalRow1 = 16 
    xlSubtotalRow2 = 17 
    xlSubtotalRow3 = 18 
    xlTimelinePeriodLabels1 = 38 
    xlTimelinePeriodLabels2 = 39 
    xlTimelineSelectedTimeBlock = 40 
    xlTimelineSelectedTimeBlockSpace = 42 
    xlTimelineSelectionLabel = 36 
    xlTimelineTimeLevel = 37 
    xlTimelineUnselectedTimeBlock = 41 
    xlTotalRow = 2
    xlWholeTable = 0

class XlTabPosition(IntEnum):
    xlTabPositionFirst = 0
    xlTabPositionLast = 1

class XlTextParsingType(IntEnum):
    xlDelimited = 1
    xlFixedWidth = 2

class XlTextQualifier(IntEnum):
    xlTextQualifierDoubleQuote = 1
    xlTextQualifierNone = -4142 
    xlTextQualifierSingleQuote = 2

class XlTextVisualLayoutType(IntEnum):
    xlTextVisualLTR = 1
    xlTextVisualRTL = 2

class XlThemeColor(IntEnum):
    xlThemeColorAccent1 = 5
    xlThemeColorAccent2 = 6
    xlThemeColorAccent3 = 7
    xlThemeColorAccent4 = 8
    xlThemeColorAccent5 = 9
    xlThemeColorAccent6 = 10
    xlThemeColorDark1 = 1
    xlThemeColorDark2 = 3
    xlThemeColorFollowedHyperlink = 12
    xlThemeColorHyperlink = 11
    xlThemeColorLight1 = 2
    xlThemeColorLight2 = 4

class XlThemeFont(IntEnum):
    xlThemeFontMajor = 1
    xlThemeFontMinor = 2
    xlThemeFontNone = 0

class XlThreadMode(IntEnum):
    xlThreadModeAutomatic = 0
    xlThreadModeManual = 1

class XlTickLabelOrientation(IntEnum):
    xlTickLabelOrientationAutomatic = -4105 
    xlTickLabelOrientationDownward = -4170 
    xlTickLabelOrientationHorizontal = -4128 
    xlTickLabelOrientationUpward = -4171 
    xlTickLabelOrientationVertical = -4166 

class XlTickLabelPosition(IntEnum):
    xlTickLabelPositionHigh = -4127 
    xlTickLabelPositionLow = -4134 
    xlTickLabelPositionNextToAxis = 4
    xlTickLabelPositionNone = -4142 

class XlTickMark(IntEnum):
    xlTickMarkCross = 4
    xlTickMarkInside = 2
    xlTickMarkNone = -4142 
    xlTickMarkOutside = 3

class XlTimelineLevel(IntEnum):
    xlTimelineLevelDays = 3
    xlTimelineLevelMonths = 2
    xlTimelineLevelQuarters = 1
    xlTimelineLevelYears = 0

class XlTimePeriods(IntEnum):
    xlLast7Days = 2
    xlLastMonth = 5
    xlLastWeek = 4
    xlNextMonth = 8
    xlNextWeek = 7
    xlThisMonth = 9
    xlThisWeek = 3
    xlToday = 0
    xlTomorrow = 6
    xlYesterday = 1

class XlTimeUnit(IntEnum):
    xlDays = 0
    xlMonths = 1
    xlYears = 2

class XlToolbarProtection(IntEnum):
    xlNoButtonChanges = 1
    xlNoChanges = 4
    xlNoDockingChanges = 3
    xlNoShapeChanges = 2
    xlToolbarProtectionNone = -4143 

class XlTopBottom(IntEnum):
    xlTop10Bottom = 0
    xlTop10Top = 1

class XlTotalsCalculation(IntEnum):
    xlTotalsCalculationAverage = 2
    xlTotalsCalculationCount = 3
    xlTotalsCalculationCountNums = 4
    xlTotalsCalculationCustom = 9
    xlTotalsCalculationMax = 6
    xlTotalsCalculationMin = 5
    xlTotalsCalculationNone = 0
    xlTotalsCalculationStdDev = 7
    xlTotalsCalculationSum = 1
    xlTotalsCalculationVar = 8

class XlTrendlineType(IntEnum):
    xlExponential = 5
    xlLinear = -4132 
    xlLogarithmic = -4133 
    xlMovingAvg = 6
    xlPolynomial = 3
    xlPower = 4

class XlUnderlineStyle(IntEnum):
    xlUnderlineStyleDouble = -4119 
    xlUnderlineStyleDoubleAccounting = 5
    xlUnderlineStyleNone = -4142 
    xlUnderlineStyleSingle = 2
    xlUnderlineStyleSingleAccounting = 4

class XlUpdateLinks(IntEnum):
    xlUpdateLinksAlways = 3
    xlUpdateLinksNever = 2
    xlUpdateLinksUserSetting = 1

class XlVAlign(IntEnum):
    xlVAlignBottom = -4107 
    xlVAlignCenter = -4108 
    xlVAlignDistributed = -4117 
    xlVAlignJustify = -4130 
    xlVAlignTop = -4160 

class XlValueSortOrder(IntEnum):
    xlValueAscending = 1
    xlValueDescending = 2
    xlValueNone = 0

class XlWBATemplate(IntEnum):
    xlWBATChart = -4109 
    xlWBATExcel4IntlMacroSheet = 4
    xlWBATExcel4MacroSheet = 3
    xlWBATWorksheet = -4167 

class XlWebFormatting(IntEnum):
    xlWebFormattingAll = 1
    xlWebFormattingNone = 3
    xlWebFormattingRTF = 2

class XlWebSelectionType(IntEnum):
    xlAllTables = 2
    xlEntirePage = 1
    xlSpecifiedTables = 3

class XlWindowState(IntEnum):
    xlMaximized = -4137 
    xlMinimized = -4140 
    xlNormal = -4143 

class XlWindowType(IntEnum):
    xlChartAsWindow = 5
    xlChartInPlace = 4
    xlClipboard = 3
    xlInfo = -4129 
    xlWorkbook = 1

class XlWindowView(IntEnum):
    xlNormalView = 1
    xlPageBreakPreview = 2
    xlPageLayoutView = 3

class XlXLMMacroType(IntEnum):
    xlCommand = 2
    xlFunction = 1
    xlNotXLM = 3

class XlXmlExportResult(IntEnum):
    xlXmlExportSuccess = 0
    xlXmlExportValidationFailed = 1

class XlXmlImportResult(IntEnum):
    xlXmlImportElementsTruncated = 1
    xlXmlImportSuccess = 0
    xlXmlImportValidationFailed = 2

class XlXmlLoadOption(IntEnum):
    xlXmlLoadImportToList = 2
    xlXmlLoadMapXml = 3
    xlXmlLoadOpenXml = 1
    xlXmlLoadPromptUser = 0

class XlYesNoGuess(IntEnum):
    xlGuess = 0
    xlNo = 2
    xlYes = 1

class XmlDataBinding:
    Application: Application
    Creator: XlCreator
    Parent = None
    SourceUrl: str
    def ClearSettings(self, ): ...
    def LoadSettings(self, Url: str): ...
    def Refresh(self, ) -> XlXmlImportResult: ...

class XmlMap:
    AdjustColumnWidth: bool
    AppendOnImport: bool
    Application: Application
    Creator: XlCreator
    DataBinding: XmlDataBinding
    IsExportable: bool
    Name: str
    Parent = None
    PreserveColumnFilter: bool
    PreserveNumberFormatting: bool
    RootElementName: str
    RootElementNamespace: 'XmlNamespace'
    SaveDataSourceDefinition: bool
    Schemas: 'XmlSchemas'
    ShowImportExportValidationErrors: bool
    WorkbookConnection: WorkbookConnection
    def Delete(self, ): ...
    def Export(self, Url: str, Overwrite = None) -> XlXmlExportResult: ...
    def ExportXml(self, Data: str) -> XlXmlExportResult: ...
    def Import(self, Url: str, Overwrite = None) -> XlXmlImportResult: ...
    def ImportXml(self, XmlData: str, Overwrite = None) -> XlXmlImportResult: ...

class XmlMaps:
    Application: Application
    Count: float
    Creator: XlCreator
    Parent = None
    def __call__(self, Index) -> 'XmlMap': ...
    def Add(self, Schema: str, RootElementName = None) -> XmlMap: ...
    @property
    def Item(self, Index) -> XmlMap: ...

class XmlNamespace:
    Application: Application
    Creator: XlCreator
    Parent = None
    Prefix: str
    Uri: str

class XmlNamespaces:
    Application: Application
    Count: float
    Creator: XlCreator
    Parent = None
    Value: str
    def __call__(self, Index) -> 'XmlNamespace': ...
    def InstallManifest(self, Path: str, InstallForAllUsers = None): ...
    @property
    def Item(self, Index) -> XmlNamespace: ...

class XmlSchema:
    Application: Application
    Creator: XlCreator
    Name: str
    Namespace: XmlNamespace
    Parent = None
    XML: str

class XmlSchemas:
    Application: Application
    Count: float
    Creator: XlCreator
    Parent = None
    def __call__(self, Index) -> 'XmlSchema': ...
    @property
    def Item(self, Index) -> XmlSchema: ...

class XPath:
    Application: Application
    Creator: XlCreator
    Map: XmlMap
    Parent = None
    Repeating: bool
    Value: str
    def Clear(self, ): ...
    def SetValue(self, Map: XmlMap, XPath: str, SelectionNamespace = None, Repeating = None): ...

