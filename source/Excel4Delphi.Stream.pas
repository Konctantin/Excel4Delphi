unit Excel4Delphi.Stream;

interface

uses
  System.SysUtils, System.Classes, System.Types,
  {$IFDEF FMX}
  FMX.Graphics,
  {$ELSE}
  VCL.Graphics,
  {$ENDIF}
  System.UITypes, Windows, System.Zip, System.IOUtils, Excel4Delphi.Formula,
  Excel4Delphi.Xml, Excel4Delphi, Excel4Delphi.Common,
  System.Generics.Collections;

type
  TRelationType = (rtNone = -1, rtWorkSheet = 0, rtStyles = 1, rtSharedStr = 2, rtDoc = 3, rtCoreProp = 4,
    rtExtProps = 5, rtHyperlink = 6, rtComments = 7, rtVmlDrawing = 8, rtDrawing = 9);

  TZFileItem = record
    name: string; // путь к файлу
    nameArx: string;
    original: string; // исходная строка
    ftype: TRelationType; // тип контента
  end;

  TZRelation = record
    id: string; // rID
    ftype: TRelationType; // тип ссылки
    target: string; // ссылка на файла
    fileid: Integer; // ссылка на запись
    name: string; // имя листа
    state: Byte; // состояние
    sheetid: Integer; // номер листа
  end;

  TZXLSXFileItem = record
    name: string; // путь к файлу
    nameArx: string;
    original: string; // исходная строка
    ftype: TRelationType; // тип контента
  end;

  TZXLSXRelations = record
    id: string; // rID
    ftype: TRelationType; // тип ссылки
    target: string; // ссылка на файла
    fileid: Integer; // ссылка на запись
    name: string; // имя листа
    state: Byte; // состояние
    sheetid: Integer; // номер листа
  end;

  TZXLSXRelationsArray = array of TZXLSXRelations;

  TZXLSXDiffBorderItemStyle = class(TPersistent)
  private
    FUseStyle: Boolean; // заменять ли стиль
    FUseColor: Boolean; // заменять ли цвет
    FColor: TColor; // цвет линий
    FLineStyle: TZBorderType; // стиль линий
    FWeight: Byte;
  protected
  public
    constructor Create();
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseStyle: Boolean read FUseStyle write FUseStyle;
    property UseColor: Boolean read FUseColor write FUseColor;
    property Color: TColor read FColor write FColor;
    property LineStyle: TZBorderType read FLineStyle write FLineStyle;
    property Weight: Byte read FWeight write FWeight;
  end;

  TZXLSXDiffBorder = class(TPersistent)
  private
    FBorder: array [0 .. 5] of TZXLSXDiffBorderItemStyle;
    procedure SetBorder(Num: TZBordersPos; Const Value: TZXLSXDiffBorderItemStyle);
    function GetBorder(Num: TZBordersPos): TZXLSXDiffBorderItemStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property Border[Num: TZBordersPos]: TZXLSXDiffBorderItemStyle read GetBorder write SetBorder; default;
  end;

  // Итем для дифференцированного форматирования
  // TODO: возможно, для excel xml тоже понадобится (перенести?)
  TZXLSXDiffFormattingItem = class(TPersistent)
  private
    FUseFont: Boolean; // заменять ли шрифт
    FUseFontColor: Boolean; // заменять ли цвет шрифта
    FUseFontStyles: Boolean; // заменять ли стиль шрифта
    FFontColor: TColor; // цвет шрифта
    FFontStyles: TFontStyles; // стиль шрифта
    FUseBorder: Boolean; // заменять ли рамку
    FBorders: TZXLSXDiffBorder; // Что менять в рамке
    FUseFill: Boolean; // заменять ли заливку
    FUseCellPattern: Boolean; // Заменять ли тип заливки
    FCellPattern: TZCellPattern; // тип заливки
    FUseBGColor: Boolean; // заменять ли цвет заливки
    FBGColor: TColor; // цвет заливки
    FUsePatternColor: Boolean; // Заменять ли цвет шаблона заливки
    FPatternColor: TColor; // Цвет шаблона заливки
  protected
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseFont: Boolean read FUseFont write FUseFont;
    property UseFontColor: Boolean read FUseFontColor write FUseFontColor;
    property UseFontStyles: Boolean read FUseFontStyles write FUseFontStyles;
    property FontColor: TColor read FFontColor write FFontColor;
    property FontStyles: TFontStyles read FFontStyles write FFontStyles;
    property UseBorder: Boolean read FUseBorder write FUseBorder;
    property Borders: TZXLSXDiffBorder read FBorders write FBorders;
    property UseFill: Boolean read FUseFill write FUseFill;
    property UseCellPattern: Boolean read FUseCellPattern write FUseCellPattern;
    property CellPattern: TZCellPattern read FCellPattern write FCellPattern;
    property UseBGColor: Boolean read FUseBGColor write FUseBGColor;
    property BGColor: TColor read FBGColor write FBGColor;
    property UsePatternColor: Boolean read FUsePatternColor write FUsePatternColor;
    property PatternColor: TColor read FPatternColor write FPatternColor;
  end;

  // Differential formating
  TZXLSXDiffFormatting = class(TPersistent)
  private
    FCount: Integer;
    FMaxCount: Integer;
    FItems: array of TZXLSXDiffFormattingItem;
  protected
    function GetItem(Num: Integer): TZXLSXDiffFormattingItem;
    procedure SetItem(Num: Integer; const Value: TZXLSXDiffFormattingItem);
    procedure SetCount(ACount: Integer);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Add();
    procedure Assign(Source: TPersistent); override;
    procedure Clear();
    property Count: Integer read FCount;
    property Items[Num: Integer]: TZXLSXDiffFormattingItem read GetItem write SetItem; default;
  end;

  // List of cell number formats (date/numbers/currencies etc formats)
  TZEXLSXNumberFormats = class
  private
    FFormatsCount: Integer;
    FFormats: array of string; // numFmts (include default formats)
    FStyleFmtID: array of Integer;
    FStyleFmtIDCount: Integer;
  protected
    function GetFormat(Num: Integer): string;
    procedure SetFormat(Num: Integer; const Value: string);
    function GetStyleFMTID(Num: Integer): Integer;
    procedure SetStyleFMTID(Num: Integer; const Value: Integer);
    procedure SetStyleFMTCount(Value: Integer);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure ReadNumFmts(const Xml: TZsspXMLReaderH);
    function IsDateFormat(StyleNum: Integer): Boolean;
    function FindFormatID(const Value: string): Integer;
    property FormatsCount: Integer read FFormatsCount;
    property Format[Num: Integer]: string read GetFormat write SetFormat; default;
    property StyleFMTID[Num: Integer]: Integer read GetStyleFMTID write SetStyleFMTID;
    property StyleFMTCount: Integer read FStyleFmtIDCount write SetStyleFMTCount;
  end;

  TZEXLSXReadHelper = class
  private
    FDiffFormatting: TZXLSXDiffFormatting;
    FNumberFormats: TZEXLSXNumberFormats;
  protected
    procedure SetDiffFormatting(const Value: TZXLSXDiffFormatting);
  public
    constructor Create();
    destructor Destroy(); override;
    property DiffFormatting: TZXLSXDiffFormatting read FDiffFormatting write SetDiffFormatting;
    property NumberFormats: TZEXLSXNumberFormats read FNumberFormats;
  end;

  // Store link item
  TZEXLSXHyperLinkItem = record
    RID: Integer;
    RelType: TRelationType;
    CellRef: string;
    target: string;
    ScreenTip: string;
    TargetMode: string;
  end;

  { TZEXLSXWriteHelper }

  TZEXLSXWriteHelper = class
  private
    FHyperLinks: array of TZEXLSXHyperLinkItem;
    FHyperLinksCount: Integer;
    FMaxHyperLinksCount: Integer;
    FCurrentRID: Integer; // Current rID number (for HyperLinks/comments etc)
    FisHaveComments: Boolean; // Is Need create comments*.xml?
    FisHaveDrawings: Boolean; // Is Need create drawings*.xml?
    FSheetHyperlinksArray: array of Integer;
    FSheetHyperlinksCount: Integer;
  protected
    function GenerateRID(): Integer;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure AddHyperLink(const ACellRef, ATarget, AScreenTip, ATargetMode: string);
    function AddDrawing(const ATarget: string): Integer;
    procedure WriteHyperLinksTag(const Xml: TZsspXMLWriterH);
    function CreateSheetRels(const Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string;
      BOM: ansistring): Integer;
    procedure AddSheetHyperlink(PageNum: Integer);
    function IsSheetHaveHyperlinks(PageNum: Integer): Boolean;
    procedure Clear();
    property HyperLinksCount: Integer read FHyperLinksCount;
    property isHaveComments: Boolean read FisHaveComments write FisHaveComments; // Is need create comments*.xml?
    property isHaveDrawings: Boolean read FisHaveDrawings write FisHaveDrawings; // Is need create drawings*.xml?
  end;

  TExcel4DelphiReader = class(TObject)
  private
    // todo: FRawStyleList: TRawStyleList;
    FSharedStrings: TObjectList<TRichText>;
    FFileList: TList<TZFileItem>;
    FRelationList: TList<TZRelation>;
    FThemaColorList: TList<TColor>;
    FWorkBook: TZWorkBook;

    function ReadRunProperties(AXml: TZsspXMLReaderH; ATagName: string): TZFont;
    procedure ReadComments(AStream: TStream; ASheetIndex: Integer);
  public
    constructor Create(AWorkBook: TZWorkBook);
    destructor Destroy(); override;

    procedure ReadTheme(AStream: TStream); overload;
    procedure ReadTheme(const AFileName: string); overload;

    procedure ReadContentTypes(AStream: TStream); overload;
    procedure ReadContentTypes(const AFileName: string); overload;

    procedure ReadSharedStrings(AStream: TStream); overload;
    procedure ReadSharedStrings(const AFileName: string); overload;

    procedure ReadStyles(AStream: TStream); overload;
    procedure ReadStyles(const AFileName: string); overload;

    procedure ReadRelationships(AStream: TStream); overload;
    procedure ReadRelationships(const AFileName: string); overload;

    procedure ReadWorkBook(AStream: TStream); overload;
    procedure ReadWorkBook(const AFileName: string); overload;

    procedure ReadWorkSheet(AStream: TStream; ASheetIndex: Integer); overload;
    procedure ReadWorkSheet(const AFileName: string; ASheetIndex: Integer); overload;

    procedure LoadFromStream(AStream: TStream);
    procedure LoadFromDir(const ADirName: string);
    procedure LoadFromFile(const AFileName: string);
  end;

  TExcel4DelphiWriter = class
  private
    FSharedStrings: TObjectDictionary<TRichText, Integer>;
    // FFileList: TList<TZXLSXFileItem>; // never used
    FCommentList: TList<string>;
    FWorkBook: TZWorkBook;

  // procedure ReadRunProperties(AXml: TZsspXMLWriterH; ATagName: string; AFont: TZFont);// never used
  public
    constructor Create(AWorkBook: TZWorkBook);
    destructor Destroy(); override;

    procedure WriteTheme(AStream: TStream); overload;
    procedure WriteTheme(const AFileName: string); overload;

    procedure WriteContentTypes(AStream: TStream); overload;
    procedure WriteContentTypes(const AFileName: string); overload;

    procedure WriteSharedStrings(AStream: TStream); overload;
    procedure WriteSharedStrings(const AFileName: string); overload;

    procedure WriteStyles(AStream: TStream); overload;
    procedure WriteStyles(const AFileName: string); overload;

    procedure WriteRelationships(AStream: TStream); overload;
    procedure WriteRelationships(const AFileName: string); overload;

    procedure WriteDrawings(AStream: TStream); overload;
    procedure WriteDrawings(const AFileName: string); overload;

    procedure WriteDrawingRels(AStream: TStream); overload;
    procedure WriteDrawingRels(const AFileName: string); overload;

    procedure WriteWorkBook(AStream: TStream); overload;
    procedure WriteWorkBook(const AFileName: string); overload;

    procedure WriteDocPropsApp(AStream: TStream); overload;
    procedure WriteDocPropsApp(const AFileName: string); overload;

    procedure WriteDocPropsCore(AStream: TStream); overload;
    procedure WriteDocPropsCore(const AFileName: string); overload;

    procedure WriteWorkSheet(AStream: TStream; ASheetIndex: Integer); overload;
    procedure WriteWorkSheet(const AFileName: string; const ASheetIndex: Integer); overload;

    procedure SaveToStream(AStream: TStream);
    procedure SaveToDir(ADirName: string);
    procedure SaveToFile(const AFileName: string);
  end;

  TZEXMLSSHelper = class helper for TZWorkBook
  public
    procedure LoadFromStream(AStream: TStream);
    procedure LoadFromFile(const AFileName: string);
    procedure SaveToStream(AStream: TStream);
    procedure SaveToFile(const AFileName: string);
  end;

// Дополнительные функции для экспорта отдельных файлов
function ZEXLSXCreateStyles(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
function ZEXLSXCreateWorkBook(var XMLSS: TZWorkBook; Stream: TStream; const _pages: TIntegerDynArray;
  const _names: TStringDynArray; PageCount: Integer; TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring): Integer;
function ZEXLSXCreateSheet(var XMLSS: TZWorkBook; Stream: TStream; SheetNum: Integer;
  var SharedStrings: TStringDynArray; const SharedStringsDictionary: TDictionary<string, Integer>;
  TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring;
  const WriteHelper: TZEXLSXWriteHelper): Integer;
function ZEXLSXCreateContentTypes(var XMLSS: TZWorkBook; Stream: TStream; PageCount: Integer; CommentCount: Integer;
  const PagesComments: TIntegerDynArray; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring;
  const WriteHelper: TZEXLSXWriteHelper): Integer;
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring): Integer;
function ZEXLSXCreateSharedStrings(var XMLSS: TZWorkBook; Stream: TStream; const SharedStrings: TStringDynArray;
  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): Integer;
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring): Integer;
function ZEXLSXCreateDocPropsCore(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
function ZEXLSXCreateDrawing(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring): Integer;
function ZEXLSXCreateDrawingRels(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: String; BOM: ansistring): Integer;
procedure ZEAddRelsRelation(Xml: TZsspXMLWriterH; const RID: string; ridType: TRelationType; const target: string;
  const TargetMode: string = '');

function ReadXLSXPath(var XMLSS: TZWorkBook; DirName: string): Integer;
function ReadXLSXFile(var XMLSS: TZWorkBook; zipStream: TStream): Integer;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = '')
  : Integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string): Integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string): Integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZWorkBook; zipStream: TStream; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring = ''): Integer;

// Дополнительные функции, на случай чтения отдельного файла
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray;
  var ThemaColorCount: Integer): Boolean;
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TArray<TZXLSXFileItem>;
  var FilesCount: Integer): Boolean;
function ZEXSLXReadSharedStrings(var Stream: TStream; out StrArray: TStringDynArray; out StrCount: Integer): Boolean;
function ZEXSLXReadStyles(var XMLSS: TZWorkBook; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray;
  var ThemaColorCount: Integer; var MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): Boolean;
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: Integer;
  var isWorkSheet: Boolean; needReplaceDelimiter: Boolean): Boolean;
function ZEXSLXReadWorkBook(var XMLSS: TZWorkBook; var Stream: TStream; var Relations: TZXLSXRelationsArray;
  var RelationsCount: Integer): Boolean;
function ZEXSLXReadSheet(var XMLSS: TZWorkBook; var Stream: TStream; const SheetName: string;
  var StrArray: TStringDynArray; StrCount: Integer; var Relations: TZXLSXRelationsArray; RelationsCount: Integer;
  MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): Boolean;
function ZEXSLXReadComments(var XMLSS: TZWorkBook; var Stream: TStream): Boolean;

implementation

uses
  System.AnsiStrings, System.StrUtils, System.Math, Excel4Delphi.NumberFormats, System.NetEncoding;

const
{$REGION 'Indexed Colors'}
  INDEXED_COLORS: array [0 .. 63] of TColor = ($00000000, // 0
    $00FFFFFF, // 1
    $00FF0000, // 2
    $0000FF00, // 3
    $000000FF, // 4
    $00FFFF00, // 5
    $00FF00FF, // 6
    $0000FFFF, // 7
    $00000000, // 8
    $00FFFFFF, // 9
    $00FF0000, // 10
    $0000FF00, // 11
    $000000FF, // 12
    $00FFFF00, // 13
    $00FF00FF, // 14
    $0000FFFF, // 15
    $00800000, // 16
    $00008000, // 17
    $00000080, // 18
    $00808000, // 19
    $00800080, // 20
    $00008080, // 21
    $00C0C0C0, // 22
    $00808080, // 23
    $009999FF, // 24
    $00993366, // 25
    $00FFFFCC, // 26
    $00CCFFFF, // 27
    $00660066, // 28
    $00FF8080, // 29
    $000066CC, // 30
    $00CCCCFF, // 31
    $00000080, // 32
    $00FF00FF, // 33
    $00FFFF00, // 34
    $0000FFFF, // 35
    $00800080, // 36
    $00800000, // 37
    $00008080, // 38
    $000000FF, // 39
    $0000CCFF, // 40
    $00CCFFFF, // 41
    $00CCFFCC, // 42
    $00FFFF99, // 43
    $0099CCFF, // 44
    $00FF99CC, // 45
    $00CC99FF, // 46
    $00FFCC99, // 47
    $003366FF, // 48
    $0033CCCC, // 49
    $0099CC00, // 50
    $00FFCC00, // 51
    $00FF9900, // 52
    $00FF6600, // 53
    $00666699, // 54
    $00969696, // 55
    $00003366, // 56
    $00339966, // 57
    $00003300, // 58
    $00333300, // 59
    $00993300, // 60
    $00993366, // 61
    $00333399, // 62
    $00333333 // 63
    );
{$ENDREGION 'Indexed Colors'}

const
  SCHEMA_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006';
  SCHEMA_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  SCHEMA_PACKAGE = 'http://schemas.openxmlformats.org/package/2006';
  SCHEMA_PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  SCHEMA_SHEET_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

type
  TZEXLSXFont = record
    name: string;
    bold: Boolean;
    italic: Boolean;
    underline: Boolean;
    strike: Boolean;
    charset: Integer;
    Color: TColor;
    ColorType: Byte;
    LumFactor: double;
    fontsize: double;
    superscript: Boolean;
    subscript: Boolean;
  end;

  TZEXLSXFontArray = array of TZEXLSXFont;

  TContentTypeRec = record
    ftype: TRelationType;
    name: string;
    rel: string;
  end;

const
  CONTENT_TYPES: array [0 .. 10] of TContentTypeRec = ((ftype: TRelationType.rtWorkSheet;
    name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'),

    (ftype: TRelationType.rtStyles; name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'),

    (ftype: TRelationType.rtSharedStr;
    name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),

    (ftype: TRelationType.rtSharedStr;
    name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),

    (ftype: TRelationType.rtDoc; name: 'application/vnd.openxmlformats-package.relationships+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'),

    (ftype: TRelationType.rtCoreProp;
    name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'),

    (ftype: TRelationType.rtExtProps; name: 'application/vnd.openxmlformats-package.core-properties+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'),

    (ftype: TRelationType.rtHyperlink; name: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'),

    (ftype: TRelationType.rtComments; name: 'application/vnd.openxmlformats-officedocument.vmlDrawing';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'),

    (ftype: TRelationType.rtVmlDrawing; name: 'application/vnd.openxmlformats-officedocument.theme+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'),

    (ftype: TRelationType.rtDrawing; name: 'application/vnd.openxmlformats-officedocument.drawing+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'));

function GetMaximumDigitWidth(fontName: string; fontsize: double): double;
const
  numbers = '0123456789';
var
  {$IFDEF FMX}
  bitmap: FMX.Graphics.TBitmap;
  {$ELSE}
  bitmap: VCL.Graphics.TBitmap;
  {$ENDIF}
  number: string;
begin
  // А.А.Валуев Расчитываем ширину самого широкого числа.
  Result := 0;
  {$IFDEF FMX}
  bitmap := FMX.Graphics.TBitmap.Create;
  bitmap.Canvas.Font.Family := fontName;
  {$ELSE}
  bitmap := VCL.Graphics.TBitmap.Create;
  bitmap.Canvas.Font.PixelsPerInch := 96;
  bitmap.Canvas.Font.name := fontName;
  {$ENDIF}
  try
    bitmap.Canvas.Font.Size := Trunc(fontsize);
    for number in numbers do
      Result := Max(Result, bitmap.Canvas.TextWidth(number));
  finally
    bitmap.Free;
  end;
end;

/// /:::::::::::::  TZXLSXDiffFormating :::::::::::::::::////

constructor TZXLSXDiffFormatting.Create();
var
  i: Integer;
begin
  FCount := 0;
  FMaxCount := 20;
  SetLength(FItems, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FItems[i] := TZXLSXDiffFormattingItem.Create();
end;

destructor TZXLSXDiffFormatting.Destroy();
var
  i: Integer;
begin
  for i := 0 to FMaxCount - 1 do
    if (Assigned(FItems[i])) then
      FreeAndNil(FItems[i]);
  inherited;
end;

procedure TZXLSXDiffFormatting.Add();
begin
  SetCount(FCount + 1);
  FItems[FCount - 1].Clear();
end;

procedure TZXLSXDiffFormatting.Assign(Source: TPersistent);
var
  df: TZXLSXDiffFormatting;
  i: Integer;
begin
  if Assigned(Source) and (Source is TZXLSXDiffFormatting) then
  begin
    df := Source as TZXLSXDiffFormatting;
    SetCount(df.Count);
    for i := 0 to Count - 1 do
      FItems[i].Assign(df[i]);
  end
  else
    inherited;
end; // Assign

procedure TZXLSXDiffFormatting.SetCount(ACount: Integer);
var
  i: Integer;
begin
  if (ACount >= FMaxCount) then
  begin
    FMaxCount := ACount + 20;
    SetLength(FItems, FMaxCount);
    // Здесь FCount + 1, потому что иначе затирается последний элемент.
    // В результате утечки и потеря считанного форматирования.
    for i := FCount + 1 to FMaxCount - 1 do
      FItems[i] := TZXLSXDiffFormattingItem.Create();
  end;
  FCount := ACount;
end;

procedure TZXLSXDiffFormatting.Clear();
begin
  FCount := 0;
end;

function TZXLSXDiffFormatting.GetItem(Num: Integer): TZXLSXDiffFormattingItem;
begin
  if ((Num >= 0) and (Num < Count)) then
    Result := FItems[Num]
  else
    Result := nil;
end;

procedure TZXLSXDiffFormatting.SetItem(Num: Integer; const Value: TZXLSXDiffFormattingItem);
begin
  if ((Num >= 0) and (Num < Count)) then
    if (Assigned(Value)) then
      FItems[Num].Assign(Value);
end;

/// /:::::::::::::  TZXLSXDiffBorderItemStyle :::::::::::::::::////

constructor TZXLSXDiffBorderItemStyle.Create();
begin
  Clear();
end;

procedure TZXLSXDiffBorderItemStyle.Assign(Source: TPersistent);
var
  bs: TZXLSXDiffBorderItemStyle;
  b: Boolean;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorderItemStyle) then
    begin
      b := false;
      bs := Source as TZXLSXDiffBorderItemStyle;

      FUseStyle := bs.UseStyle;
      FUseColor := bs.UseColor;
      FColor := bs.Color;
      FWeight := bs.Weight;
      FLineStyle := bs.LineStyle;
    end;
  if (b) then
    inherited;
end; // Assign

procedure TZXLSXDiffBorderItemStyle.Clear();
begin
  FUseStyle := false;
  FUseColor := false;
  FColor := TColorRec.Black;
  FWeight := 1;
  FLineStyle := ZENone;
end;

/// /::::::::::::: TZXLSXDiffBorder :::::::::::::::::////

constructor TZXLSXDiffBorder.Create();
var
  i: Integer;
begin
  for i := 0 to 5 do
    FBorder[i] := TZXLSXDiffBorderItemStyle.Create();
  Clear();
end;

destructor TZXLSXDiffBorder.Destroy();
var
  i: Integer;
begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);
  inherited;
end;

procedure TZXLSXDiffBorder.Assign(Source: TPersistent);
var
  brd: TZXLSXDiffBorder;
  b: Boolean;
  i: TZBordersPos;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorder) then
    begin
      b := false;
      brd := Source as TZXLSXDiffBorder;
      for i := bpLeft to bpDiagonalRight do
        FBorder[Ord(i)].Assign(brd[i]);
    end;

  if (b) then
    inherited;
end; // Assign

procedure TZXLSXDiffBorder.Clear();
var
  i: TZBordersPos;
begin
  for i := bpLeft to bpDiagonalRight do
    FBorder[Ord(i)].Clear();
end;

function TZXLSXDiffBorder.GetBorder(Num: TZBordersPos): TZXLSXDiffBorderItemStyle;
begin
  Result := nil;
  if ((Num >= bpLeft) and (Num <= bpDiagonalRight)) then
    Result := FBorder[Ord(Num)];
end;

procedure TZXLSXDiffBorder.SetBorder(Num: TZBordersPos; const Value: TZXLSXDiffBorderItemStyle);
begin
  if ((Num >= bpLeft) and (Num <= bpDiagonalRight)) then
    if (Assigned(Value)) then
      FBorder[Ord(Num)].Assign(Value);
end;

/// /::::::::::::: TZXLSXDiffFormatingItem :::::::::::::::::////

constructor TZXLSXDiffFormattingItem.Create();
begin
  FBorders := TZXLSXDiffBorder.Create();
  Clear();
end;

destructor TZXLSXDiffFormattingItem.Destroy();
begin
  FreeAndNil(FBorders);
  inherited;
end;

procedure TZXLSXDiffFormattingItem.Assign(Source: TPersistent);
var
  dxfItem: TZXLSXDiffFormattingItem;
  b: Boolean;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffFormattingItem) then
    begin
      b := false;
      dxfItem := Source as TZXLSXDiffFormattingItem;

      FUseFont := dxfItem.UseFont;
      FUseFontColor := dxfItem.UseFontColor;
      FUseFontStyles := dxfItem.UseFontStyles;
      FFontColor := dxfItem.FontColor;
      FFontStyles := dxfItem.FontStyles;
      FUseBorder := dxfItem.UseBorder;
      FUseFill := dxfItem.UseFill;
      FUseCellPattern := dxfItem.UseCellPattern;
      FCellPattern := dxfItem.CellPattern;
      FUseBGColor := dxfItem.UseBGColor;
      FBGColor := dxfItem.BGColor;
      FUsePatternColor := dxfItem.UsePatternColor;
      FPatternColor := dxfItem.PatternColor;
      FBorders.Assign(dxfItem.Borders);
    end;

  if (b) then
    inherited;
end; // Assign

procedure TZXLSXDiffFormattingItem.Clear();
begin
  FUseFont := false;
  FUseFontColor := false;
  FUseFontStyles := false;
  FFontColor := TColorRec.Black;
  FFontStyles := [];
  FUseBorder := false;
  FBorders.Clear();
  FUseFill := false;
  FUseCellPattern := false;
  FCellPattern := ZPNone;
  FUseBGColor := false;
  FBGColor := TColorRec.cWindow;
  FUsePatternColor := false;
  FPatternColor := TColorRec.cWindow;
end; // Clear

// END Differential Formatting
/// /////////////////////////////////////////////////////////////////////////////

/// /::::::::::::: TZEXLSXNumberFormats :::::::::::::::::////

constructor TZEXLSXNumberFormats.Create();
var
  i: Integer;
begin
  FStyleFmtIDCount := 0;
  FFormatsCount := 164;
  SetLength(FFormats, FFormatsCount);
  for i := 0 to FFormatsCount - 1 do
    FFormats[i] := '';

  // Some "Standart" formats for xlsx:
  FFormats[1] := '0';
  FFormats[2] := '0.00';
  FFormats[3] := '#,##0';
  FFormats[4] := '#,##0.00';
  FFormats[5] := '$#,##0;\-$#,##0';
  FFormats[6] := '$#,##0;[Red]\-$#,##0';
  FFormats[7] := '$#,##0.00;\-$#,##0.00';
  FFormats[8] := '$#,##0.00;[Red]\-$#,##0.00';
  FFormats[9] := '0%';
  FFormats[10] := '0.00%';
  FFormats[11] := '0.00E+00';
  FFormats[12] := '# ?/?';
  FFormats[13] := '# ??/??';

  FFormats[14] := 'm/d/yyyy';
  FFormats[15] := 'd-mmm-yy';
  FFormats[16] := 'd-mmm';
  FFormats[17] := 'mmm-yy';
  FFormats[18] := 'h:mm AM/PM';
  FFormats[19] := 'h:mm:ss AM/PM';
  FFormats[20] := 'h:mm';
  FFormats[21] := 'h:mm:ss';
  FFormats[22] := 'm/d/yyyy h:mm';

  FFormats[27] := '[$-404]e/m/d';
  FFormats[37] := '#,##0 ;(#,##0)';
  FFormats[38] := '#,##0 ;[Red](#,##0)';
  FFormats[39] := '#,##0.00;(#,##0.00)';
  FFormats[40] := '#,##0.00;[Red](#,##0.00)';
  FFormats[44] := '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)';
  FFormats[45] := 'mm:ss';
  FFormats[46] := '[h]:mm:ss';
  FFormats[47] := 'mmss.0';
  FFormats[48] := '##0.0E+0';
  FFormats[49] := '@';

  FFormats[59] := 't0';
  FFormats[60] := 't0.00';
  FFormats[61] := 't#,##0';
  FFormats[62] := 't#,##0.00';
  FFormats[67] := 't0%';
  FFormats[68] := 't0.00%';
  FFormats[69] := 't# ?/?';
  FFormats[70] := 't# ??/??';
  FFormats[81] := 'd/m/bb';
end;

destructor TZEXLSXNumberFormats.Destroy();
begin
  SetLength(FFormats, 0);
  SetLength(FStyleFmtID, 0);
  inherited;
end;

// Find format in formats. Return -1 if format not foud.
// INPUT
// const value: string - format
// RETURN
// integer - >= 0 - index number if store.
// -1 - not found
function TZEXLSXNumberFormats.FindFormatID(const Value: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to FFormatsCount - 1 do
    if (FFormats[i] = Value) then
    begin
      Result := i;
      break;
    end;
end;

function TZEXLSXNumberFormats.GetFormat(Num: Integer): string;
begin
  Result := '';
  if ((Num >= 0) and (Num < FFormatsCount)) then
    Result := FFormats[Num];
end;

procedure TZEXLSXNumberFormats.SetFormat(Num: Integer; const Value: string);
var
  i: Integer;
begin
  if ((Num >= 0) and (Num < FFormatsCount)) then
    FFormats[Num] := Value
  else if (Num >= 0) then
  begin
    SetLength(FFormats, Num + 1);
    for i := FFormatsCount to Num do
      FFormats[i] := '';
    FFormats[Num] := Value;
    FFormatsCount := Num + 1;
  end;
end;

function TZEXLSXNumberFormats.GetStyleFMTID(Num: Integer): Integer;
begin
  if ((Num >= 0) and (Num < FStyleFmtIDCount)) then
    Result := FStyleFmtID[Num]
  else
    Result := 0;
end;

function TZEXLSXNumberFormats.IsDateFormat(StyleNum: Integer): Boolean;
var
  fmtId: Integer;
begin
  Result := false;
  if ((StyleNum >= 0) and (StyleNum < FStyleFmtIDCount)) then
    fmtId := FStyleFmtID[StyleNum]
  else
    exit;

  // If default fmtID
  if ((fmtId >= 0) and (fmtId < 100)) then
    Result := fmtId in [14 .. 22, 27 .. 36, 45 .. 47, 50 .. 58, 71 .. 76, 78 .. 81]
  else
    Result := GetXlsxNumberFormatType(FFormats[fmtId]) = ZE_NUMFORMAT_IS_DATETIME;
end;

procedure TZEXLSXNumberFormats.ReadNumFmts(const Xml: TZsspXMLReaderH);
var
  temp: Integer;
begin
  with THTMLEncoding.Create do
    try
      while Xml.ReadToEndTagByName('numFmts') do
      begin
        if (Xml.TagName = 'numFmt') then
          if (TryStrToInt(Xml.Attributes['numFmtId'], temp)) then
            Format[temp] := Decode(Xml.Attributes['formatCode']);
      end;
    finally
      Free;
    end;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTID(Num: Integer; const Value: Integer);
begin
  if ((Num >= 0) and (Num < FStyleFmtIDCount)) then
    FStyleFmtID[Num] := Value;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTCount(Value: Integer);
begin
  if (Value >= 0) then
  begin
    if (Value > FStyleFmtIDCount) then
      SetLength(FStyleFmtID, Value);
    FStyleFmtIDCount := Value
  end;
end;

/// /::::::::::::: TZEXLSXReadHelper :::::::::::::::::////

constructor TZEXLSXReadHelper.Create();
begin
  FDiffFormatting := TZXLSXDiffFormatting.Create();
  FNumberFormats := TZEXLSXNumberFormats.Create();
end;

destructor TZEXLSXReadHelper.Destroy();
begin
  FreeAndNil(FDiffFormatting);
  FreeAndNil(FNumberFormats);
  inherited;
end;

procedure TZEXLSXReadHelper.SetDiffFormatting(const Value: TZXLSXDiffFormatting);
begin
  if (Assigned(Value)) then
    FDiffFormatting.Assign(Value);
end;

/// /::::::::::::: TZEXLSXWriteHelper :::::::::::::::::////

// Generate next RID for references
function TZEXLSXWriteHelper.GenerateRID(): Integer;
begin
  inc(FCurrentRID);
  Result := FCurrentRID;
end;

constructor TZEXLSXWriteHelper.Create();
begin
  FMaxHyperLinksCount := 10;
  FSheetHyperlinksCount := 0;
  SetLength(FHyperLinks, FMaxHyperLinksCount);
  Clear();
end;

destructor TZEXLSXWriteHelper.Destroy();
begin
  SetLength(FHyperLinks, 0);
  SetLength(FSheetHyperlinksArray, 0);
  inherited Destroy;
end;

procedure TZEXLSXWriteHelper.AddHyperLink(const ACellRef, ATarget, AScreenTip, ATargetMode: string);
var
  Num: Integer;
begin
  Num := FHyperLinksCount;
  inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then
  begin
    inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[Num].RID := GenerateRID();
  FHyperLinks[Num].RelType := TRelationType.rtHyperlink;
  FHyperLinks[Num].TargetMode := ATargetMode;
  FHyperLinks[Num].CellRef := ACellRef;
  FHyperLinks[Num].target := ATarget;
  FHyperLinks[Num].ScreenTip := AScreenTip;
end;

// Add hyperlink
// INPUT
// const ATarget: string     - drawing target (../drawings/drawing2.xml)
function TZEXLSXWriteHelper.AddDrawing(const ATarget: string): Integer;
var
  Num: Integer;
begin
  Num := FHyperLinksCount;
  inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then
  begin
    inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[Num].RID := GenerateRID();
  FHyperLinks[Num].RelType := TRelationType.rtDrawing;
  FHyperLinks[Num].TargetMode := '';
  FHyperLinks[Num].CellRef := '';
  FHyperLinks[Num].target := ATarget;
  FHyperLinks[Num].ScreenTip := '';
  Result := FHyperLinks[Num].RID;
  FisHaveDrawings := true;
end;

// Writes tag <hyperlinks> .. </hyperlinks>
// INPUT
// const xml: TZsspXMLWriterH
procedure TZEXLSXWriteHelper.WriteHyperLinksTag(const Xml: TZsspXMLWriterH);
var
  i: Integer;
begin
  if (FHyperLinksCount > 0) then
  begin
    Xml.Attributes.Clear();
    Xml.WriteTagNode('hyperlinks', true, true, true);
    for i := 0 to FHyperLinksCount - 1 do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('ref', FHyperLinks[i].CellRef);
      Xml.Attributes.Add('r:id', 'rId' + IntToStr(FHyperLinks[i].RID));

      if (FHyperLinks[i].ScreenTip <> '') then
        Xml.Attributes.Add('tooltip', FHyperLinks[i].ScreenTip);

      Xml.WriteEmptyTag('hyperlink', true);

      {
      xml.Attributes.Add('Id', 'rId' + IntToStr(FHyperLinks[i].ID));
      xml.Attributes.Add('Type', ZEXLSXGetRelationName(6));
      xml.Attributes.Add('Target', FHyperLinks[i].Target);
      if (FHyperLinks[i].TargetMode <> '')
         xml.Attributes.Add('TargetMode', FHyperLinks[i].TargetMode);
 }
    end; // for i
    Xml.WriteEndTagNode(); // hyperlinks
  end; // if
end; // WriteHyperLinksTag

// Create sheet relations
// INPUT
// const Stream: TStream
// TextConverter: TAnsiToCPConverter
// CodePageName: string
// BOM: ansistring
// RETURN
function TZEXLSXWriteHelper.CreateSheetRels(const Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  i: Integer;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    Xml.WriteTagNode('Relationships', true, true, false);

    for i := 0 to FHyperLinksCount - 1 do
      ZEAddRelsRelation(Xml, 'rId' + IntToStr(FHyperLinks[i].RID), FHyperLinks[i].RelType, FHyperLinks[i].target,
        FHyperLinks[i].TargetMode);

    Xml.WriteEndTagNode(); // Relationships
  finally
    Xml.Free();
  end;
end; // CreateSheetRels

procedure TZEXLSXWriteHelper.AddSheetHyperlink(PageNum: Integer);
begin
  SetLength(FSheetHyperlinksArray, FSheetHyperlinksCount + 1);
  FSheetHyperlinksArray[FSheetHyperlinksCount] := PageNum;
  inc(FSheetHyperlinksCount);
end;

function TZEXLSXWriteHelper.IsSheetHaveHyperlinks(PageNum: Integer): Boolean;
var
  i: Integer;
begin
  Result := false;
  for i := 0 to FSheetHyperlinksCount - 1 do
    if (FSheetHyperlinksArray[i] = PageNum) then
      exit(true);
end;

procedure TZEXLSXWriteHelper.Clear();
begin
  FHyperLinksCount := 0;
  FCurrentRID := 0;
  FisHaveComments := false;
end;

// Возвращает номер Relations из rels
// INPUT
// const name: string - текст отношения
// RETURN
// integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationNumber(const name: string): TRelationType;
var
  rec: TContentTypeRec;
begin
  Result := TRelationType.rtNone;
  for rec in CONTENT_TYPES do
  begin
    if rec.rel = name then
      exit(rec.ftype);
  end;
end; // ZEXLSXGetRelationNumber

// Возвращает текст Relations для rels
// INPUT
// num: integer - номер отношения
// RETURN
// integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationName(Num: TRelationType): string;
var
  rec: TContentTypeRec;
begin
  Result := '';
  for rec in CONTENT_TYPES do
  begin
    if rec.ftype = Num then
      exit(rec.rel);
  end;
end; // ZEXLSXGetRelationName

function XLSXBoolToStr(Value: Boolean): string;
begin
  if (Value) then
    Result := 'true'
  else
    Result := 'false';
end;

// Читает тему (themeXXX.xml)
// INPUT
// var Stream: TStream                   - поток чтения
// var ThemaFillsColors: TIntegerDynArray - массив с цветами заливки
// var ThemaColorCount: integer          - кол-во цветов заливки
// RETURN
// boolean - true - всё прочиталось успешно
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray;
  var ThemaColorCount: Integer): Boolean;
var
  Xml: TZsspXMLReaderH;
  maxCount: Integer;
  flag: Boolean;

  procedure _addFillColor(const _rgb: string);
  begin
    inc(ThemaColorCount);
    if (ThemaColorCount >= maxCount) then
    begin
      maxCount := ThemaColorCount + 20;
      SetLength(ThemaFillsColors, maxCount);
    end;
    ThemaFillsColors[ThemaColorCount - 1] := HTMLHexToColor(_rgb);
  end; // _addFillColor

begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  flag := false;
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;
    ThemaColorCount := 0;
    maxCount := -1;
    while (not Xml.Eof()) do
    begin
      Xml.ReadTag();
      if (Xml.TagName = 'a:clrScheme') then
      begin
        if (Xml.IsTagStart) then
          flag := true;
        if (Xml.IsTagEnd) then
          flag := false;
      end
      else if ((Xml.TagName = 'a:sysClr') and (flag) and (Xml.IsTagStartOrClosed)) then
      begin
        _addFillColor(Xml.Attributes.ItemsByName['lastClr']);
      end
      else if ((Xml.TagName = 'a:srgbClr') and (flag) and (Xml.IsTagStartOrClosed)) then
      begin
        _addFillColor(Xml.Attributes.ItemsByName['val']);
      end;
    end; // while

    Result := true;
  finally
    Xml.Free();
  end;
end; // ZEXSLXReadThema

// Читает список нужных файлов из [Content_Types].xml
// INPUT
// var Stream: TStream             - поток чтения
// var FileArray: TArray<TZXLSXFileItem>  - список файлов
// var FilesCount: integer         - кол-во файлов
// RETURN
// boolean - true - всё прочиталось успешно
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TArray<TZXLSXFileItem>;
  var FilesCount: Integer): Boolean;
var
  Xml: TZsspXMLReaderH;
  contType: string;
  rec: TContentTypeRec;
begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if Xml.BeginReadStream(Stream) <> 0 then
      exit;
    FilesCount := 0;
    while not Xml.Eof() do
    begin
      Xml.ReadTag();
      if Xml.IsTagClosedByName('Override') then
      begin
        contType := Xml.Attributes.ItemsByName['ContentType'];
        for rec in CONTENT_TYPES do
        begin
          if contType = rec.name then
          begin
            SetLength(FileArray, FilesCount + 1);
            FileArray[FilesCount].name := Xml.Attributes.ItemsByName['PartName'];
            FileArray[FilesCount].original := Xml.Attributes.ItemsByName['PartName'];
            FileArray[FilesCount].ftype := rec.ftype;
            inc(FilesCount);
            break;
          end;
        end;
      end;
    end;
    Result := true;
  finally
    Xml.Free();
  end;
end; // ZEXSLXReadContentTypes

// Читает строки из sharedStrings.xml
// INPUT
// var Stream: TStream           - поток для чтения
// var StrArray: TStringDynArray - возвращаемый массив со строками
// var StrCount: integer         - кол-во элементов
// RETURN
// boolean - true - всё ок
function ZEXSLXReadSharedStrings(var Stream: TStream; out StrArray: TStringDynArray; out StrCount: Integer): Boolean;
var
  Xml: TZsspXMLReaderH;
  s: string;
  k: Integer;
  rs: TRichString;
begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();

  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;
    StrCount := 0;

    while not Xml.Eof() do
    begin
      Xml.ReadTag();
      if Xml.IsTagStartByName('sst') then
      begin
        StrCount := StrToIntDef(Xml.Attributes['count'], 0);
        SetLength(StrArray, StrCount + 1);
        StrCount := 0;
      end;

      if Xml.IsTagStartByName('si') then
      begin
        s := '';
        k := 0;
        while Xml.ReadToEndTagByName('si') do
        begin
          rs := TRichString.Create();
          if Xml.IsTagStartByName('rPr') then
          begin
            rs.Font := TZFont.Create();

            while Xml.ReadToEndTagByName('rPr') do
            begin
              if Xml.IsTagClosedByName('b') then
                rs.Font.Style := rs.Font.Style + [TFontStyle.fsBold]
              else if Xml.IsTagClosedByName('u') then
                rs.Font.Style := rs.Font.Style + [TFontStyle.fsUnderline]
              else if Xml.IsTagClosedByName('i') then
                rs.Font.Style := rs.Font.Style + [TFontStyle.fsItalic]
              else if Xml.IsTagClosedByName('sz') then
                rs.Font.Size := StrToFloatDef(Xml.Attributes['val'], 11, TFormatSettings.Invariant)
              else if Xml.IsTagClosedByName('color ') then
              begin
                // theme or grb
              end
              else if Xml.IsTagClosedByName('rFont') then
                rs.Font.name := trim(Xml.Attributes['val'])
              else if Xml.IsTagClosedByName('charset') then
                rs.Font.charset := StrToIntDef(Xml.Attributes['val'], 0);
            end;
          end;

          if Xml.IsTagEndByName('t') then
          begin
            if (k > 1) then
              s := s + sLineBreak;

            s := s + Xml.TextBeforeTag;
          end;

          if Xml.IsTagEndByName('r') then
            inc(k);
        end; // while
        // SetLength(StrArray, StrCount + 1);
        StrArray[StrCount] := s;
        inc(StrCount);
      end; // if
    end; // while

    Result := true;
  finally
    Xml.Free();
  end;
end; // ZEXSLXReadSharedStrings

// Получить условное форматирование и оператор из xlsx
// INPUT
// const xlsxCfType: string              - xlsx тип условного форматирования
// const xlsxCFOperator: string          - xlsx оператор
// out CFCondition: TZCondition          - распознанное условие
// out CFOperator: TZConditionalOperator - распознанный оператор
// RETURN
// boolean - true - условное форматирование и оператор успешно распознаны
function ZEXLSX_getCFCondition(const xlsxCfType, xlsxCFOperator: string; out CFCondition: TZCondition;
  out CFOperator: TZConditionalOperator): Boolean;
var
  isCheckOperator: Boolean;
  procedure _SetCFOperator(AOperator: TZConditionalOperator);
  begin
    CFOperator := AOperator;
    CFCondition := ZCFCellContentOperator;
  end;

  // Проверить тип условного форматирования
  // out isNeddCheckOperator: boolean - возвращает, нужно ли проверять
  // оператор
  // RETURN
  // boolean - true - всё ок, можно проверять далее
  function _CheckXLSXCfType(out isNeddCheckOperator: Boolean): Boolean;
  begin
    Result := true;
    isNeddCheckOperator := true;
    if (xlsxCfType = 'cellIs') then
    begin
    end
    else if (xlsxCfType = 'containsText') then
      CFCondition := ZCFContainsText
    else if (xlsxCfType = 'notContains') then
      CFCondition := ZCFNotContainsText
    else if (xlsxCfType = 'beginsWith') then
      CFCondition := ZCFBeginsWithText
    else if (xlsxCfType = 'endsWith') then
      CFCondition := ZCFEndsWithText
    else if (xlsxCfType = 'containsBlanks') then
      isNeddCheckOperator := false
    else
      Result := false;
  end; // _CheckXLSXCfType

  // Проверить оператор
  function _CheckCFoperator(): Boolean;
  begin
    Result := true;
    if (xlsxCFOperator = 'lessThan') then
      _SetCFOperator(ZCFOpLT)
    else if (xlsxCFOperator = 'equal') then
      _SetCFOperator(ZCFOpEqual)
    else if (xlsxCFOperator = 'notEqual') then
      _SetCFOperator(ZCFOpNotEqual)
    else if (xlsxCFOperator = 'greaterThanOrEqual') then
      _SetCFOperator(ZCFOpGTE)
    else if (xlsxCFOperator = 'greaterThan') then
      _SetCFOperator(ZCFOpGT)
    else if (xlsxCFOperator = 'lessThanOrEqual') then
      _SetCFOperator(ZCFOpLTE)
    else if (xlsxCFOperator = 'between') then
      CFCondition := ZCFCellContentIsBetween
    else if (xlsxCFOperator = 'notBetween') then
      CFCondition := ZCFCellContentIsNotBetween
    else if (xlsxCFOperator = 'containsText') then
      CFCondition := ZCFContainsText
    else if (xlsxCFOperator = 'notContains') then
      CFCondition := ZCFNotContainsText
    else if (xlsxCFOperator = 'beginsWith') then
      CFCondition := ZCFBeginsWithText
    else if (xlsxCFOperator = 'endsWith') then
      CFCondition := ZCFEndsWithText
    else
      Result := false;
  end; // _CheckCFoperator

begin
  Result := false;
  CFCondition := ZCFNumberValue;
  CFOperator := ZCFOpGT;

  if (_CheckXLSXCfType(isCheckOperator)) then
  begin
    if (isCheckOperator) then
      Result := _CheckCFoperator()
    else
      Result := true;
  end;
end; // ZEXLSX_getCFCondition

function ZEXLSXReadDrawingRels(sheet: TZSheet; Stream: TStream): Boolean;
var
  Xml: TZsspXMLReaderH;
  target: string;
  i, id: Integer;
begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    while Xml.ReadToEndTagByName('Relationships') do
    begin
      if Xml.IsTagClosedByName('Relationship') then
      begin
        id := StrToInt(Xml.Attributes.ItemsByName['Id'].Substring(3));
        target := Xml.Attributes.ItemsByName['Target'];

        for i := 0 to sheet.Drawing.Count - 1 do
        begin
          if sheet.Drawing[i].RelId = id then
            sheet.Drawing[i].name := target.Substring(9);
        end;
      end;
    end;
    Result := true;
  finally
    Xml.Free();
  end;
end;

function ZEXLSXReadDrawing(sheet: TZSheet; Stream: TStream): Boolean;
var
  Xml: TZsspXMLReaderH;
  picture: TZEPicture;

  procedure ReadFrom();
  begin
    while Xml.ReadToEndTagByName('xdr:from') do
    begin
      if Xml.IsTagEndByName('xdr:col') then
        picture.FromCol := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:colOff') then
        picture.FromColOff := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:row') then
        picture.FromRow := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:rowOff') then
        picture.FromRowOff := StrToInt(Xml.TextBeforeTag);
    end;
  end;

  procedure ReadTo();
  begin
    while Xml.ReadToEndTagByName('xdr:to') do
    begin
      if Xml.IsTagEndByName('xdr:col') then
        picture.ToCol := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:colOff') then
        picture.ToColOff := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:row') then
        picture.ToRow := StrToInt(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('xdr:rowOff') then
        picture.ToRowOff := StrToInt(Xml.TextBeforeTag);
    end;
  end;

  procedure ReadPic();
  begin
    while Xml.ReadToEndTagByName('xdr:pic') do
    begin
      if Xml.IsTagClosedByName('xdr:cNvPr') then
      begin
        picture.Description := Xml.Attributes['descr'];
        picture.Title := Xml.Attributes['name'];
        picture.id := StrToInt(Xml.Attributes['id']);
      end; // else
      if Xml.IsTagStartOrClosedByName('a:blip') then
      begin
        picture.RelId := StrToInt(Xml.Attributes['r:embed'].Substring(3)); // skip "rId"
      end;

      if Xml.IsTagStartByName('a:xfrm') then
      begin
        while Xml.ReadToEndTagByName('a:xfrm') do
        begin
          if Xml.IsTagClosedByName('a:off') then
          begin
            picture.FrmOffX := StrToInt(Xml.Attributes['x']);
            picture.FrmOffY := StrToInt(Xml.Attributes['y']);
          end
          else if Xml.IsTagClosedByName('a:ext') then
          begin
            picture.FrmExtCX := StrToInt(Xml.Attributes['cx']);
            picture.FrmExtCY := StrToInt(Xml.Attributes['cy']);
          end
        end;
      end;
    end;
  end;

  procedure ReadImageItem();
  begin
    picture := sheet.Drawing.Add();
    if Xml.Attributes['editAs'] = 'absolute' then
      picture.CellAnchor := ZAAbsolute
    else
      picture.CellAnchor := ZACell;

    while Xml.ReadToEndTagByName('xdr:twoCellAnchor') do
    begin
      if Xml.IsTagStartByName('xdr:from') then
      begin
        ReadFrom();
      end; // else

      if Xml.IsTagStartByName('xdr:to') then
      begin
        ReadTo();
      end; // else

      if Xml.IsTagStartByName('xdr:pic') then
      begin
        ReadPic();
      end
    end;
  end;

begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;
    picture := nil;

    while Xml.ReadToEndTagByName('xdr:wsDr') do
    begin
      if Xml.IsTagStartByName('xdr:twoCellAnchor') then
      begin
        ReadImageItem();
      end;
    end;
  finally
    Xml.Free();
  end;
end;

// Читает страницу документа
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// var Stream: TStream                 - поток для чтения
// const SheetName: string             - название страницы
// var StrArray: TStringDynArray       - строки для подстановки
// StrCount: integer               - кол-во строк подстановки
// var Relations: TZXLSXRelationsArray - отношения
// RelationsCount: integer         - кол-во отношений
// var MaximumDigitWidth: double       - ширина самого широкого числа в пикселях
// ReadHelper: TZEXLSXReadHelper   -
// RETURN
// boolean - true - страница прочиталась успешно
function ZEXSLXReadSheet(var XMLSS: TZWorkBook; var Stream: TStream; const SheetName: string;
  var StrArray: TStringDynArray; StrCount: Integer; var Relations: TZXLSXRelationsArray; RelationsCount: Integer;
  MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): Boolean;
var
  Xml: TZsspXMLReaderH;
  currentPage: Integer;
  currentRow: Integer;
  currentCol: Integer;
  currentSheet: TZSheet;
  currentCell: TZCell;
  str: string;
  tempReal: real;
  tempInt: Integer;
  tempDate: TDateTime;
  tempFloat: double;

  // Проверить кол-во строк
  procedure CheckRow(const RowCount: Integer);
  begin
    if (currentSheet.RowCount < RowCount) then
      currentSheet.RowCount := RowCount;
  end;

  // Проверить кол-во столбцов
  procedure CheckCol(const ColCount: Integer);
  begin
    if (currentSheet.ColCount < ColCount) then
      currentSheet.ColCount := ColCount
  end;

  // Чтение строк/столбцов
  procedure _ReadSheetData();
  var
    t: Integer;
    v: string;
    _num: Integer;
    _type: string;
    _cr, _cc: Integer;
    maxCol: Integer;
  begin
    _cr := 0;
    _cc := 0;
    maxCol := 0;
    CheckRow(1);
    CheckCol(1);
    while Xml.ReadToEndTagByName('sheetData') do
    begin
      // ячейка
      if (Xml.TagName = 'c') then
      begin
        str := Xml.Attributes.ItemsByName['r']; // номер
        if (str > '') then
          if (ZEGetCellCoords(str, _cc, _cr)) then
          begin
            currentCol := _cc;
            CheckCol(_cc + 1);
          end;

        _type := Xml.Attributes.ItemsByName['t']; // тип

        // s := xml.Attributes.ItemsByName['cm'];
        // s := xml.Attributes.ItemsByName['ph'];
        // s := xml.Attributes.ItemsByName['vm'];
        v := '';
        _num := 0;
        currentCell := currentSheet.Cell[currentCol, currentRow];
        str := Xml.Attributes.ItemsByName['s']; // стиль
        if (str > '') then
          if (TryStrToInt(str, t)) then
            currentCell.CellStyle := t;
        if (Xml.IsTagStart) then
          while Xml.ReadToEndTagByName('c') do
          begin
          // is пока игнорируем
            if Xml.IsTagEndByName('v') or Xml.IsTagEndByName('t') then
            begin
              if (_num > 0) then
                v := v + sLineBreak;
              v := v + Xml.TextBeforeTag;
              inc(_num);
            end
            else if Xml.IsTagEndByName('f') then
              currentCell.Formula := ZEReplaceEntity(Xml.TextBeforeTag);

          end; // while

        // Возможные типы:
        // s - sharedstring
        // b - boolean
        // n - number
        // e - error
        // str - string
        // inlineStr - inline string ??
        // d - date
        // тип может отсутствовать. Интерпретируем в таком случае как ZEGeneral
        if (_type = '') then
          currentCell.CellType := ZEGeneral
        else if (_type = 'n') then
        begin
          currentCell.CellType := ZENumber;
          // Trouble: if cell style is number, and number format is date, then
          // cell style is date. F****** m$!
          if (ReadHelper.NumberFormats.IsDateFormat(currentCell.CellStyle)) then
            if (ZEIsTryStrToFloat(v, tempFloat)) then
            begin
              currentCell.CellType := ZEDateTime;
              v := ZEDateTimeToStr(tempFloat);
            end;
        end
        else if (_type = 's') then
        begin
          currentCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < StrCount)) then
              v := StrArray[t];
        end
        else if (_type = 'd') then
        begin
          currentCell.CellType := ZEDateTime;
          if (TryZEStrToDateTime(v, tempDate)) then
            v := ZEDateTimeToStr(tempDate)
          else if (ZEIsTryStrToFloat(v, tempFloat)) then
            v := ZEDateTimeToStr(tempFloat)
          else
            currentCell.CellType := ZEString;
        end;

        currentCell.Data := ZEReplaceEntity(v);
        inc(currentCol);
        CheckCol(currentCol + 1);
        if currentCol > maxCol then
          maxCol := currentCol;
      end
      else
      // строка
        if Xml.IsTagStartOrClosedByName('row') then
        begin
          currentCol := 0;
          str := Xml.Attributes.ItemsByName['r']; // индекс строки
          if (str > '') then
            if (TryStrToInt(str, t)) then
            begin
              currentRow := t - 1;
              CheckRow(t);
            end;
        // s := xml.Attributes.ItemsByName['collapsed'];
        // s := xml.Attributes.ItemsByName['customFormat'];
        // s := xml.Attributes.ItemsByName['customHeight'];
          currentSheet.Rows[currentRow].Hidden := ZETryStrToBoolean(Xml.Attributes.ItemsByName['hidden'], false);

          str := Xml.Attributes.ItemsByName['ht']; // в поинтах
          if (str > '') then
          begin
            tempReal := ZETryStrToFloat(str, 10);
            currentSheet.Rows[currentRow].Height := tempReal;
          // tempReal := tempReal / 2.835; //???
          // currentSheet.Rows[currentRow].HeightMM := tempReal;
          end
          else
            currentSheet.Rows[currentRow].Height := currentSheet.DefaultRowHeight;

          str := Xml.Attributes.ItemsByName['outlineLevel'];
          currentSheet.Rows[currentRow].OutlineLevel := StrToIntDef(str, 0);

        // s := xml.Attributes.ItemsByName['ph'];

          str := Xml.Attributes.ItemsByName['s']; // номер стиля
          if (str > '') then
            if (TryStrToInt(str, t)) then
            begin
            // нужно подставить нужный стиль
            end;
        // s := xml.Attributes.ItemsByName['spans'];
        // s := xml.Attributes.ItemsByName['thickBot'];
        // s := xml.Attributes.ItemsByName['thickTop'];

          if Xml.IsTagClosed then
          begin
            inc(currentRow);
            CheckRow(currentRow + 1);
          end;
        end
        else
      // конец строки
          if Xml.IsTagEndByName('row') then
          begin
            inc(currentRow);
            CheckRow(currentRow + 1);
          end;
    end; // while
    currentSheet.ColCount := maxCol;
  end; // _ReadSheetData

  // Чтение диапазона ячеек с автофильтром
  procedure _ReadAutoFilter();
  begin
    currentSheet.AutoFilter := Xml.Attributes.ItemsByName['ref'];
  end;

  // Чтение объединённых ячеек
  procedure _ReadMerge();
  var
    i, t, Num: Integer;
    x1, x2, y1, y2: Integer;
    s1, s2: string;
    b: Boolean;
    function _GetCoords(var x, y: Integer): Boolean;
    begin
      Result := true;
      x := ZEGetColByA1(s1);
      if (x < 0) then
        Result := false;
      if (not TryStrToInt(s2, y)) then
        Result := false
      else
        dec(y);
      b := Result;
    end; // _GetCoords

  begin
    x1 := 0;
    y1 := 0;
    x2 := 0;
    y2 := 0;
    while Xml.ReadToEndTagByName('mergeCells') do
    begin
      if Xml.IsTagStartOrClosedByName('mergeCell') then
      begin
        str := Xml.Attributes.ItemsByName['ref'];
        t := length(str);
        if (t > 0) then
        begin
          str := str + ':';
          s1 := '';
          s2 := '';
          b := true;
          Num := 0;
          for i := 1 to t + 1 do
            case str[i] of
              'A' .. 'Z', 'a' .. 'z':
                s1 := s1 + str[i];
              '0' .. '9':
                s2 := s2 + str[i];
              ':':
                begin
                  inc(Num);
                  if (Num > 2) then
                  begin
                    b := false;
                    break;
                  end;
                  if (Num = 1) then
                  begin
                    if (not _GetCoords(x1, y1)) then
                      break;
                  end
                  else
                  begin
                    if (not _GetCoords(x2, y2)) then
                      break;
                  end;
                  s1 := '';
                  s2 := '';
                end;
            else
              begin
                b := false;
                break;
              end;
            end; // case

          if (b) then
          begin
            CheckRow(y1 + 1);
            CheckRow(y2 + 1);
            CheckCol(x1 + 1);
            CheckCol(x2 + 1);
            currentSheet.MergeCells.AddRectXY(x1, y1, x2, y2);
          end;
        end; // if
      end; // if
    end; // while
  end; // _ReadMerge

  // Столбцы
  procedure _ReadCols();
  type
    TZColInf = record
      min, Max: Integer;
      bestFit, Hidden: Boolean;
      OutlineLevel: Integer;
      width: Integer;
    end;
  var
    i, j: Integer;
    t: real;
    colInf: TArray<TZColInf>;
  const
    MAX_COL_DIFF = 500;
  begin
    i := 0;
    while Xml.ReadToEndTagByName('cols') do
    begin
      if (Xml.TagName = 'col') and Xml.IsTagStartOrClosed then
      begin
        SetLength(colInf, i + 1);

        colInf[i].min := StrToIntDef(Xml.Attributes.ItemsByName['min'], 0);
        colInf[i].Max := StrToIntDef(Xml.Attributes.ItemsByName['max'], 0);
        // защита от сплошного диапазона
        // когда значение _мах = 16384
        // но чтобы уж наверняка, проверим на MAX_COL_DIFF колонок подряд.
        if (colInf[i].Max - colInf[i].min) > MAX_COL_DIFF then
          colInf[i].Max := colInf[i].min + MAX_COL_DIFF;

        colInf[i].OutlineLevel := StrToIntDef(Xml.Attributes.ItemsByName['outlineLevel'], 0);
        str := Xml.Attributes.ItemsByName['hidden'];
        if (str > '') then
          colInf[i].Hidden := ZETryStrToBoolean(str);
        str := Xml.Attributes.ItemsByName['bestFit'];
        if (str > '') then
          colInf[i].bestFit := ZETryStrToBoolean(str);

        str := Xml.Attributes.ItemsByName['width'];
        if (str > '') then
        begin
          t := ZETryStrToFloat(str, 5.14509803921569);
          // t := 10 * t / 5.14509803921569;
          // А.А.Валуев. Формулы расёта ширины взяты здесь - https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
          t := Trunc(((256 * t + Trunc(128 / MaximumDigitWidth)) / 256) * MaximumDigitWidth);
          colInf[i].width := Trunc(t);
        end;

        inc(i);
      end; // if
    end; // while

    for i := Low(colInf) to High(colInf) do
    begin
      for j := colInf[i].min to colInf[i].Max do
      begin
        CheckCol(j);
        currentSheet.Columns[j - 1].AutoFitWidth := colInf[i].bestFit;
        currentSheet.Columns[j - 1].Hidden := colInf[i].Hidden;
        currentSheet.Columns[j - 1].WidthPix := colInf[i].width;
      end;
    end;
  end; // _ReadCols

  function _StrToMM(const st: string; var retFloat: real): Boolean;
  begin
    Result := false;
    if (str > '') then
    begin
      retFloat := ZETryStrToFloat(st, -1);
      if (retFloat > -1) then
      begin
        Result := true;
        retFloat := retFloat * ZE_MMinInch;
      end;
    end;
  end; // _StrToMM

  procedure _GetDimension();
  var
    st, s: string;
    i, l, _maxC, _maxR, c, r: Integer;
  begin
    c := 0;
    r := 0;
    st := Xml.Attributes.ItemsByName['ref'];
    l := length(st);
    if (l > 0) then
    begin
      st := st + ':';
      inc(l);
      s := '';
      _maxC := -1;
      _maxR := -1;
      for i := 1 to l do
        if (st[i] = ':') then
        begin
          if (ZEGetCellCoords(s, c, r, true)) then
          begin;
            if (c > _maxC) then
              _maxC := c;
            if (r > _maxR) then
              _maxR := r;
          end
          else
            break;
          s := '';
        end
        else
          s := s + st[i];
      if (_maxC > 0) then
        CheckCol(_maxC);
      if (_maxR > 0) then
        CheckRow(_maxR);
    end;
  end; // _GetDimension()

  // Чтение ссылок
  procedure _ReadHyperLinks();
  var
    _c, _r, i: Integer;
  begin
    _c := 0;
    _r := 0;
    while Xml.ReadToEndTagByName('hyperlinks') do
    begin
      if Xml.IsTagClosedByName('hyperlink') then
      begin
        str := Xml.Attributes.ItemsByName['ref'];
        if (str > '') then
          if (ZEGetCellCoords(str, _c, _r, true)) then
          begin
            CheckRow(_r);
            CheckCol(_c);
            currentSheet.Cell[_c, _r].HRefScreenTip := Xml.Attributes.ItemsByName['tooltip'];
            str := Xml.Attributes.ItemsByName['r:id'];
            // по r:id подставить ссылку
            for i := 0 to RelationsCount - 1 do
              if ((Relations[i].id = str) and (Relations[i].ftype = TRelationType.rtHyperlink)) then
              begin
                currentSheet.Cell[_c, _r].Href := Relations[i].target;
                break;
              end;
          end;
        // доп. атрибуты:
        // display - ??
        // id - id <> r:id??
        // location - ??
      end;
    end; // while
  end; // _ReadHyperLinks();

  procedure _ReadSheetPr();
  begin
    while Xml.ReadToEndTagByName('sheetPr') do
    begin
      if Xml.TagName = 'tabColor' then
        currentSheet.TabColor := ARGBToColor(Xml.Attributes.ItemsByName['rgb']);

      if Xml.TagName = 'pageSetUpPr' then
        currentSheet.FitToPage := ZEStrToBoolean(Xml.Attributes.ItemsByName['fitToPage']);

      if Xml.TagName = 'outlinePr' then
      begin
        currentSheet.ApplyStyles := ZEStrToBoolean(Xml.Attributes.ItemsByName['applyStyles']);
        currentSheet.SummaryBelow := Xml.Attributes.ItemsByName['summaryBelow'] <> '0';
        currentSheet.SummaryRight := Xml.Attributes.ItemsByName['summaryRight'] <> '0';
      end;
    end;
  end; // _ReadSheetPr();

  procedure _ReadRowBreaks();
  begin
    currentSheet.RowBreaks := [];
    while Xml.ReadToEndTagByName('rowBreaks') do
    begin
      if Xml.TagName = 'brk' then
        currentSheet.RowBreaks := currentSheet.RowBreaks + [StrToIntDef(Xml.Attributes.ItemsByName['id'], 0)];
    end;
  end;

  procedure _ReadColBreaks();
  begin
    currentSheet.ColBreaks := [];
    while Xml.ReadToEndTagByName('colBreaks') do
    begin
      if Xml.TagName = 'brk' then
        currentSheet.ColBreaks := currentSheet.ColBreaks + [StrToIntDef(Xml.Attributes.ItemsByName['id'], 0)];
    end;
  end;

  // <sheetViews> ... </sheetViews>
  procedure _ReadSheetViews();
  var
    vValue, hValue: Integer;
    SplitMode: TZSplitMode;
    s: string;
  begin
    while Xml.ReadToEndTagByName('sheetViews') do
    begin
      if Xml.IsTagStartByName('sheetView') or Xml.IsTagClosedByName('sheetView') then
      begin
        s := Xml.Attributes.ItemsByName['tabSelected'];
        // тут кроется проблема с выделением нескольких листов
        currentSheet.Selected := currentSheet.SheetIndex = 0; // s = '1';
        currentSheet.ViewMode := zvmNormal;
        if Xml.Attributes.ItemsByName['view'] = 'pageBreakPreview' then
          currentSheet.ViewMode := zvmPageBreakPreview;
      end;

      if Xml.IsTagClosedByName('pane') then
      begin
        SplitMode := ZSplitSplit;
        s := Xml.Attributes.ItemsByName['state'];
        if (s = 'frozen') then
          SplitMode := ZSplitFrozen;

        s := Xml.Attributes.ItemsByName['xSplit'];
        if (not TryStrToInt(s, vValue)) then
          vValue := 0;

        s := Xml.Attributes.ItemsByName['ySplit'];
        if (not TryStrToInt(s, hValue)) then
          hValue := 0;

        currentSheet.SheetOptions.SplitVerticalValue := vValue;
        currentSheet.SheetOptions.SplitHorizontalValue := hValue;

        currentSheet.SheetOptions.SplitHorizontalMode := ZSplitNone;
        currentSheet.SheetOptions.SplitVerticalMode := ZSplitNone;
        if (hValue <> 0) then
          currentSheet.SheetOptions.SplitHorizontalMode := SplitMode;
        if (vValue <> 0) then
          currentSheet.SheetOptions.SplitVerticalMode := SplitMode;

        if (currentSheet.SheetOptions.SplitHorizontalMode = ZSplitSplit) then
          currentSheet.SheetOptions.SplitHorizontalValue := PointToPixel(hValue / 20);
        if (currentSheet.SheetOptions.SplitVerticalMode = ZSplitSplit) then
          currentSheet.SheetOptions.SplitVerticalValue := PointToPixel(vValue / 20);

      end; // if
    end; // while
  end; // _ReadSheetViews()

  procedure _ReadConditionFormatting();
  var
    MaxFormulasCount: Integer;
    _formulas: array of string;
    Count: Integer;
    _sqref: string;
    _type: string;
    _operator: string;
    _CFCondition: TZCondition;
    _CFOperator: TZConditionalOperator;
    _Style: string;
    _text: string;
    _isCFAdded: Boolean;
    _isOk: Boolean;
    // _priority: string;
    _CF: TZConditionalStyle;
    _tmpStyle: TZStyle;

    function _AddCF(): Boolean;
    var
      s, ss: string;
      _len, i, kol: Integer;
      a: array of array [0 .. 5] of Integer;
      _maxx: Integer;
      ch: char;
      w, h: Integer;

      function _GetOneArea(st: string): Boolean;
      var
        i, j: Integer;
        s: string;
        ch: char;
        _cnt: Integer;
        tmpArr: array [0 .. 1, 0 .. 1] of Integer;
        _isOk: Boolean;
        t: Integer;
        tmpB: Boolean;

      begin
        Result := false;
        if (st <> '') then
        begin
          st := st + ':';
          s := '';
          _cnt := 0;
          _isOk := true;
          for i := 1 to length(st) do
          begin
            ch := st[i];
            if (ch = ':') then
            begin
              if (_cnt < 2) then
              begin
                tmpB := ZEGetCellCoords(s, tmpArr[_cnt][0], tmpArr[_cnt][1]);
                _isOk := _isOk and tmpB;
              end;
              s := '';
              inc(_cnt);
            end
            else
              s := s + ch;
          end; // for

          if (_isOk) then
            if (_cnt > 0) then
            begin
              if (_cnt > 2) then
                _cnt := 2;

              a[kol][0] := _cnt;
              t := 1;
              for i := 0 to _cnt - 1 do
                for j := 0 to 1 do
                begin
                  a[kol][t] := tmpArr[i][j];
                  inc(t);
                end;
              Result := true;
            end;
        end; // if
      end; // _GetOneArea

    begin
      Result := false;
      if (_sqref <> '') then
        try
          _maxx := 4;
          SetLength(a, _maxx);
          ss := _sqref + ' ';
          _len := length(ss);
          kol := 0;
          s := '';
          for i := 1 to _len do
          begin
            ch := ss[i];
            if (ch = ' ') then
            begin
              if (_GetOneArea(s)) then
              begin
                inc(kol);
                if (kol >= _maxx) then
                begin
                  inc(_maxx, 4);
                  SetLength(a, _maxx);
                end;
              end;
              s := '';
            end
            else
              s := s + ch;
          end; // for

          if (kol > 0) then
          begin
            currentSheet.ConditionalFormatting.Add();
            _CF := currentSheet.ConditionalFormatting[currentSheet.ConditionalFormatting.Count - 1];
            for i := 0 to kol - 1 do
            begin
              w := 1;
              h := 1;
              if (a[i][0] >= 2) then
              begin
                w := abs(a[i][3] - a[i][1]) + 1;
                h := abs(a[i][4] - a[i][2]) + 1;
              end;
              _CF.Areas.Add(a[i][1], a[i][2], w, h);
            end;
            Result := true;
          end;
        finally
          SetLength(a, 0);
        end;
    end; // _AddCF

    // Применяем условный стиль
    procedure _TryApplyCF();
    var
      b: Boolean;
      Num: Integer;
      _id: Integer;
      procedure _CheckTextCondition();
      begin
        if (Count = 1) then
          if (_formulas[0] <> '') then
            _isOk := true;
      end;

      // Найти стиль
      // пока будем делать так: предполагаем, что все ячейки в текущей области
      // условного форматирования имеют один стиль. Берём стиль из левой верхней
      // ячейки, клонируем его, применяем дифф. стиль, добавляем в хранилище стилей
      // с учётом повторов.
      // TODO: потом нужно будет переделать
      // INPUT
      // dfNum: integer - номер дифференцированного форматирования
      // RETURN
      // integer - номер применяемого стиля
      function _getStyleIdxForDF(dfNum: Integer): Integer;
      var
        _df: TZXLSXDiffFormattingItem;
        _r, _c: Integer;
        _t: Integer;
        i: TZBordersPos;
      begin
        // _currSheet
        Result := -1;
        if ((dfNum >= 0) and (dfNum < ReadHelper.DiffFormatting.Count)) then
        begin
          _df := ReadHelper.DiffFormatting[dfNum];
          _t := -1;

          if (_CF.Areas.Count > 0) then
          begin
            _r := _CF.Areas.Items[0].Row;
            _c := _CF.Areas.Items[0].Column;
            if ((_r >= 0) and (_r < currentSheet.RowCount)) then
              if ((_c >= 0) and (_c < currentSheet.ColCount)) then
                _t := currentSheet.Cell[_c, _r].CellStyle;
          end;

          _tmpStyle.Assign(XMLSS.Styles[_t]);

          if (_df.UseFont) then
          begin
            if (_df.UseFontStyles) then
              _tmpStyle.Font.Style := _df.FontStyles;
            if (_df.UseFontColor) then
              _tmpStyle.Font.Color := _df.FontColor;
          end;
          if (_df.UseFill) then
          begin
            if (_df.UseCellPattern) then
              _tmpStyle.CellPattern := _df.CellPattern;
            if (_df.UseBGColor) then
              _tmpStyle.BGColor := _df.BGColor;
            if (_df.UsePatternColor) then
              _tmpStyle.PatternColor := _df.PatternColor;
          end;
          if (_df.UseBorder) then
            for i := bpLeft to bpDiagonalRight do
            begin
              if (_df.Borders[i].UseStyle) then
              begin
                _tmpStyle.Border[i].Weight := _df.Borders[i].Weight;
                _tmpStyle.Border[i].LineStyle := _df.Borders[i].LineStyle;
              end;
              if (_df.Borders[i].UseColor) then
                _tmpStyle.Border[i].Color := _df.Borders[i].Color;
            end; // for

          Result := XMLSS.Styles.Add(_tmpStyle, true);
        end; // if
      end; // _getStyleIdxForDF

    begin
      _isOk := false;
      case (_CFCondition) of
        ZCFIsTrueFormula:
          ;
        ZCFCellContentIsBetween, ZCFCellContentIsNotBetween:
          begin
            // только числа
            if (Count = 2) then
            begin
              ZETryStrToFloat(_formulas[0], b);
              if (b) then
                ZETryStrToFloat(_formulas[1], _isOk);
            end;
          end;
        ZCFCellContentOperator:
          begin
            // только числа
            if (Count = 1) then
              ZETryStrToFloat(_formulas[0], _isOk);
          end;
        ZCFNumberValue:
          ;
        ZCFString:
          ;
        ZCFBoolTrue:
          ;
        ZCFBoolFalse:
          ;
        ZCFFormula:
          ;
        ZCFContainsText:
          _CheckTextCondition();
        ZCFNotContainsText:
          _CheckTextCondition();
        ZCFBeginsWithText:
          _CheckTextCondition();
        ZCFEndsWithText:
          _CheckTextCondition();
      end; // case

      if (_isOk) then
      begin
        if (not _isCFAdded) then
          _isCFAdded := _AddCF();

        if ((_isCFAdded) and (Assigned(_CF))) then
        begin
          Num := _CF.Count;
          _CF.Add();
          if (_Style <> '') then
            if (TryStrToInt(_Style, _id)) then
              _CF[Num].ApplyStyleID := _getStyleIdxForDF(_id);
          _CF[Num].Condition := _CFCondition;
          _CF[Num].ConditionOperator := _CFOperator;

          _CF[Num].Value1 := _formulas[0];
          if (Count >= 2) then
            _CF[Num].Value2 := _formulas[1];
        end;
      end;
    end; // _TryApplyCF

  begin
    try
      _sqref := Xml.Attributes['sqref'];
      MaxFormulasCount := 2;
      SetLength(_formulas, MaxFormulasCount);
      _isCFAdded := false;
      _CF := nil;
      _tmpStyle := TZStyle.Create();
      while Xml.ReadToEndTagByName('conditionalFormatting') do
      begin
        // cfRule = Conditional Formatting Rule
        if Xml.IsTagStartByName('cfRule') then
        begin
         (*
          Атрибуты в cfRule:
          type	       	- тип
                            expression        - ??
                            cellIs            -
                            colorScale        - ??
                            dataBar           - ??
                            iconSet           - ??
                            top10             - ??
                            uniqueValues      - ??
                            duplicateValues   - ??
                            containsText      -    ?
                            notContainsText   -    ?
                            beginsWith        -    ?
                            endsWith          -    ?
                            containsBlanks    - ??
                            notContainsBlanks - ??
                            containsErrors    - ??
                            notContainsErrors - ??
                            timePeriod        - ??
                            aboveAverage      - ?
          dxfId	        - ID применяемого формата
          priority	    - приоритет
          stopIfTrue	  -  ??
          aboveAverage  -  ??
          percent	      -  ??
          bottom	      -  ??
          operator	    - оператор:
                              lessThan	          <
                              lessThanOrEqual	    <=
                              equal	              =
                              notEqual	          <>
                              greaterThanOrEqual  >=
                              greaterThan	        >
                              between	            Between
                              notBetween	        Not Between
                              containsText	      содержит текст
                              notContains	        не содержит
                              beginsWith	        начинается с
                              endsWith	          оканчивается на
          text	        -  ??
          timePeriod	  -  ??
          rank	        -  ??
          stdDev  	    -  ??
          equalAverage	-  ??
 *)
          _type := Xml.Attributes['type'];
          _operator := Xml.Attributes['operator'];
          _Style := Xml.Attributes['dxfId'];
          _text := ZEReplaceEntity(Xml.Attributes['text']);
          // _priority := xml.Attributes['priority'];

          Count := 0;
          while Xml.ReadToEndTagByName('cfRule') do
          begin
            if Xml.IsTagEndByName('formula') then
            begin
              if (Count >= MaxFormulasCount) then
              begin
                inc(MaxFormulasCount, 2);
                SetLength(_formulas, MaxFormulasCount);
              end;
              _formulas[Count] := ZEReplaceEntity(Xml.TextBeforeTag);
              inc(Count);
            end;
          end; // while

          if (ZEXLSX_getCFCondition(_type, _operator, _CFCondition, _CFOperator)) then
            _TryApplyCF();
        end; // if
      end; // while
    finally
      SetLength(_formulas, 0);
      FreeAndNil(_tmpStyle);
    end;
  end; // _ReadConditionFormatting

  procedure _ReadHeaderFooter();
  begin
    currentSheet.SheetOptions.IsDifferentFirst := ZEStrToBoolean(Xml.Attributes['differentFirst']);
    currentSheet.SheetOptions.IsDifferentOddEven := ZEStrToBoolean(Xml.Attributes['differentOddEven']);
    while Xml.ReadToEndTagByName('headerFooter') do
    begin
      if Xml.IsTagEndByName('oddHeader') then
        currentSheet.SheetOptions.Header := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('oddFooter') then
        currentSheet.SheetOptions.Footer := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('evenHeader') then
        currentSheet.SheetOptions.EvenHeader := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('evenFooter') then
        currentSheet.SheetOptions.EvenFooter := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('firstHeader') then
        currentSheet.SheetOptions.FirstPageHeader := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('firstFooter') then
        currentSheet.SheetOptions.FirstPageFooter := ClenuapXmlTagValue(Xml.TextBeforeTag);
    end;
  end;

begin
  Xml := TZsspXMLReaderH.Create();
  Result := false;
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    currentPage := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := XMLSS.Sheets.Count + 1;
    currentRow := 0;
    currentSheet := XMLSS.Sheets[currentPage];
    currentSheet.Title := SheetName;

    while Xml.ReadTag() do
    begin
      if Xml.IsTagStartByName('sheetData') then
        _ReadSheetData()
      else if Xml.IsTagClosedByName('autoFilter') then
        _ReadAutoFilter()
      else if Xml.IsTagStartByName('mergeCells') then
        _ReadMerge()
      else if Xml.IsTagStartByName('cols') then
        _ReadCols()
      else if Xml.IsTagClosedByName('drawing') then
      begin
        currentSheet.DrawingRid := StrToIntDef(Xml.Attributes.ItemsByName['r:id'].Substring(3), 0);
      end
      else if Xml.IsTagClosedByName('pageMargins') then
      begin
        str := Xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginBottom := round(tempReal);
        str := Xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.FooterMargins.Height := abs(round(tempReal));
        str := Xml.Attributes.ItemsByName['header'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.HeaderMargins.Height := abs(round(tempReal));
        str := Xml.Attributes.ItemsByName['left'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginLeft := round(tempReal);
        str := Xml.Attributes.ItemsByName['right'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginRight := round(tempReal);
        str := Xml.Attributes.ItemsByName['top'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginTop := round(tempReal);
      end
      else
      // Настройки страницы
        if Xml.IsTagClosedByName('pageSetup') then
        begin
        // str := xml.Attributes.ItemsByName['blackAndWhite'];
        // str := xml.Attributes.ItemsByName['cellComments'];
        // str := xml.Attributes.ItemsByName['copies'];
        // str := xml.Attributes.ItemsByName['draft'];
        // str := xml.Attributes.ItemsByName['errors'];
          str := Xml.Attributes.ItemsByName['firstPageNumber'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.StartPageNumber := tempInt;

          str := Xml.Attributes.ItemsByName['fitToHeight'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.FitToHeight := tempInt;

          str := Xml.Attributes.ItemsByName['fitToWidth'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.FitToWidth := tempInt;

        // str := xml.Attributes.ItemsByName['horizontalDpi'];
        // str := xml.Attributes.ItemsByName['id'];
          str := Xml.Attributes.ItemsByName['orientation'];
          if (str > '') then
          begin
            currentSheet.SheetOptions.PortraitOrientation := false;
            if (str = 'portrait') then
              currentSheet.SheetOptions.PortraitOrientation := true;
          end;

        // str := xml.Attributes.ItemsByName['pageOrder'];

          str := Xml.Attributes.ItemsByName['paperSize'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.PaperSize := tempInt;
        // str := xml.Attributes.ItemsByName['paperHeight']; //если утановлены paperHeight и Width, то paperSize игнорируется
        // str := xml.Attributes.ItemsByName['paperWidth'];

          str := Xml.Attributes.ItemsByName['scale'];
          currentSheet.SheetOptions.ScaleToPercent := StrToIntDef(str, 100);
        // str := xml.Attributes.ItemsByName['useFirstPageNumber'];
        // str := xml.Attributes.ItemsByName['usePrinterDefaults'];
        // str := xml.Attributes.ItemsByName['verticalDpi'];
        end
        else
      // настройки печати
          if Xml.IsTagClosedByName('printOptions') then
          begin
        // str := xml.Attributes.ItemsByName['gridLines'];
        // str := xml.Attributes.ItemsByName['gridLinesSet'];
        // str := xml.Attributes.ItemsByName['headings'];
            str := Xml.Attributes.ItemsByName['horizontalCentered'];
            if (str > '') then
              currentSheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(str);

            str := Xml.Attributes.ItemsByName['verticalCentered'];
            if (str > '') then
              currentSheet.SheetOptions.CenterVertical := ZEStrToBoolean(str);
          end
          else if Xml.IsTagClosedByName('sheetFormatPr') then
          begin
            str := Xml.Attributes.ItemsByName['defaultColWidth'];
            if (str > '') then
              currentSheet.DefaultColWidth := ZETryStrToFloat(str, currentSheet.DefaultColWidth);
            str := Xml.Attributes.ItemsByName['defaultRowHeight'];
            if (str > '') then
              currentSheet.DefaultRowHeight := ZETryStrToFloat(str, currentSheet.DefaultRowHeight);
          end
          else if Xml.IsTagClosedByName('dimension') then
            _GetDimension()
          else if Xml.IsTagStartByName('hyperlinks') then
            _ReadHyperLinks()
          else if Xml.IsTagStartByName('sheetPr') then
            _ReadSheetPr()
          else if Xml.IsTagStartByName('rowBreaks') then
            _ReadRowBreaks()
          else if Xml.IsTagStartByName('colBreaks') then
            _ReadColBreaks()
          else if Xml.IsTagStartByName('sheetViews') then
            _ReadSheetViews()
          else if Xml.IsTagStartByName('conditionalFormatting') then
            _ReadConditionFormatting()
          else if Xml.IsTagStartByName('headerFooter') then
            _ReadHeaderFooter();
    end; // while

    Result := true;
  finally
    Xml.Free();
  end;
end; // ZEXSLXReadSheet

// Прочитать стили из потока (styles.xml)
// INPUT
// var XMLSS: TZWorkBook                    - хранилище
// var Stream: TStream                    - поток
// var ThemaFillsColors: TIntegerDynArray - цвета из темы
// var ThemaColorCount: integer           - кол-во цветов заливки в теме
// var MaximumDigitWidth: double          - ширина самого широкого числа в пикселях
// ReadHelper: TZEXLSXReadHelper      -
// RETURN
// boolean - true - стили прочитались без ошибок
function ZEXSLXReadStyles(var XMLSS: TZWorkBook; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray;
  var ThemaColorCount: Integer; var MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): Boolean;
type
  TZXLSXBorderItem = record
    Color: TColor;
    isColor: Boolean;
    isEnabled: Boolean;
    Style: TZBorderType;
    Weight: Byte;
  end;

  // 0 - left           левая граница
  // 1 - Top            верхняя граница
  // 2 - Right          правая граница
  // 3 - Bottom         нижняя граница
  // 4 - DiagonalLeft   диагональ от верхнего левого угла до нижнего правого
  // 5 - DiagonalRight  диагональ от нижнего левого угла до правого верхнего
  TZXLSXBorder = array [0 .. 5] of TZXLSXBorderItem;
  TZXLSXBordersArray = array of TZXLSXBorder;

  TZXLSXCellAlignment = record
    horizontal: TZHorizontalAlignment;
    indent: Integer;
    shrinkToFit: Boolean;
    textRotation: Integer;
    vertical: TZVerticalAlignment;
    wrapText: Boolean;
  end;

  TZXLSXCellStyle = record
    applyAlignment: Boolean;
    applyBorder: Boolean;
    applyFont: Boolean;
    applyProtection: Boolean;
    borderId: Integer;
    fillId: Integer;
    fontId: Integer;
    numFmtId: Integer;
    xfId: Integer;
    Hidden: Boolean;
    locked: Boolean;
    alignment: TZXLSXCellAlignment;
  end;

  TZXLSXCellStylesArray = array of TZXLSXCellStyle;

type
  TZXLSXStyle = record
    builtinId: Integer; // ??
    customBuiltin: Boolean; // ??
    name: string; // ??
    xfId: Integer;
  end;

  TZXLSXStyleArray = array of TZXLSXStyle;

  TZXLSXFill = record
    patternfill: TZCellPattern;
    bgColorType: Byte; // 0 - rgb, 1 - indexed, 2 - theme
    BGColor: TColor;
    PatternColor: TColor;
    patternColorType: Byte;
    lumFactorBG: double;
    lumFactorPattern: double;
  end;

  TZXLSXFillArray = array of TZXLSXFill;

  TZXLSXDFFont = record
    Color: TColor;
    ColorType: Byte;
    LumFactor: double;
  end;

  TZXLSXDFFontArray = array of TZXLSXDFFont;

var
  Xml: TZsspXMLReaderH;
  s: string;
  FontArray: TZEXLSXFontArray;
  FontCount: Integer;
  BorderArray: TZXLSXBordersArray;
  BorderCount: Integer;
  CellXfsArray: TZXLSXCellStylesArray;
  CellXfsCount: Integer;
  CellStyleArray: TZXLSXCellStylesArray;
  CellStyleCount: Integer;
  StyleArray: TZXLSXStyleArray;
  StyleCount: Integer;
  FillArray: TZXLSXFillArray;
  FillCount: Integer;
  indexedColor: TIntegerDynArray;
  indexedColorCount: Integer;
  indexedColorMax: Integer;
  _Style: TZStyle;
  t, i, n: Integer;
  h1, s1, l1: double;
  _dfFonts: TZXLSXDFFontArray;
  _dfFills: TZXLSXFillArray;

  // Приводит к шрифту по-умолчанию
  // INPUT
  // var fnt: TZEXLSXFont - шрифт
  procedure ZEXLSXZeroFont(var fnt: TZEXLSXFont);
  begin
    fnt.name := 'Arial';
    fnt.bold := false;
    fnt.italic := false;
    fnt.underline := false;
    fnt.strike := false;
    fnt.charset := 204;
    fnt.Color := TColorRec.Black;
    fnt.LumFactor := 0;
    fnt.ColorType := 0;
    fnt.fontsize := 8;
    fnt.superscript := false;
    fnt.subscript := false;
  end; // ZEXLSXZeroFont

  // Обнуляет границы
  // var border: TZXLSXBorder - границы
  procedure ZEXLSXZeroBorder(var Border: TZXLSXBorder);
  var
    i: Integer;
  begin
    for i := 0 to 5 do
    begin
      Border[i].isColor := false;
      Border[i].isEnabled := false;
      Border[i].Style := ZENone;
      Border[i].Weight := 0;
    end;
  end; // ZEXLSXZeroBorder

  // Меняёт местами bgColor и fgColor при несплошных заливках
  // INPUT
  // var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXSwapPatternFillColors(var PattFill: TZXLSXFill);
  var
    t: Integer;
    _b: Byte;
  begin
    // если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
    if (not(PattFill.patternfill in [ZPNone, ZPSolid])) then
    begin
      t := PattFill.PatternColor;
      PattFill.PatternColor := PattFill.BGColor;
      PattFill.BGColor := t;
      l1 := PattFill.lumFactorPattern;
      PattFill.lumFactorPattern := PattFill.lumFactorBG;
      PattFill.lumFactorBG := l1;

      _b := PattFill.patternColorType;
      PattFill.patternColorType := PattFill.bgColorType;
      PattFill.bgColorType := _b;
    end; // if
  end; // ZEXLSXSwapPatternFillColors

  // Очистить заливку ячейки
  // INPUT
  // var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXClearPatternFill(var PattFill: TZXLSXFill);
  begin
    PattFill.patternfill := ZPNone;
    PattFill.BGColor := TColorRec.cWindow;
    PattFill.PatternColor := TColorRec.cWindow;
    PattFill.bgColorType := 0;
    PattFill.patternColorType := 0;
    PattFill.lumFactorBG := 0.0;
    PattFill.lumFactorPattern := 0.0;
  end; // ZEXLSXClearPatternFill

  // Обнуляет стиль
  // INPUT
  // var style: TZXLSXCellStyle - стиль XLSX
  procedure ZEXLSXZeroCellStyle(var Style: TZXLSXCellStyle);
  begin
    Style.applyAlignment := false;
    Style.applyBorder := false;
    Style.applyProtection := false;
    Style.Hidden := false;
    Style.locked := false;
    Style.borderId := -1;
    Style.fontId := -1;
    Style.fillId := -1;
    Style.numFmtId := -1;
    Style.xfId := -1;
    Style.alignment.horizontal := ZHAutomatic;
    Style.alignment.vertical := ZVAutomatic;
    Style.alignment.shrinkToFit := false;
    Style.alignment.wrapText := false;
    Style.alignment.textRotation := 0;
    Style.alignment.indent := 0;
  end; // ZEXLSXZeroCellStyle

  // TZEXLSXFont в TFont
  // INPUT
  // var fnt: TZEXLSXFont  - XLSX шрифт
  // var font: TFont       - стандартный шрифт
  { procedure ZEXLSXFontToFont(var fnt: TZEXLSXFont; var font: TFont);
  begin
    if (Assigned(font)) then begin
      if (fnt.bold) then
        font.Style := font.Style + [fsBold];
      if (fnt.italic) then
        font.Style := font.Style + [fsItalic];
      if (fnt.underline) then
        font.Style := font.Style + [fsUnderline];
      if (fnt.strike) then
        font.Style := font.Style + [fsStrikeOut];
      font.Charset := fnt.charset;
      font.Name := fnt.name;
      font.Size := fnt.fontsize;
    end;
  end; } // ZEXLSXFontToFont

  // Прочитать цвет
  // INPUT
  // var retColor: TColor      - возвращаемый цвет
  // var retColorType: byte    - тип цвета: 0 - rgb, 1 - indexed, 2 - theme
  // var retLumfactor: double  - яркость
  procedure ZXLSXGetColor(var retColor: TColor; var retColorType: Byte; var retLumfactor: double);
  var
    t: Integer;
  begin
    s := Xml.Attributes.ItemsByName['rgb'];
    if (length(s) > 2) then
    begin
      delete(s, 1, 2);
      if (s > '') then
        retColor := HTMLHexToColor(s);
    end;
    s := Xml.Attributes.ItemsByName['theme'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 2;
        retColor := t;
      end;
    s := Xml.Attributes.ItemsByName['indexed'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 1;
        retColor := t;
      end;
    s := Xml.Attributes.ItemsByName['tint'];
    if (s <> '') then
      retLumfactor := ZETryStrToFloat(s, 0);
  end; // ZXLSXGetColor

  procedure _ReadFonts();
  var
    _currFont: Integer;
    sz: double;
  begin
    _currFont := -1;
    while Xml.ReadToEndTagByName('fonts') do
    begin
      s := Xml.Attributes.ItemsByName['val'];
      if Xml.IsTagStartByName('font') then
      begin
        _currFont := FontCount;
        inc(FontCount);
        SetLength(FontArray, FontCount);
        ZEXLSXZeroFont(FontArray[_currFont]);
      end
      else if (_currFont >= 0) then
      begin
        if Xml.IsTagClosedByName('name') then
          FontArray[_currFont].name := s
        else if Xml.IsTagClosedByName('b') then
          FontArray[_currFont].bold := true
        else if Xml.IsTagClosedByName('charset') then
        begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].charset := t;
        end
        else if Xml.IsTagClosedByName('color') then
        begin
          ZXLSXGetColor(FontArray[_currFont].Color, FontArray[_currFont].ColorType, FontArray[_currFont].LumFactor);
        end
        else if Xml.IsTagClosedByName('i') then
          FontArray[_currFont].italic := true
        else if Xml.IsTagClosedByName('strike') then
          FontArray[_currFont].strike := true
        else if Xml.IsTagClosedByName('sz') then
        begin
          if (TryStrToFloat(s, sz, TFormatSettings.Invariant)) then
            FontArray[_currFont].fontsize := sz;
        end
        else if Xml.IsTagClosedByName('u') then
        begin
          FontArray[_currFont].underline := true;
        end
        else if Xml.IsTagClosedByName('vertAlign') then
        begin
          FontArray[_currFont].superscript := s = 'superscript';
          FontArray[_currFont].subscript := s = 'subscript';
        end;
      end; // if
      // Тэги настройки шрифта
      // *b - bold
      // *charset
      // *color
      // ?condense
      // ?extend
      // ?family
      // *i - italic
      // *name
      // ?outline
      // ?scheme
      // ?shadow
      // *strike
      // *sz - size
      // *u - underline
      // *vertAlign

    end; // while
  end; // _ReadFonts

  // Получить тип заливки
  function _GetPatternFillByStr(const s: string): TZCellPattern;
  begin
    if (s = 'solid') then
      Result := ZPSolid
    else if (s = 'none') then
      Result := ZPNone
    else if (s = 'gray125') then
      Result := ZPGray125
    else if (s = 'gray0625') then
      Result := ZPGray0625
    else if (s = 'darkUp') then
      Result := ZPDiagStripe
    else if (s = 'mediumGray') then
      Result := ZPGray50
    else if (s = 'darkGray') then
      Result := ZPGray75
    else if (s = 'lightGray') then
      Result := ZPGray25
    else if (s = 'darkHorizontal') then
      Result := ZPHorzStripe
    else if (s = 'darkVertical') then
      Result := ZPVertStripe
    else if (s = 'darkDown') then
      Result := ZPReverseDiagStripe
    else if (s = 'darkUpDark') then
      Result := ZPDiagStripe
    else if (s = 'darkGrid') then
      Result := ZPDiagCross
    else if (s = 'darkTrellis') then
      Result := ZPThickDiagCross
    else if (s = 'lightHorizontal') then
      Result := ZPThinHorzStripe
    else if (s = 'lightVertical') then
      Result := ZPThinVertStripe
    else if (s = 'lightDown') then
      Result := ZPThinReverseDiagStripe
    else if (s = 'lightUp') then
      Result := ZPThinDiagStripe
    else if (s = 'lightGrid') then
      Result := ZPThinHorzCross
    else if (s = 'lightTrellis') then
      Result := ZPThinDiagCross
    else
      Result := ZPSolid; // {tut} потом подумать насчёт стилей границ
  end; // _GetPatternFillByStr

  // Определить стиль начертания границы
  // INPUT
  // const st: string            - название стиля
  // var retWidth: byte          - возвращаемая ширина линии
  // var retStyle: TZBorderType  - возвращаемый стиль начертания линии
  // RETURN
  // boolean - true - стиль определён
  function XLSXGetBorderStyle(const st: string; var retWidth: Byte; var retStyle: TZBorderType): Boolean;
  begin
    Result := true;
    retWidth := 1;
    if (st = 'thin') then
      retStyle := ZEContinuous
    else if (st = 'hair') then
      retStyle := ZEHair
    else if (st = 'dashed') then
      retStyle := ZEDash
    else if (st = 'dotted') then
      retStyle := ZEDot
    else if (st = 'dashDot') then
      retStyle := ZEDashDot
    else if (st = 'dashDotDot') then
      retStyle := ZEDashDotDot
    else if (st = 'slantDashDot') then
      retStyle := ZESlantDashDot
    else if (st = 'double') then
      retStyle := ZEDouble
    else if (st = 'medium') then
    begin
      retStyle := ZEContinuous;
      retWidth := 2;
    end
    else if (st = 'thick') then
    begin
      retStyle := ZEContinuous;
      retWidth := 3;
    end
    else if (st = 'mediumDashed') then
    begin
      retStyle := ZEDash;
      retWidth := 2;
    end
    else if (st = 'mediumDashDot') then
    begin
      retStyle := ZEDashDot;
      retWidth := 2;
    end
    else if (st = 'mediumDashDotDot') then
    begin
      retStyle := ZEDashDotDot;
      retWidth := 2;
    end
    else if (st = 'none') then
      retStyle := ZENone
    else
      Result := false;
  end; // XLSXGetBorderStyle

  procedure _ReadBorders();
  var
    _diagDown, _diagUP: Boolean;
    _currBorder: Integer; // текущий набор границ
    _currBorderItem: Integer; // текущая граница (левая/правая ...)
    _color: TColor;
    _isColor: Boolean;

    procedure _SetCurBorder(borderNum: Integer);
    begin
      _currBorderItem := borderNum;
      s := Xml.Attributes.ItemsByName['style'];
      if (s > '') then
      begin
        BorderArray[_currBorder][borderNum].isEnabled :=
          XLSXGetBorderStyle(s, BorderArray[_currBorder][borderNum].Weight, BorderArray[_currBorder][borderNum].Style);
      end;
    end; // _SetCurBorder

  begin
    _currBorderItem := -1;
    _diagDown := false;
    _diagUP := false;
    _color := TColorRec.Black;
    while Xml.ReadToEndTagByName('borders') do
    begin
      if Xml.IsTagStartByName('border') then
      begin
        _currBorder := BorderCount;
        inc(BorderCount);
        SetLength(BorderArray, BorderCount);
        ZEXLSXZeroBorder(BorderArray[_currBorder]);
        _diagDown := false;
        _diagUP := false;
        s := Xml.Attributes.ItemsByName['diagonalDown'];
        if (s > '') then
          _diagDown := ZEStrToBoolean(s);

        s := Xml.Attributes.ItemsByName['diagonalUp'];
        if (s > '') then
          _diagUP := ZEStrToBoolean(s);
      end
      else
      begin
        if (Xml.IsTagStartOrClosed) then
        begin
          if (Xml.TagName = 'left') then
          begin
            _SetCurBorder(0);
          end
          else if (Xml.TagName = 'right') then
          begin
            _SetCurBorder(2);
          end
          else if (Xml.TagName = 'top') then
          begin
            _SetCurBorder(1);
          end
          else if (Xml.TagName = 'bottom') then
          begin
            _SetCurBorder(3);
          end
          else if (Xml.TagName = 'diagonal') then
          begin
            if (_diagUP) then
              _SetCurBorder(5);
            if (_diagDown) then
            begin
              if (_diagUP) then
                BorderArray[_currBorder][4] := BorderArray[_currBorder][5]
              else
                _SetCurBorder(4);
            end;
          end
          else if (Xml.TagName = 'end') then
          begin
          end
          else if (Xml.TagName = 'start') then
          begin
          end
          else if (Xml.TagName = 'color') then
          begin
            _isColor := false;
            s := Xml.Attributes.ItemsByName['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if (s > '') then
            begin
              _color := HTMLHexToColor(s);
              _isColor := true;
            end;
            if (_isColor and (_currBorderItem >= 0) and (_currBorderItem < 6)) then
            begin
              BorderArray[_currBorder][_currBorderItem].Color := _color;
              BorderArray[_currBorder][_currBorderItem].isColor := true;
            end;
          end;
        end;
      end;
    end;
  end;

  procedure _ReadFills();
  var
    _currFill: Integer;
  begin
    _currFill := -1;
    while Xml.ReadToEndTagByName('fills') do
    begin
      if Xml.IsTagStartByName('fill') then
      begin
        _currFill := FillCount;
        inc(FillCount);
        SetLength(FillArray, FillCount);
        ZEXLSXClearPatternFill(FillArray[_currFill]);
      end
      else if ((Xml.TagName = 'patternFill') and (Xml.IsTagStartOrClosed)) then
      begin
        if (_currFill >= 0) then
        begin
          s := Xml.Attributes.ItemsByName['patternType'];
          {
          *none	None
          *solid	Solid
          ?mediumGray	Medium Gray
          ?darkGray	Dary Gray
          ?lightGray	Light Gray
          ?darkHorizontal	Dark Horizontal
          ?darkVertical	Dark Vertical
          ?darkDown	Dark Down
          ?darkUpDark Up
          ?darkGrid	Dark Grid
          ?darkTrellis	Dark Trellis
          ?lightHorizontal	Light Horizontal
          ?lightVertical	Light Vertical
          ?lightDown	Light Down
          ?lightUp	Light Up
          ?lightGrid	Light Grid
          ?lightTrellis	Light Trellis
          *gray125	Gray 0.125
          *gray0625	Gray 0.0625
 }

          if (s > '') then
            FillArray[_currFill].patternfill := _GetPatternFillByStr(s);
        end;
      end
      else if Xml.IsTagClosedByName('bgColor') then
      begin
        if (_currFill >= 0) then
        begin
          ZXLSXGetColor(FillArray[_currFill].PatternColor, FillArray[_currFill].patternColorType,
            FillArray[_currFill].lumFactorPattern);

          // если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
          ZEXLSXSwapPatternFillColors(FillArray[_currFill]);
        end;
      end
      else if Xml.IsTagClosedByName('fgColor') then
      begin
        if (_currFill >= 0) then
          ZXLSXGetColor(FillArray[_currFill].BGColor, FillArray[_currFill].bgColorType,
            FillArray[_currFill].lumFactorBG);
      end; // fgColor
    end; // while
  end; // _ReadFills

  // Читает стили (cellXfs и cellStyleXfs)
  // INPUT
  // const TagName: string           - имя тэга
  // var CSA: TZXLSXCellStylesArray  - массив со стилями
  // var StyleCount: integer         - кол-во стилей
  procedure _ReadCellCommonStyles(const TagName: string; var CSA: TZXLSXCellStylesArray; var StyleCount: Integer);
  var
    _currCell: Integer;
    b: Boolean;
  begin
    _currCell := -1;
    while Xml.ReadToEndTagByName(TagName) do
    begin
      b := false;
      if ((Xml.TagName = 'xf') and (Xml.IsTagStartOrClosed)) then
      begin
        _currCell := StyleCount;
        inc(StyleCount);
        SetLength(CSA, StyleCount);
        ZEXLSXZeroCellStyle(CSA[_currCell]);
        s := Xml.Attributes.ItemsByName['applyAlignment'];
        if (s > '') then
          CSA[_currCell].applyAlignment := ZEStrToBoolean(s);

        s := Xml.Attributes.ItemsByName['applyBorder'];
        if (s > '') then
          CSA[_currCell].applyBorder := ZEStrToBoolean(s)
        else
          b := true;

        s := Xml.Attributes.ItemsByName['applyFont'];
        if (s > '') then
          CSA[_currCell].applyFont := ZEStrToBoolean(s);

        s := Xml.Attributes.ItemsByName['applyProtection'];
        if (s > '') then
          CSA[_currCell].applyProtection := ZEStrToBoolean(s);

        s := Xml.Attributes.ItemsByName['borderId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            CSA[_currCell].borderId := t;
            if (b and (t >= 1)) then
              CSA[_currCell].applyBorder := true;
          end;

        s := Xml.Attributes.ItemsByName['fillId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fillId := t;

        s := Xml.Attributes.ItemsByName['fontId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fontId := t;

        s := Xml.Attributes.ItemsByName['numFmtId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].numFmtId := t;

        {
          <xfId> (Format Id)
          For <xf> records contained in <cellXfs> this is the zero-based index of an <xf> record contained in <cellStyleXfs> corresponding to the cell style applied to the cell.
          Not present for <xf> records contained in <cellStyleXfs>.
          The possible values for this attribute are defined by the ST_CellStyleXfId simple type (§3.18.11).
          https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_xf_topic_ID0E13S6.html

 }

        s := Xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].xfId := t;
      end
      else if Xml.IsTagClosedByName('alignment') then
      begin
        if (_currCell >= 0) then
        begin
          s := Xml.Attributes.ItemsByName['horizontal'];
          if (s > '') then
          begin
            if (s = 'general') then
              CSA[_currCell].alignment.horizontal := ZHAutomatic
            else if (s = 'left') then
              CSA[_currCell].alignment.horizontal := ZHLeft
            else if (s = 'right') then
              CSA[_currCell].alignment.horizontal := ZHRight
            else if ((s = 'center') or (s = 'centerContinuous')) then
              CSA[_currCell].alignment.horizontal := ZHCenter
            else if (s = 'fill') then
              CSA[_currCell].alignment.horizontal := ZHFill
            else if (s = 'justify') then
              CSA[_currCell].alignment.horizontal := ZHJustify
            else if (s = 'distributed') then
              CSA[_currCell].alignment.horizontal := ZHDistributed;
          end;

          s := Xml.Attributes.ItemsByName['indent'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.indent := t;

          s := Xml.Attributes.ItemsByName['shrinkToFit'];
          if (s > '') then
            CSA[_currCell].alignment.shrinkToFit := ZEStrToBoolean(s);

          s := Xml.Attributes.ItemsByName['textRotation'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.textRotation := t;

          s := Xml.Attributes.ItemsByName['vertical'];
          if (s > '') then
          begin
            if (s = 'center') then
              CSA[_currCell].alignment.vertical := ZVCenter
            else if (s = 'top') then
              CSA[_currCell].alignment.vertical := ZVTop
            else if (s = 'bottom') then
              CSA[_currCell].alignment.vertical := ZVBottom
            else if (s = 'justify') then
              CSA[_currCell].alignment.vertical := ZVJustify
            else if (s = 'distributed') then
              CSA[_currCell].alignment.vertical := ZVDistributed;
          end;

          s := Xml.Attributes.ItemsByName['wrapText'];
          if (s > '') then
            CSA[_currCell].alignment.wrapText := ZEStrToBoolean(s);
        end;
      end
      else if Xml.IsTagClosedByName('protection') then
      begin
        if (_currCell >= 0) then
        begin
          s := Xml.Attributes.ItemsByName['hidden'];
          if (s > '') then
            CSA[_currCell].Hidden := ZEStrToBoolean(s);

          s := Xml.Attributes.ItemsByName['locked'];
          if (s > '') then
            CSA[_currCell].locked := ZEStrToBoolean(s);
        end;
      end;
    end;
  end;

  // Сами стили ?? (или для чего они вообще?)
  procedure _ReadCellStyles();
  var
    b: Boolean;
  begin
    while Xml.ReadToEndTagByName('cellStyles') do
    begin
      if Xml.IsTagClosedByName('cellStyle') then
      begin
        b := false;
        SetLength(StyleArray, StyleCount + 1);
        s := Xml.Attributes.ItemsByName['builtinId']; // ?
        if (s > '') then
          if (TryStrToInt(s, t)) then
            StyleArray[StyleCount].builtinId := t;

        s := Xml.Attributes.ItemsByName['customBuiltin']; // ?
        if (s > '') then
          StyleArray[StyleCount].customBuiltin := ZEStrToBoolean(s);

        s := Xml.Attributes.ItemsByName['name']; // ?
        StyleArray[StyleCount].name := s;

        s := Xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            StyleArray[StyleCount].xfId := t;
            b := true;
          end;

        if (b) then
          inc(StyleCount);
      end;
    end;
  end;

  procedure _ReadColors();
  begin
    while Xml.ReadToEndTagByName('colors') do
    begin
      if Xml.IsTagClosedByName('rgbColor') then
      begin
        s := Xml.Attributes.ItemsByName['rgb'];
        if (length(s) > 2) then
          delete(s, 1, 2);
        if (s > '') then
        begin
          inc(indexedColorCount);
          if (indexedColorCount >= indexedColorMax) then
          begin
            indexedColorMax := indexedColorCount + 80;
            SetLength(indexedColor, indexedColorMax);
          end;
          indexedColor[indexedColorCount - 1] := HTMLHexToColor(s);
        end;
      end;
    end;
  end;

  // Конвертирует RGB в HSL
  // http://en.wikipedia.org/wiki/HSL_color_space
  // INPUT
  // r: byte     -
  // g: byte     -
  // b: byte     -
  // out h: double   - Hue - тон цвета
  // out s: double   - Saturation - насыщенность
  // out l: double   - Lightness (Intensity) - светлота (яркость)
  procedure ZRGBToHSL(r, g, b: Byte; out h, s, l: double);
  var
    _max, _min: double;
    intMax, intMin: Integer;
    _r, _g, _b: double;
    _delta: double;
    _idx: Integer;
  begin
    _r := r / 255;
    _g := g / 255;
    _b := b / 255;

    intMax := Max(r, Max(g, b));
    intMin := min(r, min(g, b));

    _max := Max(_r, Max(_g, _b));
    _min := min(_r, min(_g, _b));

    h := (_max + _min) * 0.5;
    s := h;
    l := h;
    if (intMax = intMin) then
    begin
      h := 0;
      s := 0;
    end
    else
    begin
      _delta := _max - _min;
      if (l > 0.5) then
        s := _delta / (2 - _max - _min)
      else
        s := _delta / (_max + _min);

      if (intMax = r) then
        _idx := 1
      else if (intMax = g) then
        _idx := 2
      else
        _idx := 3;

      case (_idx) of
        1:
          begin
            h := (_g - _b) / _delta;
            if (g < b) then
              h := h + 6;
          end;
        2:
          h := (_b - _r) / _delta + 2;
        3:
          h := (_r - _g) / _delta + 4;
      end;

      h := h / 6;
    end;
  end; // ZRGBToHSL

  // Конвертирует TColor (RGB) в HSL
  // http://en.wikipedia.org/wiki/HSL_color_space
  // INPUT
  // Color: TColor - цвет
  // out h: double     - Hue - тон цвета
  // out s: double     - Saturation - насыщенность
  // out l: double     - Lightness (Intensity) - светлота (яркость)
  procedure ZColorToHSL(Color: TColor; out h, s, l: double);
  var
    _rgb: Integer;
  begin
    _rgb := TColorRec.ColorToRGB(Color);
    ZRGBToHSL(Byte(_rgb), Byte(_rgb shr 8), Byte(_rgb shr 16), h, s, l);
  end; // ZColorToHSL

  // Конвертирует HSL в RGB
  // http://en.wikipedia.org/wiki/HSL_color_space
  // INPUT
  // h: double - Hue - тон цвета
  // s: double - Saturation - насыщенность
  // l: double - Lightness (Intensity) - светлота (яркость)
  // out r: byte   -
  // out g: byte   -
  // out b: byte   -
  procedure ZHSLToRGB(h, s, l: double; out r, g, b: Byte);
  var
    _r, _g, _b, q, p: double;
    function HueToRgb(p, q, t: double): double;
    begin
      Result := p;
      if (t < 0) then
        t := t + 1;
      if (t > 1) then
        t := t - 1;
      if (t < 1 / 6) then
        Result := p + (q - p) * 6 * t
      else if (t < 0.5) then
        Result := q
      else if (t < 2 / 3) then
        Result := p + (q - p) * (2 / 3 - t) * 6;
    end; // HueToRgb

  begin
    if (s = 0) then
    begin
      // Оттенок серого
      _r := l;
      _g := l;
      _b := l;
    end
    else
    begin
      if (l < 0.5) then
        q := l * (1 + s)
      else
        q := l + s - l * s;
      p := 2 * l - q;
      _r := HueToRgb(p, q, h + 1 / 3);
      _g := HueToRgb(p, q, h);
      _b := HueToRgb(p, q, h - 1 / 3);
    end;
    r := Byte(round(_r * 255));
    g := Byte(round(_g * 255));
    b := Byte(round(_b * 255));
  end; // ZHSLToRGB

  // Конвертирует HSL в Color
  // http://en.wikipedia.org/wiki/HSL_color_space
  // INPUT
  // h: double - Hue - тон цвета
  // s: double - Saturation - насыщенность
  // l: double - Lightness (Intensity) - светлота (яркость)
  // RETURN
  // TColor - цвет
  function ZHSLToColor(h, s, l: double): TColor;
  var
    r, g, b: Byte;
  begin
    ZHSLToRGB(h, s, l, r, g, b);
    Result := (b shl 16) or (g shl 8) or r;
  end; // ZHSLToColor

  // Применить tint к цвету
  // Thanks Tomasz Wieckowski!
  // http://msdn.microsoft.com/en-us/library/ff532470%28v=office.12%29.aspx
  procedure ApplyLumFactor(var Color: TColor; var LumFactor: double);
  begin
    // +delta?
    if (LumFactor <> 0.0) then
    begin
      ZColorToHSL(Color, h1, s1, l1);
      LumFactor := 1 - LumFactor;

      if (l1 = 1) then
        l1 := l1 * (1 - LumFactor)
      else
        l1 := l1 * LumFactor + (1 - LumFactor);

      Color := ZHSLToColor(h1, s1, l1);
    end;
  end; // ApplyLumFactor

  // Differential Formatting для xlsx
  procedure _Readdxfs();
  var
    _df: TZXLSXDiffFormattingItem;
    _dfIndex: Integer;

    procedure _addFontStyle(fnts: TFontStyle);
    begin
      _df.FontStyles := _df.FontStyles + [fnts];
      _df.UseFontStyles := true;
    end;

    procedure _ReadDFFont();
    begin
      _df.UseFont := true;
      while Xml.ReadToEndTagByName('font') do
      begin
        if (Xml.TagName = 'i') then
          _addFontStyle(TFontStyle.fsItalic);
        if (Xml.TagName = 'b') then
          _addFontStyle(TFontStyle.fsBold);
        if (Xml.TagName = 'u') then
          _addFontStyle(TFontStyle.fsUnderline);
        if (Xml.TagName = 'strike') then
          _addFontStyle(TFontStyle.fsStrikeOut);

        if (Xml.TagName = 'color') then
        begin
          _df.UseFontColor := true;
          ZXLSXGetColor(_dfFonts[_dfIndex].Color, _dfFonts[_dfIndex].ColorType, _dfFonts[_dfIndex].LumFactor);
        end;
      end; // while
    end; // _ReadDFFont

    procedure _ReadDFFill();
    begin
      _df.UseFill := true;
      while not Xml.IsTagEndByName('fill') do
      begin
        Xml.ReadTag();
        if (Xml.Eof) then
          break;

        if (Xml.IsTagStartOrClosed) then
        begin
          if (Xml.TagName = 'patternFill') then
          begin
            s := Xml.Attributes.ItemsByName['patternType'];
            if (s <> '') then
            begin
              _df.UseCellPattern := true;
              _df.CellPattern := _GetPatternFillByStr(s);
            end;
          end
          else if (Xml.TagName = 'bgColor') then
          begin
            _df.UseBGColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].BGColor, _dfFills[_dfIndex].bgColorType, _dfFills[_dfIndex].lumFactorBG)
          end
          else if (Xml.TagName = 'fgColor') then
          begin
            _df.UsePatternColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].PatternColor, _dfFills[_dfIndex].patternColorType,
              _dfFills[_dfIndex].lumFactorPattern);
            ZEXLSXSwapPatternFillColors(_dfFills[_dfIndex]);
          end;
        end;
      end; // while
    end; // _ReadDFFill

    procedure _ReadDFBorder();
    var
      _borderNum: TZBordersPos;
      t: Byte;
      _bt: TZBorderType;
      procedure _SetDFBorder(borderNum: TZBordersPos);
      begin
        _borderNum := borderNum;
        s := Xml.Attributes['style'];
        if (s <> '') then
          if (XLSXGetBorderStyle(s, t, _bt)) then
          begin
            _df.UseBorder := true;
            _df.Borders[borderNum].Weight := t;
            _df.Borders[borderNum].LineStyle := _bt;
            _df.Borders[borderNum].UseStyle := true;
          end;
      end; // _SetDFBorder

    begin
      _df.UseBorder := true;
      _borderNum := bpLeft;
      while Xml.ReadToEndTagByName('border') do
      begin
        if Xml.IsTagStartOrClosed then
        begin
          if (Xml.TagName = 'left') then
            _SetDFBorder(bpLeft)
          else if (Xml.TagName = 'right') then
            _SetDFBorder(bpRight)
          else if (Xml.TagName = 'top') then
            _SetDFBorder(bpTop)
          else if (Xml.TagName = 'bottom') then
            _SetDFBorder(bpBottom)
          else if (Xml.TagName = 'vertical') then
            _SetDFBorder(bpDiagonalLeft)
          else if (Xml.TagName = 'horizontal') then
            _SetDFBorder(bpDiagonalRight)
          else if (Xml.TagName = 'color') then
          begin
            s := Xml.Attributes['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if ((_borderNum >= bpLeft) and (_borderNum <= bpDiagonalRight)) then
              if (s <> '') then
              begin
                _df.UseBorder := true;
                _df.Borders[_borderNum].UseColor := true;
                _df.Borders[_borderNum].Color := HTMLHexToColor(s);
              end;
          end;
        end; // if
      end; // while
    end; // _ReadDFBorder

    procedure _ReaddxfItem();
    begin
      _dfIndex := ReadHelper.DiffFormatting.Count;

      SetLength(_dfFonts, _dfIndex + 1);
      _dfFonts[_dfIndex].ColorType := 0;
      _dfFonts[_dfIndex].LumFactor := 0;

      SetLength(_dfFills, _dfIndex + 1);
      ZEXLSXClearPatternFill(_dfFills[_dfIndex]);

      ReadHelper.DiffFormatting.Add();
      _df := ReadHelper.DiffFormatting[_dfIndex];
      while Xml.ReadToEndTagByName('dxf') do
      begin
        if Xml.IsTagStartByName('font') then
          _ReadDFFont()
        else if Xml.IsTagStartByName('fill') then
          _ReadDFFill()
        else if Xml.IsTagStartByName('border') then
          _ReadDFBorder();
      end; // while
    end; // _ReaddxfItem

  begin
    while Xml.ReadToEndTagByName('dxfs') do
    begin
      if Xml.IsTagStartByName('dxf') then
        _ReaddxfItem();
    end; // while
  end; // _Readdxfs

  procedure XLSXApplyColor(var AColor: TColor; ColorType: Byte; LumFactor: double);
  begin
    // Thema color
    if (ColorType = 2) then
    begin
      t := AColor - 1;
      if ((t >= 0) and (t < ThemaColorCount)) then
        AColor := ThemaFillsColors[t];
    end;
    if (ColorType = 1) then
      if ((AColor >= 0) and (AColor < indexedColorCount)) then
        AColor := indexedColor[AColor];
    ApplyLumFactor(AColor, LumFactor);
  end; // XLSXApplyColor

  // Применить стиль
  // INPUT
  // var XMLSSStyle: TZStyle         - стиль в хранилище
  // var XLSXStyle: TZXLSXCellStyle  - стиль в xlsx
  procedure _ApplyStyle(var XMLSSStyle: TZStyle; var XLSXStyle: TZXLSXCellStyle);
  var
    i: Integer;
    b: TZBordersPos;
  begin
    if (XLSXStyle.numFmtId >= 0) then
      XMLSSStyle.NumberFormat := ReadHelper.NumberFormats.GetFormat(XLSXStyle.numFmtId);
    XMLSSStyle.NumberFormatId := XLSXStyle.numFmtId;

    if (XLSXStyle.applyAlignment) then
    begin
      XMLSSStyle.alignment.horizontal := XLSXStyle.alignment.horizontal;
      XMLSSStyle.alignment.vertical := XLSXStyle.alignment.vertical;
      XMLSSStyle.alignment.indent := XLSXStyle.alignment.indent;
      XMLSSStyle.alignment.shrinkToFit := XLSXStyle.alignment.shrinkToFit;
      XMLSSStyle.alignment.wrapText := XLSXStyle.alignment.wrapText;

      XMLSSStyle.alignment.Rotate := 0;
      i := XLSXStyle.alignment.textRotation;
      XMLSSStyle.alignment.VerticalText := (i = 255);
      if (i >= 0) and (i <= 180) then
      begin
        if i > 90 then
          i := 90 - i;
        XMLSSStyle.alignment.Rotate := i
      end;
    end;

    if XLSXStyle.applyBorder then
    begin
      n := XLSXStyle.borderId;
      if (n >= 0) and (n < BorderCount) then
        for b := bpLeft to bpDiagonalRight do
        begin
          if (BorderArray[n][Ord(b)].isEnabled) then
          begin
            XMLSSStyle.Border[b].LineStyle := BorderArray[n][Ord(b)].Style;
            XMLSSStyle.Border[b].Weight := BorderArray[n][Ord(b)].Weight;
            if (BorderArray[n][Ord(b)].isColor) then
              XMLSSStyle.Border[b].Color := BorderArray[n][Ord(b)].Color;
          end;
        end;
    end;

    if (XLSXStyle.applyFont) then
    begin
      n := XLSXStyle.fontId;
      if ((n >= 0) and (n < FontCount)) then
      begin
        XLSXApplyColor(FontArray[n].Color, FontArray[n].ColorType, FontArray[n].LumFactor);
        XMLSSStyle.Font.name := FontArray[n].name;
        XMLSSStyle.Font.Size := FontArray[n].fontsize;
        XMLSSStyle.Font.charset := FontArray[n].charset;
        XMLSSStyle.Font.Color := FontArray[n].Color;
        if (FontArray[n].bold) then
          XMLSSStyle.Font.Style := [TFontStyle.fsBold];
        if (FontArray[n].underline) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [TFontStyle.fsUnderline];
        if (FontArray[n].italic) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [TFontStyle.fsItalic];
        if (FontArray[n].strike) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [TFontStyle.fsStrikeOut];
        XMLSSStyle.superscript := FontArray[n].superscript;
        XMLSSStyle.subscript := FontArray[n].subscript;
      end;
    end;

    if (XLSXStyle.applyProtection) then
    begin
      XMLSSStyle.Protect := XLSXStyle.locked;
      XMLSSStyle.HideFormula := XLSXStyle.Hidden;
    end;

    n := XLSXStyle.fillId;
    if ((n >= 0) and (n < FillCount)) then
    begin
      XMLSSStyle.CellPattern := FillArray[n].patternfill;
      XMLSSStyle.BGColor := FillArray[n].BGColor;
      XMLSSStyle.PatternColor := FillArray[n].PatternColor;
    end;
  end; // _ApplyStyle

  procedure _CheckIndexedColors();
  const
    _standart: array [0 .. 63] of string = ('#000000', // 0
      '#FFFFFF', // 1
      '#FF0000', // 2
      '#00FF00', // 3
      '#0000FF', // 4
      '#FFFF00', // 5
      '#FF00FF', // 6
      '#00FFFF', // 7
      '#000000', // 8
      '#FFFFFF', // 9
      '#FF0000', // 10
      '#00FF00', // 11
      '#0000FF', // 12
      '#FFFF00', // 13
      '#FF00FF', // 14
      '#00FFFF', // 15
      '#800000', // 16
      '#008000', // 17
      '#000080', // 18
      '#808000', // 19
      '#800080', // 20
      '#008080', // 21
      '#C0C0C0', // 22
      '#808080', // 23
      '#9999FF', // 24
      '#993366', // 25
      '#FFFFCC', // 26
      '#CCFFFF', // 27
      '#660066', // 28
      '#FF8080', // 29
      '#0066CC', // 30
      '#CCCCFF', // 31
      '#000080', // 32
      '#FF00FF', // 33
      '#FFFF00', // 34
      '#00FFFF', // 35
      '#800080', // 36
      '#800000', // 37
      '#008080', // 38
      '#0000FF', // 39
      '#00CCFF', // 40
      '#CCFFFF', // 41
      '#CCFFCC', // 42
      '#FFFF99', // 43
      '#99CCFF', // 44
      '#FF99CC', // 45
      '#CC99FF', // 46
      '#FFCC99', // 47
      '#3366FF', // 48
      '#33CCCC', // 49
      '#99CC00', // 50
      '#FFCC00', // 51
      '#FF9900', // 52
      '#FF6600', // 53
      '#666699', // 54
      '#969696', // 55
      '#003366', // 56
      '#339966', // 57
      '#003300', // 58
      '#333300', // 59
      '#993300', // 60
      '#993366', // 61
      '#333399', // 62
      '#333333' // 63
      );
  var
    i: Integer;
  begin
    if (indexedColorCount = 0) then
    begin
      indexedColorCount := 63;
      indexedColorMax := indexedColorCount + 10;
      SetLength(indexedColor, indexedColorMax);
      for i := 0 to 63 do
        indexedColor[i] := HTMLHexToColor(_standart[i]);
    end;
  end; // _CheckIndexedColors

begin
  Result := false;
  MaximumDigitWidth := 0;
  Xml := nil;
  CellXfsArray := nil;
  CellStyleArray := nil;
  try
    Xml := TZsspXMLReaderH.Create();
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    FontCount := 0;
    BorderCount := 0;
    CellStyleCount := 0;
    StyleCount := 0;
    CellXfsCount := 0;
    FillCount := 0;
    indexedColorCount := 0;
    indexedColorMax := -1;

    while not Xml.Eof() do
    begin
      Xml.ReadTag();

      if Xml.IsTagStartByName('fonts') then
      begin
        _ReadFonts();
        if length(FontArray) > 0 then
          MaximumDigitWidth := GetMaximumDigitWidth(FontArray[0].name, FontArray[0].fontsize);
      end
      else if Xml.IsTagStartByName('borders') then
        _ReadBorders()
      else if Xml.IsTagStartByName('fills') then
        _ReadFills()
      else
      {
        А.А.Валуев:
        Элементы внутри cellXfs ссылаются на элементы внутри cellStyleXfs.
        Элементы внутри cellStyleXfs ни на что не ссылаются.
 }
        if Xml.IsTagStartByName('cellStyleXfs') then
          _ReadCellCommonStyles('cellStyleXfs', CellStyleArray, CellStyleCount) // _ReadCellStyleXfs()
        else if Xml.IsTagStartByName('cellXfs') then // сами стили?
          _ReadCellCommonStyles('cellXfs', CellXfsArray, CellXfsCount) // _ReadCellXfs()
        else if Xml.IsTagStartByName('cellStyles') then // ??
          _ReadCellStyles()
        else if Xml.IsTagStartByName('colors') then
          _ReadColors()
        else if Xml.IsTagStartByName('dxfs') then
          _Readdxfs()
        else if Xml.IsTagStartByName('numFmts') then
          ReadHelper.NumberFormats.ReadNumFmts(Xml);
    end; // while

    // тут незабыть применить номера цветов, если были введены

    _CheckIndexedColors();

    //
    for i := 0 to FillCount - 1 do
    begin
      XLSXApplyColor(FillArray[i].BGColor, FillArray[i].bgColorType, FillArray[i].lumFactorBG);
      XLSXApplyColor(FillArray[i].PatternColor, FillArray[i].patternColorType, FillArray[i].lumFactorPattern);
    end; // for

    // {tut}

    XMLSS.Styles.Count := CellXfsCount;
    ReadHelper.NumberFormats.StyleFMTCount := CellXfsCount;
    for i := 0 to CellXfsCount - 1 do
    begin
      t := CellXfsArray[i].xfId;
      ReadHelper.NumberFormats.StyleFMTID[i] := CellXfsArray[i].numFmtId;

      _Style := XMLSS.Styles[i];
      if ((t >= 0) and (t < CellStyleCount)) then
        _ApplyStyle(_Style, CellStyleArray[t]);
      // else
      _ApplyStyle(_Style, CellXfsArray[i]);
    end;

    // Применение цветов к DF
    for i := 0 to ReadHelper.DiffFormatting.Count - 1 do
    begin
      if (ReadHelper.DiffFormatting[i].UseFontColor) then
      begin
        XLSXApplyColor(_dfFonts[i].Color, _dfFonts[i].ColorType, _dfFonts[i].LumFactor);
        ReadHelper.DiffFormatting[i].FontColor := _dfFonts[i].Color;
      end;
      if (ReadHelper.DiffFormatting[i].UseBGColor) then
      begin
        XLSXApplyColor(_dfFills[i].BGColor, _dfFills[i].bgColorType, _dfFills[i].lumFactorBG);
        ReadHelper.DiffFormatting[i].BGColor := _dfFills[i].BGColor;
      end;
      if (ReadHelper.DiffFormatting[i].UsePatternColor) then
      begin
        XLSXApplyColor(_dfFills[i].PatternColor, _dfFills[i].patternColorType, _dfFills[i].lumFactorPattern);
        ReadHelper.DiffFormatting[i].PatternColor := _dfFills[i].PatternColor;
      end;
    end;

    Result := true;
  finally
    if (Assigned(Xml)) then
      FreeAndNil(Xml);
    SetLength(FontArray, 0);
    FontArray := nil;
    SetLength(BorderArray, 0);
    BorderArray := nil;
    SetLength(CellStyleArray, 0);
    CellStyleArray := nil;
    SetLength(StyleArray, 0);
    StyleArray := nil;
    SetLength(CellXfsArray, 0);
    CellXfsArray := nil;
    SetLength(FillArray, 0);
    FillArray := nil;
    SetLength(indexedColor, 0);
    indexedColor := nil;
    SetLength(_dfFonts, 0);
    SetLength(_dfFills, 0);
  end;
end; // ZEXSLXReadStyles

// Читает названия листов (workbook.xml)
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// var Stream: TStream                 - поток
// var Relations: TZXLSXRelationsArray - связи
// var RelationsCount: integer         - кол-во
// RETURN
// boolean - true - названия прочитались без ошибок
function ZEXSLXReadWorkBook(var XMLSS: TZWorkBook; var Stream: TStream; var Relations: TZXLSXRelationsArray;
  var RelationsCount: Integer): Boolean;
var
  Xml: TZsspXMLReaderH;
  s: string;
  i, t, dn: Integer;
begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  try
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    dn := 0;
    while (not Xml.Eof()) do
    begin
      Xml.ReadTag();

      if Xml.IsTagStartByName('definedName') then
      begin
        Xml.ReadTag();
        SetLength(XMLSS.FDefinedNames, dn + 1);
        XMLSS.FDefinedNames[dn].LocalSheetId := StrToIntDef(Xml.Attributes.ItemsByName['localSheetId'], 0);
        XMLSS.FDefinedNames[dn].name := Xml.Attributes.ItemsByName['name'];
        XMLSS.FDefinedNames[dn].Body := Xml.TagValue;
        inc(dn);
      end
      else if Xml.IsTagClosedByName('sheet') then
      begin
        s := Xml.Attributes.ItemsByName['r:id'];
        for i := 0 to RelationsCount - 1 do
          if (Relations[i].id = s) then
          begin
            Relations[i].name := ZEReplaceEntity(Xml.Attributes.ItemsByName['name']);
            s := Xml.Attributes.ItemsByName['sheetId'];
            Relations[i].sheetid := -1;
            if (TryStrToInt(s, t)) then
              Relations[i].sheetid := t;
            s := Xml.Attributes.ItemsByName['state'];
            break;
          end;
      end
      else if Xml.IsTagClosedByName('workbookView') then
      begin
        s := Xml.Attributes.ItemsByName['activeTab'];
        s := Xml.Attributes.ItemsByName['firstSheet'];
        s := Xml.Attributes.ItemsByName['showHorizontalScroll'];
        s := Xml.Attributes.ItemsByName['showSheetTabs'];
        s := Xml.Attributes.ItemsByName['showVerticalScroll'];
        s := Xml.Attributes.ItemsByName['tabRatio'];
        s := Xml.Attributes.ItemsByName['windowHeight'];
        s := Xml.Attributes.ItemsByName['windowWidth'];
        s := Xml.Attributes.ItemsByName['xWindow'];
        s := Xml.Attributes.ItemsByName['yWindow'];
      end;
    end; // while
    Result := true;
  finally
    Xml.Free();
  end;
end; // ZEXSLXReadWorkBook

// Удаляет первый символ + меняет все разделители на нужные
// INPUT
// var FileArray: TArray<TZXLSXFileItem>  - файлы
// FilesCount: integer         - кол-во файлов
procedure ZE_XSLXReplaceDelimiter(var FileArray: TArray<TZXLSXFileItem>; FilesCount: Integer);
var
  i, j, k: Integer;
begin
  for i := 0 to FilesCount - 1 do
  begin
    k := length(FileArray[i].name);
    if (k > 1) then
    begin
      if (FileArray[i].name[1] = '/') then
        delete(FileArray[i].name, 1, 1);
      if (PathDelim <> '/') then
        for j := 1 to k - 1 do
          if (FileArray[i].name[j] = '/') then
            FileArray[i].name[j] := PathDelim;
    end;
  end;
end; // ZE_XSLXReplaceDelimiter

// Читает связи страниц/стилей  (*.rels: workbook.xml.rels и .rels)
// INPUT
// var Stream: TStream                 - поток для чтения
// var Relations: TZXLSXRelationsArray - массив с отношениями
// var RelationsCount: integer         - кол-во
// var isWorkSheet: boolean            - признак workbook.xml.rels
// needReplaceDelimiter: boolean   - признак необходимости заменять разделитель
// RETURN
// boolean - true - успешно прочитано
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: Integer;
  var isWorkSheet: Boolean; needReplaceDelimiter: Boolean): Boolean;
var
  Xml: TZsspXMLReaderH;
  rt: TRelationType;
begin
  Result := false;
  Xml := TZsspXMLReaderH.Create();
  RelationsCount := 0;
  isWorkSheet := false;
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    while not Xml.Eof() do
    begin
      Xml.ReadTag();

      if Xml.IsTagClosedByName('Relationship') then
      begin
        SetLength(Relations, RelationsCount + 1);
        Relations[RelationsCount].id := Xml.Attributes.ItemsByName['Id'];

        rt := ZEXLSXGetRelationNumber(Xml.Attributes.ItemsByName['Type']);
        if ((rt >= rtWorkSheet) and (rt < rtDoc)) then
          isWorkSheet := true;

        Relations[RelationsCount].fileid := -1;
        Relations[RelationsCount].state := 0;
        Relations[RelationsCount].sheetid := 0;
        Relations[RelationsCount].name := '';
        Relations[RelationsCount].ftype := rt;
        Relations[RelationsCount].target := Xml.Attributes.ItemsByName['Target'];
        if (rt >= rtWorkSheet) then
          inc(RelationsCount);
      end;
    end; // while
    Result := true;
  finally
    FreeAndNil(Xml);
  end;
end; // ZE_XSLXReadRelationsips

// Читает примечания (добавляет примечания на последнюю страницу)
// INPUT
// var XMLSS: TZWorkBook - хранилище
// var Stream: TStream - поток для чтения
// RETURN
// boolean - true - всё нормально
function ZEXSLXReadComments(var XMLSS: TZWorkBook; var Stream: TStream): Boolean;
var
  Xml: TZsspXMLReaderH;
  authors: TList<string>;
  page: Integer;

  procedure _ReadComment();
  var
    _c, _r, _a: Integer;
    _comment, str: string;
    _kol: Integer;
  begin
    _c := 0;
    _r := 0;
    str := Xml.Attributes.ItemsByName['ref'];
    if (str = '') then
      exit;
    if (ZEGetCellCoords(str, _c, _r, true)) then
    begin
      if (_c >= XMLSS.Sheets[page].ColCount) then
        XMLSS.Sheets[page].ColCount := _c + 1;
      if (_r >= XMLSS.Sheets[page].RowCount) then
        XMLSS.Sheets[page].RowCount := _r + 1;

      if (TryStrToInt(Xml.Attributes.ItemsByName['authorId'], _a)) then
        if (_a >= 0) and (_a < authors.Count) then
          XMLSS.Sheets[page].Cell[_c, _r].CommentAuthor := authors[_a];

      _comment := '';
      _kol := 0;
      while (not((Xml.TagName = 'comment') and (Xml.IsTagEnd))) do
      begin
        Xml.ReadTag();
        if (Xml.Eof()) then
          break;
        if ((Xml.TagName = 't') and (Xml.IsTagEnd)) then
        begin
          if (_kol > 0) then
            _comment := _comment + sLineBreak + Xml.TextBeforeTag
          else
            _comment := _comment + Xml.TextBeforeTag;
          inc(_kol);
        end;
      end; // while
      XMLSS.Sheets[page].Cell[_c, _r].Comment := _comment;
      XMLSS.Sheets[page].Cell[_c, _r].ShowComment := true;
    end // if
  end; // _ReadComment();

begin
  Result := false;
  authors := TList<string>.Create();
  Xml := TZsspXMLReaderH.Create();
  page := XMLSS.Sheets.Count - 1;
  if (page < 0) then
    exit;
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(Stream) <> 0) then
      exit;

    while not Xml.Eof() do
    begin
      Xml.ReadTag();

      if Xml.IsTagStartByName('authors') then
      begin
        while Xml.ReadToEndTagByName('authors') do
        begin
          if Xml.IsTagEndByName('authors') then
            authors.Add(Xml.TextBeforeTag);
        end;
      end
      else

        if Xml.IsTagStartByName('comment') then
        _ReadComment();
    end; // while
    Result := true;
  finally
    Xml.Free();
    authors.Free();
  end;
end; // ZEXSLXReadComments

procedure XLSXSortRelationArray(var arr: TZXLSXRelationsArray; Count: Integer);
var
  tmp: TZXLSXRelations;
  i, j: Integer;
  _t1, _t2: Integer;
  s: string;
  b: Boolean;
  function _cmp(): Boolean;
  begin
    b := false;
    s := arr[j].id;
    delete(s, 1, 3);
    b := TryStrToInt(s, _t1);
    if (b) then
    begin
      s := arr[j + 1].id;
      delete(s, 1, 3);
      b := TryStrToInt(s, _t2);
    end;

    if (b) then
      Result := _t1 > _t2
    else
      Result := arr[j].id > arr[j + 1].id;
  end;

begin
  // TODO: do not forget update sorting.
  for i := 0 to Count - 2 do
    for j := 0 to Count - 2 do
      if (_cmp()) then
      begin
        tmp := arr[j];
        arr[j] := arr[j + 1];
        arr[j + 1] := tmp;
      end;
end;

// Читает распакованный xlsx
// INPUT
// var XMLSS: TZWorkBook - хранилище
// DirName: string     - имя папки
// RETURN
// integer - номер ошибки (0 - всё OK)
function ReadXLSXPath(var XMLSS: TZWorkBook; DirName: string): Integer;
var
  Stream: TStream;
  FileArray: TArray<TZXLSXFileItem>;
  FilesCount: Integer;
  StrArray: TStringDynArray;
  StrCount: Integer;
  RelationsArray: array of TZXLSXRelationsArray;
  RelationsCounts: array of Integer;
  SheetRelations: TZXLSXRelationsArray;
  SheetRelationsCount: Integer;
  RelationsCount: Integer;
  ThemaColor: TIntegerDynArray;
  ThemaColorCount: Integer;
  SheetRelationNumber: Integer;
  i, j, k: Integer;
  s: string;
  b: Boolean;
  _no_sheets: Boolean;
  MaximumDigitWidth: double;
  RH: TZEXLSXReadHelper;
  // Пытается прочитать rel для листа
  // INPUT
  // const fname: string - имя файла листа
  // RETURN
  // boolean - true - прочитал успешно
  function _CheckSheetRelations(const fname: string): Boolean;
  var
    rstream: TStream;
    s: string;
    i, Num: Integer;
    b: Boolean;
  begin
    Result := false;
    SheetRelationsCount := 0;
    Num := -1;
    b := false;
    s := '';
    for i := length(fname) downto 1 do
      if (fname[i] = PathDelim) then
      begin
        Num := i;
        s := fname;
        insert('_rels' + PathDelim, s, Num + 1);
        s := DirName + s + '.rels';
        if (not FileExists(s)) then
          Num := -1;
        break;
      end;

    if (Num > 0) then
    begin
      rstream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
      try
        Result := ZE_XSLXReadRelationships(rstream, SheetRelations, SheetRelationsCount, b, false);
      finally
        rstream.Free();
      end;
    end;
  end; // _CheckSheetRelations

  // Прочитать примечания
  procedure _ReadComments();
  var
    i, l: Integer;
    s: string;
    b: Boolean;
    Stream: TStream;
  begin
    b := false;
    s := '';
    for i := 0 to SheetRelationsCount - 1 do
      if (SheetRelations[i].ftype = TRelationType.rtComments) then
      begin
        s := SheetRelations[i].target;
        b := true;
        break;
      end;

    // Если найдены примечания
    if (b) then
    begin
      l := length(s);
      if (l >= 3) then
        if ((s[1] = '.') and (s[2] = '.')) then
          delete(s, 1, 3);
      b := false;
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtComments) then
          if (pos(s, FileArray[i].name) <> 0) then
            if (FileExists(DirName + FileArray[i].name)) then
            begin
              s := DirName + FileArray[i].name;
              b := true;
              break;
            end;
      // Если файл не найден
      if (not b) then
      begin
        s := DirName + 'xl' + PathDelim + s;
        if FileExists(s) then
          b := true;
      end;

      // Файл с примечаниями таки присутствует!
      if (b) then
      begin
        Stream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
        try
          ZEXSLXReadComments(XMLSS, Stream);
        finally
          Stream.Free();
        end;
      end;
    end;
  end; // _ReadComments

begin
  Result := 0;
  MaximumDigitWidth := 0;
  FilesCount := 0;
  FileArray := nil;

  if (not TDirectory.Exists(DirName)) then
  begin
    Result := -1;
    exit;
  end;

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  Stream := nil;
  RelationsCount := 0;
  ThemaColorCount := 0;
  SheetRelationsCount := 0;
  ThemaColor := nil;
  RH := nil;

  try
    try
      Stream := TFileStream.Create(DirName + '[Content_Types].xml', fmOpenRead or fmShareDenyNone);
      if (not ZEXSLXReadContentTypes(Stream, FileArray, FilesCount)) then
      begin
        Result := 3;
        exit;
      end;

      FreeAndNil(Stream);

      ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      SheetRelationNumber := -1;

      b := false;
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtDoc) then
        begin
          b := true;
          break;
        end;

      RH := TZEXLSXReadHelper.Create();

      if (not b) then
      begin
        s := DirName + '_rels' + PathDelim + '.rels';
        if (FileExists(s)) then
        begin
          SetLength(FileArray, FilesCount + 1);
          s := '/_rels/.rels';
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        s := DirName + 'xl' + PathDelim + '_rels' + PathDelim + 'workbook.xml.rels';
        if (FileExists(s)) then
        begin
          SetLength(FileArray, FilesCount + 1);
          s := '/xl/_rels/workbook.xml.rels';
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      end;

      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtDoc) then
        begin
          SetLength(RelationsArray, RelationsCount + 1);
          SetLength(RelationsCounts, RelationsCount + 1);

          Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZE_XSLXReadRelationships(Stream, RelationsArray[RelationsCount], RelationsCounts[RelationsCount], b,
            true)) then
          begin
            Result := 4;
            exit;
          end;
          if (b) then
          begin
            SheetRelationNumber := RelationsCount;
            for j := 0 to RelationsCounts[RelationsCount] - 1 do
              if (RelationsArray[RelationsCount][j].ftype = TRelationType.rtWorkSheet) then
                for k := 0 to FilesCount - 1 do
                  if (RelationsArray[RelationsCount][j].fileid < 0) then
                    if ((pos(RelationsArray[RelationsCount][j].target, FileArray[k].original)) > 0) then
                    begin
                      RelationsArray[RelationsCount][j].fileid := k;
                      break;
                    end;
          end; // if
          FreeAndNil(Stream);
          inc(RelationsCount);
        end;

      // sharedStrings.xml
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtCoreProp) then
        begin
          FreeAndNil(Stream);
          Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadSharedStrings(Stream, StrArray, StrCount)) then
          begin
            Result := 3;
            exit;
          end;
          break;
        end;

      // тема (если есть)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtVmlDrawing) then
        begin
          FreeAndNil(Stream);
          Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadTheme(Stream, ThemaColor, ThemaColorCount)) then
          begin
            Result := 6;
            exit;
          end;
          break;
        end;

      // стили (styles.xml)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtStyles) then
        begin
          FreeAndNil(Stream);
          Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadStyles(XMLSS, Stream, ThemaColor, ThemaColorCount, MaximumDigitWidth, RH)) then
          begin
            Result := 5;
            exit;
          end
          else
            b := true;
        end;

      // чтение страниц
      _no_sheets := true;
      if (SheetRelationNumber > 0) then
      begin
        for i := 0 to FilesCount - 1 do
          if (FileArray[i].ftype = TRelationType.rtSharedStr) then
          begin
            FreeAndNil(Stream);
            Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
            if (not ZEXSLXReadWorkBook(XMLSS, Stream, RelationsArray[SheetRelationNumber],
              RelationsCounts[SheetRelationNumber])) then
            begin
              Result := 3;
              exit;
            end;
            break;
          end; // if

        // for i := 1 to RelationsCounts[SheetRelationNumber] do
        XLSXSortRelationArray(RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
          if (RelationsArray[SheetRelationNumber][j].sheetid > 0) then
          begin
            b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
            FreeAndNil(Stream);
            Stream := TFileStream.Create(DirName + FileArray[RelationsArray[SheetRelationNumber][j].fileid].name,
              fmOpenRead or fmShareDenyNone);
            if (not ZEXSLXReadSheet(XMLSS, Stream, RelationsArray[SheetRelationNumber][j].name, StrArray, StrCount,
              SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
              Result := Result or 4;
            if (b) then
              _ReadComments();
            _no_sheets := false;
          end; // if
      end;
      // если прочитано 0 листов - пробуем прочитать все (не удалось прочитать workbook/rel)
      if (_no_sheets) then
        for i := 0 to FilesCount - 1 do
          if (FileArray[i].ftype = TRelationType.rtWorkSheet) then
          begin
            b := _CheckSheetRelations(FileArray[i].name);
            FreeAndNil(Stream);
            Stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
            if (not ZEXSLXReadSheet(XMLSS, Stream, '', StrArray, StrCount, SheetRelations, SheetRelationsCount,
              MaximumDigitWidth, RH)) then
              Result := Result or 4;
            if (b) then
              _ReadComments();
          end;
    except
      Result := 2;
    end;
  finally
    if (Assigned(Stream)) then
      FreeAndNil(Stream);

    SetLength(FileArray, 0);
    FileArray := nil;
    SetLength(StrArray, 0);
    StrArray := nil;
    for i := 0 to RelationsCount - 1 do
    begin
      SetLength(RelationsArray[i], 0);
      RelationsArray[i] := nil;
    end;
    SetLength(RelationsCounts, 0);
    RelationsCounts := nil;
    SetLength(RelationsArray, 0);
    RelationsArray := nil;
    SetLength(ThemaColor, 0);
    ThemaColor := nil;
    SetLength(SheetRelations, 0);
    SheetRelations := nil;
    if (Assigned(RH)) then
      FreeAndNil(RH);
  end;
end; // ReadXLSXPath

function ReadXLSXFile(var XMLSS: TZWorkBook; zipStream: TStream): Integer;
var
  Stream: TStream;
  FileArray: TArray<TZXLSXFileItem>;
  FileList: TList<TZXLSXFileItem>;
  FilesCount: Integer;
  StrArray: TStringDynArray;
  StrCount: Integer;
  RelationsArray: array of TZXLSXRelationsArray;
  RelationsCounts: array of Integer;
  SheetRelations: TZXLSXRelationsArray;
  SheetRelationsCount: Integer;
  RelationsCount: Integer;
  ThemaColor: TIntegerDynArray;
  ThemaColorCount: Integer;
  SheetRelationNumber: Integer;
  i, j, k: Integer;
  s: string;
  b: Boolean;
  _no_sheets: Boolean;
  RH: TZEXLSXReadHelper;
  Zip: TZipFile;
  encoding: TEncoding;
  zipHdr: TZipHeader;
  buff: TBytes;
  MaximumDigitWidth: double;
  // zfiles: TArray<string>;

  function _CheckSheetRelations(const fname: string): Boolean;
  var
    rstream: TStream;
    s: string;
    i, Num: Integer;
    b: Boolean;
  begin
    Result := false;
    SheetRelationsCount := 0;
    Num := -1;
    b := false;
    s := '';
    for i := length(fname) downto 1 do
    begin
      if (fname[i] = PathDelim) then
      begin
        Num := i;
        s := fname;
        insert('_rels' + PathDelim, s, Num + 1);
        s := s + '.rels';
        if (not FileExists(s)) then
          Num := -1;
        break;
      end;
    end;

    if (Num > 0) then
    begin
      rstream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
      try
        Result := ZE_XSLXReadRelationships(rstream, SheetRelations, SheetRelationsCount, b, false);
      finally
        rstream.Free();
      end;
    end;
  end; // _CheckSheetRelations

  // Прочитать примечания
  procedure _ReadComments();
  var
    i, l: Integer;
    s: string;
    b: Boolean;
    _stream: TStream;
  begin
    b := false;
    s := '';
    _stream := nil;
    for i := 0 to SheetRelationsCount - 1 do
      if (SheetRelations[i].ftype = TRelationType.rtComments) then
      begin
        s := SheetRelations[i].target;
        b := true;
        break;
      end;

    // Если найдены примечания
    if (b) then
    begin
      l := length(s);
      if (l >= 3) then
        if ((s[1] = '.') and (s[2] = '.')) then
          delete(s, 1, 3);
      b := false;
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtComments) then
          if (pos(s, FileArray[i].name) <> 0) then
            if (FileExists(FileArray[i].name)) then
            begin
              s := FileArray[i].name;
              b := true;
              break;
            end;
      // Если файл не найден
      if (not b) then
      begin
        s := 'xl' + PathDelim + s;
        if (FileExists(s)) then
          b := true;
      end;

      // Файл с примечаниями таки присутствует!
      if (b) then
        try
          _stream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
          ZEXSLXReadComments(XMLSS, _stream);
        finally
          if (Assigned(_stream)) then
            FreeAndNil(_stream);
        end;
    end;
  end; // _ReadComments

begin
  Result := 0;
  MaximumDigitWidth := 0;
  FilesCount := 0;
  Zip := TZipFile.Create();
  encoding := TEncoding.GetEncoding(437);
{$IFDEF VER330}
  Zip.encoding := encoding;
{$ENDIF}
  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  RelationsCount := 0;
  ThemaColorCount := 0;
  SheetRelationsCount := 0;
  RH := TZEXLSXReadHelper.Create();
  FileList := TList<TZXLSXFileItem>.Create();
  Stream := nil;
  try
    Zip.Open(zipStream, zmRead);
    try
      Zip.Read('[Content_Types].xml', Stream, zipHdr);
      try
        if (not ZEXSLXReadContentTypes(Stream, FileArray, FilesCount)) then
          raise Exception.Create('Could not read [Content_Types].xml');
      finally
        FreeAndNil(Stream);
      end;

      ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      SheetRelationNumber := -1;

      b := false;
      for i := 0 to FilesCount - 1 do
      begin
        if (FileArray[i].ftype = TRelationType.rtDoc) then
        begin
          b := true;
          break;
        end;
      end;

      if (not b) then
      begin
        s := '/_rels/.rels';
        if Zip.IndexOf(s.Substring(1)) > -1 then
        begin
          SetLength(FileArray, FilesCount + 1);
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        s := '/xl/_rels/workbook.xml.rels';
        if Zip.IndexOf(s.Substring(1)) > -1 then
        begin
          SetLength(FileArray, FilesCount + 1);
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      end;

      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtDoc) then
        begin
          SetLength(RelationsArray, RelationsCount + 1);
          SetLength(RelationsCounts, RelationsCount + 1);
          Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
          try
            if (not ZE_XSLXReadRelationships(Stream, RelationsArray[RelationsCount], RelationsCounts[RelationsCount], b,
              true)) then
            begin
              Result := 4;
              exit;
            end;
            if (b) then
            begin
              SheetRelationNumber := RelationsCount;
              for j := 0 to RelationsCounts[RelationsCount] - 1 do
                if (RelationsArray[RelationsCount][j].ftype = TRelationType.rtWorkSheet) then
                  for k := 0 to FilesCount - 1 do
                    if (RelationsArray[RelationsCount][j].fileid < 0) then
                      if ((pos(RelationsArray[RelationsCount][j].target, FileArray[k].original)) > 0) then
                      begin
                        RelationsArray[RelationsCount][j].fileid := k;
                        break;
                      end;
            end; // if
          finally
            FreeAndNil(Stream);
          end;
          inc(RelationsCount);
        end;

      // sharedStrings.xml
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtCoreProp) then
        begin
          Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
          try
            if (not ZEXSLXReadSharedStrings(Stream, StrArray, StrCount)) then
            begin
              Result := 3;
              exit;
            end;
          finally
            FreeAndNil(Stream);
          end;
          break;
        end;

      // тема (если есть)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtVmlDrawing) then
        begin
          Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
          try
            if (not ZEXSLXReadTheme(Stream, ThemaColor, ThemaColorCount)) then
            begin
              Result := 6;
              exit;
            end;
          finally
            FreeAndNil(Stream);
          end;
          break;
        end;

      // стили (styles.xml)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtStyles) then
        begin
          Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
          try
            if (not ZEXSLXReadStyles(XMLSS, Stream, ThemaColor, ThemaColorCount, MaximumDigitWidth, RH)) then
            begin
              Result := 5;
              exit;
            end
            else
              b := true;
          finally
            FreeAndNil(Stream);
          end;
        end;

      // чтение страниц
      _no_sheets := true;
      if (SheetRelationNumber > 0) then
      begin
        for i := 0 to FilesCount - 1 do
          if (FileArray[i].ftype = TRelationType.rtSharedStr) then
          begin
            Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
            try
              if (not ZEXSLXReadWorkBook(XMLSS, Stream, RelationsArray[SheetRelationNumber],
                RelationsCounts[SheetRelationNumber])) then
              begin
                Result := 3;
                exit;
              end;
            finally
              FreeAndNil(Stream);
            end;
            break;
          end; // if

        // for i := 1 to RelationsCounts[SheetRelationNumber] do
        XLSXSortRelationArray(RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
          if (RelationsArray[SheetRelationNumber][j].sheetid > 0) then
          begin
            b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
            Zip.Read(FileArray[RelationsArray[SheetRelationNumber][j].fileid].original.Substring(1), Stream, zipHdr);
            try
              if (not ZEXSLXReadSheet(XMLSS, Stream, RelationsArray[SheetRelationNumber][j].name, StrArray, StrCount,
                SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
                Result := Result or 4;
              if (b) then
                _ReadComments();
            finally
              FreeAndNil(Stream);
            end;
            _no_sheets := false;
          end; // if
      end;

      // если прочитано 0 листов - пробуем прочитать все (не удалось прочитать workbook/rel)
      if _no_sheets then
      begin
        for i := 0 to FilesCount - 1 do
        begin
          if FileArray[i].ftype = TRelationType.rtWorkSheet then
          begin
            b := _CheckSheetRelations(FileArray[i].name);
            Zip.Read(FileArray[i].original.Substring(1), Stream, zipHdr);
            try
              if (not ZEXSLXReadSheet(XMLSS, Stream, '', StrArray, StrCount, SheetRelations, SheetRelationsCount,
                MaximumDigitWidth, RH)) then
                Result := Result or 4;
              if (b) then
                _ReadComments();
            finally
              FreeAndNil(Stream);
            end;
          end;
        end;
      end;

      // drawings
      for i := 0 to XMLSS.Sheets.Count - 1 do
      begin
        if XMLSS.Sheets[i].DrawingRid > 0 then
        begin
          // load images
          s := 'xl/drawings/drawing' + IntToStr(i + 1) + '.xml';
          Zip.Read(s, Stream, zipHdr);
          try
            ZEXLSXReadDrawing(XMLSS.Sheets[i], Stream);
          finally
            Stream.Free();
          end;

          // read drawing rels
          s := 'xl/drawings/_rels/drawing' + IntToStr(i + 1) + '.xml.rels';
          Zip.Read(s, Stream, zipHdr);
          try
            ZEXLSXReadDrawingRels(XMLSS.Sheets[i], Stream);
          finally
            Stream.Free();
          end;

          // read img file
          for j := 0 to XMLSS.Sheets[i].Drawing.Count - 1 do
          begin
            s := XMLSS.Sheets[i].Drawing[j].name;
            Zip.Read('xl/media/' + s, buff);
              // only unique content
            XMLSS.AddMediaContent(s, buff, true);
          end;
        end;
      end;
    except
      Result := 2;
    end;
  finally
    Zip.Free();
    encoding.Free;
    FileList.Free();
    RH.Free();
  end;
end; // ReadXLSXPath

/// //////////////////////////////////////////////////////////////////////////
/// //////////                    запись                         /////////////
/// //////////////////////////////////////////////////////////////////////////

// Создаёт [Content_Types].xml
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// PageCount: integer                - кол-во страниц
// CommentCount: integer             - кол-во страниц с комментариями
// const PagesComments: TIntegerDynArray- номера страниц с комментариями (нумеряция с нуля)
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// const WriteHelper: TZEXLSXWriteHelper - additional data
// RETURN
// integer
function ZEXLSXCreateContentTypes(var XMLSS: TZWorkBook; Stream: TStream; PageCount: Integer; CommentCount: Integer;
  const PagesComments: TIntegerDynArray; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring;
  const WriteHelper: TZEXLSXWriteHelper): Integer;
var
  Xml: TZsspXMLWriterH;
  s: string;
  procedure _WriteOverride(const PartName: string; ct: Integer);
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('PartName', PartName);
    case ct of
      0:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
      1:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
      2:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
      3:
        s := 'application/vnd.openxmlformats-package.relationships+xml';
      4:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
      5:
        s := 'application/vnd.openxmlformats-package.core-properties+xml';
      6:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
      7:
        s := 'application/vnd.openxmlformats-officedocument.vmlDrawing';
      8:
        s := 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
      9:
        s := 'application/vnd.openxmlformats-officedocument.drawing+xml';
    end;
    Xml.Attributes.Add('ContentType', s, false);
    Xml.WriteEmptyTag('Override', true);
  end; // _WriteOverride

  procedure _WriteTypeDefault(extension, contentType: string);
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('Extension', extension);
    Xml.Attributes.Add('ContentType', contentType, false);
    Xml.WriteEmptyTag('Default', true);
  end;

  procedure _WriteTypes();
  var
    i: Integer;
  begin
    _WriteTypeDefault('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    _WriteTypeDefault('xml', 'application/xml');
    _WriteTypeDefault('png', 'image/png');
    _WriteTypeDefault('jpeg', 'image/jpeg');
    _WriteTypeDefault('wmf', 'image/x-wmf');

    // Страницы
    // _WriteOverride('/_rels/.rels', 3);
    // _WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for i := 0 to PageCount - 1 do
    begin
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
      if (WriteHelper.IsSheetHaveHyperlinks(i)) then
        _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels', 3);
    end;
    // комментарии
    for i := 0 to CommentCount - 1 do
    begin
      _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
      _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
    end;

    for i := 0 to XMLSS.Sheets.Count - 1 do
    begin
      if Assigned(XMLSS.Sheets[i].Drawing) and (XMLSS.Sheets[i].Drawing.Count > 0) then
      begin
        _WriteOverride('/xl/drawings/drawing' + IntToStr(i + 1) + '.xml', 9);
        // _WriteOverride('/xl/drawings/_rels/drawing' + IntToStr(i+1) + '.xml.rels', 3);
// for ii := 0 to _drawing.PictureStore.Count - 1 do begin
// _picture := _drawing.PictureStore.Items[ii];
// // image/ override
// xml.Attributes.Clear();
// xml.Attributes.Add('PartName', '/xl/media/' + _picture.Name);
// xml.Attributes.Add('ContentType', 'image/' + Copy(ExtractFileExt(_picture.Name), 2, 99), false);
// xml.WriteEmptyTag('Override', true);
// end;
      end;
    end;

    _WriteOverride('/xl/workbook.xml', 2);
    _WriteOverride('/xl/styles.xml', 1);
    _WriteOverride('/xl/sharedStrings.xml', 4);
    _WriteOverride('/docProps/app.xml', 8);
    _WriteOverride('/docProps/core.xml', 5);
  end; // _WriteTypes

begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE + '/content-types');
    Xml.WriteTagNode('Types', true, true, true);
    _WriteTypes();
    Xml.WriteEndTagNode(); // Types
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateContentTypes

function ZEXLSXCreateDrawing(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  pic: TZEPicture;
  i: Integer;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.NewLine := false;
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
    Xml.Attributes.Add('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    Xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    Xml.WriteTagNode('xdr:wsDr', false, false, false);

    for i := 0 to sheet.Drawing.Count - 1 do
    begin
      pic := sheet.Drawing.Items[i];
      // cell anchor
      Xml.Attributes.Clear();
      if pic.CellAnchor = ZAAbsolute then
        Xml.Attributes.Add('editAs', 'absolute')
      else
        Xml.Attributes.Add('editAs', 'oneCell');
      Xml.WriteTagNode('xdr:twoCellAnchor', false, false, false);

      // - xdr:from
      Xml.Attributes.Clear();
      Xml.WriteTagNode('xdr:from', false, false, false);
      Xml.WriteTag('xdr:col', IntToStr(pic.FromCol), false, false);
      Xml.WriteTag('xdr:colOff', IntToStr(pic.FromColOff), false, false);
      Xml.WriteTag('xdr:row', IntToStr(pic.FromRow), false, false);
      Xml.WriteTag('xdr:rowOff', IntToStr(pic.FromRowOff), false, false);
      Xml.WriteEndTagNode(); // xdr:from
      // - xdr:to
      Xml.Attributes.Clear();
      Xml.WriteTagNode('xdr:to', false, false, false);
      Xml.WriteTag('xdr:col', IntToStr(pic.ToCol), false, false);
      Xml.WriteTag('xdr:colOff', IntToStr(pic.ToColOff), false, false);
      Xml.WriteTag('xdr:row', IntToStr(pic.ToRow), false, false);
      Xml.WriteTag('xdr:rowOff', IntToStr(pic.ToRowOff), false, false);
      Xml.WriteEndTagNode(); // xdr:from
      // - xdr:pic
      Xml.WriteTagNode('xdr:pic', false, false, false);
      // -- xdr:nvPicPr
      Xml.WriteTagNode('xdr:nvPicPr', false, false, false);
      // --- xdr:cNvPr
      Xml.Attributes.Clear();
      Xml.Attributes.Add('descr', pic.Description);
      Xml.Attributes.Add('name', pic.Title);
      Xml.Attributes.Add('id', IntToStr(pic.id)); // 1
      Xml.WriteEmptyTag('xdr:cNvPr', false);
      // --- xdr:cNvPicPr
      Xml.Attributes.Clear();
      Xml.WriteEmptyTag('xdr:cNvPicPr', false);
      Xml.WriteEndTagNode(); // -- xdr:nvPicPr

      // -- xdr:blipFill
      Xml.Attributes.Clear();
      Xml.WriteTagNode('xdr:blipFill', false, false, false);
      // --- a:blip
      Xml.Attributes.Clear();
      Xml.Attributes.Add('r:embed', 'rId' + IntToStr(pic.RelId)); // rId1
      Xml.WriteEmptyTag('a:blip', false);
      // --- a:stretch
      Xml.Attributes.Clear();
      Xml.WriteEmptyTag('a:stretch', false);
      Xml.WriteEndTagNode(); // -- xdr:blipFill

      // -- xdr:spPr
      Xml.Attributes.Clear();
      Xml.WriteTagNode('xdr:spPr', false, false, false);
      // --- a:xfrm
      Xml.WriteTagNode('a:xfrm', false, false, false);
      // ----
      Xml.Attributes.Clear();
      Xml.Attributes.Add('x', IntToStr(pic.FrmOffX));
      Xml.Attributes.Add('y', IntToStr(pic.FrmOffY));
      Xml.WriteEmptyTag('a:off', false);
      // ----
      Xml.Attributes.Clear();
      Xml.Attributes.Add('cx', IntToStr(pic.FrmExtCX));
      Xml.Attributes.Add('cy', IntToStr(pic.FrmExtCY));
      Xml.WriteEmptyTag('a:ext', false);
      Xml.Attributes.Clear();
      Xml.WriteEndTagNode(); // --- a:xfrm

      // --- a:prstGeom
      Xml.Attributes.Clear();
      Xml.Attributes.Add('prst', 'rect');
      Xml.WriteTagNode('a:prstGeom', false, false, false);
      Xml.Attributes.Clear();
      Xml.WriteEmptyTag('a:avLst', false);
      Xml.WriteEndTagNode(); // --- a:prstGeom

      // --- a:ln
      Xml.WriteTagNode('a:ln', false, false, false);
      Xml.WriteEmptyTag('a:noFill', false);
      Xml.WriteEndTagNode(); // --- a:ln

      Xml.WriteEndTagNode(); // -- xdr:spPr

      Xml.WriteEndTagNode(); // - xdr:pic

      Xml.WriteEmptyTag('xdr:clientData', false);

      Xml.WriteEndTagNode(); // xdr:twoCellAnchor
    end;
    Xml.WriteEndTagNode(); // xdr:wsDr
  finally
    Xml.Free();
  end;
end;

function ZEXLSXCreateDrawingRels(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: String; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  i: Integer;
  dic: TDictionary<Integer, string>;
  pair: TPair<Integer, string>;
begin
  Result := 0;
  dic := TDictionary<Integer, string>.Create();
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL, false);
    Xml.WriteTagNode('Relationships', false, false, false);

    for i := 0 to sheet.Drawing.Count - 1 do
    begin
      dic.AddOrSetValue(sheet.Drawing[i].RelId, sheet.Drawing[i].name);
    end;

    for pair in dic do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('Id', 'rId' + IntToStr(pair.Key));
      Xml.Attributes.Add('Type', SCHEMA_DOC_REL + '/image');
      Xml.Attributes.Add('Target', '../media/' + pair.Value);
      Xml.WriteEmptyTag('Relationship', false, true);
    end;
    Xml.WriteEndTagNode(); // Relationships
  finally
    Xml.Free();
    dic.Free();
  end;
end;

// Создаёт лист документа (sheet*.xml)
// INPUT
// XMLSS: TZWorkBook                                       - хранилище
// Stream: TStream                                       - поток для записи
// SheetNum: integer                                     - номер листа в документе
// SharedStrings: TStringDynArray                        - общие строки
// SharedStringsDictionary: TDictionary<string, integer> - словарь для определения дубликатов строк
// TextConverter: TAnsiToCPConverter                     - конвертер из локальной кодировки в нужную
// CodePageName: string                                  - название кодовой страници
// BOM: ansistring                                       - BOM
// isHaveComments: boolean                               - возвращает true, если были комментарии (чтобы создать comments*.xml)
// WriteHelper: TZEXLSXWriteHelper                       - additional data
// RETURN
// integer
function ZEXLSXCreateSheet(var XMLSS: TZWorkBook; Stream: TStream; SheetNum: Integer;
  var SharedStrings: TStringDynArray; const SharedStringsDictionary: TDictionary<string, Integer>;
  TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring;
  const WriteHelper: TZEXLSXWriteHelper): Integer;
var
  Xml: TZsspXMLWriterH; // писатель
  sheet: TZSheet;
  procedure WriteXLSXSheetHeader();
  var
    s: string;
    b: Boolean;
    SheetOptions: TZSheetOptions;
    procedure _AddSplitValue(const SplitMode: TZSplitMode; const SplitValue: Integer; const AttrName: string);
    var
      s: string;
      b: Boolean;
    begin
      s := '0';
      b := true;
      case SplitMode of
        ZSplitFrozen:
          begin
            s := IntToStr(SplitValue);
            if (SplitValue = 0) then
              b := false;
          end;
        ZSplitSplit:
          s := IntToStr(round(PixelToPoint(SplitValue) * 20));
        ZSplitNone:
          b := false;
      end;
      if (b) then
        Xml.Attributes.Add(AttrName, s);
    end; // _AddSplitValue

    procedure _AddTopLeftCell(const VMode: TZSplitMode; const vValue: Integer; const HMode: TZSplitMode;
      const hValue: Integer);
    var
      isProblem: Boolean;
    begin
      isProblem := (VMode = ZSplitSplit) or (HMode = ZSplitSplit);
      isProblem := isProblem or (vValue > 1000) or (hValue > 100);
      if not isProblem then
      begin
        s := ZEGetA1byCol(vValue) + IntToStr(hValue + 1);
        Xml.Attributes.Add('topLeftCell', s);
      end;
    end; // _AddTopLeftCell

  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('filterMode', 'false');
    Xml.WriteTagNode('sheetPr', true, true, false);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(sheet.TabColor));
    Xml.WriteEmptyTag('tabColor', true, false);

    Xml.Attributes.Clear();
    if sheet.ApplyStyles then
      Xml.Attributes.Add('applyStyles', '1');
    if not sheet.SummaryBelow then
      Xml.Attributes.Add('summaryBelow', '0');
    if not sheet.SummaryRight then
      Xml.Attributes.Add('summaryRight', '0');
    Xml.WriteEmptyTag('outlinePr', true, false);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('fitToPage', XLSXBoolToStr(sheet.FitToPage));
    Xml.WriteEmptyTag('pageSetUpPr', true, false);

    Xml.WriteEndTagNode(); // sheetPr

    Xml.Attributes.Clear();
    s := 'A1';
    if (sheet.ColCount > 0) then
      s := s + ':' + ZEGetA1byCol(sheet.ColCount - 1) + IntToStr(sheet.RowCount);
    Xml.Attributes.Add('ref', s);
    Xml.WriteEmptyTag('dimension', true, false);

    Xml.Attributes.Clear();
    Xml.WriteTagNode('sheetViews', true, true, true);

    Xml.Attributes.Add('colorId', '64');
    Xml.Attributes.Add('defaultGridColor', 'true', false);
    Xml.Attributes.Add('rightToLeft', 'false', false);
    Xml.Attributes.Add('showFormulas', 'false', false);
    Xml.Attributes.Add('showGridLines', 'true', false);
    Xml.Attributes.Add('showOutlineSymbols', 'true', false);
    Xml.Attributes.Add('showRowColHeaders', 'true', false);
    Xml.Attributes.Add('showZeros', 'true', false);

    if sheet.Selected then
      Xml.Attributes.Add('tabSelected', 'true', false);

    Xml.Attributes.Add('topLeftCell', 'A1', false);

    if sheet.ViewMode = zvmPageBreakPreview then
      Xml.Attributes.Add('view', 'pageBreakPreview', false)
    else
      Xml.Attributes.Add('view', 'normal', false);

    Xml.Attributes.Add('windowProtection', 'false', false);
    Xml.Attributes.Add('workbookViewId', '0', false);
    Xml.Attributes.Add('zoomScale', '100', false);
    Xml.Attributes.Add('zoomScaleNormal', '100', false);
    Xml.Attributes.Add('zoomScalePageLayoutView', '100', false);
    Xml.WriteTagNode('sheetView', true, true, false);

{$REGION 'write sheetFormatPr'}
    if (sheet.OutlineLevelCol > 0) or (sheet.OutlineLevelRow > 0) then
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('defaultColWidth', FloatToStr(sheet.DefaultColWidth, TFormatSettings.Invariant));
      Xml.Attributes.Add('defaultRowHeight', FloatToStr(sheet.DefaultRowHeight, TFormatSettings.Invariant));
      if (sheet.OutlineLevelCol > 0) then
        Xml.Attributes.Add('outlineLevelCol', IntToStr(sheet.OutlineLevelCol));
      if (sheet.OutlineLevelRow > 0) then
        Xml.Attributes.Add('outlineLevelRow', IntToStr(sheet.OutlineLevelRow));
      Xml.WriteEmptyTag('sheetFormatPr', true, false);
    end;
{$ENDREGION}
    SheetOptions := sheet.SheetOptions;

    b := (SheetOptions.SplitVerticalMode <> ZSplitNone) or (SheetOptions.SplitHorizontalMode <> ZSplitNone);
    if (b) then
    begin
      Xml.Attributes.Clear();
      _AddSplitValue(SheetOptions.SplitVerticalMode, SheetOptions.SplitVerticalValue, 'xSplit');
      _AddSplitValue(SheetOptions.SplitHorizontalMode, SheetOptions.SplitHorizontalValue, 'ySplit');

      _AddTopLeftCell(SheetOptions.SplitVerticalMode, SheetOptions.SplitVerticalValue, SheetOptions.SplitHorizontalMode,
        SheetOptions.SplitHorizontalValue);

      Xml.Attributes.Add('activePane', 'topLeft');

      s := 'split';
      if ((SheetOptions.SplitVerticalMode = ZSplitFrozen) or (SheetOptions.SplitHorizontalMode = ZSplitFrozen)) then
        s := 'frozen';
      Xml.Attributes.Add('state', s);

      Xml.WriteEmptyTag('pane', true, false);
    end; // if
    {
    <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
    activePane (Active Pane) The pane that is active.
                The possible values for this attribute are
                defined by the ST_Pane simple type (§18.18.52).
                  bottomRight	Bottom Right Pane
                  topRight	Top Right Pane
                  bottomLeft	Bottom Left Pane
                  topLeft	Top Left Pane

    state (Split State) Indicates whether the pane has horizontal / vertical
                splits, and whether those splits are frozen.
                The possible values for this attribute are defined by the ST_PaneState simple type (§18.18.53).
                   Split
                   Frozen
                   FrozenSplit

    topLeftCell (Top Left Visible Cell) Location of the top left visible
                cell in the bottom right pane (when in Left-To-Right mode).
                The possible values for this attribute are defined by the
                ST_CellRef simple type (§18.18.7).

    xSplit (Horizontal Split Position) Horizontal position of the split,
                in 1/20th of a point; 0 (zero) if none. If the pane is frozen,
                this value indicates the number of columns visible in the
                top pane. The possible values for this attribute are defined
                by the W3C XML Schema double datatype.

    ySplit (Vertical Split Position) Vertical position of the split, in 1/20th
                of a point; 0 (zero) if none. If the pane is frozen, this
                value indicates the number of rows visible in the left pane.
                The possible values for this attribute are defined by the
                W3C XML Schema double datatype.
 }

    {
    xml.Attributes.Clear();
    xml.Attributes.Add('activePane', 'topLeft');
    xml.Attributes.Add('topLeftCell', 'A1', false);
    xml.Attributes.Add('xSplit', '0', false);
    xml.Attributes.Add('ySplit', '-1', false);
    xml.WriteEmptyTag('pane', true, false);
 }

    {
    _AddSelection('A1', 'bottomLeft');
    _AddSelection('F16', 'topLeft');
 }

    s := ZEGetA1byCol(sheet.SheetOptions.ActiveCol) + IntToStr(sheet.SheetOptions.ActiveRow + 1);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('activeCell', s);
    Xml.Attributes.Add('sqref', s);
    Xml.WriteEmptyTag('selection', true, false);

    Xml.WriteEndTagNode(); // sheetView
    Xml.WriteEndTagNode(); // sheetViews
  end; // WriteXLSXSheetHeader

  procedure WriteXLSXSheetCols();
  var
    i: Integer;
    s: string;
    ProcessedColumn: TZColOptions;
    MaximumDigitWidth: double;
    NumberOfCharacters: double;
    width: real;
  begin
    MaximumDigitWidth := GetMaximumDigitWidth(XMLSS.Styles[0].Font.name, XMLSS.Styles[0].Font.Size);
    // Если совсем нет стилей, пусть будет ошибка.
    Xml.Attributes.Clear();
    Xml.WriteTagNode('cols', true, true, true);
    for i := 0 to sheet.ColCount - 1 do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('collapsed', 'false', false);
      Xml.Attributes.Add('hidden', XLSXBoolToStr(sheet.Columns[i].Hidden), false);
      Xml.Attributes.Add('max', IntToStr(i + 1), false);
      Xml.Attributes.Add('min', IntToStr(i + 1), false);
      s := '0';
      ProcessedColumn := sheet.Columns[i];
      if ((ProcessedColumn.StyleID >= -1) and (ProcessedColumn.StyleID < XMLSS.Styles.Count)) then
        s := IntToStr(ProcessedColumn.StyleID + 1);
      Xml.Attributes.Add('style', s, false);
      // xml.Attributes.Add('width', ZEFloatSeparator(FormatFloat('0.##########', ProcessedColumn.WidthMM * 5.14509803921569 / 10)), false);
      // А.А.Валуев. Формулы расёта ширины взяты здесь - https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
      // А.А.Валуев. Получаем ширину в символах в Excel-е.
      NumberOfCharacters := Trunc((ProcessedColumn.WidthPix - 5) / MaximumDigitWidth * 100 + 0.5) / 100;
      // А.А.Валуев. Конвертируем ширину в символах в ширину для сохранения в файл.
      width := Trunc((NumberOfCharacters * MaximumDigitWidth + 5) / MaximumDigitWidth * 256) / 256;
      Xml.Attributes.Add('width', ZEFloatSeparator(FormatFloat('0.##########', width)), false);
      if ProcessedColumn.AutoFitWidth then
        Xml.Attributes.Add('bestFit', '1', false);
      if sheet.Columns[i].OutlineLevel > 0 then
        Xml.Attributes.Add('outlineLevel', IntToStr(sheet.Columns[i].OutlineLevel));
      Xml.WriteEmptyTag('col', true, false);
    end;
    Xml.WriteEndTagNode(); // cols
  end; // WriteXLSXSheetCols

  procedure WriteXLSXSheetData();
  var
    i, j, n: Integer;
    b: Boolean;
    s: string;
    _r: TZMergeArea;
    strIndex: Integer;
  begin
    Xml.Attributes.Clear();
    Xml.WriteTagNode('sheetData', true, true, true);
    n := sheet.ColCount - 1;
    for i := 0 to sheet.RowCount - 1 do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('collapsed', 'false', false); // ?
      Xml.Attributes.Add('customFormat', 'false', false); // ?
      Xml.Attributes.Add('customHeight', XLSXBoolToStr((abs(sheet.DefaultRowHeight - sheet.Rows[i].Height) > 0.001))
        { 'true' } , false); // ?
      Xml.Attributes.Add('hidden', XLSXBoolToStr(sheet.Rows[i].Hidden), false);
      Xml.Attributes.Add('ht', ZEFloatSeparator(FormatFloat('0.##', sheet.Rows[i].HeightMM * 2.835)), false);
      if sheet.Rows[i].OutlineLevel > 0 then
        Xml.Attributes.Add('outlineLevel', IntToStr(sheet.Rows[i].OutlineLevel), false);
      Xml.Attributes.Add('r', IntToStr(i + 1), false);
      Xml.WriteTagNode('row', true, true, false);
      for j := 0 to n do
      begin
        Xml.Attributes.Clear();
        if (not WriteHelper.isHaveComments) then
          if (sheet.Cell[j, i].Comment > '') then
            WriteHelper.isHaveComments := true;
        b := (sheet.Cell[j, i].Data > '') or (sheet.Cell[j, i].Formula > '');
        s := ZEGetA1byCol(j) + IntToStr(i + 1);

        if (sheet.Cell[j, i].Href <> '') then
          WriteHelper.AddHyperLink(s, sheet.Cell[j, i].Href, sheet.Cell[j, i].HRefScreenTip, 'External');

        Xml.Attributes.Add('r', s);

        if (sheet.Cell[j, i].CellStyle >= -1) and (sheet.Cell[j, i].CellStyle < XMLSS.Styles.Count) then
          s := IntToStr(sheet.Cell[j, i].CellStyle + 1)
        else
          s := '0';
        Xml.Attributes.Add('s', s, false);

        case sheet.Cell[j, i].CellType of
          ZENumber:
            s := 'n';
          ZEDateTime:
            s := 'd'; // ??
          ZEBoolean:
            s := 'b';
          ZEString:
            begin
            // А.А.Валуев Общие строки пишем только, если в строке есть
            // определённые символы. Хотя можно писать и всё подряд.
              if sheet.Cell[j, i].Data.StartsWith(' ') or sheet.Cell[j, i].Data.EndsWith(' ') or
                (sheet.Cell[j, i].Data.IndexOfAny([#10, #13]) >= 0) then
              begin
              // А.А.Валуев С помощью словаря пытаемся находить дубликаты строк.
                if SharedStringsDictionary.ContainsKey(sheet.Cell[j, i].Data) then
                  strIndex := SharedStringsDictionary[sheet.Cell[j, i].Data]
                else
                begin
                  strIndex := length(SharedStrings);
                  insert(sheet.Cell[j, i].Data, SharedStrings, strIndex);
                  SharedStringsDictionary.Add(sheet.Cell[j, i].Data, strIndex);
                end;
                s := 's';
              end
              else
                s := 'str';
            end;
          ZEError:
            s := 'e';
        end;

        // если тип ячейки ZEGeneral, то атрибут опускаем
        if (sheet.Cell[j, i].CellType <> ZEGeneral) and (sheet.Cell[j, i].CellType <> ZEDateTime) then
          Xml.Attributes.Add('t', s, false);

        if (b) then
        begin
          Xml.WriteTagNode('c', true, true, false);
          if (sheet.Cell[j, i].Formula > '') then
          begin
            Xml.Attributes.Clear();
            Xml.Attributes.Add('aca', 'false');
            Xml.WriteTag('f', sheet.Cell[j, i].Formula, true, false, true);
          end;
          if (sheet.Cell[j, i].Data > '') then
          begin
            Xml.Attributes.Clear();
            if s = 's' then
              Xml.WriteTag('v', strIndex.ToString, true, false, true)
            else
              Xml.WriteTag('v', sheet.Cell[j, i].Data, true, false, true);
          end;
          Xml.WriteEndTagNode();
        end
        else
          Xml.WriteEmptyTag('c', true);
      end;
      Xml.WriteEndTagNode(); // row
    end; // for i

    Xml.WriteEndTagNode(); // sheetData

    // autoFilter
    if trim(sheet.AutoFilter) <> '' then
    begin
      Xml.Attributes.Clear;
      Xml.Attributes.Add('ref', sheet.AutoFilter);
      Xml.WriteEmptyTag('autoFilter', true, false);
    end;

    // Merge Cells
    if sheet.MergeCells.Count > 0 then
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('count', IntToStr(sheet.MergeCells.Count));
      Xml.WriteTagNode('mergeCells', true, true, false);
      for i := 0 to sheet.MergeCells.Count - 1 do
      begin
        Xml.Attributes.Clear();
        _r := sheet.MergeCells.Items[i];
        s := ZEGetA1byCol(_r.Left) + IntToStr(_r.Top + 1) + ':' + ZEGetA1byCol(_r.Right) + IntToStr(_r.Bottom + 1);
        Xml.Attributes.Add('ref', s);
        Xml.WriteEmptyTag('mergeCell', true, false);
      end;
      Xml.WriteEndTagNode(); // mergeCells
    end; // if

    WriteHelper.WriteHyperLinksTag(Xml);
  end; // WriteXLSXSheetData

  procedure WriteColontituls();
  begin
    Xml.Attributes.Clear;
    if sheet.SheetOptions.IsDifferentOddEven then
      Xml.Attributes.Add('differentOddEven', '1');
    if sheet.SheetOptions.IsDifferentFirst then
      Xml.Attributes.Add('differentFirst', '1');
    Xml.WriteTagNode('headerFooter', true, true, false);

    Xml.Attributes.Clear;
    Xml.WriteTag('oddHeader', sheet.SheetOptions.Header, true, false, true);
    Xml.WriteTag('oddFooter', sheet.SheetOptions.Footer, true, false, true);

    if sheet.SheetOptions.IsDifferentOddEven then
    begin
      Xml.WriteTag('evenHeader', sheet.SheetOptions.EvenHeader, true, false, true);
      Xml.WriteTag('evenFooter', sheet.SheetOptions.EvenFooter, true, false, true);
    end;
    if sheet.SheetOptions.IsDifferentFirst then
    begin
      Xml.WriteTag('firstHeader', sheet.SheetOptions.FirstPageHeader, true, false, true);
      Xml.WriteTag('firstFooter', sheet.SheetOptions.FirstPageFooter, true, false, true);
    end;

    Xml.WriteEndTagNode(); // headerFooter
  end;

  procedure WriteBreakData(TagName: string; breaks: TArray<Integer>; manV, maxV: string);
  var
    brk: Integer;
  begin
    if length(breaks) > 0 then
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('count', IntToStr(length(breaks)));
      Xml.Attributes.Add('manualBreakCount', IntToStr(length(breaks)));
      Xml.WriteTagNode(TagName, true, true, true);
      for brk in breaks do
      begin
        Xml.Attributes.Clear();
        Xml.Attributes.Add('id', IntToStr(brk));
        Xml.Attributes.Add('man', manV);
        Xml.Attributes.Add('max', maxV);
        Xml.WriteEmptyTag('brk', true, false);
      end;
      Xml.WriteEndTagNode(); // (row|col)Breaks
    end;
  end;

  procedure WriteXLSXSheetFooter();
  var
    s: string;
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('headings', 'false', false);
    Xml.Attributes.Add('gridLines', 'false', false);
    Xml.Attributes.Add('gridLinesSet', 'true', false);
    Xml.Attributes.Add('horizontalCentered', XLSXBoolToStr(sheet.SheetOptions.CenterHorizontal), false);
    Xml.Attributes.Add('verticalCentered', XLSXBoolToStr(sheet.SheetOptions.CenterVertical), false);
    Xml.WriteEmptyTag('printOptions', true, false);

    Xml.Attributes.Clear();
    s := '0.##########';
    Xml.Attributes.Add('left', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginLeft / ZE_MMinInch)), false);
    Xml.Attributes.Add('right', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginRight / ZE_MMinInch)), false);
    Xml.Attributes.Add('top', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginTop / ZE_MMinInch)), false);
    Xml.Attributes.Add('bottom', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginBottom / ZE_MMinInch)
      ), false);
    Xml.Attributes.Add('header', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.HeaderMargins.Height /
      ZE_MMinInch)), false);
    Xml.Attributes.Add('footer', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.FooterMargins.Height /
      ZE_MMinInch)), false);
    Xml.WriteEmptyTag('pageMargins', true, false);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('blackAndWhite', 'false', false);
    Xml.Attributes.Add('cellComments', 'none', false);
    Xml.Attributes.Add('copies', '1', false);
    Xml.Attributes.Add('draft', 'false', false);
    Xml.Attributes.Add('firstPageNumber', '1', false);
    if sheet.SheetOptions.FitToHeight >= 0 then
      Xml.Attributes.Add('fitToHeight', IntToStr(sheet.SheetOptions.FitToHeight), false);

    if sheet.SheetOptions.FitToWidth >= 0 then
      Xml.Attributes.Add('fitToWidth', IntToStr(sheet.SheetOptions.FitToWidth), false);

    Xml.Attributes.Add('horizontalDpi', '300', false);

    // ECMA 376 ed.4 part1 18.18.50: default|portrait|landscape
    Xml.Attributes.Add('orientation', IfThen(sheet.SheetOptions.PortraitOrientation, 'portrait', 'landscape'), false);

    Xml.Attributes.Add('pageOrder', 'downThenOver', false);
    Xml.Attributes.Add('paperSize', IntToStr(sheet.SheetOptions.PaperSize), false);
    if (sheet.SheetOptions.FitToWidth = -1) and (sheet.SheetOptions.FitToWidth = -1) then
      Xml.Attributes.Add('scale', IntToStr(sheet.SheetOptions.ScaleToPercent), false);
    Xml.Attributes.Add('useFirstPageNumber', 'true', false);
    // xml.Attributes.Add('usePrinterDefaults', 'false', false); //do not use!
    Xml.Attributes.Add('verticalDpi', '300', false);
    Xml.WriteEmptyTag('pageSetup', true, false);

    WriteColontituls();

    // <legacyDrawing r:id="..."/>

    // write (row|col)Breaks
    WriteBreakData('rowBreaks', sheet.RowBreaks, '1', '16383');
    WriteBreakData('colBreaks', sheet.ColBreaks, '1', '1048575');
  end; // WriteXLSXSheetFooter

  procedure WriteXLSXSheetDrawings();
  var
    RID: Integer;
  begin
    // drawings
    if (not sheet.Drawing.IsEmpty) then
    begin
      // rels to helper
      RID := WriteHelper.AddDrawing('../drawings/drawing' + IntToStr(sheet.SheetIndex + 1) + '.xml');
      Xml.Attributes.Clear();
      Xml.Attributes.Add('r:id', 'rId' + IntToStr(RID));
      Xml.WriteEmptyTag('drawing');
    end;
  end;

begin
  WriteHelper.Clear();
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    Xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL);
    Xml.WriteTagNode('worksheet', true, true, false);

    sheet := XMLSS.Sheets[SheetNum];
    WriteXLSXSheetHeader();
    if (sheet.ColCount > 0) then
      WriteXLSXSheetCols();
    WriteXLSXSheetData();
    WriteXLSXSheetFooter();
    WriteXLSXSheetDrawings();
    Xml.WriteEndTagNode(); // worksheet
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateSheet

// Создаёт workbook.xml
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// Stream: TStream                   - поток для записи
// const _pages: TIntegerDynArray       - массив страниц
// const _names: TStringDynArray       - массив имён страниц
// PageCount: integer                - кол-во страниц
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateWorkBook(var XMLSS: TZWorkBook; Stream: TStream; const _pages: TIntegerDynArray;
  const _names: TStringDynArray; PageCount: Integer; TextConverter: TAnsiToCPConverter; CodePageName: String;
  BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  i: Integer;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';

    Xml.WriteHeader(CodePageName, BOM);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    Xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    Xml.WriteTagNode('workbook', true, true, true);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('appName', 'ZEXMLSSlib');
    Xml.WriteEmptyTag('fileVersion', true);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('backupFile', 'false');
    Xml.Attributes.Add('showObjects', 'all', false);
    Xml.Attributes.Add('date1904', 'false', false);
    Xml.WriteEmptyTag('workbookPr', true);

    Xml.Attributes.Clear();
    Xml.WriteEmptyTag('workbookProtection', true);

    Xml.WriteTagNode('bookViews', true, true, true);

    Xml.Attributes.Add('activeTab', '0');
    Xml.Attributes.Add('firstSheet', '0', false);
    Xml.Attributes.Add('showHorizontalScroll', 'true', false);
    Xml.Attributes.Add('showSheetTabs', 'true', false);
    Xml.Attributes.Add('showVerticalScroll', 'true', false);
    Xml.Attributes.Add('tabRatio', '600', false);
    Xml.Attributes.Add('windowHeight', '8192', false);
    Xml.Attributes.Add('windowWidth', '16384', false);
    Xml.Attributes.Add('xWindow', '0', false);
    Xml.Attributes.Add('yWindow', '0', false);
    Xml.WriteEmptyTag('workbookView', true);
    Xml.WriteEndTagNode(); // bookViews

    // sheets
    Xml.Attributes.Clear();
    Xml.WriteTagNode('sheets', true, true, true);
    for i := 0 to PageCount - 1 do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('name', _names[i], false);
      Xml.Attributes.Add('sheetId', IntToStr(i + 1), false);
      Xml.Attributes.Add('state', 'visible', false);
      Xml.Attributes.Add('r:id', 'rId' + IntToStr(i + 2), false);
      Xml.WriteEmptyTag('sheet', true);
    end; // for i
    Xml.WriteEndTagNode(); // sheets

    // definedNames
    Xml.Attributes.Clear();
    Xml.WriteTagNode('definedNames', true, true, true);
    for i := 0 to High(XMLSS.FDefinedNames) do
    begin
      Xml.Attributes.Clear();
      Xml.Attributes.Add('localSheetId', IntToStr(XMLSS.FDefinedNames[i].LocalSheetId), false);
      Xml.Attributes.Add('name', XMLSS.FDefinedNames[i].name, false);
      Xml.WriteTag('definedName', XMLSS.FDefinedNames[i].Body);
    end; // for i
    Xml.WriteEndTagNode(); // definedNames

    Xml.Attributes.Clear();
    Xml.Attributes.Add('iterateCount', '100');
    Xml.Attributes.Add('refMode', 'A1', false); // {tut}
    Xml.Attributes.Add('iterate', 'false', false);
    Xml.Attributes.Add('iterateDelta', '0.001', false);
    Xml.WriteEmptyTag('calcPr', true);

    Xml.WriteEndTagNode(); // workbook
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateWorkBook

// Создаёт styles.xml
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateStyles(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH; // писатель
  _FontIndex: TIntegerDynArray; // соответствия шрифтов
  _FillIndex: TIntegerDynArray; // заливки
  _BorderIndex: TIntegerDynArray; // границы
  _StylesCount: Integer;
  _NumFmtIndexes: array of Integer;
  _FmtParser: TNumFormatParser;
  _DateParser: TZDateTimeODSFormatParser;

  // <numFmts> .. </numFmts>
  procedure WritenumFmts();
  var
    kol: Integer;
    i: Integer;
    _nfmt: TZEXLSXNumberFormats;
    _is_dateTime: array of Boolean;
    s: string;
    _count: Integer;
    _idx: array of Integer;
    _fmt: array of string;
    _Style: TZStyle;
    _currSheet: Integer;
    _currRow, _currCol: Integer;
    _sheet: TZSheet;
    _currstylenum: Integer;
    _numfmt_counter: Integer;

    function _GetNumFmt(StyleNum: Integer): Integer;
    var
      i, j, k: Integer;
      b: Boolean;
      _cs, _cr, _cc: Integer;
    begin
      Result := 0;
      _Style := XMLSS.Styles[StyleNum];
      if (_Style.NumberFormatId > 0) and (_Style.NumberFormatId < 164) then
        exit(_Style.NumberFormatId);

      // If cell type is datetime and cell style is empty then need write default NumFmtId = 14.
      if ((trim(_Style.NumberFormat) = '') or (UpperCase(_Style.NumberFormat) = 'GENERAL')) then
      begin
        if (_is_dateTime[StyleNum + 1]) then
          Result := 14
        else
        begin
          b := false;
          _cs := _currSheet;
          for i := _cs to XMLSS.Sheets.Count - 1 do
          begin
            _sheet := XMLSS.Sheets[i];
            _cr := _currRow;
            for j := _cr to _sheet.RowCount - 1 do
            begin
              _cc := _currCol;
              for k := _cc to _sheet.ColCount - 1 do
              begin
                _currstylenum := _sheet[k, j].CellStyle + 1;
                if (_currstylenum >= 0) and (_currstylenum < kol) then
                  if (_sheet[k, j].CellType = ZEDateTime) then
                  begin
                    _is_dateTime[_currstylenum] := true;
                    if (_currstylenum = StyleNum + 1) then
                    begin
                      b := true;
                      break;
                    end;
                  end;
              end; // for k
              _currRow := j + 1;
              _currCol := 0;
              if (b) then
                break;
            end; // for j

            _currSheet := i + 1;
            _currRow := 0;
            _currCol := 0;
            if (b) then
              break;
          end; // for i

          if (b) then
            Result := 14;
        end;
      end // if
      else
      begin
        s := ConvertFormatNativeToXlsx(_Style.NumberFormat, _FmtParser, _DateParser);
        i := _nfmt.FindFormatID(s);
        if (i < 0) then
        begin
          i := _numfmt_counter;
          _nfmt.Format[i] := s;
          inc(_numfmt_counter);

          SetLength(_idx, _count + 1);
          SetLength(_fmt, _count + 1);
          _idx[_count] := i;
          _fmt[_count] := s;

          inc(_count);
        end;
        Result := i;
      end;
    end; // _GetNumFmt

  begin
    kol := XMLSS.Styles.Count + 1;
    SetLength(_NumFmtIndexes, kol);
    SetLength(_is_dateTime, kol);
    for i := 0 to kol - 1 do
      _is_dateTime[i] := false;

    _nfmt := nil;
    _count := 0;

    _numfmt_counter := 164;

    _currSheet := 0;
    _currRow := 0;
    _currCol := 0;

    try
      _nfmt := TZEXLSXNumberFormats.Create();
      for i := -1 to kol - 2 do
        _NumFmtIndexes[i + 1] := _GetNumFmt(i);

      if (_count > 0) then
      begin
        Xml.Attributes.Clear();
        Xml.Attributes.Add('count', IntToStr(_count));
        Xml.WriteTagNode('numFmts', true, true, false);

        for i := 0 to _count - 1 do
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('numFmtId', IntToStr(_idx[i]));
          Xml.Attributes.Add('formatCode', _fmt[i]);
          Xml.WriteEmptyTag('numFmt', true, true);
        end;

        Xml.WriteEndTagNode(); // numFmts
      end;
    finally
      FreeAndNil(_nfmt);
      SetLength(_idx, 0);
      SetLength(_fmt, 0);
      SetLength(_is_dateTime, 0);
    end;
  end; // WritenumFmts

  // Являются ли шрифты стилей одинаковыми
  function _isFontsEqual(const stl1, stl2: TZStyle): Boolean;
  begin
    Result := false;
    if (stl1.Font.Color <> stl2.Font.Color) then
      exit;

    if (stl1.Font.name <> stl2.Font.name) then
      exit;

    if (stl1.Font.Size <> stl2.Font.Size) then
      exit;

    if (stl1.Font.Style <> stl2.Font.Style) then
      exit;

    if stl1.superscript <> stl2.superscript then
      exit;

    if stl1.subscript <> stl2.subscript then
      exit;

    Result := true; // если уж до сюда добрались
  end; // _isFontsEqual

  // Обновляет индексы в массиве
  // INPUT
  // var arr: TIntegerDynArray  - массив
  // cnt: integer          - номер последнего элемента в массиве (начинает с 0)
  // (предполагается, что возникнет ситуация, когда нужно будет использовать только часть массива)
  procedure _UpdateArrayIndex(var arr: TIntegerDynArray; cnt: Integer);
  var
    res: TIntegerDynArray;
    i, j: Integer;
    Num: Integer;
  begin
    // Assert( Length(arr) - cnt = 2, 'Wow! We really may need this parameter!');
    // cnt := Length(arr) - 2;   // get ready to strip it
    SetLength(res, length(arr));

    Num := 0;
    for i := -1 to cnt do
      if (arr[i + 1] = -2) then
      begin
        res[i + 1] := Num;
        for j := i + 1 to cnt do
          if (arr[j + 1] = i) then
            res[j + 1] := Num;
        inc(Num);
      end; // if

    arr := res;
  end; // _UpdateArrayIndex

  // <fonts>...</fonts>
  procedure WriteXLSXFonts();
  var
    i, j, n: Integer;
    _fontCount: Integer;
    fnt: TZFont;
  begin
    _fontCount := 0;
    SetLength(_FontIndex, _StylesCount + 1);
    for i := 0 to _StylesCount do
      _FontIndex[i] := -2;

    for i := -1 to _StylesCount - 1 do
      if (_FontIndex[i + 1] = -2) then
      begin
        inc(_fontCount);
        n := i + 1;
        for j := n to _StylesCount - 1 do
          if (_FontIndex[j + 1] = -2) then
            if (_isFontsEqual(XMLSS.Styles[i], XMLSS.Styles[j])) then
              _FontIndex[j + 1] := i;
      end; // if

    Xml.Attributes.Clear();
    Xml.Attributes.Add('count', IntToStr(_fontCount));
    Xml.WriteTagNode('fonts', true, true, true);

    for i := 0 to _StylesCount do
      if (_FontIndex[i] = -2) then
      begin
        fnt := XMLSS.Styles[i - 1].Font;
        Xml.Attributes.Clear();
        Xml.WriteTagNode('font', true, true, true);

        Xml.Attributes.Clear();
        Xml.Attributes.Add('val', fnt.name);
        Xml.WriteEmptyTag('name', true);

        Xml.Attributes.Clear();
        Xml.Attributes.Add('val', IntToStr(fnt.charset));
        Xml.WriteEmptyTag('charset', true);

        Xml.Attributes.Clear();
        Xml.Attributes.Add('val', FloatToStr(fnt.Size, TFormatSettings.Invariant));
        Xml.WriteEmptyTag('sz', true);

        if (fnt.Color <> TColorRec.cWindowText) then
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(fnt.Color));
          Xml.WriteEmptyTag('color', true);
        end;

        if (TFontStyle.fsBold in fnt.Style) then
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('val', 'true');
          Xml.WriteEmptyTag('b', true);
        end;

        if (TFontStyle.fsItalic in fnt.Style) then
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('val', 'true');
          Xml.WriteEmptyTag('i', true);
        end;

        if (TFontStyle.fsStrikeOut in fnt.Style) then
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('val', 'true');
          Xml.WriteEmptyTag('strike', true);
        end;

        if (TFontStyle.fsUnderline in fnt.Style) then
        begin
          Xml.Attributes.Clear();
          Xml.Attributes.Add('val', 'single');
          Xml.WriteEmptyTag('u', true);
        end;

      // <vertAlign val="superscript"/>
        if XMLSS.Styles[i - 1].superscript then
        begin
          Xml.Attributes.Clear();

          Xml.Attributes.Add('val', 'superscript');
          Xml.WriteEmptyTag('vertAlign', true);
        end

      // <vertAlign val="subscript"/>

        else if XMLSS.Styles[i - 1].subscript then
        begin

          Xml.Attributes.Clear();

          Xml.Attributes.Add('val', 'subscript');
          Xml.WriteEmptyTag('vertAlign', true);
        end;

        Xml.WriteEndTagNode(); // font
      end; // if

    _UpdateArrayIndex(_FontIndex, _StylesCount - 1);

    Xml.WriteEndTagNode(); // fonts
  end; // WriteXLSXFonts

  // Являются ли заливки одинаковыми
  function _isFillsEqual(style1, style2: TZStyle): Boolean;
  begin
    Result := (style1.BGColor = style2.BGColor) and (style1.PatternColor = style2.PatternColor) and
      (style1.CellPattern = style2.CellPattern);
  end; // _isFillsEqual

  procedure _WriteBlankFill(const st: string);
  begin
    Xml.Attributes.Clear();
    Xml.WriteTagNode('fill', true, true, true);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('patternType', st);
    Xml.WriteEmptyTag('patternFill', true, false);
    Xml.WriteEndTagNode(); // fill
  end; // _WriteBlankFill

  // <fills> ... </fills>
  procedure WriteXLSXFills();
  var
    i, j: Integer;
    _fillCount: Integer;
    s: string;
    b: Boolean;
    _tmpColor: TColor;
    _reverse: Boolean;

  begin
    _fillCount := 0;
    SetLength(_FillIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _FillIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
      if (_FillIndex[i + 1] = -2) then
      begin
        inc(_fillCount);
        for j := i + 1 to _StylesCount - 1 do
          if (_FillIndex[j + 1] = -2) then
            if (_isFillsEqual(XMLSS.Styles[i], XMLSS.Styles[j])) then
              _FillIndex[j + 1] := i;
      end; // if

    Xml.Attributes.Clear();
    Xml.Attributes.Add('count', IntToStr(_fillCount + 2));
    Xml.WriteTagNode('fills', true, true, true);

    // по какой-то непонятной причине, если в начале нету двух стилей заливок (none + gray125),
    // в грёбаном 2010-ом офисе глючат заливки (то-ли чтобы сложно было сделать экспорт в xlsx, то-ли
    // кривые руки у мелкомягких программеров). LibreOffice открывает нормально.
    _WriteBlankFill('none');
    _WriteBlankFill('gray125');

    // TODO:
    // ВНИМАНИЕ!!! //{tut}
    // в некоторых случаях fgColor - это цвет заливки (вроде для solid), а в некоторых - bgColor.
    // Потом не забыть разобраться.
    for i := -1 to _StylesCount - 1 do
      if (_FillIndex[i + 1] = -2) then
      begin
        Xml.Attributes.Clear();
        Xml.WriteTagNode('fill', true, true, true);

        case XMLSS.Styles[i].CellPattern of
          ZPSolid:
            s := 'solid';
          ZPNone:
            s := 'none';
          ZPGray125:
            s := 'gray125';
          ZPGray0625:
            s := 'gray0625';
          ZPDiagStripe:
            s := 'darkUp';
          ZPGray50:
            s := 'mediumGray';
          ZPGray75:
            s := 'darkGray';
          ZPGray25:
            s := 'lightGray';
          ZPHorzStripe:
            s := 'darkHorizontal';
          ZPVertStripe:
            s := 'darkVertical';
          ZPReverseDiagStripe:
            s := 'darkDown';
          ZPDiagCross:
            s := 'darkGrid';
          ZPThickDiagCross:
            s := 'darkTrellis';
          ZPThinHorzStripe:
            s := 'lightHorizontal';
          ZPThinVertStripe:
            s := 'lightVertical';
          ZPThinReverseDiagStripe:
            s := 'lightDown';
          ZPThinDiagStripe:
            s := 'lightUp';
          ZPThinHorzCross:
            s := 'lightGrid';
          ZPThinDiagCross:
            s := 'lightTrellis';
        else
          s := 'solid';
        end; // case

        b := (XMLSS.Styles[i].PatternColor <> TColorRec.cWindow) or (XMLSS.Styles[i].BGColor <> TColorRec.cWindow);
        Xml.Attributes.Clear();
        if b and (XMLSS.Styles[i].CellPattern = ZPNone) then
          Xml.Attributes.Add('patternType', 'solid')
        else
          Xml.Attributes.Add('patternType', s);

        if (b) then
          Xml.WriteTagNode('patternFill', true, true, false)
        else
          Xml.WriteEmptyTag('patternFill', true, false);

        _reverse := not(XMLSS.Styles[i].CellPattern in [ZPNone, ZPSolid]);

        if (XMLSS.Styles[i].BGColor <> TColorRec.cWindow) then
        begin
          Xml.Attributes.Clear();
          if (_reverse) then
            _tmpColor := XMLSS.Styles[i].PatternColor
          else
            _tmpColor := XMLSS.Styles[i].BGColor;
          Xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(_tmpColor));
          Xml.WriteEmptyTag('fgColor', true);
        end;

        if (XMLSS.Styles[i].PatternColor <> TColorRec.cWindow) then
        begin
          Xml.Attributes.Clear();
          if (_reverse) then
            _tmpColor := XMLSS.Styles[i].BGColor
          else
            _tmpColor := XMLSS.Styles[i].PatternColor;
          Xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(_tmpColor));
          Xml.WriteEmptyTag('bgColor', true);
        end;

        if (b) then
          Xml.WriteEndTagNode(); // patternFill

        Xml.WriteEndTagNode(); // fill
      end; // if

    _UpdateArrayIndex(_FillIndex, _StylesCount - 1);

    Xml.WriteEndTagNode(); // fills
  end; // WriteXLSXFills();

  // единичная граница
  procedure _WriteBorderItem(StyleNum: Integer; borderNum: TZBordersPos);
  var
    s, s1: string;
    _border: TZBorderStyle;
    n: Integer;
  begin
    Xml.Attributes.Clear();
    case borderNum of
      bpLeft:
        s := 'left';
      bpTop:
        s := 'top';
      bpRight:
        s := 'right';
      bpBottom:
        s := 'bottom';
    else
      s := 'diagonal';
    end;
    _border := XMLSS.Styles[StyleNum].Border[borderNum];
    s1 := '';
    case _border.LineStyle of
      ZEContinuous:
        begin
          if (_border.Weight = 1) then
            s1 := 'thin'
          else if (_border.Weight = 2) then
            s1 := 'medium'
          else
            s1 := 'thick';
        end;
      ZEHair:
        begin
          s1 := 'hair';
        end;
      ZEDash:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashed'
          else if (_border.Weight >= 2) then
            s1 := 'mediumDashed';
        end;
      ZEDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dotted'
          else if (_border.Weight >= 2) then
            s1 := 'mediumDotted';
        end;
      ZEDashDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDot'
          else if (_border.Weight >= 2) then
            s1 := 'mediumDashDot';
        end;
      ZEDashDotDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDotDot'
          else if (_border.Weight >= 2) then
            s1 := 'mediumDashDotDot';
        end;
      ZESlantDashDot:
        begin
          s1 := 'slantDashDot';
        end;
      ZEDouble:
        begin
          s1 := 'double';
        end;
      ZENone:
        begin
        end;
    end; // case

    n := length(s1);

    if (n > 0) then
      Xml.Attributes.Add('style', s1);

    if ((_border.Color <> TColorRec.Black) and (n > 0)) then
    begin
      Xml.WriteTagNode(s, true, true, true);
      Xml.Attributes.Clear();
      Xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(_border.Color));
      Xml.WriteEmptyTag('color', true);
      Xml.WriteEndTagNode();
    end
    else
      Xml.WriteEmptyTag(s, true);
  end; // _WriteBorderItem

  // <borders> ... </borders>
  procedure WriteXLSXBorders();
  var
    i, j: Integer;
    _borderCount: Integer;
    s: string;
  begin
    _borderCount := 0;
    SetLength(_BorderIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _BorderIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
      if (_BorderIndex[i + 1] = -2) then
      begin
        inc(_borderCount);
        for j := i + 1 to _StylesCount - 1 do
          if (_BorderIndex[j + 1] = -2) then
            if (XMLSS.Styles[i].Border.isEqual(XMLSS.Styles[j].Border)) then
              _BorderIndex[j + 1] := i;
      end; // if

    Xml.Attributes.Clear();
    Xml.Attributes.Add('count', IntToStr(_borderCount));
    Xml.WriteTagNode('borders', true, true, true);

    for i := -1 to _StylesCount - 1 do
      if (_BorderIndex[i + 1] = -2) then
      begin
        Xml.Attributes.Clear();
        s := 'false';
        if (XMLSS.Styles[i].Border[bpDiagonalLeft].Weight > 0) and
          (XMLSS.Styles[i].Border[bpDiagonalLeft].LineStyle <> ZENone) then
          s := 'true';
        Xml.Attributes.Add('diagonalDown', s);
        s := 'false';
        if (XMLSS.Styles[i].Border[bpDiagonalRight].Weight > 0) and
          (XMLSS.Styles[i].Border[bpDiagonalRight].LineStyle <> ZENone) then
          s := 'true';
        Xml.Attributes.Add('diagonalUp', s, false);
        Xml.WriteTagNode('border', true, true, true);

        _WriteBorderItem(i, bpLeft);
        _WriteBorderItem(i, bpRight);
        _WriteBorderItem(i, bpTop);
        _WriteBorderItem(i, bpBottom);
        _WriteBorderItem(i, bpDiagonalLeft);
      // _WriteBorderItem(i, bpDiagonalRight);
        Xml.WriteEndTagNode(); // border
      end; // if

    _UpdateArrayIndex(_BorderIndex, _StylesCount - 1);

    Xml.WriteEndTagNode(); // borders
  end; // WriteXLSXBorders

  // Добавляет <xf> ... </xf>
  // INPUT
  // NumStyle: integer - номер стиля
  // isxfId: boolean   - нужно ли добавлять атрибут "xfId"
  // xfId: integer     - значение "xfId"
  procedure _WriteXF(NumStyle: Integer; isxfId: Boolean; xfId: Integer);
  var
    _addalignment: Boolean;
    _Style: TZStyle;
    s: string;
    i: Integer;
    _num: Integer;
  begin
    Xml.Attributes.Clear();
    _Style := XMLSS.Styles[NumStyle];
    _addalignment := _Style.alignment.wrapText or _Style.alignment.VerticalText or (_Style.alignment.Rotate <> 0) or
      (_Style.alignment.indent <> 0) or _Style.alignment.shrinkToFit or (_Style.alignment.vertical <> ZVAutomatic) or
      (_Style.alignment.horizontal <> ZHAutomatic);

    Xml.Attributes.Add('applyAlignment', XLSXBoolToStr(_addalignment));
    Xml.Attributes.Add('applyBorder', 'true', false);
    Xml.Attributes.Add('applyFont', 'true', false);
    Xml.Attributes.Add('applyProtection', 'true', false);
    Xml.Attributes.Add('borderId', IntToStr(_BorderIndex[NumStyle + 1]), false);
    Xml.Attributes.Add('fillId', IntToStr(_FillIndex[NumStyle + 1] + 2), false);
    // +2 т.к. первыми всегда идут 2 левых стиля заливки
    Xml.Attributes.Add('fontId', IntToStr(_FontIndex[NumStyle + 1]), false);

    // ECMA 376 Ed.4:  12.3.20 Styles Part; 17.9.17 numFmt (Numbering Format); 18.8.30 numFmt (Number Format)
    // http://social.msdn.microsoft.com/Forums/sa/oxmlsdk/thread/3919af8c-644b-4d56-be65-c5e1402bfcb6
    if (isxfId) then
      _num := _NumFmtIndexes[NumStyle + 1]
    else
      _num := 0;

    Xml.Attributes.Add('numFmtId', IntToStr(_num) { '164' } , false); // TODO: support formats

    if (_num > 0) then
      Xml.Attributes.Add('applyNumberFormat', '1', false);

    if (isxfId) then
      Xml.Attributes.Add('xfId', IntToStr(xfId), false);

    Xml.WriteTagNode('xf', true, true, true);

    if (_addalignment) then
    begin
      Xml.Attributes.Clear();
      case (_Style.alignment.horizontal) of
        ZHLeft:
          s := 'left';
        ZHRight:
          s := 'right';
        ZHCenter:
          s := 'center';
        ZHFill:
          s := 'fill';
        ZHJustify:
          s := 'justify';
        ZHDistributed:
          s := 'distributed';
        ZHAutomatic:
          s := 'general';
      else
        s := 'general';
        // The standard does not specify a default value for the horizontal attribute.
        // Excel uses a default value of general for this attribute.
        // MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.d
      end; // case
      Xml.Attributes.Add('horizontal', s);
      Xml.Attributes.Add('indent', IntToStr(_Style.alignment.indent), false);
      Xml.Attributes.Add('shrinkToFit', XLSXBoolToStr(_Style.alignment.shrinkToFit), false);

      if _Style.alignment.VerticalText then
        i := 255
      else
        i := ZENormalizeAngle180(_Style.alignment.Rotate);
      Xml.Attributes.Add('textRotation', IntToStr(i), false);

      case (_Style.alignment.vertical) of
        ZVCenter:
          s := 'center';
        ZVTop:
          s := 'top';
        ZVBottom:
          s := 'bottom';
        ZVJustify:
          s := 'justify';
        ZVDistributed:
          s := 'distributed';
      else
        s := 'bottom';
        // The standard does not specify a default value for the vertical attribute.
        // Excel uses a default value of bottom for this attribute.
        // MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.e
      end; // case
      Xml.Attributes.Add('vertical', s, false);
      Xml.Attributes.Add('wrapText', XLSXBoolToStr(_Style.alignment.wrapText), false);
      Xml.WriteEmptyTag('alignment', true);
    end; // if (_addalignment)

    Xml.Attributes.Clear();
    Xml.Attributes.Add('hidden', XLSXBoolToStr(XMLSS.Styles[NumStyle].Protect));
    Xml.Attributes.Add('locked', XLSXBoolToStr(XMLSS.Styles[NumStyle].HideFormula));
    Xml.WriteEmptyTag('protection', true);

    Xml.WriteEndTagNode(); // xf
  end; // _WriteXF

  // <cellStyleXfs> ... </cellStyleXfs> / <cellXfs> ... </cellXfs>
  procedure WriteCellStyleXfs(const TagName: string; isxfId: Boolean);
  var
    i: Integer;
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('count', IntToStr(XMLSS.Styles.Count + 1));
    Xml.WriteTagNode(TagName, true, true, true);
    for i := -1 to XMLSS.Styles.Count - 1 do
    begin
      // Что-то не совсем понятно, какой именно xfId нужно ставить. Пока будет 0 для всех.
      _WriteXF(i, isxfId, 0 { i + 1 } );
    end;
    Xml.WriteEndTagNode(); // cellStyleXfs
  end; // WriteCellStyleXfs

  // <cellStyles> ... </cellStyles>
  procedure WriteCellStyles();
  begin
  end; // WriteCellStyles

begin
  Result := 0;
  _FmtParser := TNumFormatParser.Create();
  _DateParser := TZDateTimeODSFormatParser.Create();
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';

    Xml.WriteHeader(CodePageName, BOM);
    _StylesCount := XMLSS.Styles.Count;

    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    Xml.WriteTagNode('styleSheet', true, true, true);

    WritenumFmts();

    WriteXLSXFonts();
    WriteXLSXFills();
    WriteXLSXBorders();
    // DO NOT remove cellStyleXfs!!!
    WriteCellStyleXfs('cellStyleXfs', false);
    WriteCellStyleXfs('cellXfs', true);
    WriteCellStyles(); // ??

    Xml.WriteEndTagNode(); // styleSheet
  finally
    Xml.Free();
    _FmtParser.Free();
    _DateParser.Free();
    SetLength(_FontIndex, 0);
    SetLength(_FillIndex, 0);
    SetLength(_BorderIndex, 0);
    SetLength(_NumFmtIndexes, 0);
  end;
end; // ZEXLSXCreateStyles

// Добавить Relationship для rels
// INPUT
// xml: TZsspXMLWriterH  - писалка
// const rid: string         - rid
// ridType: integer      - rIdType (0..8)
// const Target: string      -
// const TargetMode: string  -
procedure ZEAddRelsRelation(Xml: TZsspXMLWriterH; const RID: string; ridType: TRelationType; const target: string;
  const TargetMode: string = '');
begin
  Xml.Attributes.Clear();
  Xml.Attributes.Add('Id', RID);
  Xml.Attributes.Add('Type', ZEXLSXGetRelationName(ridType), false);
  Xml.Attributes.Add('Target', target, false);
  if (TargetMode <> '') then
    Xml.Attributes.Add('TargetMode', TargetMode, true);
  Xml.WriteEmptyTag('Relationship', true, true);
end; // ZEAddRelsID

// Создаёт _rels/.rels
// INPUT
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    Xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(Xml, 'rId1', TRelationType.rtDoc, 'xl/workbook.xml');
    ZEAddRelsRelation(Xml, 'rId2', TRelationType.rtExtProps, 'docProps/app.xml');
    ZEAddRelsRelation(Xml, 'rId3', TRelationType.rtCoreProp, 'docProps/core.xml');

    Xml.WriteEndTagNode(); // Relationships
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateRelsMain

// Создаёт xl/_rels/workbook.xml.rels
// INPUT
// PageCount: integer                - кол-во страниц
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateRelsWorkBook(PageCount: Integer; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  i: Integer;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    Xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(Xml, 'rId1', TRelationType.rtStyles, 'styles.xml');

    for i := 0 to PageCount - 1 do
      ZEAddRelsRelation(Xml, 'rId' + IntToStr(i + 2), TRelationType.rtWorkSheet, 'worksheets/sheet' + IntToStr(i + 1)
        + '.xml');

    ZEAddRelsRelation(Xml, 'rId' + IntToStr(PageCount + 2), TRelationType.rtSharedStr, 'sharedStrings.xml');
    Xml.WriteEndTagNode(); // Relationships
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateRelsWorkBook

// Создаёт sharedStrings.xml
// INPUT
// XMLSS: TZWorkBook                   - хранилище
// Stream: TStream                   - поток для записи
// SharedStrings: TStringDynArray    - общие строки
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateSharedStrings(var XMLSS: TZWorkBook; Stream: TStream; const SharedStrings: TStringDynArray;
  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  i, Count: Integer;
  str: string;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Count := length(SharedStrings);
    Xml.Attributes.Add('count', Count.ToString);
    Xml.Attributes.Add('uniqueCount', Count.ToString, false);
    Xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN, false);
    Xml.WriteTagNode('sst', true, true, false);

    { - Write out the content of Shared Strings: <si><t>Value</t></si> }
    for i := 0 to Pred(Count) do
    begin
      Xml.Attributes.Clear();
      Xml.WriteTagNode('si', false, false, false);
      str := SharedStrings[i];
      Xml.Attributes.Clear();
      if str.StartsWith(' ') or str.EndsWith(' ') then
        // А.А.Валуев Чтобы ведущие и последние пробелы не терялись,
        // добавляем атрибут xml:space="preserve".
        Xml.Attributes.Add('xml:space', 'preserve', false);
      Xml.WriteTag('t', str);
      Xml.WriteEndTagNode();
    end;

    Xml.WriteEndTagNode(); // Relationships
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateSharedStrings

// Создаёт app.xml
// INPUT
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_DOC + '/extended-properties');
    Xml.Attributes.Add('xmlns:vt', SCHEMA_DOC + '/docPropsVTypes', false);
    Xml.WriteTagNode('Properties', true, true, false);

    Xml.Attributes.Clear();
    Xml.WriteTag('TotalTime', '0', true, false, false);
    Xml.WriteTag('Application', ZE_XLSX_APPLICATION, true, false, true);
    Xml.WriteEndTagNode(); // Properties
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateDocPropsApp

// Создаёт app.xml
// INPUT
// var XMLSS: TZWorkBook                 - хранилище
// Stream: TStream                   - поток для записи
// TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
// CodePageName: string              - название кодовой страници
// BOM: ansistring                   - BOM
// RETURN
// integer
function ZEXLSXCreateDocPropsCore(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): Integer;
var
  Xml: TZsspXMLWriterH;
  creationDate: string;
begin
  Result := 0;
  Xml := TZsspXMLWriterH.Create(Stream);
  try
    Xml.TabLength := 1;
    Xml.TextConverter := TextConverter;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader(CodePageName, BOM);
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns:cp', SCHEMA_PACKAGE + '/metadata/core-properties');
    Xml.Attributes.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
    Xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    Xml.Attributes.Add('xmlns:dcterms', 'http://purl.org/dc/terms/', false);
    Xml.Attributes.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
    Xml.WriteTagNode('cp:coreProperties', true, true, false);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    creationDate := ZEDateTimeToStr(XMLSS.DocumentProperties.Created) + 'Z';
    Xml.WriteTag('dcterms:created', creationDate, true, false, false);
    Xml.WriteTag('dcterms:modified', creationDate, true, false, false);

    Xml.Attributes.Clear();
    Xml.WriteTag('cp:revision', '1', true, false, false);

    Xml.WriteEndTagNode(); // cp:coreProperties
  finally
    Xml.Free();
  end;
end; // ZEXLSXCreateDocPropsCore

// Сохраняет незапакованный документ в формате Office Open XML (OOXML)
// INPUT
// var XMLSS: TZWorkBook                   - хранилище
// PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
// const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
// const SheetsNames: array of string    - массив названий страниц
// (количество элементов в двух массивах должны совпадать)
// TextConverter: TAnsiToCPConverter - конвертер
// CodePageName: string              - имя кодировки
// BOM: ansistring                   - Byte Order Mark
// RETURN
// integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = '')
  : Integer; overload;
var
  _pages: TIntegerDynArray; // номера страниц
  _names: TStringDynArray; // названия страниц
  kol, i { , ii } : Integer;
  Stream: TStream;
  _WriteHelper: TZEXLSXWriteHelper;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  s: string;
  SharedStrings: TStringDynArray;
  SharedStringsDictionary: TDictionary<string, Integer>;
  // iDrawingsCount: Integer;
  // path_draw, path_draw_rel, path_media: string;
  // _drawing: TZEDrawing;
  // _pic: TZEPicture;
begin
  Result := 0;
  Stream := nil;
  _WriteHelper := nil;
  kol := 0;
  SharedStrings := [];
  SharedStringsDictionary := TDictionary<string, Integer>.Create;
  try
    if (not TDirectory.Exists(PathName)) then
    begin
      Result := 3;
      exit;
    end;

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
    begin
      Result := 2;
      exit;
    end;

    _WriteHelper := TZEXLSXWriteHelper.Create();

    path_xl := TPath.Combine(PathName, 'xl') + PathDelim;
    if (not DirectoryExists(path_xl)) then
      ForceDirectories(path_xl);

    // styles
    Stream := TFileStream.Create(path_xl + 'styles.xml', fmCreate);
    try
      ZEXLSXCreateStyles(XMLSS, Stream, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    // sharedStrings.xml
    Stream := TFileStream.Create(path_xl + 'sharedStrings.xml', fmCreate);
    try
      ZEXLSXCreateSharedStrings(XMLSS, Stream, SharedStrings, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    // _rels/.rels
    path_relsmain := PathName + PathDelim + '_rels' + PathDelim;
    if (not DirectoryExists(path_relsmain)) then
      ForceDirectories(path_relsmain);
    Stream := TFileStream.Create(path_relsmain + '.rels', fmCreate);
    try
      ZEXLSXCreateRelsMain(Stream, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    // xl/_rels/workbook.xml.rels
    path_relsw := path_xl + '_rels' + PathDelim;
    if (not DirectoryExists(path_relsw)) then
      ForceDirectories(path_relsw);
    Stream := TFileStream.Create(path_relsw + 'workbook.xml.rels', fmCreate);
    try
      ZEXLSXCreateRelsWorkBook(kol, Stream, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    path_sheets := path_xl + 'worksheets' + PathDelim;
    if (not DirectoryExists(path_sheets)) then
      ForceDirectories(path_sheets);

    // iDrawingsCount := XMLSS.DrawingCount();
    // sheets of workbook
    for i := 0 to kol - 1 do
    begin
      Stream := TFileStream.Create(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml', fmCreate);
      try
        ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], SharedStrings, SharedStringsDictionary, TextConverter, CodePageName,
          BOM, _WriteHelper);
      finally
        FreeAndNil(Stream);
      end;

      if (_WriteHelper.HyperLinksCount > 0) then
      begin
        _WriteHelper.AddSheetHyperlink(i);
        s := path_sheets + '_rels' + PathDelim;
        if (not DirectoryExists(s)) then
          ForceDirectories(s);
        Stream := TFileStream.Create(s + 'sheet' + IntToStr(i + 1) + '.xml.rels', fmCreate);
        try
          _WriteHelper.CreateSheetRels(Stream, TextConverter, CodePageName, BOM);
        finally
          FreeAndNil(Stream);
        end;
      end;
    end; // for i

    // iDrawingsCount := XMLSS.DrawingCount();
// if iDrawingsCount <> 0 then begin
// path_draw := path_xl + 'drawings' + PathDelim;
// if (not DirectoryExists(path_draw)) then
// ForceDirectories(path_draw);
//
// path_draw_rel := path_draw + '_rels' + PathDelim;
// if (not DirectoryExists(path_draw_rel)) then
// ForceDirectories(path_draw_rel);
//
// path_media := path_xl + 'media' + PathDelim;
// if (not DirectoryExists(path_media)) then
// ForceDirectories(path_media);
//
// for i := 0 to iDrawingsCount - 1 do begin
// _drawing := XMLSS.GetDrawing(i);
// // drawings/drawingN.xml
// Stream := TFileStream.Create(path_draw + 'drawing' + IntToStr(_drawing.Id) + '.xml', fmCreate);
// try
// ZEXLSXCreateDrawing(XMLSS, Stream, _drawing, TextConverter, CodePageName, BOM);
// finally
// FreeAndNil(Stream);
// end;
//
// // drawings/_rels/drawingN.xml.rels
// Stream := TFileStream.Create(path_draw_rel + 'drawing' + IntToStr(i + 1) + '.xml.rels', fmCreate);
// try
// ZEXLSXCreateDrawingRels(XMLSS, Stream, _drawing, TextConverter, CodePageName, BOM);
// finally
// FreeAndNil(Stream);
// end;
//
// // media/imageN.png
// for ii := 0 to _drawing.PictureStore.Count - 1 do begin
// _pic := _drawing.PictureStore[ii];
// if not Assigned(_pic.DataStream) then Continue;
// Stream := TFileStream.Create(path_media + _pic.Name, fmCreate);
// try
// _pic.DataStream.Position := 0;
// Stream.CopyFrom(_pic.DataStream, _pic.DataStream.Size);
// finally
// FreeAndNil(Stream);
// end;
// end;
// end;
// end;

    // workbook.xml - list of shhets
    Stream := TFileStream.Create(path_xl + 'workbook.xml', fmCreate);
    try
      ZEXLSXCreateWorkBook(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    // [Content_Types].xml
    Stream := TFileStream.Create(TPath.Combine(PathName, '[Content_Types].xml'), fmCreate);
    try
      ZEXLSXCreateContentTypes(XMLSS, Stream, kol, 0, nil, TextConverter, CodePageName, BOM, _WriteHelper);
    finally
      FreeAndNil(Stream);
    end;

    path_docprops := TPath.Combine(PathName, 'docProps') + PathDelim;

    if (not DirectoryExists(path_docprops)) then
      ForceDirectories(path_docprops);

    // docProps/app.xml
    Stream := TFileStream.Create(path_docprops + 'app.xml', fmCreate);
    try
      ZEXLSXCreateDocPropsApp(Stream, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    // docProps/core.xml
    Stream := TFileStream.Create(path_docprops + 'core.xml', fmCreate);
    try
      ZEXLSXCreateDocPropsCore(XMLSS, Stream, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;
  finally
    ZESClearArrays(_pages, _names);
    if (Assigned(Stream)) then
      FreeAndNil(Stream);
    FreeAndNil(_WriteHelper);
    SharedStringsDictionary.Free;
  end;
end; // SaveXmlssToXLSXPath

// SaveXmlssToXLSXPath
// Сохраняет незапакованный документ в формате Office Open XML (OOXML)
// INPUT
// var XMLSS: TZWorkBook                   - хранилище
// PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
// const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
// const SheetsNames: array of string    - массив названий страниц
// (количество элементов в двух массивах должны совпадать)
// RETURN
// integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string): Integer; overload;
begin
  Result := SaveXmlssToXLSXPath(XMLSS, PathName, SheetsNumbers, SheetsNames, nil, 'UTF-8', '');
end; // SaveXmlssToXLSXPath

// SaveXmlssToXLSXPath
// Сохраняет незапакованный документ в формате Office Open XML (OOXML)
// INPUT
// var XMLSS: TZWorkBook                   - хранилище
// PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
// RETURN
// integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string): Integer; overload;
begin
  Result := SaveXmlssToXLSXPath(XMLSS, PathName, [], []);
end; // SaveXmlssToXLSXPath

// Сохраняет документ в формате Open Office XML (xlsx)
// INPUT
// var XMLSS: TZWorkBook                   - хранилище
// FileName: string                  - имя файла для сохранения
// const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
// const SheetsNames: array of string    - массив названий страниц
// (количество элементов в двух массивах должны совпадать)
// TextConverter: TAnsiToCPConverter - конвертер
// CodePageName: string              - имя кодировки
// BOM: ansistring                   - Byte Order Mark
// RETURN
// integer
function SaveXmlssToXLSX(var XMLSS: TZWorkBook; zipStream: TStream; const SheetsNumbers: array of Integer;
  const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string;
  BOM: ansistring = ''): Integer;
var
  _pages: TIntegerDynArray; // numbers of sheets
  _names: TStringDynArray; // names of sheets
  kol, i: Integer;
  Zip: TZipFile;
  Stream: TStream;
  WriteHelper: TZEXLSXWriteHelper;
  SharedStrings: TStringDynArray;
  SharedStringsDictionary: TDictionary<string, Integer>;
begin
  Result := 0;
  SharedStrings := [];
  Zip := TZipFile.Create();
  try
    WriteHelper := TZEXLSXWriteHelper.Create();
    try
      SharedStringsDictionary := TDictionary<string, Integer>.Create;
      try
        if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
          exit(2);

        Zip.Open(zipStream, zmReadWrite);

        // styles
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateStyles(XMLSS, Stream, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'xl/styles.xml');
        finally
          Stream.Free();
        end;

        // _rels/.rels
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateRelsMain(Stream, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, '_rels/.rels');
        finally
          Stream.Free();
        end;

        // xl/_rels/workbook.xml.rels
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateRelsWorkBook(kol, Stream, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'xl/_rels/workbook.xml.rels');
        finally
          Stream.Free();
        end;

        // sheets of workbook
        for i := 0 to kol - 1 do
        begin
          if XMLSS.Sheets[_pages[i]].RowCount > 60000 then
          begin
            Stream := TTempFileStream.Create();
            try
              ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], SharedStrings, SharedStringsDictionary, TextConverter,
                CodePageName, BOM, WriteHelper);
              Stream.Position := 0;
              Zip.Add(Stream, 'xl/worksheets/sheet' + IntToStr(i + 1) + '.xml');
            finally
              Stream.Free();
            end;
          end
          else
          begin
            Stream := TMemoryStream.Create();
            try
              ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], SharedStrings, SharedStringsDictionary, TextConverter,
                CodePageName, BOM, WriteHelper);
              Stream.Position := 0;
              Zip.Add(Stream, 'xl/worksheets/sheet' + IntToStr(i + 1) + '.xml');
            finally
              Stream.Free();
            end;
          end;

          if (WriteHelper.HyperLinksCount > 0) then
          begin
            WriteHelper.AddSheetHyperlink(i);
            Stream := TMemoryStream.Create();
            try
              WriteHelper.CreateSheetRels(Stream, TextConverter, CodePageName, BOM);
              Stream.Position := 0;
              Zip.Add(Stream, 'xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels');
            finally
              Stream.Free();
            end;
          end;
        end; // for i

        // sharedStrings.xml
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateSharedStrings(XMLSS, Stream, SharedStrings, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'xl/sharedStrings.xml');
        finally
          Stream.Free();
        end;

        for i := 0 to XMLSS.Sheets.Count - 1 do
        begin
          if not XMLSS.Sheets[i].Drawing.IsEmpty then
          begin
            // drawings/drawingN.xml
            Stream := TMemoryStream.Create();
            try
              ZEXLSXCreateDrawing(XMLSS.Sheets[i], Stream, TextConverter, CodePageName, BOM);
              Stream.Position := 0;
              Zip.Add(Stream, 'xl/drawings/drawing' + IntToStr(i + 1) + '.xml');
            finally
              Stream.Free();
            end;

            // drawings/_rels/drawingN.xml.rels
            Stream := TMemoryStream.Create();
            try
              ZEXLSXCreateDrawingRels(XMLSS.Sheets[i], Stream, TextConverter, CodePageName, BOM);
              Stream.Position := 0;
              Zip.Add(Stream, 'xl/drawings/_rels/drawing' + IntToStr(i + 1) + '.xml.rels');
            finally
              Stream.Free();
            end;
          end;
        end;

        // media/imageN.png
        for i := 0 to High(XMLSS.MediaList) do
        begin
          Zip.Add(XMLSS.MediaList[i].Content, 'xl/media/' + XMLSS.MediaList[i].fileName);
        end;

        // workbook.xml - sheets count
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateWorkBook(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'xl/workbook.xml');
        finally
          Stream.Free();
        end;

        // [Content_Types].xml
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateContentTypes(XMLSS, Stream, kol, 0, nil, TextConverter, CodePageName, BOM, WriteHelper);
          Stream.Position := 0;
          Zip.Add(Stream, '[Content_Types].xml');
        finally
          Stream.Free();
        end;

        // docProps/app.xml
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateDocPropsApp(Stream, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'docProps/app.xml');
        finally
          Stream.Free();
        end;

        // docProps/core.xml
        Stream := TMemoryStream.Create();
        try
          ZEXLSXCreateDocPropsCore(XMLSS, Stream, TextConverter, CodePageName, BOM);
          Stream.Position := 0;
          Zip.Add(Stream, 'docProps/core.xml');
        finally
          Stream.Free();
        end;
      finally
        SharedStringsDictionary.Free;
      end;
    finally
      WriteHelper.Free();
    end;
  finally
    Zip.Free();
    ZESClearArrays(_pages, _names);
  end;
end; // SaveXmlssToXSLX

{ TZEXMLSSHelper }

procedure TZEXMLSSHelper.LoadFromFile(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TZEXMLSSHelper.LoadFromStream(AStream: TStream);
begin
  ReadXLSXFile(self, AStream);
end;

procedure TZEXMLSSHelper.SaveToFile(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate or fmOpenReadWrite);
  try
    SaveToStream(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TZEXMLSSHelper.SaveToStream(AStream: TStream);
begin
  SaveXmlssToXLSX(self, AStream, [], [], nil, 'UTF-8');
end;

{ TExcel4DelphiWriter }

constructor TExcel4DelphiWriter.Create(AWorkBook: TZWorkBook);
begin
  FSharedStrings := TObjectDictionary<TRichText, Integer>.Create([doOwnsKeys], TRichTextComparer.Create);
end;

destructor TExcel4DelphiWriter.Destroy;
begin
  FSharedStrings.Free();
  inherited;
end;

// procedure TExcel4DelphiWriter.ReadRunProperties(AXml: TZsspXMLWriterH; ATagName: string; AFont: TZFont);
// begin
/// /
// end;

procedure TExcel4DelphiWriter.SaveToDir(ADirName: string);
begin

end;

procedure TExcel4DelphiWriter.SaveToFile(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.SaveToStream(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteContentTypes(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteContentTypes(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteContentTypes(AStream: TStream);
var
  Xml: TZsspXMLWriterH;
  s: string;
  i: Integer;
  procedure _WriteOverride(const PartName: string; ct: Integer);
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('PartName', PartName);
    case ct of
      0:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
      1:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
      2:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
      3:
        s := 'application/vnd.openxmlformats-package.relationships+xml';
      4:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
      5:
        s := 'application/vnd.openxmlformats-package.core-properties+xml';
      6:
        s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
      7:
        s := 'application/vnd.openxmlformats-officedocument.vmlDrawing';
      8:
        s := 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
      9:
        s := 'application/vnd.openxmlformats-officedocument.drawing+xml';
    end;
    Xml.Attributes.Add('ContentType', s, false);
    Xml.WriteEmptyTag('Override', true);
  end;

  procedure _WriteTypeDefault(extension, contentType: string);
  begin
    Xml.Attributes.Clear();
    Xml.Attributes.Add('Extension', extension);
    Xml.Attributes.Add('ContentType', contentType, false);
    Xml.WriteEmptyTag('Default', true);
  end;

begin
  Xml := TZsspXMLWriterH.Create(AStream);
  try
    Xml.TabLength := 1;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader('utf-8', '');
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_PACKAGE + '/content-types');
    Xml.WriteTagNode('Types', true, true, true);

    _WriteTypeDefault('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    _WriteTypeDefault('xml', 'application/xml');
    _WriteTypeDefault('png', 'image/png');
    _WriteTypeDefault('jpeg', 'image/jpeg');
    _WriteTypeDefault('wmf', 'image/x-wmf');

    // Страницы
    // _WriteOverride('/_rels/.rels', 3);
    // _WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for i := 0 to FWorkBook.Sheets.Count - 1 do
    begin
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
      // if (WriteHelper.IsSheetHaveHyperlinks(i)) then
      // _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels', 3);
    end;

    // комментарии
    for i := 0 to FCommentList.Count - 1 do
    begin
      // _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
      // _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
    end;

    for i := 0 to FWorkBook.Sheets.Count - 1 do
    begin
      if Assigned(FWorkBook.Sheets[i].Drawing) and (FWorkBook.Sheets[i].Drawing.Count > 0) then
      begin
        _WriteOverride('/xl/drawings/drawing' + IntToStr(i + 1) + '.xml', 9);
        // _WriteOverride('/xl/drawings/_rels/drawing' + IntToStr(i+1) + '.xml.rels', 3);
// for ii := 0 to _drawing.PictureStore.Count - 1 do begin
// _picture := _drawing.PictureStore.Items[ii];
// // image/ override
// xml.Attributes.Clear();
// xml.Attributes.Add('PartName', '/xl/media/' + _picture.Name);
// xml.Attributes.Add('ContentType', 'image/' + Copy(ExtractFileExt(_picture.Name), 2, 99), false);
// xml.WriteEmptyTag('Override', true);
// end;
      end;
    end;

    _WriteOverride('/xl/workbook.xml', 2);
    _WriteOverride('/xl/styles.xml', 1);
    _WriteOverride('/xl/sharedStrings.xml', 4);
    _WriteOverride('/docProps/app.xml', 8);
    _WriteOverride('/docProps/core.xml', 5);

    Xml.WriteEndTagNode(); // Types
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDocPropsApp(AStream: TStream);
var
  Xml: TZsspXMLWriterH;
begin
  Xml := TZsspXMLWriterH.Create(AStream);
  try
    Xml.TabLength := 1;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader('utf-8', '');
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns', SCHEMA_DOC + '/extended-properties');
    Xml.Attributes.Add('xmlns:vt', SCHEMA_DOC + '/docPropsVTypes', false);
    Xml.WriteTagNode('Properties', true, true, false);

    Xml.Attributes.Clear();
    Xml.WriteTag('TotalTime', '0', true, false, false);
    Xml.WriteTag('Application', ZE_XLSX_APPLICATION, true, false, true);
    Xml.WriteEndTagNode();
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDocPropsApp(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteDocPropsApp(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDocPropsCore(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteDocPropsCore(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDocPropsCore(AStream: TStream);
var
  Xml: TZsspXMLWriterH;
  creationDate: string;
begin
  Xml := TZsspXMLWriterH.Create(AStream);
  try
    Xml.TabLength := 1;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader('utf-8', '');
    Xml.Attributes.Clear();
    Xml.Attributes.Add('xmlns:cp', SCHEMA_PACKAGE + '/metadata/core-properties');
    Xml.Attributes.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
    Xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    Xml.Attributes.Add('xmlns:dcterms', 'http://purl.org/dc/terms/', false);
    Xml.Attributes.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
    Xml.WriteTagNode('cp:coreProperties', true, true, false);

    Xml.Attributes.Clear();
    Xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    creationDate := ZEDateTimeToStr(FWorkBook.DocumentProperties.Created) + 'Z';
    Xml.WriteTag('dcterms:created', creationDate, true, false, false);
    Xml.WriteTag('dcterms:modified', creationDate, true, false, false);

    Xml.Attributes.Clear();
    Xml.WriteTag('cp:revision', '1', true, false, false);

    Xml.WriteEndTagNode();
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDrawingRels(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteDrawingRels(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteDrawingRels(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteDrawings(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteRelationships(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteRelationships(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteRelationships(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteSharedStrings(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteSharedStrings(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteSharedStrings(AStream: TStream);
var
  Xml: TZsspXMLWriterH;
begin
  Xml := TZsspXMLWriterH.Create(AStream);
  try
    Xml.TabLength := 1;
    Xml.TabSymbol := ' ';
    Xml.WriteHeader('utf-8', '');
    Xml.Attributes.Clear();

    Xml.Attributes.Add('count', IntToStr(FSharedStrings.Count), false);
    Xml.Attributes.Add('uniqueCount', IntToStr(FSharedStrings.Count), false);
    Xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN, false);
    Xml.WriteTagNode('sst', true, true, false);

    // todo: write

    Xml.WriteEndTagNode();
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteStyles(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteStyles(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteStyles(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteTheme(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteTheme(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteTheme(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteDrawings(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteDrawings(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteWorkBook(AStream: TStream);
begin

end;

procedure TExcel4DelphiWriter.WriteWorkBook(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteWorkBook(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiWriter.WriteWorkSheet(AStream: TStream; ASheetIndex: Integer);
begin

end;

procedure TExcel4DelphiWriter.WriteWorkSheet(const AFileName: string; const ASheetIndex: Integer);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmCreate);
  try
    WriteWorkSheet(Stream, ASheetIndex);
  finally
    Stream.Free();
  end;
end;

{ TExcel4DelphiReader }

constructor TExcel4DelphiReader.Create(AWorkBook: TZWorkBook);
begin
  FSharedStrings := TObjectList<TRichText>.Create(true);
end;

destructor TExcel4DelphiReader.Destroy;
begin
  FSharedStrings.Free();
  inherited;
end;

procedure TExcel4DelphiReader.LoadFromDir(const ADirName: string);
begin

end;

procedure TExcel4DelphiReader.LoadFromFile(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.LoadFromStream(AStream: TStream);
var
  Stream: TStream;
  path: string;
  Zip: TZipFile;
  encoding: TEncoding;
  zipHdr: TZipHeader;
  index: Integer;
  fileRec: TZFileItem;
begin
  Zip := TZipFile.Create();
  encoding := TEncoding.GetEncoding(437);
{$IFDEF VER330}
  Zip.encoding := encoding;
{$ENDIF}
  FWorkBook.Styles.Clear();
  FWorkBook.Sheets.Count := 0;
  try
    Zip.Open(AStream, zmRead);

    Zip.Read('[Content_Types].xml', Stream, zipHdr);
    try
      ReadContentTypes(Stream)
    finally
      Stream.Free();
    end;

    path := '/_rels/.rels';
    if Zip.IndexOf(path.Substring(1)) > -1 then
    begin
      fileRec.original := path;
      fileRec.name := path;
      fileRec.ftype := TRelationType.rtDoc;
      FFileList.Add(fileRec);
    end;

    path := '/xl/_rels/workbook.xml.rels';
    if Zip.IndexOf(path.Substring(1)) > -1 then
    begin
      fileRec.original := path;
      fileRec.name := path;
      fileRec.ftype := TRelationType.rtDoc;
      FFileList.Add(fileRec);
    end;

      // FFileList.Sort(prority)

    index := 0;
    for fileRec in FFileList do
    begin
      Zip.Read(fileRec.original.Substring(1), Stream, zipHdr);
      try
        if fileRec.ftype = TRelationType.rtDoc then
          ReadRelationships(Stream)
        else if fileRec.ftype = TRelationType.rtCoreProp then
          ReadSharedStrings(Stream)
        else if fileRec.ftype = TRelationType.rtVmlDrawing then
          ReadTheme(Stream)
        else if fileRec.ftype = TRelationType.rtStyles then
          ReadStyles(Stream)
        else if fileRec.ftype = TRelationType.rtComments then
          ReadComments(Stream, 0)
        else if fileRec.ftype = TRelationType.rtWorkSheet then
        begin
          ReadWorkSheet(Stream, index);
          inc(index);
        end;
      finally
        Stream.Free();
      end;
    end;
  finally
    Zip.Free();
    encoding.Free;
  end;
end;

procedure TExcel4DelphiReader.ReadContentTypes(AStream: TStream);
var
  Xml: TZsspXMLReaderH;
  contType: string;
  rec: TContentTypeRec;
  fileItem: TZFileItem;
begin
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if Xml.BeginReadStream(AStream) <> 0 then
      exit;
    while Xml.ReadToEndTagByName('Types') do
    begin
      if Xml.IsTagClosedByName('Override') then
      begin
        contType := Xml.Attributes.ItemsByName['ContentType'];
        for rec in CONTENT_TYPES do
        begin
          if contType = rec.name then
          begin
            fileItem.name := Xml.Attributes.ItemsByName['PartName'];
            fileItem.original := Xml.Attributes.ItemsByName['PartName'];
            fileItem.ftype := rec.ftype;
            FFileList.Add(fileItem);
            break;
          end;
        end;
      end
      else if Xml.IsTagClosedByName('Default') then
      begin
        // ignore
      end;
    end;
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadComments(AStream: TStream; ASheetIndex: Integer);
begin

end;

procedure TExcel4DelphiReader.ReadContentTypes(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadContentTypes(Stream);
  finally
    Stream.Free();
  end;
end;

function TExcel4DelphiReader.ReadRunProperties(AXml: TZsspXMLReaderH; ATagName: string): TZFont;
begin
  // Temporary we are using TZFont
  // todo: full implement Run Properties
  // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_rPr_topic_ID0ENXS5.html
  Result := TZFont.Create();
  while AXml.ReadToEndTagByName(ATagName) do
  begin
    if AXml.IsTagClosedByName('b') then
      Result.Style := Result.Style + [TFontStyle.fsBold]
    else if AXml.IsTagClosedByName('u') then
      Result.Style := Result.Style + [TFontStyle.fsUnderline]
    else if AXml.IsTagClosedByName('i') then
      Result.Style := Result.Style + [TFontStyle.fsItalic]
    else if AXml.IsTagClosedByName('sz') then
      Result.Size := StrToFloatDef(AXml.Attributes['val'], { todo: default } 11, TFormatSettings.Invariant)
    else if AXml.IsTagClosedByName('rFont') then
      Result.name := trim(AXml.Attributes['val'])
    else if AXml.IsTagClosedByName('charset') then
      Result.charset := StrToIntDef(AXml.Attributes['val'], 0)
    else if AXml.IsTagClosedByName('color ') then
      Result.ExcelColor.Load(AXml.Attributes['rgb'], AXml.Attributes['indexed'], AXml.Attributes['theme'],
        AXml.Attributes['tint']);
  end;
end;

procedure TExcel4DelphiReader.ReadRelationships(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadRelationships(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadRelationships(AStream: TStream);
var
  Xml: TZsspXMLReaderH;
  relation: TZRelation;
begin
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(AStream) <> 0) then
      exit;

    while Xml.ReadToEndTagByName('Relationships') do
    begin
      if Xml.IsTagClosedByName('Relationship') then
      begin
        relation.id := Xml.Attributes.ItemsByName['Id'];
        relation.fileid := -1;
        relation.state := 0;
        relation.sheetid := 0;
        relation.name := '';
        relation.ftype := ZEXLSXGetRelationNumber(Xml.Attributes.ItemsByName['Type']);
        relation.target := Xml.Attributes.ItemsByName['Target'];
        FRelationList.Add(relation);
      end;
    end;
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadSharedStrings(AStream: TStream);
var
  Xml: TZsspXMLReaderH;
  richText: TRichText;
  richString: TRichString;
begin
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(AStream) <> 0) then
      exit;

    while Xml.ReadToEndTagByName('sst') do
    begin
      if Xml.IsTagStartByName('sst') then
        FSharedStrings.Capacity := StrToIntDef(Xml.Attributes['count'], 0);

      if Xml.IsTagStartByName('si') then
      begin
        richText := TRichText.Create();
        while Xml.ReadToEndTagByName('si') do
        begin
          // Rich string with formatting
          if Xml.IsTagStartByName('r') then
          begin
            richString := TRichString.Create();
            while Xml.ReadToEndTagByName('r') do
            begin
              if Xml.IsTagStartByName('rPr') then
                richString.Font := ReadRunProperties(Xml, 'rPr');
              richText.List.Add(richString);
            end;
          end
          // Simple or multyline strings without formatting
          else if Xml.IsTagEndByName('t') then
          begin
            richString := TRichString.Create();
            richString.Text := Xml.TextBeforeTag;
            richText.List.Add(richString);
          end;
        end;
        FSharedStrings.Add(richText);
      end;
    end;
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadSharedStrings(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadSharedStrings(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadStyles(AStream: TStream);
begin

end;

procedure TExcel4DelphiReader.ReadStyles(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadStyles(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadTheme(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadTheme(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadTheme(AStream: TStream);
var
  Xml: TZsspXMLReaderH;
  Value: string;
begin
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if Xml.BeginReadStream(AStream) <> 0 then
      exit;

    while Xml.ReadToEndTagByName('clrScheme') do
    begin
      if Xml.IsTagClosedByName('a:sysClr') then
      begin
        Value := Xml.Attributes.ItemsByName['lastClr'];
        FThemaColorList.Add(HTMLHexToColor(Value))
      end
      else if Xml.IsTagClosedByName('a:srgbClr') then
      begin
        Value := Xml.Attributes.ItemsByName['val'];
        FThemaColorList.Add(HTMLHexToColor(Value));
      end;
    end;
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadWorkBook(AStream: TStream);
var
  Xml: TZsspXMLReaderH;
  // s: string;
  { i, t, } dn: Integer;
begin
  Xml := TZsspXMLReaderH.Create();
  try
    if (Xml.BeginReadStream(AStream) <> 0) then
      exit;

    dn := 0;
    while Xml.ReadToEndTagByName('workbook') do
    begin
      if Xml.IsTagStartByName('definedName') then
      begin
        Xml.ReadTag();
        SetLength(FWorkBook.FDefinedNames, dn + 1);
        FWorkBook.FDefinedNames[dn].LocalSheetId := StrToIntDef(Xml.Attributes.ItemsByName['localSheetId'], 0);
        FWorkBook.FDefinedNames[dn].name := Xml.Attributes.ItemsByName['name'];
        FWorkBook.FDefinedNames[dn].Body := Xml.TagValue;
        inc(dn);
// end else
// if xml.IsTagClosedByName('sheet') then begin
// s := xml.Attributes.ItemsByName['r:id'];
// for i := 0 to RelationsCount - 1 do
// if (Relations[i].id = s) then begin
// Relations[i].name := ZEReplaceEntity(xml.Attributes.ItemsByName['name']);
// s := xml.Attributes.ItemsByName['sheetId'];
// relations[i].sheetid := -1;
// if (TryStrToInt(s, t)) then
// relations[i].sheetid := t;
// s := xml.Attributes.ItemsByName['state'];
// break;
// end;
// end else
// if xml.IsTagClosedByName('workbookView') then begin
// s := xml.Attributes.ItemsByName['activeTab'];
// s := xml.Attributes.ItemsByName['firstSheet'];
// s := xml.Attributes.ItemsByName['showHorizontalScroll'];
// s := xml.Attributes.ItemsByName['showSheetTabs'];
// s := xml.Attributes.ItemsByName['showVerticalScroll'];
// s := xml.Attributes.ItemsByName['tabRatio'];
// s := xml.Attributes.ItemsByName['windowHeight'];
// s := xml.Attributes.ItemsByName['windowWidth'];
// s := xml.Attributes.ItemsByName['xWindow'];
// s := xml.Attributes.ItemsByName['yWindow'];
      end;
    end;
  finally
    Xml.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadWorkBook(const AFileName: string);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadWorkBook(Stream);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadWorkSheet(const AFileName: string; ASheetIndex: Integer);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFileName, fmOpenRead);
  try
    ReadWorkSheet(Stream, ASheetIndex);
  finally
    Stream.Free();
  end;
end;

procedure TExcel4DelphiReader.ReadWorkSheet(AStream: TStream; ASheetIndex: Integer);
const
  MaximumDigitWidth: real = 5.0;
var
  Xml: TZsspXMLReaderH;
  // currentPage: integer;
  currentRow: Integer;
  currentCol: Integer;
  currentSheet: TZSheet;
  currentCell: TZCell;
  str: string;
  tempReal: real;
  tempInt: Integer;
  tempDate: TDateTime;
  tempFloat: double;

  procedure CheckRow(const RowCount: Integer);
  begin
    if (currentSheet.RowCount < RowCount) then
      currentSheet.RowCount := RowCount;
  end;

  procedure CheckCol(const ColCount: Integer);
  begin
    if (currentSheet.ColCount < ColCount) then
      currentSheet.ColCount := ColCount
  end;

  procedure _ReadSheetData();
  var
    t: Integer;
    v: string;
    _num: Integer;
    _type: string;
    _cr, _cc: Integer;
    maxCol: Integer;
  begin
    _cr := 0;
    _cc := 0;
    maxCol := 0;
    CheckRow(1);
    CheckCol(1);
    while Xml.ReadToEndTagByName('sheetData') do
    begin
      // ячейка
      if (Xml.TagName = 'c') then
      begin
        str := Xml.Attributes.ItemsByName['r']; // номер
        if (str > '') then
          if (ZEGetCellCoords(str, _cc, _cr)) then
          begin
            currentCol := _cc;
            CheckCol(_cc + 1);
          end;

        _type := Xml.Attributes.ItemsByName['t']; // тип

        // s := xml.Attributes.ItemsByName['cm'];
        // s := xml.Attributes.ItemsByName['ph'];
        // s := xml.Attributes.ItemsByName['vm'];
        v := '';
        _num := 0;
        currentCell := currentSheet.Cell[currentCol, currentRow];
        str := Xml.Attributes.ItemsByName['s']; // стиль
        if (str > '') then
          if (TryStrToInt(str, t)) then
            currentCell.CellStyle := t;
        if (Xml.IsTagStart) then
          while Xml.ReadToEndTagByName('c') do
          begin
          // is пока игнорируем
            if Xml.IsTagEndByName('v') or Xml.IsTagEndByName('t') then
            begin
              if (_num > 0) then
                v := v + sLineBreak;
              v := v + Xml.TextBeforeTag;
              inc(_num);
            end
            else if Xml.IsTagEndByName('f') then
              currentCell.Formula := ZEReplaceEntity(Xml.TextBeforeTag);

          end; // while

        // Возможные типы:
        // s - sharedstring
        // b - boolean
        // n - number
        // e - error
        // str - string
        // inlineStr - inline string ??
        // d - date
        // тип может отсутствовать. Интерпретируем в таком случае как ZEGeneral
        if (_type = '') then
          currentCell.CellType := ZEGeneral
        else if (_type = 'n') then
        begin
          currentCell.CellType := ZENumber;
          // Trouble: if cell style is number, and number format is date, then
          // cell style is date. F****** m$!
// if (ReadHelper.NumberFormats.IsDateFormat(currentCell.CellStyle)) then
// if (ZEIsTryStrToFloat(v, tempFloat)) then begin
// currentCell.CellType := ZEDateTime;
// v := ZEDateTimeToStr(tempFloat);
// end;
        end
        else if (_type = 's') then
        begin
          currentCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < FSharedStrings.Count)) then
              v := FSharedStrings[t].ToString();
        end
        else if (_type = 'd') then
        begin
          currentCell.CellType := ZEDateTime;
          if (TryZEStrToDateTime(v, tempDate)) then
            v := ZEDateTimeToStr(tempDate)
          else if (ZEIsTryStrToFloat(v, tempFloat)) then
            v := ZEDateTimeToStr(tempFloat)
          else
            currentCell.CellType := ZEString;
        end;

        currentCell.Data := ZEReplaceEntity(v);
        inc(currentCol);
        CheckCol(currentCol + 1);
        if currentCol > maxCol then
          maxCol := currentCol;
      end
      else
      // строка
        if Xml.IsTagStartOrClosedByName('row') then
        begin
          currentCol := 0;
          str := Xml.Attributes.ItemsByName['r']; // индекс строки
          if (str > '') then
            if (TryStrToInt(str, t)) then
            begin
              currentRow := t - 1;
              CheckRow(t);
            end;
        // s := xml.Attributes.ItemsByName['collapsed'];
        // s := xml.Attributes.ItemsByName['customFormat'];
        // s := xml.Attributes.ItemsByName['customHeight'];
          currentSheet.Rows[currentRow].Hidden := ZETryStrToBoolean(Xml.Attributes.ItemsByName['hidden'], false);

          str := Xml.Attributes.ItemsByName['ht']; // в поинтах
          if (str > '') then
          begin
            tempReal := ZETryStrToFloat(str, 10);
            currentSheet.Rows[currentRow].Height := tempReal;
          // tempReal := tempReal / 2.835; //???
          // currentSheet.Rows[currentRow].HeightMM := tempReal;
          end
          else
            currentSheet.Rows[currentRow].Height := currentSheet.DefaultRowHeight;

          str := Xml.Attributes.ItemsByName['outlineLevel'];
          currentSheet.Rows[currentRow].OutlineLevel := StrToIntDef(str, 0);

        // s := xml.Attributes.ItemsByName['ph'];

          str := Xml.Attributes.ItemsByName['s']; // номер стиля
          if (str > '') then
            if (TryStrToInt(str, t)) then
            begin
            // нужно подставить нужный стиль
            end;
        // s := xml.Attributes.ItemsByName['spans'];
        // s := xml.Attributes.ItemsByName['thickBot'];
        // s := xml.Attributes.ItemsByName['thickTop'];

          if Xml.IsTagClosed then
          begin
            inc(currentRow);
            CheckRow(currentRow + 1);
          end;
        end
        else
      // конец строки
          if Xml.IsTagEndByName('row') then
          begin
            inc(currentRow);
            CheckRow(currentRow + 1);
          end;
    end; // while
    currentSheet.ColCount := maxCol;
  end; // _ReadSheetData

  procedure _ReadAutoFilter();
  begin
    currentSheet.AutoFilter := Xml.Attributes.ItemsByName['ref'];
  end;

  procedure _ReadMerge();
  var
    i, t, Num: Integer;
    x1, x2, y1, y2: Integer;
    s1, s2: string;
    b: Boolean;
    function _GetCoords(var x, y: Integer): Boolean;
    begin
      Result := true;
      x := ZEGetColByA1(s1);
      if (x < 0) then
        Result := false;
      if (not TryStrToInt(s2, y)) then
        Result := false
      else
        dec(y);
      b := Result;
    end; // _GetCoords

  begin
    x1 := 0;
    y1 := 0;
    x2 := 0;
    y2 := 0;
    while Xml.ReadToEndTagByName('mergeCells') do
    begin
      if Xml.IsTagStartOrClosedByName('mergeCell') then
      begin
        str := Xml.Attributes.ItemsByName['ref'];
        t := length(str);
        if (t > 0) then
        begin
          str := str + ':';
          s1 := '';
          s2 := '';
          b := true;
          Num := 0;
          for i := 1 to t + 1 do
            case str[i] of
              'A' .. 'Z', 'a' .. 'z':
                s1 := s1 + str[i];
              '0' .. '9':
                s2 := s2 + str[i];
              ':':
                begin
                  inc(Num);
                  if (Num > 2) then
                  begin
                    b := false;
                    break;
                  end;
                  if (Num = 1) then
                  begin
                    if (not _GetCoords(x1, y1)) then
                      break;
                  end
                  else
                  begin
                    if (not _GetCoords(x2, y2)) then
                      break;
                  end;
                  s1 := '';
                  s2 := '';
                end;
            else
              begin
                b := false;
                break;
              end;
            end; // case

          if (b) then
          begin
            CheckRow(y1 + 1);
            CheckRow(y2 + 1);
            CheckCol(x1 + 1);
            CheckCol(x2 + 1);
            currentSheet.MergeCells.AddRectXY(x1, y1, x2, y2);
          end;
        end; // if
      end; // if
    end; // while
  end; // _ReadMerge

  procedure _ReadCols();
  type
    TZColInf = record
      min, Max: Integer;
      bestFit, Hidden: Boolean;
      OutlineLevel: Integer;
      width: Integer;
    end;
  var
    i, j: Integer;
    t: real;
    colInf: TArray<TZColInf>;
  const
    MAX_COL_DIFF = 500;
  begin
    i := 0;
    while Xml.ReadToEndTagByName('cols') do
    begin
      if (Xml.TagName = 'col') and Xml.IsTagStartOrClosed then
      begin
        SetLength(colInf, i + 1);

        colInf[i].min := StrToIntDef(Xml.Attributes.ItemsByName['min'], 0);
        colInf[i].Max := StrToIntDef(Xml.Attributes.ItemsByName['max'], 0);
        // защита от сплошного диапазона
        // когда значение _мах = 16384
        // но чтобы уж наверняка, проверим на MAX_COL_DIFF колонок подряд.
        if (colInf[i].Max - colInf[i].min) > MAX_COL_DIFF then
          colInf[i].Max := colInf[i].min + MAX_COL_DIFF;

        colInf[i].OutlineLevel := StrToIntDef(Xml.Attributes.ItemsByName['outlineLevel'], 0);
        str := Xml.Attributes.ItemsByName['hidden'];
        if (str > '') then
          colInf[i].Hidden := ZETryStrToBoolean(str);
        str := Xml.Attributes.ItemsByName['bestFit'];
        if (str > '') then
          colInf[i].bestFit := ZETryStrToBoolean(str);

        str := Xml.Attributes.ItemsByName['width'];
        if (str > '') then
        begin
          t := ZETryStrToFloat(str, 5.14509803921569);
          // t := 10 * t / 5.14509803921569;
          // А.А.Валуев. Формулы расёта ширины взяты здесь - https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
          t := Trunc(((256 * t + Trunc(128 / MaximumDigitWidth)) / 256) * MaximumDigitWidth);
          colInf[i].width := Trunc(t);
        end;

        inc(i);
      end; // if
    end; // while

    for i := Low(colInf) to High(colInf) do
    begin
      for j := colInf[i].min to colInf[i].Max do
      begin
        CheckCol(j);
        currentSheet.Columns[j - 1].AutoFitWidth := colInf[i].bestFit;
        currentSheet.Columns[j - 1].Hidden := colInf[i].Hidden;
        currentSheet.Columns[j - 1].WidthPix := colInf[i].width;
      end;
    end;
  end; // _ReadCols

  function _StrToMM(const st: string; var retFloat: real): Boolean;
  begin
    Result := false;
    if (str > '') then
    begin
      retFloat := ZETryStrToFloat(st, -1);
      if (retFloat > -1) then
      begin
        Result := true;
        retFloat := retFloat * ZE_MMinInch;
      end;
    end;
  end; // _StrToMM

  procedure _GetDimension();
  var
    st, s: string;
    i, l, _maxC, _maxR, c, r: Integer;
  begin
    c := 0;
    r := 0;
    st := Xml.Attributes.ItemsByName['ref'];
    l := length(st);
    if (l > 0) then
    begin
      st := st + ':';
      inc(l);
      s := '';
      _maxC := -1;
      _maxR := -1;
      for i := 1 to l do
        if (st[i] = ':') then
        begin
          if (ZEGetCellCoords(s, c, r, true)) then
          begin;
            if (c > _maxC) then
              _maxC := c;
            if (r > _maxR) then
              _maxR := r;
          end
          else
            break;
          s := '';
        end
        else
          s := s + st[i];
      if (_maxC > 0) then
        CheckCol(_maxC);
      if (_maxR > 0) then
        CheckRow(_maxR);
    end;
  end; // _GetDimension()

  procedure _ReadHyperLinks();
  var
    _c, _r, i: Integer;
  begin
    _c := 0;
    _r := 0;
    while Xml.ReadToEndTagByName('hyperlinks') do
    begin
      if Xml.IsTagClosedByName('hyperlink') then
      begin
        str := Xml.Attributes.ItemsByName['ref'];
        if (str > '') then
          if (ZEGetCellCoords(str, _c, _r, true)) then
          begin
            CheckRow(_r);
            CheckCol(_c);
            currentSheet.Cell[_c, _r].HRefScreenTip := Xml.Attributes.ItemsByName['tooltip'];
            str := Xml.Attributes.ItemsByName['r:id'];
            // по r:id подставить ссылку
            for i := 0 to FRelationList.Count - 1 do
              if ((FRelationList[i].id = str) and (FRelationList[i].ftype = TRelationType.rtHyperlink)) then
              begin
                currentSheet.Cell[_c, _r].Href := FRelationList[i].target;
                break;
              end;
          end;
        // доп. атрибуты:
        // display - ??
        // id - id <> r:id??
        // location - ??
      end;
    end; // while
  end; // _ReadHyperLinks();

  procedure _ReadSheetPr();
  begin
    while Xml.ReadToEndTagByName('sheetPr') do
    begin
      if Xml.TagName = 'tabColor' then
        currentSheet.TabColor := ARGBToColor(Xml.Attributes.ItemsByName['rgb']);

      if Xml.TagName = 'pageSetUpPr' then
        currentSheet.FitToPage := ZEStrToBoolean(Xml.Attributes.ItemsByName['fitToPage']);

      if Xml.TagName = 'outlinePr' then
      begin
        currentSheet.ApplyStyles := ZEStrToBoolean(Xml.Attributes.ItemsByName['applyStyles']);
        currentSheet.SummaryBelow := Xml.Attributes.ItemsByName['summaryBelow'] <> '0';
        currentSheet.SummaryRight := Xml.Attributes.ItemsByName['summaryRight'] <> '0';
      end;
    end;
  end; // _ReadSheetPr();

  procedure _ReadRowBreaks();
  begin
    currentSheet.RowBreaks := [];
    while Xml.ReadToEndTagByName('rowBreaks') do
    begin
      if Xml.TagName = 'brk' then
        currentSheet.RowBreaks := currentSheet.RowBreaks + [StrToIntDef(Xml.Attributes.ItemsByName['id'], 0)];
    end;
  end;

  procedure _ReadColBreaks();
  begin
    currentSheet.ColBreaks := [];
    while Xml.ReadToEndTagByName('colBreaks') do
    begin
      if Xml.TagName = 'brk' then
        currentSheet.ColBreaks := currentSheet.ColBreaks + [StrToIntDef(Xml.Attributes.ItemsByName['id'], 0)];
    end;
  end;

  procedure _ReadSheetViews();
  var
    vValue, hValue: Integer;
    SplitMode: TZSplitMode;
    s: string;
  begin
    while Xml.ReadToEndTagByName('sheetViews') do
    begin
      if Xml.IsTagStartByName('sheetView') or Xml.IsTagClosedByName('sheetView') then
      begin
        s := Xml.Attributes.ItemsByName['tabSelected'];
        // тут кроется проблема с выделением нескольких листов
        currentSheet.Selected := currentSheet.SheetIndex = 0; // s = '1';
        currentSheet.ViewMode := zvmNormal;
        if Xml.Attributes.ItemsByName['view'] = 'pageBreakPreview' then
          currentSheet.ViewMode := zvmPageBreakPreview;
      end;

      if Xml.IsTagClosedByName('pane') then
      begin
        SplitMode := ZSplitSplit;
        s := Xml.Attributes.ItemsByName['state'];
        if (s = 'frozen') then
          SplitMode := ZSplitFrozen;

        s := Xml.Attributes.ItemsByName['xSplit'];
        if (not TryStrToInt(s, vValue)) then
          vValue := 0;

        s := Xml.Attributes.ItemsByName['ySplit'];
        if (not TryStrToInt(s, hValue)) then
          hValue := 0;

        currentSheet.SheetOptions.SplitVerticalValue := vValue;
        currentSheet.SheetOptions.SplitHorizontalValue := hValue;

        currentSheet.SheetOptions.SplitHorizontalMode := ZSplitNone;
        currentSheet.SheetOptions.SplitVerticalMode := ZSplitNone;
        if (hValue <> 0) then
          currentSheet.SheetOptions.SplitHorizontalMode := SplitMode;
        if (vValue <> 0) then
          currentSheet.SheetOptions.SplitVerticalMode := SplitMode;

        if (currentSheet.SheetOptions.SplitHorizontalMode = ZSplitSplit) then
          currentSheet.SheetOptions.SplitHorizontalValue := PointToPixel(hValue / 20);
        if (currentSheet.SheetOptions.SplitVerticalMode = ZSplitSplit) then
          currentSheet.SheetOptions.SplitVerticalValue := PointToPixel(vValue / 20);

      end; // if
    end; // while
  end; // _ReadSheetViews()

  procedure _ReadConditionFormatting();
  var
    MaxFormulasCount: Integer;
    _formulas: array of string;
    Count: Integer;
    _sqref: string;
    _type: string;
    _operator: string;
    _CFCondition: TZCondition;
    _CFOperator: TZConditionalOperator;
    _Style: string;
    _text: string;
    _isCFAdded: Boolean;
    _isOk: Boolean;
    // _priority: string;
    _CF: TZConditionalStyle;
    _tmpStyle: TZStyle;

    function _AddCF(): Boolean;
    var
      s, ss: string;
      _len, i, kol: Integer;
      a: array of array [0 .. 5] of Integer;
      _maxx: Integer;
      ch: char;
      w, h: Integer;

      function _GetOneArea(st: string): Boolean;
      var
        i, j: Integer;
        s: string;
        ch: char;
        _cnt: Integer;
        tmpArr: array [0 .. 1, 0 .. 1] of Integer;
        _isOk: Boolean;
        t: Integer;
        tmpB: Boolean;

      begin
        Result := false;
        if (st <> '') then
        begin
          st := st + ':';
          s := '';
          _cnt := 0;
          _isOk := true;
          for i := 1 to length(st) do
          begin
            ch := st[i];
            if (ch = ':') then
            begin
              if (_cnt < 2) then
              begin
                tmpB := ZEGetCellCoords(s, tmpArr[_cnt][0], tmpArr[_cnt][1]);
                _isOk := _isOk and tmpB;
              end;
              s := '';
              inc(_cnt);
            end
            else
              s := s + ch;
          end; // for

          if (_isOk) then
            if (_cnt > 0) then
            begin
              if (_cnt > 2) then
                _cnt := 2;

              a[kol][0] := _cnt;
              t := 1;
              for i := 0 to _cnt - 1 do
                for j := 0 to 1 do
                begin
                  a[kol][t] := tmpArr[i][j];
                  inc(t);
                end;
              Result := true;
            end;
        end; // if
      end; // _GetOneArea

    begin
      Result := false;
      if (_sqref <> '') then
        try
          _maxx := 4;
          SetLength(a, _maxx);
          ss := _sqref + ' ';
          _len := length(ss);
          kol := 0;
          s := '';
          for i := 1 to _len do
          begin
            ch := ss[i];
            if (ch = ' ') then
            begin
              if (_GetOneArea(s)) then
              begin
                inc(kol);
                if (kol >= _maxx) then
                begin
                  inc(_maxx, 4);
                  SetLength(a, _maxx);
                end;
              end;
              s := '';
            end
            else
              s := s + ch;
          end; // for

          if (kol > 0) then
          begin
            currentSheet.ConditionalFormatting.Add();
            _CF := currentSheet.ConditionalFormatting[currentSheet.ConditionalFormatting.Count - 1];
            for i := 0 to kol - 1 do
            begin
              w := 1;
              h := 1;
              if (a[i][0] >= 2) then
              begin
                w := abs(a[i][3] - a[i][1]) + 1;
                h := abs(a[i][4] - a[i][2]) + 1;
              end;
              _CF.Areas.Add(a[i][1], a[i][2], w, h);
            end;
            Result := true;
          end;
        finally
          SetLength(a, 0);
        end;
    end; // _AddCF

    // Применяем условный стиль
    procedure _TryApplyCF();
    var
      b: Boolean;
      Num: Integer;
      _id: Integer;
      procedure _CheckTextCondition();
      begin
        if (Count = 1) then
          if (_formulas[0] <> '') then
            _isOk := true;
      end;

      // Найти стиль
      // пока будем делать так: предполагаем, что все ячейки в текущей области
      // условного форматирования имеют один стиль. Берём стиль из левой верхней
      // ячейки, клонируем его, применяем дифф. стиль, добавляем в хранилище стилей
      // с учётом повторов.
      // TODO: потом нужно будет переделать
      // INPUT
      // dfNum: integer - номер дифференцированного форматирования
      // RETURN
      // integer - номер применяемого стиля
      function _getStyleIdxForDF(dfNum: Integer): Integer;
      var
        _df: TZXLSXDiffFormattingItem;
        _r, _c: Integer;
        _t: Integer;
        i: TZBordersPos;
      begin
        // _currSheet
        Result := -1;
        if ((dfNum >= 0)
        // and (dfNum < ReadHelper.DiffFormatting.Count))
          ) then
        begin
          // _df := ReadHelper.DiffFormatting[dfNum];
          _t := -1;

          if (_CF.Areas.Count > 0) then
          begin
            _r := _CF.Areas.Items[0].Row;
            _c := _CF.Areas.Items[0].Column;
            if ((_r >= 0) and (_r < currentSheet.RowCount)) then
              if ((_c >= 0) and (_c < currentSheet.ColCount)) then
                _t := currentSheet.Cell[_c, _r].CellStyle;
          end;

          _tmpStyle.Assign(FWorkBook.Styles[_t]);

          if (_df.UseFont) then
          begin
            if (_df.UseFontStyles) then
              _tmpStyle.Font.Style := _df.FontStyles;
            if (_df.UseFontColor) then
              _tmpStyle.Font.Color := _df.FontColor;
          end;
          if (_df.UseFill) then
          begin
            if (_df.UseCellPattern) then
              _tmpStyle.CellPattern := _df.CellPattern;
            if (_df.UseBGColor) then
              _tmpStyle.BGColor := _df.BGColor;
            if (_df.UsePatternColor) then
              _tmpStyle.PatternColor := _df.PatternColor;
          end;
          if (_df.UseBorder) then
            for i := bpLeft to bpDiagonalRight do
            begin
              if (_df.Borders[i].UseStyle) then
              begin
                _tmpStyle.Border[i].Weight := _df.Borders[i].Weight;
                _tmpStyle.Border[i].LineStyle := _df.Borders[i].LineStyle;
              end;
              if (_df.Borders[i].UseColor) then
                _tmpStyle.Border[i].Color := _df.Borders[i].Color;
            end; // for

          Result := FWorkBook.Styles.Add(_tmpStyle, true);
        end; // if
      end; // _getStyleIdxForDF

    begin
      _isOk := false;
      case (_CFCondition) of
        ZCFIsTrueFormula:
          ;
        ZCFCellContentIsBetween, ZCFCellContentIsNotBetween:
          begin
            // только числа
            if (Count = 2) then
            begin
              ZETryStrToFloat(_formulas[0], b);
              if (b) then
                ZETryStrToFloat(_formulas[1], _isOk);
            end;
          end;
        ZCFCellContentOperator:
          begin
            // только числа
            if (Count = 1) then
              ZETryStrToFloat(_formulas[0], _isOk);
          end;
        ZCFNumberValue:
          ;
        ZCFString:
          ;
        ZCFBoolTrue:
          ;
        ZCFBoolFalse:
          ;
        ZCFFormula:
          ;
        ZCFContainsText:
          _CheckTextCondition();
        ZCFNotContainsText:
          _CheckTextCondition();
        ZCFBeginsWithText:
          _CheckTextCondition();
        ZCFEndsWithText:
          _CheckTextCondition();
      end; // case

      if (_isOk) then
      begin
        if (not _isCFAdded) then
          _isCFAdded := _AddCF();

        if ((_isCFAdded) and (Assigned(_CF))) then
        begin
          Num := _CF.Count;
          _CF.Add();
          if (_Style <> '') then
            if (TryStrToInt(_Style, _id)) then
              _CF[Num].ApplyStyleID := _getStyleIdxForDF(_id);
          _CF[Num].Condition := _CFCondition;
          _CF[Num].ConditionOperator := _CFOperator;

          _CF[Num].Value1 := _formulas[0];
          if (Count >= 2) then
            _CF[Num].Value2 := _formulas[1];
        end;
      end;
    end;

  begin
    try
      _sqref := Xml.Attributes['sqref'];
      MaxFormulasCount := 2;
      SetLength(_formulas, MaxFormulasCount);
      _isCFAdded := false;
      _CF := nil;
      _tmpStyle := TZStyle.Create();
      while Xml.ReadToEndTagByName('conditionalFormatting') do
      begin
        // cfRule = Conditional Formatting Rule
        if Xml.IsTagStartByName('cfRule') then
        begin
          _type := Xml.Attributes['type'];
          _operator := Xml.Attributes['operator'];
          _Style := Xml.Attributes['dxfId'];
          _text := ZEReplaceEntity(Xml.Attributes['text']);
          // _priority := xml.Attributes['priority'];

          Count := 0;
          while Xml.ReadToEndTagByName('cfRule') do
          begin
            if Xml.IsTagEndByName('formula') then
            begin
              if (Count >= MaxFormulasCount) then
              begin
                inc(MaxFormulasCount, 2);
                SetLength(_formulas, MaxFormulasCount);
              end;
              _formulas[Count] := ZEReplaceEntity(Xml.TextBeforeTag);
              inc(Count);
            end;
          end;

          if (ZEXLSX_getCFCondition(_type, _operator, _CFCondition, _CFOperator)) then
            _TryApplyCF();
        end;
      end;
    finally
      SetLength(_formulas, 0);
      FreeAndNil(_tmpStyle);
    end;
  end;

  procedure _ReadHeaderFooter();
  begin
    currentSheet.SheetOptions.IsDifferentFirst := ZEStrToBoolean(Xml.Attributes['differentFirst']);
    currentSheet.SheetOptions.IsDifferentOddEven := ZEStrToBoolean(Xml.Attributes['differentOddEven']);
    while Xml.ReadToEndTagByName('headerFooter') do
    begin
      if Xml.IsTagEndByName('oddHeader') then
        currentSheet.SheetOptions.Header := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('oddFooter') then
        currentSheet.SheetOptions.Footer := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('evenHeader') then
        currentSheet.SheetOptions.EvenHeader := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('evenFooter') then
        currentSheet.SheetOptions.EvenFooter := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('firstHeader') then
        currentSheet.SheetOptions.FirstPageHeader := ClenuapXmlTagValue(Xml.TextBeforeTag)
      else if Xml.IsTagEndByName('firstFooter') then
        currentSheet.SheetOptions.FirstPageFooter := ClenuapXmlTagValue(Xml.TextBeforeTag);
    end;
  end;

begin
  Xml := TZsspXMLReaderH.Create();
  try
    Xml.AttributesMatch := false;
    if (Xml.BeginReadStream(AStream) <> 0) then
      exit;

    // currentPage := FWorkBook.Sheets.Count;
    FWorkBook.Sheets.Count := FWorkBook.Sheets.Count + 1;
    currentRow := 0;
    currentSheet := FWorkBook.Sheets[ASheetIndex];
    // currentSheet.Title := SheetName;

    while Xml.ReadTag() do
    begin
      if Xml.IsTagStartByName('sheetData') then
        _ReadSheetData()
      else if Xml.IsTagClosedByName('autoFilter') then
        _ReadAutoFilter()
      else if Xml.IsTagStartByName('mergeCells') then
        _ReadMerge()
      else if Xml.IsTagStartByName('cols') then
        _ReadCols()
      else if Xml.IsTagClosedByName('drawing') then
      begin
        currentSheet.DrawingRid := StrToIntDef(Xml.Attributes.ItemsByName['r:id'].Substring(3), 0);
      end
      else if Xml.IsTagClosedByName('pageMargins') then
      begin
        str := Xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginBottom := round(tempReal);
        str := Xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.FooterMargins.Height := abs(round(tempReal));
        str := Xml.Attributes.ItemsByName['header'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.HeaderMargins.Height := abs(round(tempReal));
        str := Xml.Attributes.ItemsByName['left'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginLeft := round(tempReal);
        str := Xml.Attributes.ItemsByName['right'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginRight := round(tempReal);
        str := Xml.Attributes.ItemsByName['top'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginTop := round(tempReal);
      end
      else
      // Настройки страницы
        if Xml.IsTagClosedByName('pageSetup') then
        begin
        // str := xml.Attributes.ItemsByName['blackAndWhite'];
        // str := xml.Attributes.ItemsByName['cellComments'];
        // str := xml.Attributes.ItemsByName['copies'];
        // str := xml.Attributes.ItemsByName['draft'];
        // str := xml.Attributes.ItemsByName['errors'];
          str := Xml.Attributes.ItemsByName['firstPageNumber'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.StartPageNumber := tempInt;

          str := Xml.Attributes.ItemsByName['fitToHeight'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.FitToHeight := tempInt;

          str := Xml.Attributes.ItemsByName['fitToWidth'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.FitToWidth := tempInt;

        // str := xml.Attributes.ItemsByName['horizontalDpi'];
        // str := xml.Attributes.ItemsByName['id'];
          str := Xml.Attributes.ItemsByName['orientation'];
          if (str > '') then
          begin
            currentSheet.SheetOptions.PortraitOrientation := false;
            if (str = 'portrait') then
              currentSheet.SheetOptions.PortraitOrientation := true;
          end;

        // str := xml.Attributes.ItemsByName['pageOrder'];

          str := Xml.Attributes.ItemsByName['paperSize'];
          if (str > '') then
            if (TryStrToInt(str, tempInt)) then
              currentSheet.SheetOptions.PaperSize := tempInt;
        // str := xml.Attributes.ItemsByName['paperHeight']; //если утановлены paperHeight и Width, то paperSize игнорируется
        // str := xml.Attributes.ItemsByName['paperWidth'];

          str := Xml.Attributes.ItemsByName['scale'];
          currentSheet.SheetOptions.ScaleToPercent := StrToIntDef(str, 100);
        // str := xml.Attributes.ItemsByName['useFirstPageNumber'];
        // str := xml.Attributes.ItemsByName['usePrinterDefaults'];
        // str := xml.Attributes.ItemsByName['verticalDpi'];
        end
        else
      // настройки печати
          if Xml.IsTagClosedByName('printOptions') then
          begin
        // str := xml.Attributes.ItemsByName['gridLines'];
        // str := xml.Attributes.ItemsByName['gridLinesSet'];
        // str := xml.Attributes.ItemsByName['headings'];
            str := Xml.Attributes.ItemsByName['horizontalCentered'];
            if (str > '') then
              currentSheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(str);

            str := Xml.Attributes.ItemsByName['verticalCentered'];
            if (str > '') then
              currentSheet.SheetOptions.CenterVertical := ZEStrToBoolean(str);
          end
          else if Xml.IsTagClosedByName('sheetFormatPr') then
          begin
            str := Xml.Attributes.ItemsByName['defaultColWidth'];
            if (str > '') then
              currentSheet.DefaultColWidth := ZETryStrToFloat(str, currentSheet.DefaultColWidth);
            str := Xml.Attributes.ItemsByName['defaultRowHeight'];
            if (str > '') then
              currentSheet.DefaultRowHeight := ZETryStrToFloat(str, currentSheet.DefaultRowHeight);
          end
          else if Xml.IsTagClosedByName('dimension') then
            _GetDimension()
          else if Xml.IsTagStartByName('hyperlinks') then
            _ReadHyperLinks()
          else if Xml.IsTagStartByName('sheetPr') then
            _ReadSheetPr()
          else if Xml.IsTagStartByName('rowBreaks') then
            _ReadRowBreaks()
          else if Xml.IsTagStartByName('colBreaks') then
            _ReadColBreaks()
          else if Xml.IsTagStartByName('sheetViews') then
            _ReadSheetViews()
          else if Xml.IsTagStartByName('conditionalFormatting') then
            _ReadConditionFormatting()
          else if Xml.IsTagStartByName('headerFooter') then
            _ReadHeaderFooter();
    end;
  finally
    Xml.Free();
  end;
end;

end.
