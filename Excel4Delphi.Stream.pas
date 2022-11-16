unit Excel4Delphi.Stream;

interface

uses
  SysUtils,
  Classes,
  Types,
  Graphics,
  UITypes,
  Windows,
  Zip,
  IOUtils,
  Excel4Delphi.Formula,
  Excel4Delphi.Xml,
  Excel4Delphi,
  Excel4Delphi.Common,
  Generics.Collections;

type
  TRelationType = (
    rtNone       = -1,
    rtWorkSheet  = 0,
    rtStyles     = 1,
    rtSharedStr  = 2,
    rtDoc        = 3,
    rtCoreProp   = 4,
    rtExtProps   = 5,
    rtHyperlink  = 6,
    rtComments   = 7,
    rtVmlDrawing = 8,
    rtDrawing    = 9
 );

  TZXLSXFileItem = record
    name: string;     //путь к файлу
    nameArx: string;
    original: string; //исходная строка
    ftype: TRelationType;   //тип контента
  end;

  TZXLSXRelations = record
    id: string;           //rID
    ftype: TRelationType; //тип ссылки
    target: string;       //ссылка на файла
    fileid: integer;      //ссылка на запись
    name: string;         //имя листа
    state: string;        //cостояние, возможные значения см. https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/74cb1d22-b931-4bf8-997d-17517e2416e9
    sheetid: integer;     //номер листа
  end;

  PZXLSXRelations = ^TZXLSXRelations;

  TZXLSXRelationsArray = array of TZXLSXRelations;

  TZXLSXDiffBorderItemStyle = class(TPersistent)
  private
    FUseStyle: boolean;             //заменять ли стиль
    FUseColor: boolean;             //заменять ли цвет
    FColor: TColor;                 //цвет линий
    FLineStyle: TZBorderType;       //стиль линий
    FWeight: byte;
  protected
  public
    constructor Create();
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseStyle: boolean read FUseStyle write FUseStyle;
    property UseColor: boolean read FUseColor write FUseColor;
    property Color: TColor read FColor write FColor;
    property LineStyle: TZBorderType read FLineStyle write FLineStyle;
    property Weight: byte read FWeight write FWeight;
  end;

  TZXLSXDiffBorder = class(TPersistent)
  private
    FBorder: array [0..5] of TZXLSXDiffBorderItemStyle;
    procedure SetBorder(Num: TZBordersPos; Const Value: TZXLSXDiffBorderItemStyle);
    function GetBorder(Num: TZBordersPos): TZXLSXDiffBorderItemStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent);override;
    property Border[Num: TZBordersPos]: TZXLSXDiffBorderItemStyle read GetBorder write SetBorder; default;
  end;

  //Итем для дифференцированного форматирования
  //TODO:
  //      возможно, для excel xml тоже понадобится (перенести?)
  TZXLSXDiffFormattingItem = class(TPersistent)
  private
    FUseFont: boolean;              //заменять ли шрифт
    FUseFontColor: boolean;         //заменять ли цвет шрифта
    FUseFontStyles: boolean;        //заменять ли стиль шрифта
    FFontColor: TColor;             //цвет шрифта
    FFontStyles: TFontStyles;       //стиль шрифта
    FUseBorder: boolean;            //заменять ли рамку
    FBorders: TZXLSXDiffBorder;     //Что менять в рамке
    FUseFill: boolean;              //заменять ли заливку
    FUseCellPattern: boolean;       //Заменять ли тип заливки
    FCellPattern: TZCellPattern;    //тип заливки
    FUseBGColor: boolean;           //заменять ли цвет заливки
    FBGColor: TColor;               //цвет заливки
    FUsePatternColor: boolean;      //Заменять ли цвет шаблона заливки
    FPatternColor: TColor;          //Цвет шаблона заливки
  protected
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    property UseFont: boolean read FUseFont write FUseFont;
    property UseFontColor: boolean read FUseFontColor write FUseFontColor;
    property UseFontStyles: boolean read FUseFontStyles write FUseFontStyles;
    property FontColor: TColor read FFontColor write FFontColor;
    property FontStyles: TFontStyles read FFontStyles write FFontStyles;
    property UseBorder: boolean read FUseBorder write FUseBorder;
    property Borders: TZXLSXDiffBorder read FBorders write FBorders;
    property UseFill: boolean read FUseFill write FUseFill;
    property UseCellPattern: boolean read FUseCellPattern write FUseCellPattern;
    property CellPattern: TZCellPattern read FCellPattern write FCellPattern;
    property UseBGColor: boolean read FUseBGColor write FUseBGColor;
    property BGColor: TColor read FBGColor write FBGColor;
    property UsePatternColor: boolean read FUsePatternColor write FUsePatternColor;
    property PatternColor: TColor read FPatternColor write FPatternColor;
  end;

  // Differential formating
  TZXLSXDiffFormatting = class(TPersistent)
  private
    FCount: integer;
    FMaxCount: integer;
    FItems: array of TZXLSXDiffFormattingItem;
  protected
    function GetItem(num: integer): TZXLSXDiffFormattingItem;
    procedure SetItem(num: integer; const Value: TZXLSXDiffFormattingItem);
    procedure SetCount(ACount: integer);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Add();
    procedure Assign(Source: TPersistent); override;
    procedure Clear();
    property Count: integer read FCount;
    property Items[num: integer]: TZXLSXDiffFormattingItem read GetItem write SetItem; default;
  end;

  //List of cell number formats (date/numbers/currencies etc formats)
  TZEXLSXNumberFormats = class
  private
    FFormatsCount: integer;
    FFormats: array of string; //numFmts (include default formats)
    FStyleFmtID: array of integer;
    FStyleFmtIDCount: integer;
  protected
    function GetFormat(num: integer): string;
    procedure SetFormat(num: integer; const value: string);
    function GetStyleFMTID(num: integer): integer;
    procedure SetStyleFMTID(num: integer; const value: integer);
    procedure SetStyleFMTCount(value: integer);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure ReadNumFmts(const xml: TZsspXMLReaderH);
    function IsDateFormat(StyleNum: integer): boolean;
    function FindFormatID(const value: string): integer;
    property FormatsCount: integer read FFormatsCount;
    property Format[num: integer]: string read GetFormat write SetFormat; default;
    property StyleFMTID[num: integer]: integer read GetStyleFMTID write SetStyleFMTID;
    property StyleFMTCount: integer read FStyleFmtIDCount write SetStyleFMTCount;
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

  //Store link item
  TZEXLSXHyperLinkItem = record
    RID: integer;
    RelType: TRelationType;
    CellRef: string;
    Target: string;
    ScreenTip: string;
    TargetMode: string;
  end;

  { TZEXLSXWriteHelper }

  TZEXLSXWriteHelper = class
  private
    FHyperLinks: array of TZEXLSXHyperLinkItem;
    FHyperLinksCount: integer;
    FMaxHyperLinksCount: integer;
    FCurrentRID: integer;                      //Current rID number (for HyperLinks/comments etc)
    FisHaveComments: boolean;                  //Is Need create comments*.xml?
    FisHaveDrawings: boolean;                  //Is Need create drawings*.xml?
    FSheetHyperlinksArray: array of integer;
    FSheetHyperlinksCount: integer;
  protected
    function GenerateRID(): integer;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure AddHyperLink(const ACellRef, ATarget, AScreenTip, ATargetMode: string);
    function AddDrawing(const ATarget: string): integer;
    procedure WriteHyperLinksTag(const xml: TZsspXMLWriterH);
    function CreateSheetRels(const Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
    procedure AddSheetHyperlink(PageNum: integer);
    function IsSheetHaveHyperlinks(PageNum: integer): boolean;
    procedure Clear();
    property HyperLinksCount: integer read FHyperLinksCount;
    property isHaveComments: boolean read FisHaveComments write FisHaveComments; //Is need create comments*.xml?
    property isHaveDrawings: boolean read FisHaveDrawings write FisHaveDrawings; //Is need create drawings*.xml?
  end;

  TZEXMLSSHelper = class helper for TZWorkBook
    private
    {}

    public
    procedure LoadFromStream(stream: TStream);
    procedure LoadFromFile(fileName: string);
    procedure SaveToStream(stream: TStream);
    procedure SaveToFile(fileName: string);
  end;

//Дополнительные функции для экспорта отдельных файлов
function ZEXLSXCreateStyles(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateWorkBook(var XMLSS: TZWorkBook; Stream: TStream; const _pages: TIntegerDynArray; const _names: TStringDynArray; PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
function ZEXLSXCreateSheet(var XMLSS: TZWorkBook; Stream: TStream; SheetNum: integer; staredStrings: TSharedStringBank; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring; const WriteHelper: TZEXLSXWriteHelper): integer;
function ZEXLSXCreateContentTypes(var XMLSS: TZWorkBook; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring; const WriteHelper: TZEXLSXWriteHelper): integer;
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateSharedStrings(var XMLSS: TZWorkBook; Stream: TStream; const SharedStrings: TSharedStringBank; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDocPropsCore(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
function ZEXLSXCreateDrawing(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
function ZEXLSXCreateDrawingRels(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
procedure ZEAddRelsRelation(xml: TZsspXMLWriterH; const rid: string; ridType: TRelationType; const Target: string; const TargetMode: string = '');


function ReadXLSXPath(var XMLSS: TZWorkBook; DirName: string): integer;
function ReadXLSXFile(var XMLSS: TZWorkBook; zipStream: TStream): integer;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of integer; const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of integer; const SheetsNames: array of string): integer; overload;
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string): integer; overload;
function SaveXmlssToXLSX(var XMLSS: TZWorkBook; zipStream: TStream; const SheetsNumbers: array of integer; const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer;

//Дополнительные функции, на случай чтения отдельного файла
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TArray<TZXLSXFileItem>; var FilesCount: integer): boolean;
function ZEXSLXReadSharedStrings(var Stream: TStream; sharedStrings: TObjectList<TRichText>): boolean;
function ZEXSLXReadStyles(var XMLSS: TZWorkBook; var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer; var MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): boolean;
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer; var isWorkSheet: boolean; needReplaceDelimiter: boolean): boolean;
function ZEXSLXReadWorkBook(var XMLSS: TZWorkBook; var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer): boolean;
function ZEXSLXReadSheet(var XMLSS: TZWorkBook; var Stream: TStream; const SheetRelations: PZXLSXRelations; sharedStrings: TObjectList<TRichText>; var Relations: TZXLSXRelationsArray; RelationsCount: integer; MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): boolean;
function ZEXSLXReadComments(var XMLSS: TZWorkBook; var Stream: TStream): boolean;



implementation

uses AnsiStrings, StrUtils, Math, Excel4Delphi.NumberFormat, NetEncoding;


const
  SCHEMA_DOC         = 'http://schemas.openxmlformats.org/officeDocument/2006';
  SCHEMA_DOC_REL     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  SCHEMA_PACKAGE     = 'http://schemas.openxmlformats.org/package/2006';
  SCHEMA_PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  SCHEMA_SHEET_MAIN  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

type
  TZEXLSXFont = record
    name:      string;
    bold:      boolean;
    italic:    boolean;
    underline: boolean;
    strike:    boolean;
    charset:   integer;
    color:     TColor;
    ColorType: byte;
    LumFactor: double;
    fontsize:  double;
    superscript: boolean;
    subscript: boolean;
  end;

  TZEXLSXFontArray = array of TZEXLSXFont;

  TContentTypeRec=record
    ftype: TRelationType;
    name : string;
    rel  : string;
  end;

const
CONTENT_TYPES: array[0..10] of TContentTypeRec = (
 (ftype: TRelationType.rtWorkSheet;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'),

 (ftype: TRelationType.rtStyles;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'),

 (ftype: TRelationType.rtSharedStr;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),

 (ftype: TRelationType.rtSharedStr;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),

 (ftype: TRelationType.rtDoc;
    name:'application/vnd.openxmlformats-package.relationships+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'),

 (ftype: TRelationType.rtCoreProp;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'),

 (ftype: TRelationType.rtExtProps;
    name:'application/vnd.openxmlformats-package.core-properties+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'),

 (ftype: TRelationType.rtHyperlink;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'),

 (ftype: TRelationType.rtComments;
    name:'application/vnd.openxmlformats-officedocument.vmlDrawing';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'),

 (ftype: TRelationType.rtVmlDrawing;
    name:'application/vnd.openxmlformats-officedocument.theme+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'),

 (ftype: TRelationType.rtDrawing;
    name:'application/vnd.openxmlformats-officedocument.drawing+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing')
);

function GetMaximumDigitWidth(fontName: string; fontSize: double): double;
const
  numbers = '0123456789';
var
  bitmap: Graphics.TBitmap;
  number: string;
begin
  //А.А.Валуев Расчитываем ширину самого широкого числа.
  Result := 0;
  bitmap := Graphics.TBitmap.Create;
  try
    bitmap.Canvas.Font.PixelsPerInch := 96;
    bitmap.Canvas.Font.Size := Trunc(fontSize);
    bitmap.Canvas.Font.Name := fontName;
    for number in numbers do
      Result := Max(Result, bitmap.Canvas.TextWidth(number));
  finally
    bitmap.Free;
  end;
end;


////:::::::::::::  TZXLSXDiffFormating :::::::::::::::::////

constructor TZXLSXDiffFormatting.Create();
var i: integer;
begin
  FCount := 0;
  FMaxCount := 20;
  SetLength(FItems, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FItems[i] := TZXLSXDiffFormattingItem.Create();
end;

destructor TZXLSXDiffFormatting.Destroy();
var i: integer;
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
var df: TZXLSXDiffFormatting; i: integer; b: boolean;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffFormatting) then begin
      b := false;
      df := Source as TZXLSXDiffFormatting;
      SetCount(df.Count);
      for i := 0 to Count - 1 do
        FItems[i].Assign(df[i]);
    end;
  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffFormatting.SetCount(ACount: integer);
var i: integer;
begin
  if (ACount >= FMaxCount) then begin
    FMaxCount := ACount + 20;
    SetLength(FItems, FMaxCount);
    //Здесь FCount + 1, потому что иначе затирается последний элемент.
    //В результате утечки и потеря считанного форматирования.
    for i := FCount + 1 to FMaxCount - 1 do
      FItems[i] := TZXLSXDiffFormattingItem.Create();
  end;
  FCount := ACount;
end;

procedure TZXLSXDiffFormatting.Clear();
begin
  FCount := 0;
end;

function TZXLSXDiffFormatting.GetItem(num: integer): TZXLSXDiffFormattingItem;
begin
  if ((num >= 0) and (num < Count)) then
    result := FItems[num]
  else
    result := nil;
end;

procedure TZXLSXDiffFormatting.SetItem(num: integer; const Value: TZXLSXDiffFormattingItem);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FItems[num].Assign(Value);
end;

////:::::::::::::  TZXLSXDiffBorderItemStyle :::::::::::::::::////

constructor TZXLSXDiffBorderItemStyle.Create();
begin
  Clear();
end;

procedure TZXLSXDiffBorderItemStyle.Assign(Source: TPersistent);
var bs: TZXLSXDiffBorderItemStyle; b: boolean;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorderItemStyle) then begin
      b := false;
      bs := Source as TZXLSXDiffBorderItemStyle;

      FUseStyle  := bs.UseStyle;
      FUseColor  := bs.UseColor;
      FColor     := bs.Color;
      FWeight    := bs.Weight;
      FLineStyle := bs.LineStyle;
    end;
  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffBorderItemStyle.Clear();
begin
  FUseStyle  := false;
  FUseColor  := false;
  FColor     := clBlack;
  FWeight    := 1;
  FLineStyle := ZENone;
end;

////::::::::::::: TZXLSXDiffBorder :::::::::::::::::////

constructor TZXLSXDiffBorder.Create();
var i: integer;
begin
  for i := 0 to 5 do
    FBorder[i] := TZXLSXDiffBorderItemStyle.Create();
  Clear();
end;

destructor TZXLSXDiffBorder.Destroy();
var i: integer;
begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);
  inherited;
end;

procedure TZXLSXDiffBorder.Assign(Source: TPersistent);
var brd: TZXLSXDiffBorder; b: boolean; i: TZBordersPos;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffBorder) then begin
      b := false;
      brd := Source as TZXLSXDiffBorder;
      for i := bpLeft to bpDiagonalRight do
        FBorder[Ord(i)].Assign(brd[i]);
    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffBorder.Clear();
var i: TZBordersPos;
begin
  for i := bpLeft to bpDiagonalRight do
    FBorder[Ord(i)].Clear();
end;

function TZXLSXDiffBorder.GetBorder(Num: TZBordersPos): TZXLSXDiffBorderItemStyle;
begin
  result := nil;
  if ((num >= bpLeft) and (num <= bpDiagonalRight)) then
    result := FBorder[ord(num)];
end;

procedure TZXLSXDiffBorder.SetBorder(Num: TZBordersPos; const Value: TZXLSXDiffBorderItemStyle);
begin
  if ((num >= bpLeft) and (num <= bpDiagonalRight)) then
    if (Assigned(Value)) then
      FBorder[Ord(num)].Assign(Value);
end;

////::::::::::::: TZXLSXDiffFormatingItem :::::::::::::::::////

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
var dxfItem: TZXLSXDiffFormattingItem; b: boolean;
begin
  b := true;
  if (Assigned(Source)) then
    if (Source is TZXLSXDiffFormattingItem) then begin
      b := false;
      dxfItem := Source as TZXLSXDiffFormattingItem;

      FUseFont         := dxfItem.UseFont;
      FUseFontColor    := dxfItem.UseFontColor;
      FUseFontStyles   := dxfItem.UseFontStyles;
      FFontColor       := dxfItem.FontColor;
      FFontStyles      := dxfItem.FontStyles;
      FUseBorder       := dxfItem.UseBorder;
      FUseFill         := dxfItem.UseFill;
      FUseCellPattern  := dxfItem.UseCellPattern;
      FCellPattern     := dxfItem.CellPattern;
      FUseBGColor      := dxfItem.UseBGColor;
      FBGColor         := dxfItem.BGColor;
      FUsePatternColor := dxfItem.UsePatternColor;
      FPatternColor    := dxfItem.PatternColor;
      FBorders.Assign(dxfItem.Borders);
    end;

  if (b) then
    inherited;
end; //Assign

procedure TZXLSXDiffFormattingItem.Clear();
begin
  FUseFont := false;
  FUseFontColor := false;
  FUseFontStyles := false;
  FFontColor := clBlack;
  FFontStyles := [];
  FUseBorder := false;
  FBorders.Clear();
  FUseFill := false;
  FUseCellPattern := false;
  FCellPattern := ZPNone;
  FUseBGColor := false;
  FBGColor := clWindow;
  FUsePatternColor := false;
  FPatternColor := clWindow;
end; //Clear

// END Differential Formatting
////////////////////////////////////////////////////////////////////////////////

////::::::::::::: TZEXLSXNumberFormats :::::::::::::::::////

constructor TZEXLSXNumberFormats.Create();
var i: integer;
begin
  FStyleFmtIDCount := 0;
  FFormatsCount := 164;
  SetLength(FFormats, FFormatsCount);
  for i := 0 to FFormatsCount - 1 do
    FFormats[i] := '';

  //Some "Standart" formats for xlsx:
  FFormats[1]  := '0';
  FFormats[2]  := '0.00';
  FFormats[3]  := '#,##0';
  FFormats[4]  := '#,##0.00';
  FFormats[5]  := '$#,##0;\-$#,##0';
  FFormats[6]  := '$#,##0;[Red]\-$#,##0';
  FFormats[7]  := '$#,##0.00;\-$#,##0.00';
  FFormats[8]  := '$#,##0.00;[Red]\-$#,##0.00';
  FFormats[9]  := '0%';
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

  FFormats[30] := 'm/d/yy';

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

//Find format in formats. Return -1 if format not foud.
//INPUT
//  const value: string - format
//RETURN
//      integer - >= 0 - index number if store.
//                 -1 - not found
function TZEXLSXNumberFormats.FindFormatID(const value: string): integer;
var i: integer;
begin
  Result := -1;
  for i := 0 to FFormatsCount - 1 do
    if (FFormats[i] = value) then begin
      Result := i;
      break;
    end;
end;

function TZEXLSXNumberFormats.GetFormat(num: integer): string;
begin
  Result := '';
  if ((num >= 0) and (num < FFormatsCount)) then
    Result := FFormats[num];
end;

procedure TZEXLSXNumberFormats.SetFormat(num: integer; const value: string);
var i: integer;
begin
  if ((num >= 0) and (num < FFormatsCount)) then
    FFormats[num] := value
  else if (num >= 0) then begin
    SetLength(FFormats, num + 1);
    for i := FFormatsCount to num do
      FFormats[i] := '';
    FFormats[num] := value;
    FFormatsCount := num + 1;
  end;
end;

function TZEXLSXNumberFormats.GetStyleFMTID(num: integer): integer;
begin
  if ((num >= 0) and (num < FStyleFmtIDCount)) then
    Result := FStyleFmtID[num]
  else
    Result := 0;
end;

function TZEXLSXNumberFormats.IsDateFormat(StyleNum: integer): boolean;
var fmtId: integer;
begin
  Result := false;
  if ((StyleNum >= 0) and (StyleNum < FStyleFmtIDCount)) then
    fmtId := FStyleFmtID[StyleNum]
  else
    exit;

  //If default fmtID
  if ((fmtId >= 0) and (fmtId < 100)) then
    Result := fmtId in [14..22, 27..36, 45..47, 50..58, 71..76, 78..81]
  else
    Result := GetXlsxNumberFormatType(FFormats[fmtId]) = ZE_NUMFORMAT_IS_DATETIME;
end;

procedure TZEXLSXNumberFormats.ReadNumFmts(const xml: TZsspXMLReaderH);
var temp: integer;
begin
  with THTMLEncoding.Create do
    try
      while xml.ReadToEndTagByName('numFmts') do begin
        if (xml.TagName = 'numFmt') then
          if (TryStrToInt(xml.Attributes['numFmtId'], temp)) then
            Format[temp] := Decode(xml.Attributes['formatCode']);
      end;
    finally
      Free;
    end;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTID(num: integer; const value: integer);
begin
  if ((num >= 0) and (num < FStyleFmtIDCount)) then
    FStyleFmtID[num] := value;
end;

procedure TZEXLSXNumberFormats.SetStyleFMTCount(value: integer);
begin
  if (value >= 0) then begin
    if (value > FStyleFmtIDCount) then
      SetLength(FStyleFmtID, value);
    FStyleFmtIDCount := value
  end;
end;

////::::::::::::: TZEXLSXReadHelper :::::::::::::::::////

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

////::::::::::::: TZEXLSXWriteHelper :::::::::::::::::////

//Generate next RID for references
function TZEXLSXWriteHelper.GenerateRID(): integer;
begin
  inc(FCurrentRID);
  result := FCurrentRID;
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
var num: integer;
begin
  num := FHyperLinksCount;
  inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then begin
    inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[num].RID        := GenerateRID();
  FHyperLinks[num].RelType    := TRelationType.rtHyperlink;
  FHyperLinks[num].TargetMode := ATargetMode;
  FHyperLinks[num].CellRef    := ACellRef;
  FHyperLinks[num].Target     := ATarget;
  FHyperLinks[num].ScreenTip  := AScreenTip;
end;

//Add hyperlink
//INPUT
//     const ATarget: string     - drawing target (../drawings/drawing2.xml)
function TZEXLSXWriteHelper.AddDrawing(const ATarget: string): integer;
var num: integer;
begin
  num := FHyperLinksCount;
  Inc(FHyperLinksCount);

  if (FHyperLinksCount >= FMaxHyperLinksCount) then begin
    Inc(FMaxHyperLinksCount, 20);
    SetLength(FHyperLinks, FMaxHyperLinksCount);
  end;

  FHyperLinks[num].RID := GenerateRID();
  FHyperLinks[num].RelType := TRelationType.rtDrawing;
  FHyperLinks[num].TargetMode := '';
  FHyperLinks[num].CellRef := '';
  FHyperLinks[num].Target := ATarget;
  FHyperLinks[num].ScreenTip := '';
  Result := FHyperLinks[num].RID;
  FisHaveDrawings := True;
end;

//Writes tag <hyperlinks> .. </hyperlinks>
//INPUT
//     const xml: TZsspXMLWriterH
procedure TZEXLSXWriteHelper.WriteHyperLinksTag(const xml: TZsspXMLWriterH);
var i: integer;
begin
  if (FHyperLinksCount > 0) then begin
    xml.Attributes.Clear();
    xml.WriteTagNode('hyperlinks', true, true, true);
    for i := 0 to FHyperLinksCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('ref', FHyperLinks[i].CellRef);
      xml.Attributes.Add('r:id', 'rId' + IntToStr(FHyperLinks[i].RID));

      if (FHyperLinks[i].ScreenTip <> '') then
        xml.Attributes.Add('tooltip', FHyperLinks[i].ScreenTip);

      xml.WriteEmptyTag('hyperlink', true);

      {
      xml.Attributes.Add('Id', 'rId' + IntToStr(FHyperLinks[i].ID));
      xml.Attributes.Add('Type', ZEXLSXGetRelationName(6));
      xml.Attributes.Add('Target', FHyperLinks[i].Target);
      if (FHyperLinks[i].TargetMode <> '')
         xml.Attributes.Add('TargetMode', FHyperLinks[i].TargetMode);
      }
    end; //for i
    xml.WriteEndTagNode(); //hyperlinks
  end; //if
end; //WriteHyperLinksTag

//Create sheet relations
//INPUT
//  const Stream: TStream
//        TextConverter: TAnsiToCPConverter
//        CodePageName: string
//        BOM: ansistring
//RETURN
function TZEXLSXWriteHelper.CreateSheetRels(const Stream: TStream;
                                            TextConverter: TAnsiToCPConverter;
                                            CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH; i: integer;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);

    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    xml.WriteTagNode('Relationships', true, true, false);

    for i := 0 to FHyperLinksCount - 1 do
      ZEAddRelsRelation(xml,
        'rId' + IntToStr(FHyperLinks[i].RID),
        FHyperLinks[i].RelType,
        FHyperLinks[i].Target,
        FHyperLinks[i].TargetMode);

    xml.WriteEndTagNode(); //Relationships
  finally
    xml.Free();
  end;
end; //CreateSheetRels

procedure TZEXLSXWriteHelper.AddSheetHyperlink(PageNum: integer);
begin
  SetLength(FSheetHyperlinksArray, FSheetHyperlinksCount + 1);
  FSheetHyperlinksArray[FSheetHyperlinksCount] := PageNum;
  inc(FSheetHyperlinksCount);
end;

function TZEXLSXWriteHelper.IsSheetHaveHyperlinks(PageNum: integer): boolean;
var i: integer;
begin
  result := false;
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

//Возвращает номер Relations из rels
//INPUT
//  const name: string - текст отношения
//RETURN
//      integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationNumber(const name: string): TRelationType;
var rec: TContentTypeRec;
begin
  result := TRelationType.rtNone;
  for rec in CONTENT_TYPES do begin
    if rec.rel = name then
      exit(rec.ftype);
  end;
end; //ZEXLSXGetRelationNumber

//Возвращает текст Relations для rels
//INPUT
//      num: integer - номер отношения
//RETURN
//      integer - номер отношения. -1 - не определено
function ZEXLSXGetRelationName(num: TRelationType): string;
var rec: TContentTypeRec;
begin
  result := '';
  for rec in CONTENT_TYPES do begin
    if rec.ftype = num then
      exit(rec.rel);
  end;
end; //ZEXLSXGetRelationName

function XLSXBoolToStr(value: boolean): string;
begin
  if (value) then
    result := 'true'
  else
    result := 'false';
end;

//Читает тему (themeXXX.xml)
//INPUT
//  var Stream: TStream                   - поток чтения
//  var ThemaFillsColors: TIntegerDynArray - массив с цветами заливки
//  var ThemaColorCount: integer          - кол-во цветов заливки
//RETURN
//      boolean - true - всё прочиталось успешно
function ZEXSLXReadTheme(var Stream: TStream; var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  maxCount: integer;
  flag: boolean;

  procedure _addFillColor(const _rgb: string);
  begin
    inc(ThemaColorCount);
    if (ThemaColorCount >= maxCount) then begin
      maxCount := ThemaColorCount + 20;
      SetLength(ThemaFillsColors, maxCount);
    end;
    ThemaFillsColors[ThemaColorCount - 1] := HTMLHexToColor(_rgb);
  end; //_addFillColor

begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  flag := false;
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;
    ThemaColorCount := 0;
    maxCount := -1;
    while (not xml.Eof()) do begin
      xml.ReadTag();
      if (xml.TagName = 'a:clrScheme') then  begin
        if (xml.IsTagStart) then
          flag := true;
        if (xml.IsTagEnd) then
          flag := false;
      end else if ((xml.TagName = 'a:sysClr') and (flag) and (xml.IsTagStartOrClosed)) then begin
        _addFillColor(xml.Attributes.ItemsByName['lastClr']);
      end else if ((xml.TagName = 'a:srgbClr') and (flag) and (xml.IsTagStartOrClosed)) then begin
        _addFillColor(xml.Attributes.ItemsByName['val']);
      end;
    end; //while

    result := true;
  finally
    xml.Free();
  end;
end; //ZEXSLXReadThema

//Читает список нужных файлов из [Content_Types].xml
//INPUT
//  var Stream: TStream             - поток чтения
//  var FileArray: TArray<TZXLSXFileItem>  - список файлов
//  var FilesCount: integer         - кол-во файлов
//RETURN
//      boolean - true - всё прочиталось успешно
function ZEXSLXReadContentTypes(var Stream: TStream; var FileArray: TArray<TZXLSXFileItem>; var FilesCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  contType: string;
  rec: TContentTypeRec;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    if xml.BeginReadStream(Stream) <> 0 then
      exit;
    FilesCount := 0;
    while not xml.Eof() do begin
      xml.ReadTag();
      if xml.IsTagClosedByName('Override') then begin
        contType := xml.Attributes.ItemsByName['ContentType'];
        for rec in CONTENT_TYPES do begin
          if contType = rec.name then begin
            SetLength(FileArray, FilesCount + 1);
            FileArray[FilesCount].name     := xml.Attributes.ItemsByName['PartName'];
            FileArray[FilesCount].original := xml.Attributes.ItemsByName['PartName'];
            FileArray[FilesCount].ftype    := rec.ftype;
            inc(FilesCount);
            break;
          end;
        end;
      end;
    end;
    result := true;
  finally
    xml.Free();
  end;
end; //ZEXSLXReadContentTypes

//Читает строки из sharedStrings.xml
//INPUT
//  var Stream: TStream           - поток для чтения
//  var StrArray: TStringDynArray - возвращаемый массив со строками
//  var StrCount: integer         - кол-во элементов
//RETURN
//      boolean - true - всё ок
function ZEXSLXReadSharedStrings(var Stream: TStream; sharedStrings: TObjectList<TRichText>): boolean;
var
  xml: TZsspXMLReaderH;
  richText: TRichText;
  part: TRichString;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while not xml.Eof() do begin
      xml.ReadTag();
      if xml.IsTagStartByName('si') then begin
        richText := TRichText.Create();

        while xml.ReadToEndTagByName('si') do begin
          if xml.IsTagEndByName('t') then begin
            part := TRichString.Create();
            part.Font := nil;
            part.Text := xml.TextBeforeTag;
            richText.List.Add(part);
          end else
          if xml.IsTagStartByName('r') then begin
            part := TRichString.Create();
            while xml.ReadToEndTagByName('r') do begin
              if xml.IsTagEndByName('t') then
                part.Text := xml.TextBeforeTag
              else
              if xml.IsTagStartByName('rPr') then begin
                part.Font := TZFont.Create();
                while xml.ReadToEndTagByName('rPr') do begin
                  if xml.IsTagClosedByName('b') then
                    part.Font.Style := part.Font.Style + [fsBold]
                  else
                  if xml.IsTagClosedByName('i') then
                    part.Font.Style := part.Font.Style + [fsItalic]
                  else
                  if xml.IsTagClosedByName('u') then begin
                    part.Font.Style := part.Font.Style + [fsUnderline];
                    part.Font.DoubleInderLine := xml.Attributes['val'] = 'double';
                  end else
                  if xml.IsTagClosedByName('sz') then
                    part.Font.Size := ZETryStrToFloat(xml.Attributes['val'], 0)
                  else
                  if xml.IsTagClosedByName('rFont') then
                    part.Font.Name := xml.Attributes['val']
                  else
                  if xml.IsTagClosedByName('family') then
                    part.Font.Family := StrToIntDef(xml.Attributes['val'], 0)
                  else
                  if xml.IsTagClosedByName('charset') then
                    part.Font.Charset := StrToIntDef(xml.Attributes['val'], 0)
                  else
                  if xml.IsTagClosedByName('scheme') then
                    part.Font.Scheme := xml.Attributes['val']
                  else
                  if xml.IsTagClosedByName('color') then begin
                    if not xml.Attributes['theme'].IsEmpty then
                      part.Font.ColorTheme := StrToIntDef(xml.Attributes['theme'], 0)
                    else
                    if not xml.Attributes['rgb'].IsEmpty then
                      part.Font.Color := ARGBToColor(xml.Attributes['rgb'])
                  end;
                end;
              end;
            end;
            richText.List.Add(part);
          end;
        end; //si
        sharedStrings.Add(richText);
      end;
    end;

    result := true;
  finally
    xml.Free();
  end;
end; //ZEXSLXReadSharedStrings

//Получить условное форматирование и оператор из xlsx
//INPUT
//  const xlsxCfType: string              - xlsx тип условного форматирования
//  const xlsxCFOperator: string          - xlsx оператор
//  out CFCondition: TZCondition          - распознанное условие
//  out CFOperator: TZConditionalOperator - распознанный оператор
//RETURN
//      boolean - true - условное форматирование и оператор успешно распознаны
function ZEXLSX_getCFCondition(const xlsxCfType, xlsxCFOperator: string; out CFCondition: TZCondition; out CFOperator: TZConditionalOperator): boolean;
var isCheckOperator: boolean;
  procedure _SetCFOperator(AOperator: TZConditionalOperator);
  begin
    CFOperator := AOperator;
    CFCondition := ZCFCellContentOperator;
  end;

  //Проверить тип условного форматирования
  //  out isNeddCheckOperator: boolean - возвращает, нужно ли проверять
  //                                     оператор
  //RETURN
  //      boolean - true - всё ок, можно проверять далее
  function _CheckXLSXCfType(out isNeddCheckOperator: boolean): boolean;
  begin
    result := true;
    isNeddCheckOperator := true;
    if (xlsxCfType = 'cellIs') then begin
    end else if (xlsxCfType = 'containsText') then
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
      result := false;
  end; //_CheckXLSXCfType

  //Проверить оператор
  function _CheckCFoperator(): boolean;
  begin
    result := true;
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
      result := false;
  end; //_CheckCFoperator

begin
  result := false;
  CFCondition := ZCFNumberValue;
  CFOperator := ZCFOpGT;

  if (_CheckXLSXCfType(isCheckOperator)) then begin
    if (isCheckOperator) then
      result := _CheckCFoperator()
    else
      result := true;
  end;
end; //ZEXLSX_getCFCondition

function ZEXLSXReadDrawingRels(sheet: TZSheet; Stream: TStream): boolean;
var xml: TZsspXMLReaderH;
    target: string;
    I, id: Integer;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while xml.ReadToEndTagByName('Relationships') do begin
      if xml.IsTagClosedByName('Relationship') then begin
        id := StrToInt(xml.Attributes.ItemsByName['Id'].Substring(3));
        target := xml.Attributes.ItemsByName['Target'];

        for I := 0 to sheet.Drawing.Count-1 do begin
           if sheet.Drawing[i].RelId = id then
              sheet.Drawing[i].Name := target.Substring(9);
        end;
      end;
    end;
    result := true;
  finally
    xml.Free();
  end;
end;

function ZEXLSXReadDrawing(sheet: TZSheet; Stream: TStream): boolean;
var xml: TZsspXMLReaderH;
    picture: TZEPicture;

    procedure ReadFrom();
    begin
      while xml.ReadToEndTagByName('xdr:from') do begin
        if xml.IsTagEndByName('xdr:col') then
          picture.FromCol := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:colOff') then
          picture.FromColOff := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:row') then
          picture.FromRow := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:rowOff') then
          picture.FromRowOff := StrToInt(xml.TextBeforeTag);
      end;
    end;

    procedure ReadTo();
    begin
      while xml.ReadToEndTagByName('xdr:to') do begin
        if xml.IsTagEndByName('xdr:col') then
          picture.ToCol := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:colOff') then
          picture.ToColOff := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:row') then
          picture.ToRow := StrToInt(xml.TextBeforeTag)
        else if xml.IsTagEndByName('xdr:rowOff') then
          picture.ToRowOff := StrToInt(xml.TextBeforeTag);
      end;
    end;

    procedure ReadPic();
    begin
      while xml.ReadToEndTagByName('xdr:pic') do begin
        if xml.IsTagClosedByName('xdr:cNvPr') then begin
          picture.Description := xml.Attributes['descr'];
          picture.Title := xml.Attributes['name'];
          picture.Id := StrToInt(xml.Attributes['id']);
        end;// else
        if xml.IsTagStartOrClosedByName('a:blip') then begin
          picture.RelId := StrToInt(xml.Attributes['r:embed'].Substring(3)); // skip "rId"
        end;

        if xml.IsTagStartByName('a:xfrm') then begin
          while xml.ReadToEndTagByName('a:xfrm') do begin
            if xml.IsTagClosedByName('a:off') then begin
              picture.FrmOffX := StrToInt(xml.Attributes['x']);
              picture.FrmOffY := StrToInt(xml.Attributes['y']);
            end
            else if xml.IsTagClosedByName('a:ext') then begin
              picture.FrmExtCX := StrToInt(xml.Attributes['cx']);
              picture.FrmExtCY := StrToInt(xml.Attributes['cy']);
            end
          end;
        end;
      end;
    end;

    procedure ReadImageItem();
    begin
      picture := sheet.Drawing.Add();
      if xml.Attributes['editAs'] = 'absolute' then
        picture.CellAnchor := ZAAbsolute
      else
        picture.CellAnchor := ZACell;

      while xml.ReadToEndTagByName('xdr:twoCellAnchor') do begin
        if xml.IsTagStartByName('xdr:from') then begin
          ReadFrom();
        end;// else

        if xml.IsTagStartByName('xdr:to') then begin
          ReadTo();
        end;// else

        if xml.IsTagStartByName('xdr:pic') then begin
          ReadPic();
        end
      end;
    end;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;
    picture := nil;

    while xml.ReadToEndTagByName('xdr:wsDr') do begin
      if xml.IsTagStartByName('xdr:twoCellAnchor') then begin
        ReadImageItem();
      end;
    end;
  finally
    xml.Free();
  end;
end;

//Читает лист книги
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//  var Stream: TStream                   - поток для чтения
//  const SheetRelations: TZXLSXRelations - свойства страницы
//  var StrArray: TStringDynArray         - строки для подстановки
//      StrCount: integer                 - кол-во строк подстановки
//  var Relations: TZXLSXRelationsArray   - отношения
//      RelationsCount: integer           - кол-во отношений
//  var MaximumDigitWidth: double         - ширина самого широкого числа в пикселях
//      ReadHelper: TZEXLSXReadHelper     -
//RETURN
//      boolean - true - страница прочиталась успешно
function ZEXSLXReadSheet(var XMLSS: TZWorkBook;
                         var Stream: TStream;
                         const SheetRelations: PZXLSXRelations;
                         sharedStrings: TObjectList<TRichText>;
                         var Relations: TZXLSXRelationsArray;
                         RelationsCount: integer;
                         MaximumDigitWidth: double;
                         ReadHelper: TZEXLSXReadHelper): boolean;
var
  xml: TZsspXMLReaderH;
  currentPage: integer;
  currentRow: integer;
  currentCol: integer;
  currentSheet: TZSheet;
  currentCell: TZCell;
  str: string;
  tempReal: real;
  tempInt: integer;
  tempDate: TDateTime;
  tempFloat: Double;

  //Проверить кол-во строк
  procedure CheckRow(const RowCount: integer);
  begin
    if (currentSheet.RowCount < RowCount) then
      currentSheet.RowCount := RowCount;
  end;

  //Проверить кол-во столбцов
  procedure CheckCol(const ColCount: integer);
  begin
    if (currentSheet.ColCount < ColCount) then
      currentSheet.ColCount := ColCount
  end;

  //Чтение строк/столбцов
  procedure _ReadSheetData();
  var
    t: integer;
    v: string;
    _num: integer;
    _type: string;
    _cr, _cc: integer;
    maxCol: integer;
  begin
    _cr := 0;
    _cc := 0;
    maxCol := 0;
    CheckRow(1);
    CheckCol(1);
    while xml.ReadToEndTagByName('sheetData') do begin
      //ячейка
      if (xml.TagName = 'c') then begin
        str := xml.Attributes.ItemsByName['r']; //номер
        if (str > '') then
          if TZEFormula.GetCellCoords(str, _cc, _cr) then begin
            currentCol := _cc;
            CheckCol(_cc + 1);
          end;

        _type := xml.Attributes.ItemsByName['t']; //тип

        //s := xml.Attributes.ItemsByName['cm'];
        //s := xml.Attributes.ItemsByName['ph'];
        //s := xml.Attributes.ItemsByName['vm'];
        v := '';
        _num := 0;
        currentCell := currentSheet.Cell[currentCol, currentRow];
        str := xml.Attributes.ItemsByName['s']; //стиль
        if (str > '') then
          if (tryStrToInt(str, t)) then
            currentCell.CellStyle := t;
        if (xml.IsTagStart) then
        while xml.ReadToEndTagByName('c') do begin
          //is пока игнорируем
          if xml.IsTagEndByName('v') or xml.IsTagEndByName('t') then begin
            if (_num > 0) then
              v := v + sLineBreak;
            v := v + xml.TextBeforeTag;
            inc(_num);
          end else if xml.IsTagEndByName('f') then
            currentCell.Formula := ZEReplaceEntity(xml.TextBeforeTag);

        end; //while

        //Возможные типы:
        //  s - sharedstring
        //  b - boolean
        //  n - number
        //  e - error
        //  str - string
        //  inlineStr - inline string ??
        //  d - date
        //  тип может отсутствовать. Интерпретируем в таком случае как ZEGeneral
        if (_type = '') then
          currentCell.CellType := ZEGeneral
        else if (_type = 'n') then begin
          currentCell.CellType := ZENumber;
          //Trouble: if cell style is number, and number format is date, then
          // cell style is date. F****** m$!
          if (ReadHelper.NumberFormats.IsDateFormat(currentCell.CellStyle)) then
            if (ZEIsTryStrToFloat(v, tempFloat)) then begin
              currentCell.CellType := ZEDateTime;
              v := ZEDateTimeToStr(tempFloat);
            end;
        end else if (_type = 's') then begin
          currentCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < sharedStrings.Count)) then begin
              if (sharedStrings[t].List.Count > 1) then begin
                currentCell.RichText := TRichText.Create();
                currentCell.RichText.Assign(sharedStrings[t]);
              end else
                v := sharedStrings[t].ToString();
            end;
        end else if (_type = 'd') then begin
          currentCell.CellType := ZEDateTime;
          if (TryZEStrToDateTime(v, tempDate)) then
            v := ZEDateTimeToStr(tempDate)
          else
          if (ZEIsTryStrToFloat(v, tempFloat)) then
            v := ZEDateTimeToStr(tempFloat)
          else
            currentCell.CellType := ZEString;
        end;

        if assigned(currentCell.RichText) then
          currentCell.Data := currentCell.RichText.ToString()
        else
          currentCell.Data := ZEReplaceEntity(v);

        inc(currentCol);
        CheckCol(currentCol + 1);
        if currentCol > maxCol then
           maxCol := currentCol;
      end else
      //строка
      if xml.IsTagStartOrClosedByName('row') then begin
        currentCol := 0;
        str := xml.Attributes.ItemsByName['r']; //индекс строки
        if (str > '') then
          if (TryStrToInt(str, t)) then begin
            currentRow := t - 1;
            CheckRow(t);
          end;
        //s := xml.Attributes.ItemsByName['collapsed'];
        //s := xml.Attributes.ItemsByName['customFormat'];
        //s := xml.Attributes.ItemsByName['customHeight'];
        currentSheet.Rows[currentRow].Hidden := ZETryStrToBoolean(xml.Attributes.ItemsByName['hidden'], false);

        str := xml.Attributes.ItemsByName['ht']; //в поинтах
        if (str > '') then begin
          tempReal := ZETryStrToFloat(str, 10);
          currentSheet.Rows[currentRow].Height := tempReal;
          //tempReal := tempReal / 2.835; //???
          //currentSheet.Rows[currentRow].HeightMM := tempReal;
        end
        else
          currentSheet.Rows[currentRow].Height := currentSheet.DefaultRowHeight;

        str := xml.Attributes.ItemsByName['outlineLevel'];
        currentSheet.Rows[currentRow].OutlineLevel := StrToIntDef(str, 0);

        //s := xml.Attributes.ItemsByName['ph'];

        str := xml.Attributes.ItemsByName['s']; //номер стиля
        if (str > '') then
          if (TryStrToInt(str, t)) then begin
            //нужно подставить нужный стиль
          end;
        //s := xml.Attributes.ItemsByName['spans'];
        //s := xml.Attributes.ItemsByName['thickBot'];
        //s := xml.Attributes.ItemsByName['thickTop'];

        if xml.IsTagClosed then begin
          inc(currentRow);
          CheckRow(currentRow + 1);
        end;
      end else
      //конец строки
      if xml.IsTagEndByName('row') then begin
        inc(currentRow);
        CheckRow(currentRow + 1);
      end;
    end; //while
    currentSheet.ColCount := maxCol;
  end; //_ReadSheetData

  //Чтение диапазона ячеек с автофильтром
  procedure _ReadAutoFilter();
  begin
    currentSheet.AutoFilter := xml.Attributes.ItemsByName['ref'];
  end;

  //Чтение объединённых ячеек
  procedure _ReadMerge();
  var left, top, right, bottom: integer;
  begin
    while xml.ReadToEndTagByName('mergeCells') do begin
      if xml.IsTagStartOrClosedByName('mergeCell') then begin
        str := xml.Attributes.ItemsByName['ref'];
        if TZEFormula.GetCellRange(str, left, top, right, bottom) then begin
          CheckCol(right+1);
          CheckRow(bottom+1);
          currentSheet.MergeCells.AddRectXY(left, top, right, bottom);
        end;
      end;
    end;
  end;

  //Столбцы
  procedure _ReadCols();
  type
    TZColInf = record
      min,max: integer;
      bestFit,hidden: boolean;
      outlineLevel: integer;
      width: integer;
    end;
  var
    i, j: integer; t: real;
    colInf: TArray<TZColInf>;
  const MAX_COL_DIFF = 500;
  begin
    i := 0;
    while xml.ReadToEndTagByName('cols') do begin
      if (xml.TagName = 'col') and xml.IsTagStartOrClosed then begin
        SetLength(colInf, i + 1);

        colInf[i].min := StrToIntDef(xml.Attributes.ItemsByName['min'], 0);
        colInf[i].max := StrToIntDef(xml.Attributes.ItemsByName['max'], 0);
        // защита от сплошного диапазона
        // когда значение _мах = 16384
        // но чтобы уж наверняка, проверим на MAX_COL_DIFF колонок подряд.
        if (colInf[i].max - colInf[i].min) > MAX_COL_DIFF then
            colInf[i].max := colInf[i].min + MAX_COL_DIFF;

        colInf[i].outlineLevel := StrToIntDef(xml.Attributes.ItemsByName['outlineLevel'], 0);
        str := xml.Attributes.ItemsByName['hidden'];
        if (str > '') then colInf[i].hidden := ZETryStrToBoolean(str);
        str := xml.Attributes.ItemsByName['bestFit'];
        if (str > '') then colInf[i].bestFit := ZETryStrToBoolean(str);

        str := xml.Attributes.ItemsByName['width'];
        if (str > '') then begin
          t := ZETryStrToFloat(str, 5.14509803921569);
          //t := 10 * t / 5.14509803921569;
          //А.А.Валуев. Формулы расёта ширины взяты здесь - https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
          t := Trunc(((256 * t + Trunc(128 / MaximumDigitWidth)) / 256) * MaximumDigitWidth);
          colInf[i].width := Trunc(t);
        end;

        inc(i);
      end; //if
    end; //while

    for I := Low(colInf) to High(colInf) do begin
      for j := colInf[i].min to colInf[i].max do begin
        CheckCol(j);
        currentSheet.Columns[j-1].AutoFitWidth := colInf[i].bestFit;
        currentSheet.Columns[j-1].Hidden := colInf[i].hidden;
        currentSheet.Columns[j-1].WidthPix := colInf[i].width;
      end;
    end;
  end; //_ReadCols

  function _StrToMM(const st: string; var retFloat: real): boolean;
  begin
    result := false;
    if (str > '') then begin
      retFloat := ZETryStrToFloat(st, -1);
      if (retFloat > -1) then begin
        result := true;
        retFloat := retFloat * ZE_MMinInch;
      end;
    end;
  end; //_StrToMM

  procedure _GetDimension();
  var left, top, right, bottom: integer;
  begin
    var str := xml.Attributes.ItemsByName['ref'];
    if TZEFormula.GetCellRange(str, left, top, right, bottom) then begin
      CheckCol(right);
      CheckRow(bottom);
    end;
  end;

  //Чтение ссылок
  procedure _ReadHyperLinks();
  var _c, _r, i: integer;
  begin
    _c := 0;
    _r := 0;
    while xml.ReadToEndTagByName('hyperlinks') do begin
      if xml.IsTagClosedByName('hyperlink') then begin
        str := xml.Attributes.ItemsByName['ref'];
        if (str > '') then
          if TZEFormula.GetCellCoords(str, _c, _r) then begin
            CheckRow(_r);
            CheckCol(_c);
            currentSheet.Cell[_c, _r].HRefScreenTip := xml.Attributes.ItemsByName['tooltip'];
            str := xml.Attributes.ItemsByName['r:id'];
            //по r:id подставить ссылку
            for i := 0 to RelationsCount - 1 do
              if ((Relations[i].id = str) and (Relations[i].ftype = TRelationType.rtHyperlink)) then begin
                currentSheet.Cell[_c, _r].Href := Relations[i].target;
                break;
              end;
          end;
      end;
    end;
  end;

  procedure _ReadSheetPr();
  begin
    while xml.ReadToEndTagByName('sheetPr') do begin
      if xml.TagName = 'tabColor' then
        currentSheet.TabColor := ARGBToColor(xml.Attributes.ItemsByName['rgb']);

      if xml.TagName = 'pageSetUpPr' then
        currentSheet.FitToPage := xml.AttribBool('fitToPage');

      if xml.TagName = 'outlinePr' then begin
        currentSheet.ApplyStyles  := xml.AttribBool('applyStyles');
        currentSheet.SummaryBelow := xml.AttribBool('summaryBelow', true);
        currentSheet.SummaryRight := xml.AttribBool('summaryRight', true);
      end;
    end;
  end; //_ReadSheetPr();

  procedure _ReadRowBreaks();
  begin
    currentSheet.RowBreaks := [];
    while xml.ReadToEndTagByName('rowBreaks') do begin
      if xml.TagName = 'brk' then
        currentSheet.RowBreaks := currentSheet.RowBreaks + [ xml.AttribInt('id', 0) ];
    end;
  end;

  procedure _ReadColBreaks();
  begin
    currentSheet.ColBreaks := [];
    while xml.ReadToEndTagByName('colBreaks') do begin
      if xml.TagName = 'brk' then
        currentSheet.ColBreaks := currentSheet.ColBreaks + [ xml.AttribInt('id', 0) ];
    end;
  end;

  //<sheetViews> ... </sheetViews>
  procedure _ReadSheetViews();
  var
    vValue, hValue: integer;
    SplitMode: TZSplitMode;
    s: string;
  begin
    while xml.ReadToEndTagByName('sheetViews') do begin
      if xml.IsTagStartByName('sheetView') or xml.IsTagClosedByName('sheetView') then begin
        s := xml.Attributes.ItemsByName['tabSelected'];
        // тут кроется проблема с выделением нескольких листов
        currentSheet.Selected := currentSheet.SheetIndex = 0;// s = '1';
        currentSheet.ViewMode := zvmNormal;
        if xml.Attributes.ItemsByName['view'] = 'pageBreakPreview' then
            currentSheet.ViewMode := zvmPageBreakPreview;
        currentSheet.ShowZeros := ZETryStrToBoolean(xml.Attributes.ItemsByName['showZeros'], true);
      end;

      if xml.IsTagClosedByName('pane') then begin
        SplitMode := ZSplitSplit;
        s := xml.Attributes.ItemsByName['state'];
        if (s = 'frozen') then
          SplitMode := ZSplitFrozen;

        s := xml.Attributes.ItemsByName['xSplit'];
        if (not TryStrToInt(s, vValue)) then
          vValue := 0;

        s := xml.Attributes.ItemsByName['ySplit'];
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
          currentSheet.SheetOptions.SplitHorizontalValue := PointToPixel(hValue/20);
        if (currentSheet.SheetOptions.SplitVerticalMode = ZSplitSplit) then
          currentSheet.SheetOptions.SplitVerticalValue := PointToPixel(vValue/20);

      end; //if
    end; //while
  end; //_ReadSheetViews()

  procedure _ReadConditionFormatting();
  var
    MaxFormulasCount: integer;
    _formulas: array of string;
    count: integer;
    _sqref: string;
    _type: string;
    _operator: string;
    _CFCondition: TZCondition;
    _CFOperator: TZConditionalOperator;
    _Style: string;
    _text: string;
    _isCFAdded: boolean;
    _isOk: boolean;
    //_priority: string;
    _CF: TZConditionalStyle;
    _tmpStyle: TZStyle;

    function _AddCF(): boolean;
    var
      s, ss: string;
      _len, i, kol: integer;
      a: array of array[0..5] of integer;
      _maxx: integer;
      ch: char;
      w, h: integer;

      function _GetOneArea(st: string): boolean;
      var
        i, j: integer;
        s: string;
        ch: char;
        _cnt: integer;
        tmpArr: array [0..1, 0..1] of integer;
        _isOk: boolean;
        t: integer;
        tmpB: boolean;

      begin
        result := false;
        if (st <> '') then begin
          st := st + ':';
          s := '';
          _cnt := 0;
          _isOk := true;
          for i := 1 to length(st) do begin
            ch := st[i];
            if (ch = ':') then begin
              if (_cnt < 2) then begin
                tmpB := TZEFormula.GetCellCoords(s, tmpArr[_cnt][0], tmpArr[_cnt][1]);
                _isOk := _isOk and tmpB;
              end;
              s := '';
              inc(_cnt);
            end else
              s := s + ch;
          end; //for

          if (_isOk) then
            if (_cnt > 0) then begin
              if (_cnt > 2) then
                _cnt := 2;

              a[kol][0] := _cnt;
              t := 1;
              for i := 0 to _cnt - 1 do
                for j := 0 to 1 do begin
                  a[kol][t] := tmpArr[i][j];
                  inc(t);
                end;
              result := true;
            end;
        end; //if
      end; //_GetOneArea

    begin
      result := false;
      if (_sqref <> '') then
      try
        _maxx := 4;
        SetLength(a, _maxx);
        ss := _sqref + ' ';
        _len := Length(ss);
        kol := 0;
        s := '';
        for i := 1 to _len do begin
          ch := ss[i];
          if (ch = ' ') then begin
            if (_GetOneArea(s)) then begin
              inc(kol);
              if (kol >= _maxx) then begin
                inc(_maxx, 4);
                SetLength(a, _maxx);
              end;
            end;
            s := '';
          end else
            s := s + ch;
        end; //for

        if (kol > 0) then begin
          currentSheet.ConditionalFormatting.Add();
          _CF := currentSheet.ConditionalFormatting[currentSheet.ConditionalFormatting.Count - 1];
          for i := 0 to kol - 1 do begin
            w := 1;
            h := 1;
            if (a[i][0] >= 2) then begin
              w := abs(a[i][3] - a[i][1]) + 1;
              h := abs(a[i][4] - a[i][2]) + 1;
            end;
            _CF.Areas.Add(a[i][1], a[i][2], w, h);
          end;
          result := true;
        end;
      finally
        SetLength(a, 0);
      end;
    end; //_AddCF

    //Применяем условный стиль
    procedure _TryApplyCF();
    var  b: boolean;
      num: integer;
      _id: integer;
      procedure _CheckTextCondition();
      begin
        if (count = 1) then
          if (_formulas[0] <> '') then
            _isOk := true;
      end;

      //Найти стиль
      //  пока будем делать так: предполагаем, что все ячейки в текущей области
      //  условного форматирования имеют один стиль. Берём стиль из левой верхней
      //  ячейки, клонируем его, применяем дифф. стиль, добавляем в хранилище стилей
      //  с учётом повторов.
      //TODO: потом нужно будет переделать
      //INPUT
      //      dfNum: integer - номер дифференцированного форматирования
      //RETURN
      //      integer - номер применяемого стиля
      function _getStyleIdxForDF(dfNum: integer): integer;
      var
        _df: TZXLSXDiffFormattingItem;
        _r, _c: integer;
        _t: integer;
        i: TZBordersPos;
      begin
        //_currSheet
        result := -1;
        if ((dfNum >= 0) and (dfNum < ReadHelper.DiffFormatting.Count)) then begin
          _df := ReadHelper.DiffFormatting[dfNum];
          _t := -1;

          if (_cf.Areas.Count > 0) then begin
            _r := _cf.Areas.Items[0].Row;
            _c := _cf.Areas.Items[0].Column;
            if ((_r >= 0) and (_r < currentSheet.RowCount)) then
              if ((_c >= 0) and (_c < currentSheet.ColCount)) then
                _t := currentSheet.Cell[_c, _r].CellStyle;
          end;

          _tmpStyle.Assign(XMLSS.Styles[_t]);

          if (_df.UseFont) then begin
            if (_df.UseFontStyles) then
              _tmpStyle.Font.Style := _df.FontStyles;
            if (_df.UseFontColor) then
              _tmpStyle.Font.Color := _df.FontColor;
          end;
          if (_df.UseFill) then begin
            if (_df.UseCellPattern) then
              _tmpStyle.CellPattern := _df.CellPattern;
            if (_df.UseBGColor) then
              _tmpStyle.BGColor := _df.BGColor;
            if (_df.UsePatternColor) then
              _tmpStyle.PatternColor := _df.PatternColor;
          end;
          if (_df.UseBorder) then
            for i := bpLeft to bpDiagonalRight do begin
              if (_df.Borders[i].UseStyle) then begin
                _tmpStyle.Border[i].Weight := _df.Borders[i].Weight;
                _tmpStyle.Border[i].LineStyle := _df.Borders[i].LineStyle;
              end;
              if (_df.Borders[i].UseColor) then
                _tmpStyle.Border[i].Color := _df.Borders[i].Color;
            end; //for

          result := XMLSS.Styles.Add(_tmpStyle, true);
        end; //if
      end; //_getStyleIdxForDF

    begin
      _isOk := false;
      case (_CFCondition) of
        ZCFIsTrueFormula:;
        ZCFCellContentIsBetween, ZCFCellContentIsNotBetween:
          begin
            //только числа
            if (count = 2) then
            begin
              ZETryStrToFloat(_formulas[0], b);
              if (b) then
                ZETryStrToFloat(_formulas[1], _isOk);
            end;
          end;
        ZCFCellContentOperator:
          begin
            //только числа
            if (count = 1) then
              ZETryStrToFloat(_formulas[0], _isOk);
          end;
        ZCFNumberValue:;
        ZCFString:;
        ZCFBoolTrue:;
        ZCFBoolFalse:;
        ZCFFormula:;
        ZCFContainsText: _CheckTextCondition();
        ZCFNotContainsText: _CheckTextCondition();
        ZCFBeginsWithText: _CheckTextCondition();
        ZCFEndsWithText: _CheckTextCondition();
      end; //case

      if (_isOk) then begin
        if (not _isCFAdded) then
          _isCFAdded := _AddCF();

        if ((_isCFAdded) and (Assigned(_CF))) then begin
          num := _CF.Count;
          _CF.Add();
          if (_Style <> '') then
            if (TryStrToInt(_Style, _id)) then
             _CF[num].ApplyStyleID := _getStyleIdxForDF(_id);
          _CF[num].Condition := _CFCondition;
          _CF[num].ConditionOperator := _CFOperator;

          _cf[num].Value1 := _formulas[0];
          if (count >= 2) then
            _cf[num].Value2 := _formulas[1];
        end;
      end;
    end; //_TryApplyCF

  begin
    try
      _sqref := xml.Attributes['sqref'];
      MaxFormulasCount := 2;
      SetLength(_formulas, MaxFormulasCount);
      _isCFAdded := false;
      _CF := nil;
      _tmpStyle := TZStyle.Create();
      while xml.ReadToEndTagByName('conditionalFormatting') do begin
        // cfRule = Conditional Formatting Rule
        if xml.IsTagStartByName('cfRule') then begin
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
          _type     := xml.Attributes['type'];
          _operator := xml.Attributes['operator'];
          _Style    := xml.Attributes['dxfId'];
          _text     := ZEReplaceEntity(xml.Attributes['text']);
          //_priority := xml.Attributes['priority'];

          count := 0;
          while xml.ReadToEndTagByName('cfRule')  do begin
            if xml.IsTagEndByName('formula') then begin
              if (count >= MaxFormulasCount) then begin
                inc(MaxFormulasCount, 2);
                SetLength(_formulas, MaxFormulasCount);
              end;
              _formulas[count] := ZEReplaceEntity(xml.TextBeforeTag);
              inc(count);
            end;
          end; //while

          if (ZEXLSX_getCFCondition(_type, _operator, _CFCondition, _CFOperator)) then
            _TryApplyCF();
        end; //if
      end; //while
    finally
      SetLength(_formulas, 0);
      FreeAndNil(_tmpStyle);
    end;
  end; //_ReadConditionFormatting

  procedure _ReadHeaderFooter();
  begin
    currentSheet.SheetOptions.IsDifferentFirst   := ZEStrToBoolean(xml.Attributes['differentFirst']);
    currentSheet.SheetOptions.IsDifferentOddEven := ZEStrToBoolean(xml.Attributes['differentOddEven']);
    while xml.ReadToEndTagByName('headerFooter') do begin
      if xml.IsTagEndByName('oddHeader') then
        currentSheet.SheetOptions.Header := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('oddFooter') then
        currentSheet.SheetOptions.Footer := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('evenHeader') then
        currentSheet.SheetOptions.EvenHeader := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('evenFooter') then
        currentSheet.SheetOptions.EvenFooter := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('firstHeader') then
        currentSheet.SheetOptions.FirstPageHeader := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('firstFooter') then
        currentSheet.SheetOptions.FirstPageFooter := ClenuapXmlTagValue(xml.TextBeforeTag);
    end;
  end;
begin
  xml := TZsspXMLReaderH.Create();
  result := false;
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    currentPage := XMLSS.Sheets.Count;
    XMLSS.Sheets.Count := XMLSS.Sheets.Count + 1;
    currentRow := 0;
    currentSheet := XMLSS.Sheets[currentPage];
    if assigned(SheetRelations) then
    begin
      currentSheet.Title := SheetRelations^.name;
      if SameText(SheetRelations^.state, 'hidden') then
        currentSheet.Visible := svHidden
      else if SameText(SheetRelations^.state, 'veryhidden') then
        currentSheet.Visible := svVeryHidden
      else
        currentSheet.Visible := svVisible;
    end
    else
      currentSheet.Title := '';

    while xml.ReadTag() do begin
      if xml.IsTagStartByName('sheetData') then
        _ReadSheetData()
      else
      if xml.IsTagClosedByName('autoFilter') then
        _ReadAutoFilter()
      else
      if xml.IsTagStartByName('mergeCells') then
        _ReadMerge()
      else
      if xml.IsTagStartByName('cols') then
        _ReadCols()
      else
      if xml.IsTagClosedByName('drawing') then begin
         currentSheet.DrawingRid := StrtoIntDef(xml.Attributes.ItemsByName['r:id'].Substring(3), 0);
      end else
      if xml.IsTagClosedByName('pageMargins') then begin
        str := xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginBottom := round(tempReal);
        str := xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.FooterMargins.Height := abs(round(tempReal));
        str := xml.Attributes.ItemsByName['header'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.HeaderMargins.Height := abs(round(tempReal));
        str := xml.Attributes.ItemsByName['left'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginLeft := round(tempReal);
        str := xml.Attributes.ItemsByName['right'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginRight := round(tempReal);
        str := xml.Attributes.ItemsByName['top'];
        if (_StrToMM(str, tempReal)) then
          currentSheet.SheetOptions.MarginTop := round(tempReal);
      end else
      //Настройки страницы
      if xml.IsTagClosedByName('pageSetup') then begin
        //str := xml.Attributes.ItemsByName['blackAndWhite'];
        //str := xml.Attributes.ItemsByName['cellComments'];
        //str := xml.Attributes.ItemsByName['copies'];
        //str := xml.Attributes.ItemsByName['draft'];
        //str := xml.Attributes.ItemsByName['errors'];
        str := xml.Attributes.ItemsByName['firstPageNumber'];
        if (str > '') and (TryStrToInt(str, tempInt)) then
          currentSheet.SheetOptions.StartPageNumber := tempInt;

        str := xml.Attributes.ItemsByName['fitToHeight'];
        if (str > '') and (TryStrToInt(str, tempInt)) then
          currentSheet.SheetOptions.FitToHeight := tempInt;

        str := xml.Attributes.ItemsByName['fitToWidth'];
        if (str > '') and (TryStrToInt(str, tempInt)) then
          currentSheet.SheetOptions.FitToWidth := tempInt;

        //str := xml.Attributes.ItemsByName['horizontalDpi'];
        //str := xml.Attributes.ItemsByName['id'];
        str := xml.Attributes.ItemsByName['orientation'];
        if (str > '') then begin
          currentSheet.SheetOptions.PortraitOrientation := false;
          if (str = 'portrait') then
            currentSheet.SheetOptions.PortraitOrientation := true;
        end;

        // overThenDown / downThenOver (might empty)
        currentSheet.SheetOptions.PageOrderDown := xml.AttribString('pageOrder') <> 'overThenDown';

        str := xml.Attributes.ItemsByName['paperSize'];
        if (str > '') and (TryStrToInt(str, tempInt)) then
          currentSheet.SheetOptions.PaperSize := tempInt;
        //str := xml.Attributes.ItemsByName['paperHeight']; //если утановлены paperHeight и Width, то paperSize игнорируется
        //str := xml.Attributes.ItemsByName['paperWidth'];

        str := xml.Attributes.ItemsByName['scale'];
        currentSheet.SheetOptions.ScaleToPercent := StrToIntDef(str, 100);
        //str := xml.Attributes.ItemsByName['useFirstPageNumber'];
        //str := xml.Attributes.ItemsByName['usePrinterDefaults'];
        //str := xml.Attributes.ItemsByName['verticalDpi'];
      end else
      //настройки печати
      if xml.IsTagClosedByName('printOptions') then begin
        //str := xml.Attributes.ItemsByName['gridLines'];
        //str := xml.Attributes.ItemsByName['gridLinesSet'];
        //str := xml.Attributes.ItemsByName['headings'];
        str := xml.Attributes.ItemsByName['horizontalCentered'];
        if (str > '') then
          currentSheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(str);

        str := xml.Attributes.ItemsByName['verticalCentered'];
        if (str > '') then
          currentSheet.SheetOptions.CenterVertical := ZEStrToBoolean(str);
      end
      else
      if xml.IsTagClosedByName('sheetFormatPr') then
      begin
        str := xml.Attributes.ItemsByName['defaultColWidth'];
        if (str > '') then
          currentSheet.DefaultColWidth := ZETryStrToFloat(str, currentSheet.DefaultColWidth);
        str := xml.Attributes.ItemsByName['defaultRowHeight'];
        if (str > '') then
          currentSheet.DefaultRowHeight := ZETryStrToFloat(str, currentSheet.DefaultRowHeight);
      end
      else
      if xml.IsTagClosedByName('dimension') then
        _GetDimension()
      else
      if xml.IsTagStartByName('hyperlinks') then
        _ReadHyperLinks()
      else
      if xml.IsTagStartByName('sheetPr') then
        _ReadSheetPr()
      else
      if xml.IsTagStartByName('rowBreaks')then
        _ReadRowBreaks()
      else
      if xml.IsTagStartByName('colBreaks') then
        _ReadColBreaks()
      else
      if xml.IsTagStartByName('sheetViews') then
        _ReadSheetViews()
      else
      if xml.IsTagStartByName('conditionalFormatting') then
        _ReadConditionFormatting()
      else
      if xml.IsTagStartByName('headerFooter') then
        _ReadHeaderFooter();
    end; //while

    result := true;
  finally
    xml.Free();
  end;
end; //ZEXSLXReadSheet

//Прочитать стили из потока (styles.xml)
//INPUT
//  var XMLSS: TZWorkBook                    - хранилище
//  var Stream: TStream                    - поток
//  var ThemaFillsColors: TIntegerDynArray - цвета из темы
//  var ThemaColorCount: integer           - кол-во цветов заливки в теме
//  var MaximumDigitWidth: double          - ширина самого широкого числа в пикселях
//      ReadHelper: TZEXLSXReadHelper      -
//RETURN
//      boolean - true - стили прочитались без ошибок
function ZEXSLXReadStyles(var XMLSS: TZWorkBook; var Stream: TStream;
    var ThemaFillsColors: TIntegerDynArray; var ThemaColorCount: integer;
    var MaximumDigitWidth: double; ReadHelper: TZEXLSXReadHelper): boolean;
type
  TZXLSXBorderItem = record
    color: TColor;
    isColor: boolean;
    isEnabled: boolean;
    style: TZBorderType;
    Weight: byte;
  end;

  //   0 - left           левая граница
  //   1 - Top            верхняя граница
  //   2 - Right          правая граница
  //   3 - Bottom         нижняя граница
  //   4 - DiagonalLeft   диагональ от верхнего левого угла до нижнего правого
  //   5 - DiagonalRight  диагональ от нижнего левого угла до правого верхнего
  TZXLSXBorder = array[0..5] of TZXLSXBorderItem;
  TZXLSXBordersArray = array of TZXLSXBorder;

  TZXLSXCellAlignment = record
    horizontal: TZHorizontalAlignment;
    indent: integer;
    shrinkToFit: boolean;
    textRotation: integer;
    vertical: TZVerticalAlignment;
    wrapText: boolean;
  end;

  TZXLSXCellStyle = record
    applyAlignment: boolean;
    applyBorder: boolean;
    applyFont: boolean;
    applyProtection: boolean;
    borderId: integer;
    fillId: integer;
    fontId: integer;
    numFmtId: integer;
    xfId: integer;
    hidden: boolean;
    locked: boolean;
    alignment: TZXLSXCellAlignment;
  end;

  TZXLSXCellStylesArray = array of TZXLSXCellStyle;

  type TZXLSXStyle = record
    builtinId: integer;     //??
    customBuiltin: boolean; //??
    name: string;           //??
    xfId: integer;
  end;

  TZXLSXStyleArray = array of TZXLSXStyle;

  TZXLSXFill = record
    patternfill: TZCellPattern;
    bgColorType: byte;  //0 - rgb, 1 - indexed, 2 - theme
    bgcolor: TColor;
    patterncolor: TColor;
    patternColorType: byte;
    lumFactorBG: double;
    lumFactorPattern: double;
  end;

  TZXLSXFillArray = array of TZXLSXFill;

  TZXLSXDFFont = record
    Color: TColor;
    ColorType: byte;
    LumFactor: double;
  end;

  TZXLSXDFFontArray = array of TZXLSXDFFont;

var
  xml: TZsspXMLReaderH;
  s: string;
  FontArray: TZEXLSXFontArray;
  FontCount: integer;
  BorderArray: TZXLSXBordersArray;
  BorderCount: integer;
  CellXfsArray: TZXLSXCellStylesArray;
  CellXfsCount: integer;
  CellStyleArray: TZXLSXCellStylesArray;
  CellStyleCount: integer;
  StyleArray: TZXLSXStyleArray;
  StyleCount: integer;
  FillArray: TZXLSXFillArray;
  FillCount: integer;
  indexedColor: TIntegerDynArray;
  indexedColorCount: integer;
  indexedColorMax: integer;
  _Style: TZStyle;
  t, i, n: integer;
  h1, s1, l1: double;
  _dfFonts: TZXLSXDFFontArray;
  _dfFills: TZXLSXFillArray;

  //Приводит к шрифту по-умолчанию
  //INPUT
  //  var fnt: TZEXLSXFont - шрифт
  procedure ZEXLSXZeroFont(var fnt: TZEXLSXFont);
  begin
    fnt.name := 'Arial';
    fnt.bold := false;
    fnt.italic := false;
    fnt.underline := false;
    fnt.strike := false;
    fnt.charset := 204;
    fnt.color := clBlack;
    fnt.LumFactor := 0;
    fnt.ColorType := 0;
    fnt.fontsize := 8;
    fnt.superscript := false;
    fnt.subscript := false;
  end; //ZEXLSXZeroFont

  //Обнуляет границы
  //  var border: TZXLSXBorder - границы
  procedure ZEXLSXZeroBorder(var border: TZXLSXBorder);
  var i: integer;
  begin
    for i := 0 to 5 do begin
      border[i].isColor := false;
      border[i].isEnabled := false;
      border[i].style := ZENone;
      border[i].Weight := 0;
    end;
  end; //ZEXLSXZeroBorder

  //Меняёт местами bgColor и fgColor при несплошных заливках
  //INPUT
  //  var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXSwapPatternFillColors(var PattFill: TZXLSXFill);
  var t: integer; _b: byte;
  begin
    //если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
    if (not (PattFill.patternfill in [ZPNone, ZPSolid])) then begin
      t := PattFill.patterncolor;
      PattFill.patterncolor := PattFill.bgcolor;
      PattFill.bgColor := t;
      l1 := PattFill.lumFactorPattern;
      PattFill.lumFactorPattern := PattFill.lumFactorBG;
      PattFill.lumFactorBG := l1;

      _b := PattFill.patternColorType;
      PattFill.patternColorType := PattFill.bgColorType;
      PattFill.bgColorType := _b;
    end; //if
  end; //ZEXLSXSwapPatternFillColors

  //Очистить заливку ячейки
  //INPUT
  //  var PattFill: TZXLSXFill - заливка
  procedure ZEXLSXClearPatternFill(var PattFill: TZXLSXFill);
  begin
    PattFill.patternfill := ZPNone;
    PattFill.bgcolor := clWindow;
    PattFill.patterncolor := clWindow;
    PattFill.bgColorType := 0;
    PattFill.patternColorType := 0;
    PattFill.lumFactorBG := 0.0;
    PattFill.lumFactorPattern := 0.0;
  end; //ZEXLSXClearPatternFill

  //Обнуляет стиль
  //INPUT
  //  var style: TZXLSXCellStyle - стиль XLSX
  procedure ZEXLSXZeroCellStyle(var style: TZXLSXCellStyle);
  begin
    style.applyAlignment := false;
    style.applyBorder := false;
    style.applyProtection := false;
    style.hidden := false;
    style.locked := false;
    style.borderId := -1;
    style.fontId := -1;
    style.fillId := -1;
    style.numFmtId := -1;
    style.xfId := -1;
    style.alignment.horizontal := ZHAutomatic;
    style.alignment.vertical := ZVAutomatic;
    style.alignment.shrinkToFit := false;
    style.alignment.wrapText := false;
    style.alignment.textRotation := 0;
    style.alignment.indent := 0;
  end; //ZEXLSXZeroCellStyle

  //TZEXLSXFont в TFont
  //INPUT
  //  var fnt: TZEXLSXFont  - XLSX шрифт
  //  var font: TFont       - стандартный шрифт
  {procedure ZEXLSXFontToFont(var fnt: TZEXLSXFont; var font: TFont);
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
  end;} //ZEXLSXFontToFont

  //Прочитать цвет
  //INPUT
  //  var retColor: TColor      - возвращаемый цвет
  //  var retColorType: byte    - тип цвета: 0 - rgb, 1 - indexed, 2 - theme
  //  var retLumfactor: double  - яркость
  procedure ZXLSXGetColor(var retColor: TColor; var retColorType: byte; var retLumfactor: double);
  var t: integer;
  begin
    //I hate this f****** format! m$ office >= 2007 is big piece of shit! Arrgh!
    s := xml.Attributes.ItemsByName['rgb'];
    if (length(s) > 2) then
    begin
      delete(s, 1, 2);
      if (s > '') then
        retColor := HTMLHexToColor(s);
    end;
    s := xml.Attributes.ItemsByName['theme'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 2;
        retColor := t;
      end;
    s := xml.Attributes.ItemsByName['indexed'];
    if (s > '') then
      if (TryStrToInt(s, t)) then
      begin
        retColorType := 1;
        retColor := t;
      end;
    s := xml.Attributes.ItemsByName['tint'];
    if (s <> '') then
      retLumfactor := ZETryStrToFloat(s, 0);
  end; //ZXLSXGetColor

  procedure _ReadFonts();
  var _currFont: integer; sz: double;
  begin
    _currFont := -1;
    while xml.ReadToEndTagByName('fonts') do begin
      s := xml.Attributes.ItemsByName['val'];
      if xml.IsTagStartByName('font') then begin
        _currFont := FontCount;
        inc(FontCount);
        SetLength(FontArray, FontCount);
        ZEXLSXZeroFont(FontArray[_currFont]);
      end else if (_currFont >= 0) then begin
        if xml.IsTagClosedByName('name') then
          FontArray[_currFont].name := s
        else if xml.IsTagClosedByName('b') then
          FontArray[_currFont].bold := true
        else if xml.IsTagClosedByName('charset') then begin
          if (TryStrToInt(s, t)) then
            FontArray[_currFont].charset := t;
        end else if xml.IsTagClosedByName('color') then begin
          ZXLSXGetColor(FontArray[_currFont].color,
                        FontArray[_currFont].ColorType,
                        FontArray[_currFont].LumFactor);
        end else if xml.IsTagClosedByName('i') then
          FontArray[_currFont].italic := true
        else if xml.IsTagClosedByName('strike') then
          FontArray[_currFont].strike := true
        else
        if xml.IsTagClosedByName('sz') then begin
          if (TryStrToFloat(s, sz, TFormatSettings.Invariant)) then
            FontArray[_currFont].fontsize := sz;
        end else if xml.IsTagClosedByName('u') then begin
          FontArray[_currFont].underline := true;
        end else if xml.IsTagClosedByName('vertAlign') then begin
          FontArray[_currFont].superscript := s = 'superscript';
          FontArray[_currFont].subscript := s = 'subscript';
        end;
      end; //if
      //Тэги настройки шрифта
      //*b - bold
      //*charset
      //*color
      //?condense
      //?extend
      //?family
      //*i - italic
      //*name
      //?outline
      //?scheme
      //?shadow
      //*strike
      //*sz - size
      //*u - underline
      //*vertAlign

    end; //while
  end; //_ReadFonts

  //Получить тип заливки
  function _GetPatternFillByStr(const s: string): TZCellPattern;
  begin
    if (s = 'solid') then
      result := ZPSolid
    else if (s = 'none') then
      result := ZPNone
    else if (s = 'gray125') then
      result := ZPGray125
    else if (s = 'gray0625') then
      result := ZPGray0625
    else if (s = 'darkUp') then
      result := ZPDiagStripe
    else if (s = 'mediumGray') then
      result := ZPGray50
    else if (s = 'darkGray') then
      result := ZPGray75
    else if (s = 'lightGray') then
      result := ZPGray25
    else if (s = 'darkHorizontal') then
      result := ZPHorzStripe
    else if (s = 'darkVertical') then
      result := ZPVertStripe
    else if (s = 'darkDown') then
      result := ZPReverseDiagStripe
    else if (s = 'darkUpDark') then
      result := ZPDiagStripe
    else if (s = 'darkGrid') then
      result := ZPDiagCross
    else if (s = 'darkTrellis') then
      result := ZPThickDiagCross
    else if (s = 'lightHorizontal') then
      result := ZPThinHorzStripe
    else if (s = 'lightVertical') then
      result := ZPThinVertStripe
    else if (s = 'lightDown') then
      result := ZPThinReverseDiagStripe
    else if (s = 'lightUp') then
      result := ZPThinDiagStripe
    else if (s = 'lightGrid') then
      result := ZPThinHorzCross
    else if (s = 'lightTrellis') then
      result := ZPThinDiagCross
    else
      result := ZPSolid; //{tut} потом подумать насчёт стилей границ
  end; //_GetPatternFillByStr

  //Определить стиль начертания границы
  //INPUT
  //  const st: string            - название стиля
  //  var retWidth: byte          - возвращаемая ширина линии
  //  var retStyle: TZBorderType  - возвращаемый стиль начертания линии
  //RETURN
  //      boolean - true - стиль определён
  function XLSXGetBorderStyle(const st: string; var retWidth: byte; var retStyle: TZBorderType): boolean;
  begin
    result := true;
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
    else if (st = 'medium') then  begin
      retStyle := ZEContinuous;
      retWidth := 2;
    end else if (st = 'thick') then begin
      retStyle := ZEContinuous;
      retWidth := 3;
    end else if (st = 'mediumDashed') then begin
      retStyle := ZEDash;
      retWidth := 2;
    end else if (st = 'mediumDashDot') then begin
      retStyle := ZEDashDot;
      retWidth := 2;
    end else if (st = 'mediumDashDotDot') then begin
      retStyle := ZEDashDotDot;
      retWidth := 2;
    end else if (st = 'none') then
      retStyle := ZENone
    else
      result := false;
  end; //XLSXGetBorderStyle

  procedure _ReadBorders();
  var
    _diagDown, _diagUP: boolean;
    _currBorder: integer; //текущий набор границ
    _currBorderItem: integer; //текущая граница (левая/правая ...)
    _color: TColor;
    _isColor: boolean;

    procedure _SetCurBorder(borderNum: integer);
    begin
      _currBorderItem := borderNum;
      s := xml.Attributes.ItemsByName['style'];
      if (s > '') then begin
        BorderArray[_currBorder][borderNum].isEnabled :=
          XLSXGetBorderStyle(s,
                             BorderArray[_currBorder][borderNum].Weight,
                             BorderArray[_currBorder][borderNum].style);
      end;
    end; //_SetCurBorder

  begin
    _currBorderItem := -1;
    _diagDown := false;
    _diagUP := false;
    _color := clBlack;
    while xml.ReadToEndTagByName('borders') do begin
      if xml.IsTagStartByName('border') then begin
        _currBorder := BorderCount;
        inc(BorderCount);
        SetLength(BorderArray, BorderCount);
        ZEXLSXZeroBorder(BorderArray[_currBorder]);
        _diagDown := false;
        _diagUP := false;
        s := xml.Attributes.ItemsByName['diagonalDown'];
        if (s > '') then
          _diagDown := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['diagonalUp'];
        if (s > '') then
          _diagUP := ZEStrToBoolean(s);
      end else begin
        if (xml.IsTagStartOrClosed) then begin
          if (xml.TagName = 'left') then begin
            _SetCurBorder(0);
          end else if (xml.TagName = 'right') then begin
            _SetCurBorder(2);
          end else if (xml.TagName = 'top') then begin
            _SetCurBorder(1);
          end else if (xml.TagName = 'bottom') then begin
            _SetCurBorder(3);
          end else if (xml.TagName = 'diagonal') then begin
            if (_diagUp) then
              _SetCurBorder(5);
            if (_diagDown) then begin
              if (_diagUp) then
                BorderArray[_currBorder][4] := BorderArray[_currBorder][5]
              else
                _SetCurBorder(4);
            end;
          end else if (xml.TagName = 'end') then begin
          end else if (xml.TagName = 'start') then begin
          end else if (xml.TagName = 'color') then begin
            _isColor := false;
            s := xml.Attributes.ItemsByName['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if (s > '') then begin
              _color := HTMLHexToColor(s);
              _isColor := true;
            end;
            if (_isColor and (_currBorderItem >= 0) and (_currBorderItem < 6)) then begin
              BorderArray[_currBorder][_currBorderItem].color := _color;
              BorderArray[_currBorder][_currBorderItem].isColor := true;
            end;
          end;
        end; //if
      end; //else
    end; //while
  end; //_ReadBorders

  procedure _ReadFills();
  var _currFill: integer;
  begin
    _currFill := -1;
    while xml.ReadToEndTagByName('fills') do begin
      if xml.IsTagStartByName('fill') then begin
        _currFill := FillCount;
        inc(FillCount);
        SetLength(FillArray, FillCount);
        ZEXLSXClearPatternFill(FillArray[_currFill]);
      end else if ((xml.TagName = 'patternFill') and (xml.IsTagStartOrClosed)) then begin
        if (_currFill >= 0) then begin
          s := xml.Attributes.ItemsByName['patternType'];
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
      end else
      if xml.IsTagClosedByName('bgColor') then begin
        if (_currFill >= 0) then  begin
          ZXLSXGetColor(FillArray[_currFill].patterncolor,
                        FillArray[_currFill].patternColorType,
                        FillArray[_currFill].lumFactorPattern);

          //если не сплошная заливка - нужно поменять местами цвета (bgColor <-> fgColor)
          ZEXLSXSwapPatternFillColors(FillArray[_currFill]);
        end;
      end else if xml.IsTagClosedByName('fgColor') then begin
        if (_currFill >= 0) then
          ZXLSXGetColor(FillArray[_currFill].bgcolor,
                        FillArray[_currFill].bgColorType,
                        FillArray[_currFill].lumFactorBG);
      end; //fgColor
    end; //while
  end; //_ReadFills

  //Читает стили (cellXfs и cellStyleXfs)
  //INPUT
  //  const TagName: string           - имя тэга
  //  var CSA: TZXLSXCellStylesArray  - массив со стилями
  //  var StyleCount: integer         - кол-во стилей
  procedure _ReadCellCommonStyles(const TagName: string; var CSA: TZXLSXCellStylesArray; var StyleCount: integer);
  var _currCell: integer; b: boolean;
  begin
    _currCell := -1;
    while xml.ReadToEndTagByName(TagName)  do begin
      b := false;
      if ((xml.TagName = 'xf') and (xml.IsTagStartOrClosed)) then begin
        _currCell := StyleCount;
        inc(StyleCount);
        SetLength(CSA, StyleCount);
        ZEXLSXZeroCellStyle(CSA[_currCell]);
        s := xml.Attributes.ItemsByName['applyAlignment'];
        if (s > '') then
          CSA[_currCell].applyAlignment := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyBorder'];
        if (s > '') then
          CSA[_currCell].applyBorder := ZEStrToBoolean(s)
        else
          b := true;

        s := xml.Attributes.ItemsByName['applyFont'];
        if (s > '') then
          CSA[_currCell].applyFont := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['applyProtection'];
        if (s > '') then
          CSA[_currCell].applyProtection := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['borderId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then begin
            CSA[_currCell].borderId := t;
            if (b and (t >= 1)) then
              CSA[_currCell].applyBorder := true;
          end;

        s := xml.Attributes.ItemsByName['fillId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fillId := t;

        s := xml.Attributes.ItemsByName['fontId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].fontId := t;

        s := xml.Attributes.ItemsByName['numFmtId'];
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

        s := xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
            CSA[_currCell].xfId := t;
      end else
      if xml.IsTagClosedByName('alignment') then begin
        if (_currCell >= 0) then begin
          s := xml.Attributes.ItemsByName['horizontal'];
          if (s > '') then begin
            if (s = 'general') then
              CSA[_currCell].alignment.horizontal := ZHAutomatic
            else
            if (s = 'left') then
              CSA[_currCell].alignment.horizontal := ZHLeft
            else
            if (s = 'right') then
              CSA[_currCell].alignment.horizontal := ZHRight
            else
            if ((s = 'center') or (s = 'centerContinuous')) then
              CSA[_currCell].alignment.horizontal := ZHCenter
            else
            if (s = 'fill') then
              CSA[_currCell].alignment.horizontal := ZHFill
            else
            if (s = 'justify') then
              CSA[_currCell].alignment.horizontal := ZHJustify
            else
            if (s = 'distributed') then
              CSA[_currCell].alignment.horizontal := ZHDistributed;
          end;

          s := xml.Attributes.ItemsByName['indent'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.indent := t;

          s := xml.Attributes.ItemsByName['shrinkToFit'];
          if (s > '') then
            CSA[_currCell].alignment.shrinkToFit := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['textRotation'];
          if (s > '') then
            if (TryStrToInt(s, t)) then
              CSA[_currCell].alignment.textRotation := t;

          s := xml.Attributes.ItemsByName['vertical'];
          if (s > '') then begin
            if (s = 'center') then
              CSA[_currCell].alignment.vertical := ZVCenter
            else
            if (s = 'top') then
              CSA[_currCell].alignment.vertical := ZVTop
            else
            if (s = 'bottom') then
              CSA[_currCell].alignment.vertical := ZVBottom
            else
            if (s = 'justify') then
              CSA[_currCell].alignment.vertical := ZVJustify
            else
            if (s = 'distributed') then
              CSA[_currCell].alignment.vertical := ZVDistributed;
          end;

          s := xml.Attributes.ItemsByName['wrapText'];
          if (s > '') then
            CSA[_currCell].alignment.wrapText := ZEStrToBoolean(s);
        end; //if
      end else if xml.IsTagClosedByName('protection') then begin
        if (_currCell >= 0) then begin
          s := xml.Attributes.ItemsByName['hidden'];
          if (s > '') then
            CSA[_currCell].hidden := ZEStrToBoolean(s);

          s := xml.Attributes.ItemsByName['locked'];
          if (s > '') then
            CSA[_currCell].locked := ZEStrToBoolean(s);
        end;
      end;
    end; //while
  end; //_ReadCellCommonStyles

  //Сами стили ?? (или для чего они вообще?)
  procedure _ReadCellStyles();
  var b: boolean;
  begin
    while xml.ReadToEndTagByName('cellStyles') do begin
      if xml.IsTagClosedByName('cellStyle') then begin
        b := false;
        SetLength(StyleArray, StyleCount + 1);
        s := xml.Attributes.ItemsByName['builtinId']; //?
        if (s > '') then
          if (TryStrToInt(s, t)) then
            StyleArray[StyleCount].builtinId := t;

        s := xml.Attributes.ItemsByName['customBuiltin']; //?
        if (s > '') then
          StyleArray[StyleCount].customBuiltin := ZEStrToBoolean(s);

        s := xml.Attributes.ItemsByName['name']; //?
          StyleArray[StyleCount].name := s;

        s := xml.Attributes.ItemsByName['xfId'];
        if (s > '') then
          if (TryStrToInt(s, t)) then
          begin
            StyleArray[StyleCount].xfId := t;
            b := true;
          end;

        if (b) then
          inc(StyleCount);
      end;
    end; //while
  end; //_ReadCellStyles

  procedure _ReadColors();
  begin
    while xml.ReadToEndTagByName('colors') do begin
      if xml.IsTagClosedByName('rgbColor') then begin
        s := xml.Attributes.ItemsByName['rgb'];
        if (length(s) > 2) then
          delete(s, 1, 2);
        if (s > '') then begin
          inc(indexedColorCount);
          if (indexedColorCount >= indexedColorMax) then begin
            indexedColorMax := indexedColorCount + 80;
            SetLength(indexedColor, indexedColorMax);
          end;
          indexedColor[indexedColorCount - 1] := HTMLHexToColor(s);
        end;
      end;
    end; //while
  end; //_ReadColors

  //Конвертирует RGB в HSL
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      r: byte     -
  //      g: byte     -
  //      b: byte     -
  //  out h: double   - Hue - тон цвета
  //  out s: double   - Saturation - насыщенность
  //  out l: double   - Lightness (Intensity) - светлота (яркость)
  procedure ZRGBToHSL(r, g, b: byte; out h, s, l: double);
  var
    _max, _min: double;
    intMax, intMin: integer;
    _r, _g, _b: double;
    _delta: double;
    _idx: integer;
  begin
    _r := r / 255;
    _g := g / 255;
    _b := b / 255;

    intMax := Max(r, Max(g, b));
    intMin := Min(r, Min(g, b));

    _max := Max(_r, Max(_g, _b));
    _min := Min(_r, Min(_g, _b));

    h := (_max + _min) * 0.5;
    s := h;
    l := h;
    if (intMax = intMin) then begin
      h := 0;
      s := 0;
    end else begin
      _delta := _max - _min;
      if (l > 0.5) then
        s := _delta / (2 - _max - _min)
      else
        s := _delta / (_max + _min);

        if (intMax = r) then
          _idx := 1
        else
        if (intMax = g) then
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
          2: h := (_b - _r) / _delta + 2;
          3: h := (_r - _g) / _delta + 4;
        end;

        h := h / 6;
    end;
  end; //ZRGBToHSL

  //Конвертирует TColor (RGB) в HSL
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      Color: TColor - цвет
  //  out h: double     - Hue - тон цвета
  //  out s: double     - Saturation - насыщенность
  //  out l: double     - Lightness (Intensity) - светлота (яркость)
  procedure ZColorToHSL(Color: TColor; out h, s, l: double);
  var _RGB: integer;
  begin
    _RGB := ColorToRGB(Color);
    ZRGBToHSL(byte(_RGB), byte(_RGB shr 8), byte(_RGB shr 16), h, s, l);
  end; //ZColorToHSL

  //Конвертирует HSL в RGB
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      h: double - Hue - тон цвета
  //      s: double - Saturation - насыщенность
  //      l: double - Lightness (Intensity) - светлота (яркость)
  //  out r: byte   -
  //  out g: byte   -
  //  out b: byte   -
  procedure ZHSLToRGB(h, s, l: double; out r, g, b: byte);
  var _r, _g, _b, q, p: double;
    function HueToRgb(p, q, t: double): double;
    begin
      result := p;
      if (t < 0) then
        t := t + 1;
      if (t > 1) then
        t := t - 1;
      if (t < 1/6) then
        result := p + (q - p) * 6 * t
      else
      if (t < 0.5) then
        result := q
      else
      if (t < 2/3) then
        result := p + (q - p) * (2/3 - t) * 6;
    end; //HueToRgb

  begin
    if (s = 0) then begin
      //Оттенок серого
      _r := l;
      _g := l;
      _b := l;
    end else begin
      if (l < 0.5) then
        q := l * (1 + s)
      else
        q := l + s - l * s;
      p := 2 * l - q;
      _r := HueToRgb(p, q, h + 1/3);
      _g := HueToRgb(p, q, h);
      _b := HueToRgb(p, q, h - 1/3);
    end;
    r := byte(round(_r * 255));
    g := byte(round(_g * 255));
    b := byte(round(_b * 255));
  end; //ZHSLToRGB

  //Конвертирует HSL в Color
  //http://en.wikipedia.org/wiki/HSL_color_space
  //INPUT
  //      h: double - Hue - тон цвета
  //      s: double - Saturation - насыщенность
  //      l: double - Lightness (Intensity) - светлота (яркость)
  //RETURN
  //      TColor - цвет
  function ZHSLToColor(h, s, l: double): TColor;
  var r, g, b: byte;
  begin
    ZHSLToRGB(h, s, l, r, g, b);
    result := (b shl 16) or (g shl 8) or r;
  end; //ZHSLToColor

  //Применить tint к цвету
  // Thanks Tomasz Wieckowski!
  //   http://msdn.microsoft.com/en-us/library/ff532470%28v=office.12%29.aspx
  procedure ApplyLumFactor(var Color: TColor; var lumFactor: double);
  begin
    //+delta?
    if (lumFactor <> 0.0) then begin
      ZColorToHSL(Color, h1, s1, l1);
      lumFactor := 1 - lumFactor;

      if (l1 = 1) then
        l1 := l1 * (1 - lumFactor)
      else
        l1 := l1 * lumFactor + (1 - lumFactor);

      Color := ZHSLtoColor(h1, s1, l1);
    end;
  end; //ApplyLumFactor

  //Differential Formatting для xlsx
  procedure _Readdxfs();
  var
    _df: TZXLSXDiffFormattingItem;
    _dfIndex: integer;

    procedure _addFontStyle(fnts: TFontStyle);
    begin
      _df.FontStyles := _df.FontStyles + [fnts];
      _df.UseFontStyles := true;
    end;

    procedure _ReadDFFont();
    begin
      _df.UseFont := true;
      while xml.ReadToEndTagByName('font') do begin
        if (xml.TagName = 'i') then
          _addFontStyle(fsItalic);
        if (xml.TagName = 'b') then
          _addFontStyle(fsBold);
        if (xml.TagName = 'u') then
          _addFontStyle(fsUnderline);
        if (xml.TagName = 'strike') then
          _addFontStyle(fsStrikeOut);

        if (xml.TagName = 'color') then begin
          _df.UseFontColor := true;
          ZXLSXGetColor(_dfFonts[_dfIndex].Color,
                        _dfFonts[_dfIndex].ColorType,
                        _dfFonts[_dfIndex].LumFactor);
        end;
      end; //while
    end; //_ReadDFFont

    procedure _ReadDFFill();
    begin
      _df.UseFill := true;
      while not xml.IsTagEndByName('fill') do begin
        xml.ReadTag();
        if (xml.Eof) then
          break;

        if (xml.IsTagStartOrClosed) then begin
          if (xml.TagName = 'patternFill') then begin
            s := xml.Attributes.ItemsByName['patternType'];
            if (s <> '') then begin
              _df.UseCellPattern := true;
              _df.CellPattern := _GetPatternFillByStr(s);
            end;
          end else
          if (xml.TagName = 'bgColor') then begin
            _df.UseBGColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].bgcolor,
                          _dfFills[_dfIndex].bgColorType,
                          _dfFills[_dfIndex].lumFactorBG)
          end else
          if (xml.TagName = 'fgColor') then begin
            _df.UsePatternColor := true;
            ZXLSXGetColor(_dfFills[_dfIndex].patterncolor,
                          _dfFills[_dfIndex].patternColorType,
                          _dfFills[_dfIndex].lumFactorPattern);
            ZEXLSXSwapPatternFillColors(_dfFills[_dfIndex]);
          end;
        end;
      end; //while
    end; //_ReadDFFill

    procedure _ReadDFBorder();
    var _borderNum: TZBordersPos;
      t: byte;
      _bt: TZBorderType;
      procedure _SetDFBorder(BorderNum: TZBordersPos);
      begin
        _borderNum := BorderNum;
        s := xml.Attributes['style'];
        if (s <> '') then
          if (XLSXGetBorderStyle(s, t, _bt)) then begin
            _df.UseBorder := true;
            _df.Borders[BorderNum].Weight := t;
            _df.Borders[BorderNum].LineStyle := _bt;
            _df.Borders[BorderNum].UseStyle := true;
          end;
      end; //_SetDFBorder

    begin
      _df.UseBorder := true;
      _borderNum := bpLeft;
      while xml.ReadToEndTagByName('border') do begin
        if xml.IsTagStartOrClosed then begin
          if (xml.TagName = 'left') then
            _SetDFBorder(bpLeft)
          else
          if (xml.TagName = 'right') then
            _SetDFBorder(bpRight)
          else
          if (xml.TagName = 'top') then
            _SetDFBorder(bpTop)
          else
          if (xml.TagName = 'bottom') then
            _SetDFBorder(bpBottom)
          else
          if (xml.TagName = 'vertical') then
            _SetDFBorder(bpDiagonalLeft)
          else
          if (xml.TagName = 'horizontal') then
            _SetDFBorder(bpDiagonalRight)
          else
          if (xml.TagName = 'color') then
          begin
            s := xml.Attributes['rgb'];
            if (length(s) > 2) then
              delete(s, 1, 2);
            if ((_borderNum >= bpLeft) and (_borderNum <= bpDiagonalRight)) then
              if (s <> '') then begin
                _df.UseBorder := true;
                _df.Borders[_borderNum].UseColor := true;
                _df.Borders[_borderNum].Color := HTMLHexToColor(s);
              end;
          end;
        end; //if
      end; //while
    end; //_ReadDFBorder

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
      while xml.ReadToEndTagByName('dxf') do begin
        if xml.IsTagStartByName('font') then
          _ReadDFFont()
        else
        if xml.IsTagStartByName('fill') then
          _ReadDFFill()
        else
        if xml.IsTagStartByName('border') then
          _ReadDFBorder();
      end; //while
    end; //_ReaddxfItem

  begin
    while xml.ReadToEndTagByName('dxfs') do begin
      if xml.IsTagStartByName('dxf') then
        _ReaddxfItem();
    end; //while
  end; //_Readdxfs

  procedure XLSXApplyColor(var AColor: TColor; ColorType: byte; LumFactor: double);
  begin
    //Thema color
    if (ColorType = 2) then begin
      t := AColor - 1;
      if ((t >= 0) and (t < ThemaColorCount)) then
        AColor := ThemaFillsColors[t];
    end;
    if (ColorType = 1) then
      if ((AColor >= 0) and (AColor < indexedColorCount))  then
        AColor := indexedColor[AColor];
    ApplyLumFactor(AColor, LumFactor);
  end; //XLSXApplyColor

  //Применить стиль
  //INPUT
  //  var XMLSSStyle: TZStyle         - стиль в хранилище
  //  var XLSXStyle: TZXLSXCellStyle  - стиль в xlsx
  procedure _ApplyStyle(var XMLSSStyle: TZStyle; var XLSXStyle: TZXLSXCellStyle);
  var i: integer; b: TZBordersPos;
  begin
    if (XLSXStyle.numFmtId >= 0) then
      XMLSSStyle.NumberFormat := ReadHelper.NumberFormats.GetFormat(XLSXStyle.numFmtId);
    XMLSSStyle.NumberFormatId := XLSXStyle.numFmtId;

    if (XLSXStyle.applyAlignment) then begin
      XMLSSStyle.Alignment.Horizontal  := XLSXStyle.alignment.horizontal;
      XMLSSStyle.Alignment.Vertical    := XLSXStyle.alignment.vertical;
      XMLSSStyle.Alignment.Indent      := XLSXStyle.alignment.indent;
      XMLSSStyle.Alignment.ShrinkToFit := XLSXStyle.alignment.shrinkToFit;
      XMLSSStyle.Alignment.WrapText    := XLSXStyle.alignment.wrapText;

      XMLSSStyle.Alignment.Rotate := 0;
      i := XLSXStyle.alignment.textRotation;
      XMLSSStyle.Alignment.VerticalText := (i = 255);
      if (i >= 0) and (i <= 180) then begin
        if i > 90 then i := 90 - i;
        XMLSSStyle.Alignment.Rotate := i
      end;
    end;

    if XLSXStyle.applyBorder then begin
      n := XLSXStyle.borderId;
      if (n >= 0) and (n < BorderCount) then
        for b := bpLeft to bpDiagonalRight do begin
          if (BorderArray[n][Ord(b)].isEnabled) then begin
            XMLSSStyle.Border[b].LineStyle := BorderArray[n][Ord(b)].style;
            XMLSSStyle.Border[b].Weight := BorderArray[n][Ord(b)].Weight;
            if (BorderArray[n][Ord(b)].isColor) then
              XMLSSStyle.Border[b].Color := BorderArray[n][Ord(b)].color;
          end;
        end;
    end;

    if (XLSXStyle.applyFont) then begin
      n := XLSXStyle.fontId;
      if ((n >= 0) and (n < FontCount)) then begin
        XLSXApplyColor(FontArray[n].color,
                       FontArray[n].ColorType,
                       FontArray[n].LumFactor);
        XMLSSStyle.Font.Name := FontArray[n].name;
        XMLSSStyle.Font.Size := FontArray[n].fontsize;
        XMLSSStyle.Font.Charset := FontArray[n].charset;
        XMLSSStyle.Font.Color := FontArray[n].color;
        if (FontArray[n].bold) then
          XMLSSStyle.Font.Style := [fsBold];
        if (FontArray[n].underline) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsUnderline];
        if (FontArray[n].italic) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsItalic];
        if (FontArray[n].strike) then
          XMLSSStyle.Font.Style := XMLSSStyle.Font.Style + [fsStrikeOut];
        XMLSSStyle.Superscript := FontArray[n].superscript;
        XMLSSStyle.Subscript := FontArray[n].subscript;
      end;
    end;

    if (XLSXStyle.applyProtection) then begin
      XMLSSStyle.Protect := XLSXStyle.locked;
      XMLSSStyle.HideFormula := XLSXStyle.hidden;
    end;

    n := XLSXStyle.fillId;
    if ((n >= 0) and (n < FillCount)) then begin
      XMLSSStyle.CellPattern := FillArray[n].patternfill;
      XMLSSStyle.BGColor := FillArray[n].bgcolor;
      XMLSSStyle.PatternColor := FillArray[n].patterncolor;
    end;
  end; //_ApplyStyle

  procedure _CheckIndexedColors();
  const
    _standart: array [0..63] of string = (
      '#000000', // 0
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
      '#333333'  // 63
    );
  var i: integer;
  begin
    if (indexedColorCount = 0) then begin
      indexedColorCount := 63;
      indexedColorMax := indexedColorCount + 10;
      SetLength(indexedColor, indexedColorMax);
      for i := 0 to 63 do
        indexedColor[i] := HTMLHexToColor(_standart[i]);
    end;
  end; //_CheckIndexedColors

begin
  result := false;
  MaximumDigitWidth := 0;
  xml := nil;
  CellXfsArray := nil;
  CellStyleArray := nil;
  try
    xml := TZsspXMLReaderH.Create();
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    FontCount := 0;
    BorderCount := 0;
    CellStyleCount := 0;
    StyleCount := 0;
    CellXfsCount := 0;
    FillCount := 0;
    indexedColorCount := 0;
    indexedColorMax := -1;

    while not xml.Eof() do begin
      xml.ReadTag();

      if xml.IsTagStartByName('fonts') then
      begin
        _ReadFonts();
        if Length(FontArray) > 0 then
          MaximumDigitWidth := GetMaximumDigitWidth(FontArray[0].Name, FontArray[0].fontsize);
      end
      else
      if xml.IsTagStartByName('borders') then
        _ReadBorders()
      else
      if xml.IsTagStartByName('fills') then
        _ReadFills()
      else
      {
        А.А.Валуев:
        Элементы внутри cellXfs ссылаются на элементы внутри cellStyleXfs.
        Элементы внутри cellStyleXfs ни на что не ссылаются.
      }
      if xml.IsTagStartByName('cellStyleXfs') then
        _ReadCellCommonStyles('cellStyleXfs', CellStyleArray, CellStyleCount)//_ReadCellStyleXfs()
      else
      if xml.IsTagStartByName('cellXfs') then  //сами стили?
        _ReadCellCommonStyles('cellXfs', CellXfsArray, CellXfsCount) //_ReadCellXfs()
      else
      if xml.IsTagStartByName('cellStyles') then //??
        _ReadCellStyles()
      else
      if xml.IsTagStartByName('colors') then
        _ReadColors()
      else
      if xml.IsTagStartByName('dxfs') then
        _Readdxfs()
      else
      if xml.IsTagStartByName('numFmts') then
        ReadHelper.NumberFormats.ReadNumFmts(xml);
    end; //while

    //тут незабыть применить номера цветов, если были введены

    _CheckIndexedColors();

    //
    for i := 0 to FillCount - 1 do begin
      XLSXApplyColor(FillArray[i].bgcolor, FillArray[i].bgColorType, FillArray[i].lumFactorBG);
      XLSXApplyColor(FillArray[i].patterncolor, FillArray[i].patternColorType, FillArray[i].lumFactorPattern);
    end; //for

    //{tut}

    XMLSS.Styles.Count := CellXfsCount;
    ReadHelper.NumberFormats.StyleFMTCount := CellXfsCount;
    for i := 0 to CellXfsCount - 1 do begin
      t := CellXfsArray[i].xfId;
      ReadHelper.NumberFormats.StyleFMTID[i] := CellXfsArray[i].numFmtId;

      _Style := XMLSS.Styles[i];
      if ((t >= 0) and (t < CellStyleCount)) then
        _ApplyStyle(_Style, CellStyleArray[t]);
      //else
        _ApplyStyle(_Style, CellXfsArray[i]);
    end;

    //Применение цветов к DF
    for i := 0 to ReadHelper.DiffFormatting.Count - 1 do begin
      if (ReadHelper.DiffFormatting[i].UseFontColor) then begin
        XLSXApplyColor(_dfFonts[i].Color, _dfFonts[i].ColorType, _dfFonts[i].LumFactor);
        ReadHelper.DiffFormatting[i].FontColor := _dfFonts[i].Color;
      end;
      if (ReadHelper.DiffFormatting[i].UseBGColor) then begin
        XLSXApplyColor(_dfFills[i].bgcolor, _dfFills[i].bgColorType, _dfFills[i].lumFactorBG);
        ReadHelper.DiffFormatting[i].BGColor := _dfFills[i].bgcolor;
      end;
      if (ReadHelper.DiffFormatting[i].UsePatternColor) then begin
        XLSXApplyColor(_dfFills[i].patterncolor, _dfFills[i].patternColorType, _dfFills[i].lumFactorPattern);
        ReadHelper.DiffFormatting[i].PatternColor := _dfFills[i].patterncolor;
      end;
    end;

    result := true;
  finally
    if (Assigned(xml)) then
      FreeAndNil(xml);
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
end; //ZEXSLXReadStyles

//Читает названия листов (workbook.xml)
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//  var Stream: TStream                 - поток
//  var Relations: TZXLSXRelationsArray - связи
//  var RelationsCount: integer         - кол-во
//RETURN
//      boolean - true - названия прочитались без ошибок
function ZEXSLXReadWorkBook(var XMLSS: TZWorkBook; var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer): boolean;
var
  xml: TZsspXMLReaderH;
  s: string;
  i, t, dn: integer;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  try
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    dn := 0;
    while (not xml.Eof()) do begin
      xml.ReadTag();

      if xml.IsTagStartByName('definedName') then begin
        SetLength(XMLSS.FDefinedNames, dn + 1);
        XMLSS.FDefinedNames[dn].LocalSheetId := StrToIntDef(xml.Attributes.ItemsByName['localSheetId'], 0);
        XMLSS.FDefinedNames[dn].Name := xml.Attributes.ItemsByName['name'];
        inc(dn);
      end
      else if xml.IsTagEndByName('definedName') and (Length(XMLSS.FDefinedNames) = dn) then
        XMLSS.FDefinedNames[dn - 1].Body := xml.TagValue
      else if xml.IsTagClosedByName('sheet') then begin
        s := xml.Attributes.ItemsByName['r:id'];
        for i := 0 to RelationsCount - 1 do
          if (Relations[i].id = s) then begin
            Relations[i].name := ZEReplaceEntity(xml.Attributes.ItemsByName['name']);
            s := xml.Attributes.ItemsByName['sheetId'];
            relations[i].sheetid := -1;
            if (TryStrToInt(s, t)) then
              relations[i].sheetid := t;
            Relations[i].state := xml.Attributes.ItemsByName['state'];
            break;
          end;
      end else
      if xml.IsTagClosedByName('workbookView') then begin
        s := xml.Attributes.ItemsByName['activeTab'];
        s := xml.Attributes.ItemsByName['firstSheet'];
        s := xml.Attributes.ItemsByName['showHorizontalScroll'];
        s := xml.Attributes.ItemsByName['showSheetTabs'];
        s := xml.Attributes.ItemsByName['showVerticalScroll'];
        s := xml.Attributes.ItemsByName['tabRatio'];
        s := xml.Attributes.ItemsByName['windowHeight'];
        s := xml.Attributes.ItemsByName['windowWidth'];
        s := xml.Attributes.ItemsByName['xWindow'];
        s := xml.Attributes.ItemsByName['yWindow'];
      end;
    end; //while
    result := true;
  finally
    xml.Free();
  end;
end; //ZEXSLXReadWorkBook

//Удаляет первый символ + меняет все разделители на нужные
//INPUT
//  var FileArray: TArray<TZXLSXFileItem>  - файлы
//      FilesCount: integer         - кол-во файлов
procedure ZE_XSLXReplaceDelimiter(var FileArray: TArray<TZXLSXFileItem>; FilesCount: integer);
var i, j, k: integer;
begin
  for i := 0 to FilesCount - 1 do begin
    k := length(FileArray[i].name);
    if (k > 1) then begin
      if (FileArray[i].name[1] = '/') then
        Delete(FileArray[i].name, 1, 1);
      if (PathDelim <> '/') then
        for j := 1 to k - 1 do
          if (FileArray[i].name[j] = '/') then
            FileArray[i].name[j] := PathDelim;
    end;
  end;
end; //ZE_XSLXReplaceDelimiter

//Читает связи страниц/стилей  (*.rels: workbook.xml.rels и .rels)
//INPUT
//  var Stream: TStream                 - поток для чтения
//  var Relations: TZXLSXRelationsArray - массив с отношениями
//  var RelationsCount: integer         - кол-во
//  var isWorkSheet: boolean            - признак workbook.xml.rels
//      needReplaceDelimiter: boolean   - признак необходимости заменять разделитель
//RETURN
//      boolean - true - успешно прочитано
function ZE_XSLXReadRelationships(var Stream: TStream; var Relations: TZXLSXRelationsArray; var RelationsCount: integer; var isWorkSheet: boolean; needReplaceDelimiter: boolean): boolean;
var xml: TZsspXMLReaderH; rt: TRelationType;
begin
  result := false;
  xml := TZsspXMLReaderH.Create();
  RelationsCount := 0;
  isWorkSheet := false;
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while not xml.Eof() do begin
      xml.ReadTag();

      if xml.IsTagClosedByName('Relationship') then begin
        SetLength(Relations, RelationsCount + 1);
        Relations[RelationsCount].id := xml.Attributes.ItemsByName['Id'];

        rt := ZEXLSXGetRelationNumber(xml.Attributes.ItemsByName['Type']);
        if ((rt >= rtWorkSheet) and (rt < rtDoc)) then
          isWorkSheet := true;

        Relations[RelationsCount].fileid := -1;
        Relations[RelationsCount].state := '';
        Relations[RelationsCount].sheetid := 0;
        Relations[RelationsCount].name := '';
        Relations[RelationsCount].ftype := rt;
        Relations[RelationsCount].target := xml.Attributes.ItemsByName['Target'];
        if (rt >= rtWorkSheet) then
          inc(RelationsCount);
      end;
    end; //while
    result := true;
  finally
    FreeAndNil(xml);
  end;
end; //ZE_XSLXReadRelationsips

//Читает примечания (добавляет примечания на последнюю страницу)
//INPUT
//  var XMLSS: TZWorkBook - хранилище
//  var Stream: TStream - поток для чтения
//RETURN
//      boolean - true - всё нормально
function ZEXSLXReadComments(var XMLSS: TZWorkBook; var Stream: TStream): boolean;
var xml: TZsspXMLReaderH;
  authors: TList<string>;
  page: integer;

  procedure _ReadComment();
  var _c, _r, _a: integer;
    _comment, str: string;
    _kol: integer;
  begin
    _c := 0;
    _r := 0;
    str := xml.Attributes.ItemsByName['ref'];
    if (str = '') then
      exit;
    if TZEFormula.GetCellCoords(str, _c, _r) then
    begin
      if (_c >= XMLSS.Sheets[page].ColCount) then
        XMLSS.Sheets[page].ColCount := _c + 1;
      if (_r >= XMLSS.Sheets[page].RowCount) then
        XMLSS.Sheets[page].RowCount := _r + 1;

      if (TryStrToInt(xml.Attributes.ItemsByName['authorId'], _a)) then
        if (_a >= 0) and (_a < authors.Count) then
          XMLSS.Sheets[page].Cell[_c, _r].CommentAuthor := authors[_a];

      _comment := '';
      _kol := 0;
      while (not((xml.TagName = 'comment') and (xml.IsTagEnd))) do begin
        xml.ReadTag();
        if (xml.Eof()) then
          break;
        if ((xml.TagName = 't') and (xml.IsTagEnd)) then
        begin
          if (_kol > 0) then
            _comment := _comment + sLineBreak + xml.TextBeforeTag
          else
            _comment := _comment + xml.TextBeforeTag;
          inc(_kol);
        end;
      end; //while
      XMLSS.Sheets[page].Cell[_c, _r].Comment := _comment;
      XMLSS.Sheets[page].Cell[_c, _r].ShowComment := true;
    end //if
  end; //_ReadComment();

begin
  result := false;
  authors:= TList<string>.Create();
  xml := TZsspXMLReaderH.Create();
  page := XMLSS.Sheets.Count - 1;
  if (page < 0) then
    exit;
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    while not xml.Eof() do begin
      xml.ReadTag();

      if xml.IsTagStartByName('authors') then begin
        while xml.ReadToEndTagByName('authors') do begin
          if xml.IsTagEndByName('authors') then
            authors.Add(xml.TextBeforeTag);
        end;
      end else

      if xml.IsTagStartByName('comment') then
        _ReadComment();
    end; //while
    result := true;
  finally
    xml.Free();
    authors.Free();
  end;
end; //ZEXSLXReadComments

procedure XLSXSortRelationArray(var arr: TZXLSXRelationsArray; count: integer);
var tmp: TZXLSXRelations;
  i, j: integer;
  _t1, _t2: integer;
  s: string;
  b: boolean;
  function _cmp(): boolean;
  begin
    b := false;
    s := arr[j].id;
    delete(s, 1, 3);
    b := TryStrToInt(s, _t1);
    if (b) then begin
      s := arr[j + 1].id;
      delete(s, 1, 3);
      b := TryStrToInt(s, _t2);
    end;

    if (b) then
      result := _t1 > _t2
    else
      result := arr[j].id > arr[j + 1].id;
  end;

begin
  //TODO: do not forget update sorting.
  for i := 0 to count - 2 do
    for j := 0 to count - 2 do
      if (_cmp()) then begin
        tmp := arr[j];
        arr[j] := arr[j + 1];
        arr[j + 1] := tmp;
      end;
end;

//Читает распакованный xlsx
//INPUT
//  var XMLSS: TZWorkBook - хранилище
//  DirName: string     - имя папки
//RETURN
//      integer - номер ошибки (0 - всё OK)
function ReadXLSXPath(var XMLSS: TZWorkBook; DirName: string): integer;
var
  stream: TStream;
  FileArray: TArray<TZXLSXFileItem>;
  FilesCount: integer;
  sharedStrings: TObjectList<TRichText>;
  RelationsArray: array of TZXLSXRelationsArray;
  RelationsCounts: array of integer;
  SheetRelations: TZXLSXRelationsArray;
  SheetRelationsCount: integer;
  RelationsCount: integer;
  ThemaColor: TIntegerDynArray;
  ThemaColorCount: integer;
  SheetRelationNumber: integer;
  i, j, k: integer;
  s: string;
  b: boolean;
  _no_sheets: boolean;
  MaximumDigitWidth: double;
  RH: TZEXLSXReadHelper;
  //Пытается прочитать rel для листа
  //INPUT
  //  const fname: string - имя файла листа
  //RETURN
  //      boolean - true - прочитал успешно
  function _CheckSheetRelations(const fname: string): boolean;
  var rstream: TStream;
    s: string;
    i, num: integer;
    b: boolean;
  begin
    result := false;
    SheetRelationsCount := 0;
    num := -1;
    b := false;
    s := '';
    for i := length(fname) downto 1 do
    if (fname[i] = PathDelim) then
    begin
      num := i;
      s := fname;
      insert('_rels' + PathDelim, s, num + 1);
      s := DirName + s + '.rels';
      if (not FileExists(s)) then
        num := -1;
      break;
    end;

    if (num > 0) then begin
      rstream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
      try
        result := ZE_XSLXReadRelationships(rstream, SheetRelations, SheetRelationsCount, b, false);
      finally
        rstream.Free();
      end;
    end;
  end; //_CheckSheetRelations

  //Прочитать примечания
  procedure _ReadComments();
  var i, l: integer;
    s: string;
    b: boolean;
    stream: TStream;
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

    //Если найдены примечания
    if (b) then begin
      l := length(s);
      if (l >= 3) then
        if ((s[1] = '.') and (s[2] = '.')) then
          delete(s, 1, 3);
      b := false;
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtComments) then
          if (pos(s, FileArray[i].name) <> 0) then
            if (FileExists(DirName + FileArray[i].name)) then begin
              s := DirName + FileArray[i].name;
              b := true;
              break;
            end;
      //Если файл не найден
      if (not b) then begin
        s := DirName + 'xl' + PathDelim + s;
        if FileExists(s) then
          b := true;
      end;

      //Файл с примечаниями таки присутствует!
      if (b) then begin
        stream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
        try
          ZEXSLXReadComments(XMLSS, stream);
        finally
          stream.Free();
        end;
      end;
    end;
  end; //_ReadComments

begin
  result := 0;
  MaximumDigitWidth := 0;
  FilesCount := 0;
  FileArray := nil;

  if (not TDirectory.Exists(DirName)) then begin
    result := -1;
    exit;
  end;

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  stream := nil;
  RelationsCount := 0;
  ThemaColorCount := 0;
  SheetRelationsCount := 0;
  ThemaColor := nil;
  RH := nil;

  sharedStrings := TObjectList<TRichText>.Create(true);
  try
    try
      stream := TFileStream.Create(DirName + '[Content_Types].xml', fmOpenRead or fmShareDenyNone);
      if (not ZEXSLXReadContentTypes(stream,  FileArray, FilesCount)) then begin
        result := 3;
        exit;
      end;

      FreeAndNil(stream);

      ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      SheetRelationNumber := -1;

      b := false;
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = TRelationType.rtDoc) then begin
        b := true;
        break;
      end;

      RH := TZEXLSXReadHelper.Create();

      if (not b) then begin
        s := DirName + '_rels' + PathDelim + '.rels';
        if (FileExists(s)) then begin
          SetLength(FileArray, FilesCount + 1);
          s := '/_rels/.rels';
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        s := DirName + 'xl' + PathDelim + '_rels' + PathDelim + 'workbook.xml.rels';
        if (FileExists(s)) then begin
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
      if (FileArray[i].ftype = TRelationType.rtDoc) then begin
        SetLength(RelationsArray, RelationsCount + 1);
        SetLength(RelationsCounts, RelationsCount + 1);

        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZE_XSLXReadRelationships(stream, RelationsArray[RelationsCount], RelationsCounts[RelationsCount], b, true)) then
        begin
          result := 4;
          exit;
        end;
        if (b) then begin
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
        end; //if
        FreeAndNil(stream);
        inc(RelationsCount);
      end;

      //sharedStrings.xml
      for i:= 0 to FilesCount - 1 do
      if (FileArray[i].ftype = TRelationType.rtCoreProp) then begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadSharedStrings(stream, sharedStrings)) then begin
          result := 3;
          exit;
        end;
        break;
      end;

      //тема (если есть)
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = TRelationType.rtVmlDrawing) then begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadTheme(stream, ThemaColor, ThemaColorCount)) then
        begin
          result := 6;
          exit;
        end;
        break;
      end;

      //стили (styles.xml)
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = TRelationType.rtStyles) then begin
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadStyles(XMLSS, stream, ThemaColor, ThemaColorCount, MaximumDigitWidth, RH)) then begin
          result := 5;
          exit;
        end else
          b := true;
      end;

      //чтение страниц
      _no_sheets := true;
      if (SheetRelationNumber > 0) then begin
        for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtSharedStr) then begin
          FreeAndNil(stream);
          stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
          if (not ZEXSLXReadWorkBook(XMLSS, stream, RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber])) then
          begin
            result := 3;
            exit;
          end;
          break;
        end; //if

        //for i := 1 to RelationsCounts[SheetRelationNumber] do
        XLSXSortRelationArray(RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
          if (RelationsArray[SheetRelationNumber][j].sheetid > 0) then begin
            b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
            FreeAndNil(stream);
            stream := TFileStream.Create(DirName + FileArray[RelationsArray[SheetRelationNumber][j].fileid].name, fmOpenRead or fmShareDenyNone);
            if (not ZEXSLXReadSheet(XMLSS, stream, @RelationsArray[SheetRelationNumber][j], sharedStrings, SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
              result := result or 4;
            if (b) then
              _ReadComments();
            _no_sheets := false;
          end; //if
      end;
      //если прочитано 0 листов - пробуем прочитать все (не удалось прочитать workbook/rel)
      if (_no_sheets) then
      for i := 0 to FilesCount - 1 do
      if (FileArray[i].ftype = TRelationType.rtWorkSheet) then begin
        b := _CheckSheetRelations(FileArray[i].name);
        FreeAndNil(stream);
        stream := TFileStream.Create(DirName + FileArray[i].name, fmOpenRead or fmShareDenyNone);
        if (not ZEXSLXReadSheet(XMLSS, stream, nil, sharedStrings, SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
          result := result or 4;
        if (b) then
            _ReadComments();
      end;
    except
      result := 2;
    end;
  finally
    if (Assigned(stream)) then
      FreeAndNil(stream);

    SetLength(FileArray, 0);
    FileArray := nil;
    sharedStrings.Free();
    for i := 0 to RelationsCount - 1 do begin
      Setlength(RelationsArray[i], 0);
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
end; //ReadXLSXPath

function ReadXLSXFile(var XMLSS: TZWorkBook; zipStream: TStream): integer;
var
  stream: TStream;
  FileArray: TArray<TZXLSXFileItem>;
  FileList: TList<TZXLSXFileItem>;
  FilesCount: integer;
  sharedStrings: TObjectList<TRichText>;
  RelationsArray: array of TZXLSXRelationsArray;
  RelationsCounts: array of integer;
  SheetRelations: TZXLSXRelationsArray;
  SheetRelationsCount: integer;
  RelationsCount: integer;
  ThemaColor: TIntegerDynArray;
  ThemaColorCount: integer;
  SheetRelationNumber: integer;
  i, j, k: integer;
  s: string;
  b: boolean;
  _no_sheets: boolean;
  RH: TZEXLSXReadHelper;
  zip: TZipFile;
  encoding: TEncoding;
  zipHdr: TZipHeader;
  buff: TBytes;
  MaximumDigitWidth: double;
  //zfiles: TArray<string>;

  function _CheckSheetRelations(const fname: string): boolean;
  var rstream: TStream;
    s: string;
    i, num: integer;
    b: boolean;
  begin
    result := false;
    SheetRelationsCount := 0;
    num := -1;
    b := false;
    s := '';
    for i := length(fname) downto 1 do begin
      if (fname[i] = PathDelim) then begin
        num := i;
        s := fname;
        insert('_rels' + PathDelim, s, num + 1);
        s := s + '.rels';
        if (not FileExists(s)) then
          num := -1;
        break;
      end;
    end;

    if (num > 0) then begin
      rstream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
      try
        result := ZE_XSLXReadRelationships(rstream, SheetRelations, SheetRelationsCount, b, false);
      finally
        rstream.Free();
      end;
    end;
  end; //_CheckSheetRelations

  //Прочитать примечания
  procedure _ReadComments();
  var i, l: integer;
    s: string;
    b: boolean;
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

    //Если найдены примечания
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
      //Если файл не найден
      if (not b) then begin
        s := 'xl' + PathDelim + s;
        if (FileExists(s)) then
          b := true;
      end;

      //Файл с примечаниями таки присутствует!
      if (b) then
      try
        _stream := TFileStream.Create(s, fmOpenRead or fmShareDenyNone);
        ZEXSLXReadComments(XMLSS, _stream);
      finally
        if (Assigned(_stream)) then
          FreeAndNil(_stream);
      end;
    end;
  end; //_ReadComments

begin
  result := 0;
  MaximumDigitWidth := 0;
  FilesCount := 0;
  zip := TZipFile.Create();
  encoding := TEncoding.GetEncoding(437);
{$IFDEF VER330}
  zip.Encoding := encoding;
{$ENDIF}

  XMLSS.Styles.Clear();
  XMLSS.Sheets.Count := 0;
  RelationsCount := 0;
  ThemaColorCount := 0;
  SheetRelationsCount := 0;
  RH := TZEXLSXReadHelper.Create();
  FileList := TList<TZXLSXFileItem>.Create();
  sharedStrings := TObjectList<TRichText>.Create(true);
  stream := nil;
  try
    zip.Open(zipStream, zmRead);
    try
      zip.Read('[Content_Types].xml', stream, zipHdr);
      try
        if (not ZEXSLXReadContentTypes(stream, FileArray, FilesCount)) then
          raise Exception.Create('Could not read [Content_Types].xml');
      finally
        FreeAndNil(stream);
      end;

      ZE_XSLXReplaceDelimiter(FileArray, FilesCount);
      SheetRelationNumber := -1;

      b := false;
      for i := 0 to FilesCount - 1 do begin
        if (FileArray[i].ftype = TRelationType.rtDoc) then begin
          b := true;
          break;
        end;
      end;

      if (not b) then
      begin
        s := '/_rels/.rels';
        if zip.IndexOf(s.Substring(1)) > -1 then
        begin
          SetLength(FileArray, FilesCount + 1);
          FileArray[FilesCount].original := s;
          FileArray[FilesCount].name := s;
          FileArray[FilesCount].ftype := TRelationType.rtDoc;
          inc(FilesCount);
        end;

        s := '/xl/_rels/workbook.xml.rels';
        if zip.IndexOf(s.Substring(1)) > -1 then
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
          zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
          try
            if (not ZE_XSLXReadRelationships(stream, RelationsArray[RelationsCount], RelationsCounts[RelationsCount], b, true)) then
            begin
              result := 4;
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
            end; //if
          finally
            FreeAndNil(stream);
          end;
          inc(RelationsCount);
        end;

      //sharedStrings.xml
      for i:= 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtCoreProp) then
        begin
          zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
          try
            if (not ZEXSLXReadSharedStrings(stream, sharedStrings)) then begin
              result := 3;
              exit;
            end;
          finally
            FreeAndNil(stream);
          end;
          break;
        end;

      //тема (если есть)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtVmlDrawing) then
        begin
          zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
          try
            if (not ZEXSLXReadTheme(stream, ThemaColor, ThemaColorCount)) then
            begin
              result := 6;
              exit;
            end;
          finally
            FreeAndNil(stream);
          end;
          break;
        end;

      //стили (styles.xml)
      for i := 0 to FilesCount - 1 do
        if (FileArray[i].ftype = TRelationType.rtStyles) then
        begin
          zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
          try
            if (not ZEXSLXReadStyles(XMLSS, stream, ThemaColor, ThemaColorCount, MaximumDigitWidth, RH)) then begin
              result := 5;
              exit;
            end
            else
              b := true;
          finally
            FreeAndNil(stream);
          end;
        end;

      //чтение страниц
      _no_sheets := true;
      if (SheetRelationNumber > 0) then
      begin
        for i := 0 to FilesCount - 1 do
          if (FileArray[i].ftype = TRelationType.rtSharedStr) then
          begin
            zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
            try
              if (not ZEXSLXReadWorkBook(XMLSS, stream, RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber])) then
              begin
                result := 3;
                exit;
              end;
            finally
              FreeAndNil(stream);
            end;
            break;
          end; //if

        //for i := 1 to RelationsCounts[SheetRelationNumber] do
        XLSXSortRelationArray(RelationsArray[SheetRelationNumber], RelationsCounts[SheetRelationNumber]);
        for j := 0 to RelationsCounts[SheetRelationNumber] - 1 do
          if (RelationsArray[SheetRelationNumber][j].sheetid > 0) then
          begin
            b := _CheckSheetRelations(FileArray[RelationsArray[SheetRelationNumber][j].fileid].name);
            zip.Read(FileArray[RelationsArray[SheetRelationNumber][j].fileid].original.Substring(1), stream, zipHdr);
            try
              if (not ZEXSLXReadSheet(XMLSS, stream, @RelationsArray[SheetRelationNumber][j], sharedStrings, SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
                result := result or 4;
              if (b) then
                _ReadComments();
            finally
              FreeAndNil(stream);
            end;
            _no_sheets := false;
          end; //if
      end;

      //если прочитано 0 листов - пробуем прочитать все (не удалось прочитать workbook/rel)
      if _no_sheets then begin
        for i := 0 to FilesCount - 1 do begin
          if FileArray[i].ftype = TRelationType.rtWorkSheet then begin
            b := _CheckSheetRelations(FileArray[i].name);
            zip.Read(FileArray[i].original.Substring(1), stream, zipHdr);
            try
              if (not ZEXSLXReadSheet(XMLSS, stream, nil, sharedStrings, SheetRelations, SheetRelationsCount, MaximumDigitWidth, RH)) then
                result := result or 4;
              if (b) then
                _ReadComments();
            finally
              FreeAndNil(stream);
            end;
          end;
        end;
      end;

      // drawings
      for I := 0 to XMLSS.Sheets.Count-1 do begin
        if XMLSS.Sheets[i].DrawingRid > 0 then begin
          // load images
          s := 'xl/drawings/drawing'+IntToStr(i+1)+'.xml';
          if zip.IndexOf(s) > -1 then begin
            zip.Read(s, stream, zipHdr);
            try
              ZEXLSXReadDrawing(XMLSS.Sheets[i], stream);
            finally
              stream.Free();
            end;

            // read drawing rels
            s := 'xl/drawings/_rels/drawing'+IntToStr(i+1)+'.xml.rels';
            zip.Read(s, stream, zipHdr);
            try
              ZEXLSXReadDrawingRels(XMLSS.Sheets[i], stream);
            finally
              stream.Free();
            end;

            // read img file
            for j := 0 to XMLSS.Sheets[i].Drawing.Count-1 do begin
              s := XMLSS.Sheets[i].Drawing[j].Name;
              zip.Read('xl/media/' + s, buff);
              // only unique content
              XMLSS.AddMediaContent(s, buff, true);
            end;
          end;
        end;
      end;
    except
      result := 2;
    end;
  finally
    zip.Free();
    encoding.Free;
    FileList.Free();
    sharedStrings.Free();
    RH.Free();
  end;
end; //ReadXLSXPath

/////////////////////////////////////////////////////////////////////////////
/////////////                    запись                         /////////////
/////////////////////////////////////////////////////////////////////////////

//Создаёт [Content_Types].xml
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    PageCount: integer                - кол-во страниц
//    CommentCount: integer             - кол-во страниц с комментариями
//  const PagesComments: TIntegerDynArray- номера страниц с комментариями (нумеряция с нуля)
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//  const WriteHelper: TZEXLSXWriteHelper - additional data
//RETURN
//      integer
function ZEXLSXCreateContentTypes(var XMLSS: TZWorkBook; Stream: TStream; PageCount: integer; CommentCount: integer; const PagesComments: TIntegerDynArray;
                                  TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring;
                                  const WriteHelper: TZEXLSXWriteHelper): integer;
var xml: TZsspXMLWriterH; s: string;
  procedure _WriteOverride(const PartName: string; ct: integer);
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('PartName', PartName);
    case ct of
      0: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
      1: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
      2: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
      3: s := 'application/vnd.openxmlformats-package.relationships+xml';
      4: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
      5: s := 'application/vnd.openxmlformats-package.core-properties+xml';
      6: s := 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
      7: s := 'application/vnd.openxmlformats-officedocument.vmlDrawing';
      8: s := 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
      9: s := 'application/vnd.openxmlformats-officedocument.drawing+xml';
    end;
    xml.Attributes.Add('ContentType', s, false);
    xml.WriteEmptyTag('Override', true);
  end; //_WriteOverride

  procedure _WriteTypeDefault(extension, contentType: string);
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('Extension', extension);
    xml.Attributes.Add('ContentType', contentType, false);
    xml.WriteEmptyTag('Default', true);
  end;

  procedure _WriteTypes();
  var i: integer;
  begin
    _WriteTypeDefault('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    _WriteTypeDefault('xml',  'application/xml');
    _WriteTypeDefault('png',  'image/png');
    _WriteTypeDefault('jpeg', 'image/jpeg');
    _WriteTypeDefault('wmf',  'image/x-wmf');

    //Страницы
    //_WriteOverride('/_rels/.rels', 3);
    //_WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for i := 0 to PageCount - 1 do begin
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
      if (WriteHelper.IsSheetHaveHyperlinks(i)) then
        _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels', 3);
    end;
    //комментарии
    for i := 0 to CommentCount - 1 do begin
      _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
      _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
    end;

    for i := 0 to XMLSS.Sheets.Count - 1 do begin
      if Assigned(XMLSS.Sheets[i].Drawing) and (XMLSS.Sheets[i].Drawing.Count > 0) then begin
        _WriteOverride('/xl/drawings/drawing' + IntToStr(i+1) + '.xml', 9);
        //_WriteOverride('/xl/drawings/_rels/drawing' + IntToStr(i+1) + '.xml.rels', 3);
//        for ii := 0 to _drawing.PictureStore.Count - 1 do begin
//          _picture := _drawing.PictureStore.Items[ii];
//          // image/ override
//          xml.Attributes.Clear();
//          xml.Attributes.Add('PartName', '/xl/media/' + _picture.Name);
//          xml.Attributes.Add('ContentType', 'image/' + Copy(ExtractFileExt(_picture.Name), 2, 99), false);
//          xml.WriteEmptyTag('Override', true);
//        end;
      end;
    end;

    _WriteOverride('/xl/workbook.xml', 2);
    _WriteOverride('/xl/styles.xml', 1);
    _WriteOverride('/xl/sharedStrings.xml', 4);
    _WriteOverride('/docProps/app.xml', 8);
    _WriteOverride('/docProps/core.xml', 5);
  end; //_WriteTypes

begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE + '/content-types');
    xml.WriteTagNode('Types', true, true, true);
    _WriteTypes();
    xml.WriteEndTagNode(); //Types
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateContentTypes

function ZEXLSXCreateDrawing(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var xml: TZsspXMLWriterH;
    pic: TZEPicture;
    i: integer;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.NewLine := false;
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
    xml.Attributes.Add('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    xml.WriteTagNode('xdr:wsDr', false, false, false);

    for i := 0 to sheet.Drawing.Count - 1 do begin
      pic := sheet.Drawing.Items[i];
      // cell anchor
      xml.Attributes.Clear();
      if pic.CellAnchor = ZAAbsolute then
        xml.Attributes.Add('editAs', 'absolute')
      else
        xml.Attributes.Add('editAs', 'oneCell');
      xml.WriteTagNode('xdr:twoCellAnchor', false, false, false);

      // - xdr:from
      xml.Attributes.Clear();
      xml.WriteTagNode('xdr:from', false, false, false);
      xml.WriteTag('xdr:col',    IntToStr(pic.FromCol), false, false);
      xml.WriteTag('xdr:colOff', IntToStr(pic.FromColOff), false, false);
      xml.WriteTag('xdr:row',    IntToStr(pic.FromRow), false, false);
      xml.WriteTag('xdr:rowOff', IntToStr(pic.FromRowOff), false, false);
      xml.WriteEndTagNode(); // xdr:from
      // - xdr:to
      xml.Attributes.Clear();
      xml.WriteTagNode('xdr:to', false, false, false);
      xml.WriteTag('xdr:col',    IntToStr(pic.ToCol), false, false);
      xml.WriteTag('xdr:colOff', IntToStr(pic.ToColOff), false, false);
      xml.WriteTag('xdr:row',    IntToStr(pic.ToRow), false, false);
      xml.WriteTag('xdr:rowOff', IntToStr(pic.ToRowOff), false, false);
      xml.WriteEndTagNode(); // xdr:from
      // - xdr:pic
      xml.WriteTagNode('xdr:pic', false, false, false);
      // -- xdr:nvPicPr
      xml.WriteTagNode('xdr:nvPicPr', false, false, false);
      // --- xdr:cNvPr
      xml.Attributes.Clear();
      xml.Attributes.Add('descr', pic.Description);
      xml.Attributes.Add('name', pic.Title);
      xml.Attributes.Add('id', IntToStr(pic.Id));  // 1
      xml.WriteEmptyTag('xdr:cNvPr', false);
      // --- xdr:cNvPicPr
      xml.Attributes.Clear();
      xml.WriteEmptyTag('xdr:cNvPicPr', false);
      xml.WriteEndTagNode(); // -- xdr:nvPicPr

      // -- xdr:blipFill
      xml.Attributes.Clear();
      xml.WriteTagNode('xdr:blipFill', false, false, false);
      // --- a:blip
      xml.Attributes.Clear();
      xml.Attributes.Add('r:embed', 'rId' + IntToStr(pic.RelId)); // rId1
      xml.WriteEmptyTag('a:blip', false);
      // --- a:stretch
      xml.Attributes.Clear();
      xml.WriteEmptyTag('a:stretch', false);
      xml.WriteEndTagNode(); // -- xdr:blipFill

      // -- xdr:spPr
      xml.Attributes.Clear();
      xml.WriteTagNode('xdr:spPr', false, false, false);
      // --- a:xfrm
      xml.WriteTagNode('a:xfrm', false, false, false);
      // ----
      xml.Attributes.Clear();
      xml.Attributes.Add('x', IntToStr(pic.FrmOffX));
      xml.Attributes.Add('y', IntToStr(pic.FrmOffY));
      xml.WriteEmptyTag('a:off', false);
      // ----
      xml.Attributes.Clear();
      xml.Attributes.Add('cx', IntToStr(pic.FrmExtCX));
      xml.Attributes.Add('cy', IntToStr(pic.FrmExtCY));
      xml.WriteEmptyTag('a:ext', false);
      xml.Attributes.Clear();
      xml.WriteEndTagNode(); // --- a:xfrm

      // --- a:prstGeom
      xml.Attributes.Clear();
      xml.Attributes.Add('prst', 'rect');
      xml.WriteTagNode('a:prstGeom', false, false, false);
      xml.Attributes.Clear();
      xml.WriteEmptyTag('a:avLst', false);
      xml.WriteEndTagNode(); // --- a:prstGeom

      // --- a:ln
      xml.WriteTagNode('a:ln', false, false, false);
      xml.WriteEmptyTag('a:noFill', false);
      xml.WriteEndTagNode(); // --- a:ln

      xml.WriteEndTagNode(); // -- xdr:spPr

      xml.WriteEndTagNode(); // - xdr:pic

      xml.WriteEmptyTag('xdr:clientData', false);

      xml.WriteEndTagNode(); // xdr:twoCellAnchor
    end;
    xml.WriteEndTagNode(); // xdr:wsDr
  finally
    xml.Free();
  end;
end;

function ZEXLSXCreateDrawingRels(sheet: TZSheet; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring): integer;
var xml: TZsspXMLWriterH;
    i: integer;
    dic: TDictionary<integer, string>;
    pair: TPair<integer, string>;
begin
  result := 0;
  dic := TDictionary<integer, string>.Create();
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL, false);
    xml.WriteTagNode('Relationships', false, false, false);

    for i := 0 to sheet.Drawing.Count - 1 do begin
      dic.AddOrSetValue(sheet.Drawing[i].RelId, sheet.Drawing[i].Name);
    end;

    for pair in dic do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('Id', 'rId' + IntToStr(pair.Key));
      xml.Attributes.Add('Type', SCHEMA_DOC_REL + '/image');
      xml.Attributes.Add('Target', '../media/' + pair.Value);
      xml.WriteEmptyTag('Relationship', false, true);
    end;
    xml.WriteEndTagNode(); // Relationships
  finally
    xml.Free();
    dic.Free();
  end;
end;

//Создаёт лист документа (sheet*.xml)
//INPUT
//    XMLSS: TZWorkBook                                       - хранилище
//    Stream: TStream                                       - поток для записи
//    SheetNum: integer                                     - номер листа в документе
//    SharedStrings: TStringDynArray                        - общие строки
//    SharedStringsDictionary: TDictionary<string, integer> - словарь для определения дубликатов строк
//    TextConverter: TAnsiToCPConverter                     - конвертер из локальной кодировки в нужную
//    CodePageName: string                                  - название кодовой страници
//    BOM: ansistring                                       - BOM
//    isHaveComments: boolean                               - возвращает true, если были комментарии (чтобы создать comments*.xml)
//    WriteHelper: TZEXLSXWriteHelper                       - additional data
//RETURN
//      integer
function ZEXLSXCreateSheet(var XMLSS: TZWorkBook; Stream: TStream; SheetNum: integer; staredStrings: TSharedStringBank;
  TextConverter: TAnsiToCPConverter; CodePageName: String; BOM: ansistring; const WriteHelper: TZEXLSXWriteHelper): integer;
var xml: TZsspXMLWriterH;    //писатель
  sheet: TZSheet;
  procedure WriteXLSXSheetHeader();
  var s: string;
    b: boolean;
    sheetOptions: TZSheetOptions;
    procedure _AddSplitValue(const SplitMode: TZSplitMode; const SplitValue: integer; const AttrName: string);
    var s: string; b: boolean;
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
        ZSplitSplit: s := IntToStr(round(PixelToPoint(SplitValue) * 20));
        ZSplitNone: b := false;
      end;
      if (b) then
        xml.Attributes.Add(AttrName, s);
    end; //_AddSplitValue

    procedure _AddTopLeftCell(const VMode: TZSplitMode; const VValue: integer; const HMode: TZSplitMode; const HValue: integer);
    var isProblem: boolean;
    begin
      isProblem := (VMode = ZSplitSplit) or (HMode = ZSplitSplit);
      isProblem := isProblem or (VValue > 1000) or (HValue > 100);
      if not isProblem then begin
        s := TZEFormula.GetColAddres(VValue) + IntToSTr(HValue + 1);
        xml.Attributes.Add('topLeftCell', s);
      end;
    end; //_AddTopLeftCell

  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('filterMode', 'false');
    xml.WriteTagNode('sheetPr', true, true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('rgb', 'FF'+ColorToHTMLHex(sheet.TabColor));
    xml.WriteEmptyTag('tabColor', true, false);

    xml.Attributes.Clear();
    if sheet.ApplyStyles      then xml.Attributes.Add('applyStyles', '1');
    if not sheet.SummaryBelow then xml.Attributes.Add('summaryBelow', '0');
    if not sheet.SummaryRight then xml.Attributes.Add('summaryRight', '0');
    xml.WriteEmptyTag('outlinePr', true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('fitToPage', XLSXBoolToStr(sheet.FitToPage));
    xml.WriteEmptyTag('pageSetUpPr', true, false);

    xml.WriteEndTagNode(); //sheetPr

    xml.Attributes.Clear();
    s := 'A1';
    if (sheet.ColCount > 0) then
      s := s + ':' + TZEFormula.GetColAddres(sheet.ColCount - 1) + IntToStr(sheet.RowCount);
    xml.Attributes.Add('ref', s);
    xml.WriteEmptyTag('dimension', true, false);

    xml.Attributes.Clear();
    xml.WriteTagNode('sheetViews', true, true, true);

    xml.Attributes.Add('colorId', '64');
    xml.Attributes.Add('defaultGridColor', 'true', false);
    xml.Attributes.Add('rightToLeft', 'false', false);
    xml.Attributes.Add('showFormulas', 'false', false);
    xml.Attributes.Add('showGridLines', 'true', false);
    xml.Attributes.Add('showOutlineSymbols', 'true', false);
    xml.Attributes.Add('showRowColHeaders', 'true', false);
    xml.Attributes.Add('showZeros', ifthen(sheet.ShowZeros, '1', '0'), false);

    if sheet.Selected then
      xml.Attributes.Add('tabSelected', 'true', false);

    xml.Attributes.Add('topLeftCell', 'A1', false);

    if sheet.ViewMode = zvmPageBreakPreview then
      xml.Attributes.Add('view', 'pageBreakPreview', false)
    else
      xml.Attributes.Add('view', 'normal', false);

    xml.Attributes.Add('windowProtection', 'false', false);
    xml.Attributes.Add('workbookViewId', '0', false);
    xml.Attributes.Add('zoomScale', '100', false);
    xml.Attributes.Add('zoomScaleNormal', '100', false);
    xml.Attributes.Add('zoomScalePageLayoutView', '100', false);
    xml.WriteTagNode('sheetView', true, true, false);

    {$REGION 'write sheetFormatPr'}
    if (sheet.OutlineLevelCol > 0) or (sheet.OutlineLevelRow > 0) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('defaultColWidth', FloatToStr(sheet.DefaultColWidth, TFormatSettings.Invariant));
        xml.Attributes.Add('defaultRowHeight', FloatToStr(sheet.DefaultRowHeight, TFormatSettings.Invariant));
        if (sheet.OutlineLevelCol > 0) then
            xml.Attributes.Add('outlineLevelCol', IntToStr(sheet.OutlineLevelCol));
        if (sheet.OutlineLevelRow > 0) then
            xml.Attributes.Add('outlineLevelRow', IntToStr(sheet.OutlineLevelRow));
        xml.WriteEmptyTag('sheetFormatPr', true, false);
    end;
    {$ENDREGION}

    sheetOptions := sheet.SheetOptions;

    b := (sheetOptions.SplitVerticalMode <> ZSplitNone) or
         (sheetOptions.SplitHorizontalMode <> ZSplitNone);
    if (b) then begin
      xml.Attributes.Clear();
      _AddSplitValue(sheetOptions.SplitVerticalMode,
                     sheetOptions.SplitVerticalValue,
                     'xSplit');
      _AddSplitValue(sheetOptions.SplitHorizontalMode,
                     sheetOptions.SplitHorizontalValue,
                     'ySplit');

      _AddTopLeftCell(sheetOptions.SplitVerticalMode, sheetOptions.SplitVerticalValue,
                      sheetOptions.SplitHorizontalMode, sheetOptions.SplitHorizontalValue);

      xml.Attributes.Add('activePane', 'topLeft');

      s := 'split';
      if ((sheetOptions.SplitVerticalMode = ZSplitFrozen) or (sheetOptions.SplitHorizontalMode = ZSplitFrozen)) then
        s := 'frozen';
      xml.Attributes.Add('state', s);

      xml.WriteEmptyTag('pane', true, false);
    end; //if
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

    s := TZEFormula.GetColAddres(sheet.SheetOptions.ActiveCol) + IntToSTr(sheet.SheetOptions.ActiveRow + 1);
    xml.Attributes.Clear();
    xml.Attributes.Add('activeCell', s);
    xml.Attributes.Add('sqref', s);
    xml.WriteEmptyTag('selection', true, false);

    xml.WriteEndTagNode(); //sheetView
    xml.WriteEndTagNode(); //sheetViews
  end; //WriteXLSXSheetHeader

  procedure WriteXLSXSheetCols();
  var i: integer;
    s: string;
    ProcessedColumn: TZColOptions;
    MaximumDigitWidth: double;
    NumberOfCharacters: double;
    width: real;
  begin
    MaximumDigitWidth := GetMaximumDigitWidth(XMLSS.Styles[0].Font.Name, XMLSS.Styles[0].Font.Size); //Если совсем нет стилей, пусть будет ошибка.
    xml.Attributes.Clear();
    xml.WriteTagNode('cols', true, true, true);
    for i := 0 to sheet.ColCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('collapsed', 'false', false);
      xml.Attributes.Add('hidden', XLSXBoolToStr(sheet.Columns[i].Hidden), false);
      xml.Attributes.Add('max', IntToStr(i + 1), false);
      xml.Attributes.Add('min', IntToStr(i + 1), false);
      s := '0';
      ProcessedColumn := sheet.Columns[i];
      if ((ProcessedColumn.StyleID >= -1) and (ProcessedColumn.StyleID < XMLSS.Styles.Count)) then
        s := IntToStr(ProcessedColumn.StyleID + 1);
      xml.Attributes.Add('style', s, false);
      //xml.Attributes.Add('width', ZEFloatSeparator(FormatFloat('0.##########', ProcessedColumn.WidthMM * 5.14509803921569 / 10)), false);
      //А.А.Валуев. Формулы расёта ширины взяты здесь - https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
      //А.А.Валуев. Получаем ширину в символах в Excel-е.
      NumberOfCharacters := Trunc((ProcessedColumn.WidthPix - 5) / MaximumDigitWidth * 100 + 0.5) / 100;
      //А.А.Валуев. Конвертируем ширину в символах в ширину для сохранения в файл.
      width := Trunc((NumberOfCharacters * MaximumDigitWidth + 5) / MaximumDigitWidth * 256) / 256;
      xml.Attributes.Add('width', ZEFloatSeparator(FormatFloat('0.##########', width)), false);
      if ProcessedColumn.AutoFitWidth then
        xml.Attributes.Add('bestFit', '1', false);
      if sheet.Columns[i].OutlineLevel > 0 then
        xml.Attributes.Add('outlineLevel', IntToStr(sheet.Columns[i].OutlineLevel));
      xml.WriteEmptyTag('col', true, false);
    end;
    xml.WriteEndTagNode(); //cols
  end; //WriteXLSXSheetCols

  procedure WriteXLSXSheetData();
  var i, j, n: integer;
    b: boolean;
    s, cellValue: string;
    _r: TRect;
    cell: TZCell;
  begin
    xml.Attributes.Clear();
    xml.WriteTagNode('sheetData', true, true, true);
    n := sheet.ColCount - 1;
    for i := 0 to sheet.RowCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('collapsed', 'false', false); //?
      xml.Attributes.Add('customFormat', 'false', false); //?
      xml.Attributes.Add('customHeight', XLSXBoolToStr((abs(sheet.DefaultRowHeight - sheet.Rows[i].Height) > 0.001)){'true'}, false); //?
      xml.Attributes.Add('hidden', XLSXBoolToStr(sheet.Rows[i].Hidden), false);
      xml.Attributes.Add('ht', ZEFloatSeparator(FormatFloat('0.##', sheet.Rows[i].HeightMM * 2.835)), false);
      if sheet.Rows[i].OutlineLevel > 0 then
        xml.Attributes.Add('outlineLevel', IntToStr(sheet.Rows[i].OutlineLevel), false);
      xml.Attributes.Add('r', IntToStr(i + 1), false);
      xml.WriteTagNode('row', true, true, false);
      for j := 0 to n do begin
        cell := sheet.Cell[j, i];
        xml.Attributes.Clear();
        if (not WriteHelper.isHaveComments) and (cell.Comment > '') then
          WriteHelper.isHaveComments := true;
        b := (cell.Data > '') or (cell.Formula > '');
        s := TZEFormula.GetColAddres(j) + IntToStr(i + 1);

        if (cell.HRef <> '') then
          WriteHelper.AddHyperLink(s, cell.HRef, cell.HRefScreenTip, 'External');

        xml.Attributes.Add('r', s);

        if (cell.CellStyle >= -1) and (cell.CellStyle < XMLSS.Styles.Count) then
          s := IntToStr(cell.CellStyle + 1)
        else
          s := '0';
        xml.Attributes.Add('s', s, false);

        cellValue := cell.Data;
        case cell.CellType of
          ZENumber:   s := 'n';
          ZEDateTime: s := 'd'; //??
          ZEBoolean:  s := 'b';
          ZEString:
          begin
            // А.А.Валуев Общие строки пишем только, если в строке есть
            // определённые символы. Хотя можно писать и всё подряд.
            if Assigned(cell.RichText) then begin
              cellValue := IntToStr(staredStrings.Add(cell.RichText));
              s := 's';
            end else
            if cellValue.StartsWith(' ') or cellValue.EndsWith(' ')
            or (cellValue.IndexOfAny([#10, #13]) >= 0) then begin
              cellValue := IntToStr(staredStrings.Add(cellValue));
              s := 's';
            end else
              s := 'str';
          end;
          ZEError: s := 'e';
        end;

        // если тип ячейки ZEGeneral, то атрибут опускаем
        if (cell.CellType <> ZEGeneral) and (cell.CellType <> ZEDateTime) then
          xml.Attributes.Add('t', s, false);

        if (b) then begin
          xml.WriteTagNode('c', true, true, false);
          if (cell.Formula > '') then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('aca', 'false');
            xml.WriteTag('f', cell.Formula, true, false, true);
          end;
          if (cellValue > '') then begin
            xml.Attributes.Clear();
            xml.WriteTag('v', cellValue, true, false, true);
          end;
          xml.WriteEndTagNode();
        end else
          xml.WriteEmptyTag('c', true);
      end;
      xml.WriteEndTagNode(); //row
    end; //for i

    xml.WriteEndTagNode(); //sheetData

    // autoFilter
    if trim(sheet.AutoFilter)<>'' then begin
      xml.Attributes.Clear;
      xml.Attributes.Add('ref', sheet.AutoFilter);
      xml.WriteEmptyTag('autoFilter', true, false);
    end;

    //Merge Cells
    if sheet.MergeCells.Count > 0 then begin
      xml.Attributes.Clear();
      xml.Attributes.Add('count', IntToStr(sheet.MergeCells.Count));
      xml.WriteTagNode('mergeCells', true, true, false);
      for i := 0 to sheet.MergeCells.Count - 1 do begin
        xml.Attributes.Clear();
        _r := sheet.MergeCells.Items[i];
        s := TZEFormula.GetColAddres(_r.Left) + IntToStr(_r.Top + 1) + ':' + TZEFormula.GetColAddres(_r.Right) + IntToStr(_r.Bottom + 1);
        xml.Attributes.Add('ref', s);
        xml.WriteEmptyTag('mergeCell', true, false);
      end;
      xml.WriteEndTagNode(); //mergeCells
    end; //if

    WriteHelper.WriteHyperLinksTag(xml);
  end; //WriteXLSXSheetData

  procedure WriteColontituls();
  begin
    xml.Attributes.Clear;
    if sheet.SheetOptions.IsDifferentOddEven then
      xml.Attributes.Add('differentOddEven', '1');
    if sheet.SheetOptions.IsDifferentFirst then
      xml.Attributes.Add('differentFirst', '1');
    xml.WriteTagNode('headerFooter', true, true, false);

    xml.Attributes.Clear;
    xml.WriteTag('oddHeader', sheet.SheetOptions.Header, true, false, true);
    xml.WriteTag('oddFooter', sheet.SheetOptions.Footer, true, false, true);

    if sheet.SheetOptions.IsDifferentOddEven then begin
      xml.WriteTag('evenHeader', sheet.SheetOptions.EvenHeader, true, false, true);
      xml.WriteTag('evenFooter', sheet.SheetOptions.EvenFooter, true, false, true);
    end;
    if sheet.SheetOptions.IsDifferentFirst then begin
      xml.WriteTag('firstHeader', sheet.SheetOptions.FirstPageHeader, true, false, true);
      xml.WriteTag('firstFooter', sheet.SheetOptions.FirstPageFooter, true, false, true);
    end;

    xml.WriteEndTagNode(); //headerFooter
  end;

  procedure WriteBreakData(tagName: string; breaks: TArray<Integer>; manV, maxV: string);
  var brk: Integer;
  begin
    if Length(breaks) > 0 then begin
      xml.Attributes.Clear();
      xml.Attributes.Add('count', IntToStr(Length(breaks)));
      xml.Attributes.Add('manualBreakCount', IntToStr(Length(breaks)));
      xml.WriteTagNode(tagName, true, true, true);
      for brk in breaks do begin
        xml.Attributes.Clear();
        xml.Attributes.Add('id', IntToStr(brk));
        xml.Attributes.Add('man', manV);
        xml.Attributes.Add('max', maxV);
        xml.WriteEmptyTag('brk', true, false);
      end;
      xml.WriteEndTagNode(); //(row|col)Breaks
    end;
  end;

  procedure WriteXLSXSheetFooter();
  var s: string;
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('headings', 'false', false);
    xml.Attributes.Add('gridLines', 'false', false);
    xml.Attributes.Add('gridLinesSet', 'true', false);
    xml.Attributes.Add('horizontalCentered', XLSXBoolToStr(sheet.SheetOptions.CenterHorizontal), false);
    xml.Attributes.Add('verticalCentered', XLSXBoolToStr(sheet.SheetOptions.CenterVertical), false);
    xml.WriteEmptyTag('printOptions', true, false);

    xml.Attributes.Clear();
    s := '0.##########';
    xml.Attributes.Add('left',   ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginLeft / ZE_MMinInch)),   false);
    xml.Attributes.Add('right',  ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginRight / ZE_MMinInch)),  false);
    xml.Attributes.Add('top',    ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginTop / ZE_MMinInch)),    false);
    xml.Attributes.Add('bottom', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.MarginBottom / ZE_MMinInch)), false);
    xml.Attributes.Add('header', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.HeaderMargins.Height / ZE_MMinInch)), false);
    xml.Attributes.Add('footer', ZEFloatSeparator(FormatFloat(s, sheet.SheetOptions.FooterMargins.Height / ZE_MMinInch)), false);
    xml.WriteEmptyTag('pageMargins', true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('blackAndWhite', 'false', false);
    xml.Attributes.Add('cellComments', 'none', false);
    xml.Attributes.Add('copies', '1', false);
    xml.Attributes.Add('draft', 'false', false);
    xml.Attributes.Add('firstPageNumber', '1', false);
    if sheet.SheetOptions.FitToHeight >= 0 then
      xml.Attributes.Add('fitToHeight', intToStr(sheet.SheetOptions.FitToHeight), false);

    if sheet.SheetOptions.FitToWidth >= 0 then
      xml.Attributes.Add('fitToWidth', IntToStr(sheet.SheetOptions.FitToWidth), false);

    xml.Attributes.Add('horizontalDpi', '300', false);

    // ECMA 376 ed.4 part1 18.18.50: default|portrait|landscape
    xml.Attributes.Add('orientation',
        IfThen(sheet.SheetOptions.PortraitOrientation, 'portrait', 'landscape'),
        false);

    if not sheet.SheetOptions.PageOrderDown then
      xml.Attributes.Add('pageOrder', 'overThenDown', false);
    xml.Attributes.Add('paperSize', intToStr(sheet.SheetOptions.PaperSize), false);
    if (sheet.SheetOptions.FitToWidth=-1)and(sheet.SheetOptions.FitToWidth=-1) then
      xml.Attributes.Add('scale', IntToStr(sheet.SheetOptions.ScaleToPercent), false);
    xml.Attributes.Add('useFirstPageNumber', 'true', false);
    //xml.Attributes.Add('usePrinterDefaults', 'false', false); //do not use!
    xml.Attributes.Add('verticalDpi', '300', false);
    xml.WriteEmptyTag('pageSetup', true, false);

    WriteColontituls();

    //  <legacyDrawing r:id="..."/>

    // write (row|col)Breaks
    WriteBreakData('rowBreaks', sheet.RowBreaks, '1', '16383');
    WriteBreakData('colBreaks', sheet.ColBreaks, '1', '1048575');
  end; //WriteXLSXSheetFooter

  procedure WriteXLSXSheetDrawings();
  var rId: Integer;
  begin
    // drawings
    if (not sheet.Drawing.IsEmpty) then begin
      // rels to helper
      rId := WriteHelper.AddDrawing('../drawings/drawing' + IntToStr(sheet.SheetIndex + 1) + '.xml');
      xml.Attributes.Clear();
      xml.Attributes.Add('r:id', 'rId' + IntToStr(rId));
      xml.WriteEmptyTag('drawing');
    end;
  end;
begin
  WriteHelper.Clear();
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL);
    xml.WriteTagNode('worksheet', true, true, false);

    sheet := XMLSS.Sheets[SheetNum];
    WriteXLSXSheetHeader();
    if (sheet.ColCount > 0) then
      WriteXLSXSheetCols();
    WriteXLSXSheetData();
    WriteXLSXSheetFooter();
    WriteXLSXSheetDrawings();
    xml.WriteEndTagNode(); //worksheet
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateSheet

//Создаёт workbook.xml
//INPUT
//  var XMLSS: TZWorkBook               - хранилище
//    Stream: TStream                   - поток для записи
//  const _pages: TIntegerDynArray      - массив страниц
//  const _names: TStringDynArray       - массив имён страниц
//    PageCount: integer                - кол-во страниц
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateWorkBook(var XMLSS: TZWorkBook; Stream: TStream;
    const _pages: TIntegerDynArray; const _names: TStringDynArray;
    PageCount: integer; TextConverter: TAnsiToCPConverter; CodePageName: String;
    BOM: ansistring): integer;
var xml: TZsspXMLWriterH; i,id: integer;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';

    xml.WriteHeader(CodePageName, BOM);

    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL, false);
    xml.WriteTagNode('workbook', true, true, true);

    xml.Attributes.Clear();
    xml.Attributes.Add('appName', 'ZEXMLSSlib');
    xml.WriteEmptyTag('fileVersion', true);

    xml.Attributes.Clear();
    xml.Attributes.Add('backupFile', 'false');
    xml.Attributes.Add('showObjects', 'all', false);
    xml.Attributes.Add('date1904', 'false', false);
    xml.WriteEmptyTag('workbookPr', true);

    xml.Attributes.Clear();
    xml.WriteEmptyTag('workbookProtection', true);

    xml.WriteTagNode('bookViews', true, true, true);

    id := 0;
    for i := 0 to PageCount - 1 do begin
      if XMLSS.Sheets[_pages[i]].Visible = svVisible then begin
        id := i;
        break;
      end;
    end;
    xml.Attributes.Add('activeTab', IntToStr(id));
    xml.Attributes.Add('firstSheet', IntToStr(id), false);
    xml.Attributes.Add('showHorizontalScroll', 'true', false);
    xml.Attributes.Add('showSheetTabs', 'true', false);
    xml.Attributes.Add('showVerticalScroll', 'true', false);
    xml.Attributes.Add('tabRatio', '600', false);
    xml.Attributes.Add('windowHeight', '8192', false);
    xml.Attributes.Add('windowWidth', '16384', false);
    xml.Attributes.Add('xWindow', '0', false);
    xml.Attributes.Add('yWindow', '0', false);
    xml.WriteEmptyTag('workbookView', true);
    xml.WriteEndTagNode(); // bookViews

    // sheets
    xml.Attributes.clear();
    xml.WriteTagNode('sheets', true, true, true);
    for i := 0 to PageCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('name', _names[i], false);
      xml.Attributes.Add('sheetId', IntToStr(i + 1), false);
      case XMLSS.Sheets[_pages[i]].Visible of
        svHidden: xml.Attributes.Add('state', 'hidden', false);
        svVeryHidden: xml.Attributes.Add('state', 'veryhidden', false);
      else
        xml.Attributes.Add('state', 'visible', false);
      end;
      xml.Attributes.Add('r:id', 'rId' + IntToStr(i + 2), false);
      xml.WriteEmptyTag('sheet', true);
    end; //for i
    xml.WriteEndTagNode(); //sheets

    // definedNames
    xml.Attributes.clear();
    xml.WriteTagNode('definedNames', true, true, true);
    for i := 0 to High(XMLSS.FDefinedNames) do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('localSheetId', IntToStr(XMLSS.FDefinedNames[i].LocalSheetId), false);
      xml.Attributes.Add('name', XMLSS.FDefinedNames[i].Name, false);
      xml.WriteTag('definedName', XMLSS.FDefinedNames[i].Body);
    end; //for i
    xml.WriteEndTagNode(); //definedNames

    xml.Attributes.Clear();
    xml.Attributes.Add('iterateCount', '100');
    xml.Attributes.Add('refMode', 'A1', false); //{tut}
    xml.Attributes.Add('iterate', 'false', false);
    xml.Attributes.Add('iterateDelta', '0.001', false);
    xml.WriteEmptyTag('calcPr', true);

    xml.WriteEndTagNode(); //workbook
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateWorkBook

//Создаёт styles.xml
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateStyles(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var
  xml: TZsspXMLWriterH;        //писатель
  _FontIndex: TIntegerDynArray;  //соответствия шрифтов
  _FillIndex: TIntegerDynArray;  //заливки
  _BorderIndex: TIntegerDynArray;//границы
  _StylesCount: integer;
  _NumFmtIndexes: array of integer;
  _FmtParser: TNumFormatParser;
  _DateParser: TZDateTimeODSFormatParser;

  // <numFmts> .. </numFmts>
  procedure WritenumFmts();
  var kol: integer;
    i: integer;
    _nfmt: TZEXLSXNumberFormats;
    _is_dateTime: array of boolean;
    s: string;
    _count: integer;
    _idx: array of integer;
    _fmt: array of string;
    _style: TZStyle;
    _currSheet: integer;
    _currRow, _currCol: integer;
    _sheet: TZSheet;
    _currstylenum: integer;
    _numfmt_counter: integer;

    function _GetNumFmt(StyleNum: integer): integer;
    var i, j, k: integer; b: boolean;
      _cs, _cr, _cc: integer;
    begin
      Result := 0;
      _style := XMLSS.Styles[StyleNum];
      if (_style.NumberFormatId > 0) and (_style.NumberFormatId < 164) then
        Exit(_style.NumberFormatId);

      //If cell type is datetime and cell style is empty then need write default NumFmtId = 14.
      if ((Trim(_style.NumberFormat) = '') or (UpperCase(_style.NumberFormat) = 'GENERAL')) then begin
        if (_is_dateTime[StyleNum + 1]) then
          Result := 14
        else begin
          b := false;
          _cs := _currSheet;
          for i := _cs to XMLSS.Sheets.Count - 1 do begin
            _sheet := XMLSS.Sheets[i];
            _cr := _currRow;
            for j := _cr to _sheet.RowCount - 1 do begin
              _cc := _currCol;
              for k := _cc to _sheet.ColCount - 1 do begin
                _currstylenum := _sheet[k, j].CellStyle + 1;
                if (_currstylenum >= 0) and (_currstylenum < kol) then
                  if (_sheet[k, j].CellType = ZEDateTime) then begin
                    _is_dateTime[_currstylenum] := true;
                    if (_currstylenum = StyleNum + 1) then begin
                      b := true;
                      break;
                    end;
                  end;
              end; //for k
              _currRow := j + 1;
              _currCol := 0;
              if (b) then
                break;
            end; //for j

            _currSheet := i + 1;
            _currRow := 0;
            _currCol := 0;
            if (b) then
              break;
          end; //for i

          if (b) then
            Result := 14;
        end;
      end //if
      else begin
        s := ConvertFormatNativeToXlsx(_style.NumberFormat, _FmtParser, _DateParser);
        i := _nfmt.FindFormatID(s);
        if (i < 0) then begin
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
    end; //_GetNumFmt

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

      if (_count > 0) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('count', IntToStr(_count));
        xml.WriteTagNode('numFmts', true, true, false);

        for i := 0 to _count - 1 do begin
          xml.Attributes.Clear();
          xml.Attributes.Add('numFmtId', IntToStr(_idx[i]));
          xml.Attributes.Add('formatCode', _fmt[i]);
          xml.WriteEmptyTag('numFmt', true, true);
        end;

        xml.WriteEndTagNode(); //numFmts
      end;
    finally
      FreeAndNil(_nfmt);
      SetLength(_idx, 0);
      SetLength(_fmt, 0);
      SetLength(_is_dateTime, 0);
    end;
  end; //WritenumFmts

  //Являются ли шрифты стилей одинаковыми
  function _isFontsEqual(const stl1, stl2: TZStyle): boolean;
  begin
    result := False;
    if (stl1.Font.Color <> stl2.Font.Color) then
      exit;

    if (stl1.Font.Name <> stl2.Font.Name) then
      exit;

    if (stl1.Font.Size <> stl2.Font.Size) then
      exit;

    if (stl1.Font.Style <> stl2.Font.Style) then
      exit;

    if stl1.Superscript <> stl2.Superscript then
      exit;

    if stl1.Subscript <> stl2.Subscript then
      exit;

    Result := true; // если уж до сюда добрались
  end; //_isFontsEqual

  //Обновляет индексы в массиве
  //INPUT
  //  var arr: TIntegerDynArray  - массив
  //      cnt: integer          - номер последнего элемента в массиве (начинает с 0)
  //                              (предполагается, что возникнет ситуация, когда нужно будет использовать только часть массива)
  procedure _UpdateArrayIndex(var arr: TIntegerDynArray; cnt: integer);
  var res: TIntegerDynArray;
    i, j: integer;
    num: integer;
  begin
    //Assert( Length(arr) - cnt = 2, 'Wow! We really may need this parameter!');
    //cnt := Length(arr) - 2;   // get ready to strip it
    SetLength(res, Length(arr));

    num := 0;
    for i := -1 to cnt do
    if (arr[i + 1] = -2) then begin
      res[i + 1] := num;
      for j := i + 1 to cnt do
      if (arr[j + 1] = i) then
        res[j + 1] := num;
      inc(num);
    end; //if

    arr := res;
  end; //_UpdateArrayIndex

  //<fonts>...</fonts>
  procedure WriteXLSXFonts();
  var i, j, n: integer;
    _fontCount: integer;
    fnt: TZFont;
  begin
    _fontCount := 0;
    SetLength(_FontIndex, _StylesCount + 1);
    for i := 0 to _StylesCount do
      _FontIndex[i] := -2;

    for i := -1 to _StylesCount - 1 do
    if (_FontIndex[i + 1] = -2) then begin
      inc (_fontCount);
      n := i + 1;
      for j := n to _StylesCount - 1 do
      if (_FontIndex[j + 1] = -2) then
        if (_isFontsEqual(XMLSS.Styles[i], XMLSS.Styles[j])) then
          _FontIndex[j + 1] := i;
    end; //if

    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(_fontCount));
    xml.WriteTagNode('fonts', true, true, true);

    for i := 0 to _StylesCount do
    if (_FontIndex[i] = - 2) then begin
      fnt := XMLSS.Styles[i - 1].Font;
      xml.Attributes.Clear();
      xml.WriteTagNode('font', true, true, true);

      if (fsBold in fnt.Style) then begin
        xml.Attributes.Clear();
        //xml.Attributes.Add('val', 'true');
        xml.WriteEmptyTag('b', true);
      end;

      if (fsItalic in fnt.Style) then begin
        xml.Attributes.Clear();
        //xml.Attributes.Add('val', 'true');
        xml.WriteEmptyTag('i', true);
      end;

      if (fsStrikeOut in fnt.Style) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('val', 'true');
        xml.WriteEmptyTag('strike', true);
      end;

      if (fsUnderline in fnt.Style) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('val', 'single');
        xml.WriteEmptyTag('u', true);
      end;

      xml.Attributes.Clear();
      xml.Attributes.Add('val', FloatToStr(fnt.Size, TFormatSettings.Invariant));
      xml.WriteEmptyTag('sz', true);

      if (fnt.Color <> clWindowText) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(fnt.Color));
        xml.WriteEmptyTag('color', true);
      end;

      xml.Attributes.Clear();
      xml.Attributes.Add('val', fnt.Name);
      xml.WriteEmptyTag('name', true);

      xml.Attributes.Clear();
      xml.Attributes.Add('val', IntToStr(fnt.Charset));
      xml.WriteEmptyTag('charset', true);

      //<vertAlign val="superscript"/>
      if XMLSS.Styles[i - 1].Superscript then begin
        xml.Attributes.Clear();

        xml.Attributes.Add('val', 'superscript');
        xml.WriteEmptyTag('vertAlign', true);
      end

      //<vertAlign val="subscript"/>
      else if XMLSS.Styles[i - 1].Subscript then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('val', 'subscript');
        xml.WriteEmptyTag('vertAlign', true);
      end;

      xml.WriteEndTagNode(); //font
    end; //if

    _UpdateArrayIndex(_FontIndex, _StylesCount - 1);

    xml.WriteEndTagNode(); //fonts
  end; //WriteXLSXFonts

  //Являются ли заливки одинаковыми
  function _isFillsEqual(style1, style2: TZStyle): boolean;
  begin
    result := (style1.BGColor = style2.BGColor) and
              (style1.PatternColor = style2.PatternColor) and
              (style1.CellPattern = style2.CellPattern);
  end; //_isFillsEqual

  procedure _WriteBlankFill(const st: string);
  begin
    xml.Attributes.Clear();
    xml.WriteTagNode('fill', true, true, true);
    xml.Attributes.Clear();
    xml.Attributes.Add('patternType', st);
    xml.WriteEmptyTag('patternFill', true, false);
    xml.WriteEndTagNode(); //fill
  end; //_WriteBlankFill

  //<fills> ... </fills>
  procedure WriteXLSXFills();
  var
    i, j: integer;
    _fillCount: integer;
    s: string;
    b: boolean;
    _tmpColor: TColor;
    _reverse: boolean;

  begin
    _fillCount := 0;
    SetLength(_FillIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _FillIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
    if (_FillIndex[i + 1] = - 2) then begin
      inc(_fillCount);
      for j := i + 1 to _StylesCount - 1 do
      if (_FillIndex[j + 1] = -2) then
        if (_isFillsEqual(XMLSS.Styles[i], XMLSS.Styles[j])) then
          _FillIndex[j + 1] := i;
    end; //if

    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(_fillCount + 2));
    xml.WriteTagNode('fills', true, true, true);

    //по какой-то непонятной причине, если в начале нету двух стилей заливок (none + gray125),
    //в грёбаном 2010-ом офисе глючат заливки (то-ли чтобы сложно было сделать экспорт в xlsx, то-ли
    //кривые руки у мелкомягких программеров). LibreOffice открывает нормально.
    _WriteBlankFill('none');
    _WriteBlankFill('gray125');

    //TODO:
    //ВНИМАНИЕ!!! //{tut}
    //в некоторых случаях fgColor - это цвет заливки (вроде для solid), а в некоторых - bgColor.
    //Потом не забыть разобраться.
    for i := -1 to _StylesCount - 1 do
    if (_FillIndex[i + 1] = -2) then begin
      xml.Attributes.Clear();
      xml.WriteTagNode('fill', true, true, true);

      case XMLSS.Styles[i].CellPattern of
        ZPSolid:                  s := 'solid';
        ZPNone:                   s := 'none';
        ZPGray125:                s := 'gray125';
        ZPGray0625:               s := 'gray0625';
        ZPDiagStripe:             s := 'darkUp';
        ZPGray50:                 s := 'mediumGray';
        ZPGray75:                 s := 'darkGray';
        ZPGray25:                 s := 'lightGray';
        ZPHorzStripe:             s := 'darkHorizontal';
        ZPVertStripe:             s := 'darkVertical';
        ZPReverseDiagStripe:      s := 'darkDown';
        ZPDiagCross:              s := 'darkGrid';
        ZPThickDiagCross:         s := 'darkTrellis';
        ZPThinHorzStripe:         s := 'lightHorizontal';
        ZPThinVertStripe:         s := 'lightVertical';
        ZPThinReverseDiagStripe:  s := 'lightDown';
        ZPThinDiagStripe:         s := 'lightUp';
        ZPThinHorzCross:          s := 'lightGrid';
        ZPThinDiagCross:          s := 'lightTrellis';
        else
          s := 'solid';
      end; //case

      b := (XMLSS.Styles[i].PatternColor <> clWindow) or (XMLSS.Styles[i].BGColor <> clWindow);
      xml.Attributes.Clear();
      if b and (XMLSS.Styles[i].CellPattern = ZPNone) then
        xml.Attributes.Add('patternType', 'solid')
      else
        xml.Attributes.Add('patternType', s);

      if (b) then
        xml.WriteTagNode('patternFill', true, true, false)
      else
        xml.WriteEmptyTag('patternFill', true, false);

      _reverse := not (XMLSS.Styles[i].CellPattern in [ZPNone, ZPSolid]);

      if (XMLSS.Styles[i].BGColor <> clWindow) then
      begin
        xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := XMLSS.Styles[i].PatternColor
        else
          _tmpColor := XMLSS.Styles[i].BGColor;
        xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(_tmpColor));
        xml.WriteEmptyTag('fgColor', true);
      end;

      if (XMLSS.Styles[i].PatternColor <> clWindow) then
      begin
        xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := XMLSS.Styles[i].BGColor
        else
          _tmpColor := XMLSS.Styles[i].PatternColor;
        xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(_tmpColor));
        xml.WriteEmptyTag('bgColor', true);
      end;

      if (b) then
        xml.WriteEndTagNode(); //patternFill

      xml.WriteEndTagNode(); //fill
    end; //if

    _UpdateArrayIndex(_FillIndex, _StylesCount - 1);

    xml.WriteEndTagNode(); //fills
  end; //WriteXLSXFills();

  //единичная граница
  procedure _WriteBorderItem(StyleNum: integer; BorderNum: TZBordersPos);
  var s, s1: string;
    _border: TZBorderStyle;
    n: integer;
  begin
    xml.Attributes.Clear();
    case BorderNum of
      bpLeft:   s := 'left';
      bpTop:    s := 'top';
      bpRight:  s := 'right';
      bpBottom: s := 'bottom';
      else
        s := 'diagonal';
    end;
    _border := XMLSS.Styles[StyleNum].Border[BorderNum];
    s1 := '';
    case _border.LineStyle of
      ZEContinuous:
        begin
          if (_border.Weight <= 1) then
            s1 := 'thin'
          else
          if (_border.Weight = 2) then
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
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDashed';
        end;
      ZEDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dotted'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDotted';
        end;
      ZEDashDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDot'
          else
          if (_border.Weight >= 2) then
            s1 := 'mediumDashDot';
        end;
      ZEDashDotDot:
        begin
          if (_border.Weight = 1) then
            s1 := 'dashDotDot'
          else
          if (_border.Weight >= 2) then
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
    end; //case

    n := length(s1);

    if (n > 0) then
      xml.Attributes.Add('style', s1);

    if ((_border.Color <> clBlack) and (n > 0)) then begin
      xml.WriteTagNode(s, true, true, true);
      xml.Attributes.Clear();
      xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(_border.Color));
      xml.WriteEmptyTag('color', true);
      xml.WriteEndTagNode();
    end else
      xml.WriteEmptyTag(s, true);
  end; //_WriteBorderItem

  //<borders> ... </borders>
  procedure WriteXLSXBorders();
  var  i, j: integer;
    _borderCount: integer;
    s: string;
  begin
    _borderCount := 0;
    SetLength(_BorderIndex, _StylesCount + 1);
    for i := -1 to _StylesCount - 1 do
      _BorderIndex[i + 1] := -2;
    for i := -1 to _StylesCount - 1 do
    if (_BorderIndex[i + 1] = - 2) then begin
      inc(_borderCount);
      for j := i + 1 to _StylesCount - 1 do
      if (_BorderIndex[j + 1] = -2) then
        if (XMLSS.Styles[i].Border.isEqual(XMLSS.Styles[j].Border)) then
          _BorderIndex[j + 1] := i;
    end; //if

    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(_borderCount));
    xml.WriteTagNode('borders', true, true, true);

    for i := -1 to _StylesCount - 1 do
    if (_BorderIndex[i + 1] = -2) then begin
      xml.Attributes.Clear();
      s := 'false';
      if (XMLSS.Styles[i].Border[bpDiagonalLeft].Weight > 0) and (XMLSS.Styles[i].Border[bpDiagonalLeft].LineStyle <> ZENone) then
        s := 'true';
      xml.Attributes.Add('diagonalDown', s);
      s := 'false';
      if (XMLSS.Styles[i].Border[bpDiagonalRight].Weight > 0) and (XMLSS.Styles[i].Border[bpDiagonalRight].LineStyle <> ZENone) then
        s := 'true';
      xml.Attributes.Add('diagonalUp', s, false);
      xml.WriteTagNode('border', true, true, true);

      _WriteBorderItem(i, bpLeft);
      _WriteBorderItem(i, bpRight);
      _WriteBorderItem(i, bpTop);
      _WriteBorderItem(i, bpBottom);
      _WriteBorderItem(i, bpDiagonalLeft);
      //_WriteBorderItem(i, bpDiagonalRight);
      xml.WriteEndTagNode(); //border
    end; //if

    _UpdateArrayIndex(_BorderIndex, _StylesCount - 1);

    xml.WriteEndTagNode(); //borders
  end; //WriteXLSXBorders

  //Добавляет <xf> ... </xf>
  //INPUT
  //      NumStyle: integer - номер стиля
  //      isxfId: boolean   - нужно ли добавлять атрибут "xfId"
  //      xfId: integer     - значение "xfId"
  procedure _WriteXF(NumStyle: integer; isxfId: boolean; xfId: integer);
  var _addalignment: boolean;
    _style: TZStyle;
    s: string;
    i: integer;
    _num: integer;
  begin
    xml.Attributes.Clear();
    _style := XMLSS.Styles[NumStyle];
    _addalignment := _style.Alignment.WrapText or
                     _style.Alignment.VerticalText or
                    (_style.Alignment.Rotate <> 0) or
                    (_style.Alignment.Indent <> 0) or
                    _style.Alignment.ShrinkToFit or
                    (_style.Alignment.Vertical <> ZVAutomatic) or
                    (_style.Alignment.Horizontal <> ZHAutomatic);

    xml.Attributes.Add('applyAlignment', XLSXBoolToStr(_addalignment));
    xml.Attributes.Add('applyBorder', 'true', false);
    xml.Attributes.Add('applyFont', 'true', false);
    xml.Attributes.Add('applyProtection', 'true', false);
    xml.Attributes.Add('borderId', IntToStr(_BorderIndex[NumStyle + 1]), false);
    xml.Attributes.Add('fillId', IntToStr(_FillIndex[NumStyle + 1] + 2), false); //+2 т.к. первыми всегда идут 2 левых стиля заливки
    xml.Attributes.Add('fontId', IntToStr(_FontIndex[NumStyle + 1]), false);

    // ECMA 376 Ed.4:  12.3.20 Styles Part; 17.9.17 numFmt (Numbering Format); 18.8.30 numFmt (Number Format)
    // http://social.msdn.microsoft.com/Forums/sa/oxmlsdk/thread/3919af8c-644b-4d56-be65-c5e1402bfcb6
    if (isxfId) then
      _num := _NumFmtIndexes[NumStyle + 1]
    else
      _num := 0;

    xml.Attributes.Add('numFmtId', IntToStr(_num) {'164'}, false); // TODO: support formats

    if (_num > 0) then
      xml.Attributes.Add('applyNumberFormat', '1', false);

    if (isxfId) then
      xml.Attributes.Add('xfId', IntToStr(xfId), false);

    xml.WriteTagNode('xf', true, true, true);

    if (_addalignment) then
    begin
      xml.Attributes.Clear();
      case (_style.Alignment.Horizontal) of
        ZHLeft:        s := 'left';
        ZHRight:       s := 'right';
        ZHCenter:      s := 'center';
        ZHFill:        s := 'fill';
        ZHJustify:     s := 'justify';
        ZHDistributed: s := 'distributed';
        ZHAutomatic:   s := 'general';
        else
          s := 'general';
        // The standard does not specify a default value for the horizontal attribute.
        // Excel uses a default value of general for this attribute.
        // MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.d
      end; //case
      xml.Attributes.Add('horizontal', s);
      xml.Attributes.Add('indent',      IntToStr(_style.Alignment.Indent), false);
      xml.Attributes.Add('shrinkToFit', XLSXBoolToStr(_style.Alignment.ShrinkToFit), false);


      if _style.Alignment.VerticalText then i := 255
         else i := ZENormalizeAngle180(_style.Alignment.Rotate);
      xml.Attributes.Add('textRotation', IntToStr(i), false);

      case (_style.Alignment.Vertical) of
        ZVCenter:      s := 'center';
        ZVTop:         s := 'top';
        ZVBottom:      s := 'bottom';
        ZVJustify:     s := 'justify';
        ZVDistributed: s := 'distributed';
        else
          s := 'bottom';
        // The standard does not specify a default value for the vertical attribute.
        // Excel uses a default value of bottom for this attribute.
        // MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500, 18.8.1.e
      end; //case
      xml.Attributes.Add('vertical', s, false);
      xml.Attributes.Add('wrapText', XLSXBoolToStr(_style.Alignment.WrapText), false);
      xml.WriteEmptyTag('alignment', true);
    end; //if (_addalignment)

    xml.Attributes.Clear();
    xml.Attributes.Add('hidden', XLSXBoolToStr(XMLSS.Styles[NumStyle].Protect));
    xml.Attributes.Add('locked', XLSXBoolToStr(XMLSS.Styles[NumStyle].HideFormula));
    xml.WriteEmptyTag('protection', true);

    xml.WriteEndTagNode(); //xf
  end; //_WriteXF

  //<cellStyleXfs> ... </cellStyleXfs> / <cellXfs> ... </cellXfs>
  procedure WriteCellStyleXfs(const TagName: string; isxfId: boolean);
  var i: integer;
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(XMLSS.Styles.Count + 1));
    xml.WriteTagNode(TagName, true, true, true);
    for i := -1 to XMLSS.Styles.Count - 1 do  begin
      //Что-то не совсем понятно, какой именно xfId нужно ставить. Пока будет 0 для всех.
      _WriteXF(i, isxfId, 0{i + 1});
    end;
    xml.WriteEndTagNode(); //cellStyleXfs
  end; //WriteCellStyleXfs

  //<cellStyles> ... </cellStyles>
  procedure WriteCellStyles();
  begin
  end; //WriteCellStyles

begin
  result := 0;
  _FmtParser := TNumFormatParser.Create();
  _DateParser := TZDateTimeODSFormatParser.Create();
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';

    xml.WriteHeader(CodePageName, BOM);
    _StylesCount := XMLSS.Styles.Count;

    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    xml.WriteTagNode('styleSheet', true, true, true);

    WritenumFmts();

    WriteXLSXFonts();
    WriteXLSXFills();
    WriteXLSXBorders();
    //DO NOT remove cellStyleXfs!!!
    WriteCellStyleXfs('cellStyleXfs', false);
    WriteCellStyleXfs('cellXfs', true);
    WriteCellStyles(); //??

    xml.WriteEndTagNode(); //styleSheet
  finally
    xml.Free();
    _FmtParser.Free();
    _DateParser.Free();
    SetLength(_FontIndex, 0);
    SetLength(_FillIndex, 0);
    SetLength(_BorderIndex, 0);
    SetLength(_NumFmtIndexes, 0);
  end;
end; //ZEXLSXCreateStyles

//Добавить Relationship для rels
//INPUT
//      xml: TZsspXMLWriterH  - писалка
//  const rid: string         - rid
//      ridType: integer      - rIdType (0..8)
//  const Target: string      -
//  const TargetMode: string  -
procedure ZEAddRelsRelation(xml: TZsspXMLWriterH; const rid: string; ridType: TRelationType; const Target: string; const TargetMode: string = '');
begin
  xml.Attributes.Clear();
  xml.Attributes.Add('Id', rid);
  xml.Attributes.Add('Type',  ZEXLSXGetRelationName(ridType), false);
  xml.Attributes.Add('Target', Target, false);
  if (TargetMode <> '') then
     xml.Attributes.Add('TargetMode', TargetMode, true);
  xml.WriteEmptyTag('Relationship', true, true);
end; //ZEAddRelsID

//Создаёт _rels/.rels
//INPUT
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateRelsMain(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(xml, 'rId1', TRelationType.rtDoc,      'xl/workbook.xml');
    ZEAddRelsRelation(xml, 'rId2', TRelationType.rtExtProps, 'docProps/app.xml');
    ZEAddRelsRelation(xml, 'rId3', TRelationType.rtCoreProp, 'docProps/core.xml');

    xml.WriteEndTagNode(); //Relationships
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateRelsMain

//Создаёт xl/_rels/workbook.xml.rels
//INPUT
//    PageCount: integer                - кол-во страниц
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateRelsWorkBook(PageCount: integer; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH; i: integer;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    xml.WriteTagNode('Relationships', true, true, false);

    ZEAddRelsRelation(xml, 'rId1', TRelationType.rtStyles, 'styles.xml');

    for i := 0 to PageCount - 1 do
      ZEAddRelsRelation(xml, 'rId' + IntToStr(i + 2), TRelationType.rtWorkSheet, 'worksheets/sheet' + IntToStr(i + 1) + '.xml');

    ZEAddRelsRelation(xml, 'rId' + IntToStr(PageCount + 2), TRelationType.rtSharedStr, 'sharedStrings.xml');
    xml.WriteEndTagNode(); //Relationships
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateRelsWorkBook

//Создаёт sharedStrings.xml
//INPUT
//    XMLSS: TZWorkBook                     - хранилище
//    Stream: TStream                       - поток для записи
//    SharedStrings: TObjectList<TRichText> - общие строки
//    TextConverter: TAnsiToCPConverter     - конвертер из локальной кодировки в нужную
//    CodePageName: string                  - название кодовой страници
//    BOM: ansistring                       - BOM
//RETURN
//      integer
function ZEXLSXCreateSharedStrings(var XMLSS: TZWorkBook; Stream: TStream;
  const SharedStrings: TSharedStringBank; TextConverter: TAnsiToCPConverter;
  CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('count', SharedStrings.List.Count.ToString);
    xml.Attributes.Add('uniqueCount', SharedStrings.List.Count.ToString, false);
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN, false);
    xml.WriteTagNode('sst', true, true, false);

    for var sharedText in SharedStrings.List do begin
      xml.Attributes.Clear();
      xml.WriteTagNode('si', false, false, false);

      for var part in sharedText.List do begin
        xml.Attributes.Clear();
        if not Assigned(part.Font) then begin
          xml.Attributes.Clear();
          if part.Text.StartsWith(' ') or part.Text.EndsWith(' ') then
            xml.Attributes.Add('xml:space', 'preserve', false);
          xml.WriteTag('t', part.Text);
        end else begin
          xml.WriteTagNode('r', false, false, false);
          xml.WriteTagNode('rPr', false, false, false);

          if TFontStyle.fsBold in part.Font.Style then
            xml.WriteEmptyTag('b', true);
          if TFontStyle.fsItalic in part.Font.Style then
            xml.WriteEmptyTag('i', true);
          if TFontStyle.fsUnderline in part.Font.Style then
            xml.WriteEmptyTag('u', true);

          if part.Font.Size > 0 then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('val', FloatToStr(part.Font.Size), false);
            xml.WriteEmptyTag('sz', true);
          end;

          if (part.Font.ColorTheme > 0) or (part.Font.Color <> 0) then begin
            xml.Attributes.Clear();
            if part.Font.ColorTheme > 0 then
              xml.Attributes.Add('theme', IntToStr(part.Font.ColorTheme), false)
            else if part.Font.Color <> 0 then
              xml.Attributes.Add('rgb', ColorToHTMLHex(part.Font.Color), false);
            xml.WriteEmptyTag('color', true);
          end;

          if not trim(part.Font.Name).IsEmpty then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('val', part.Font.Name, false);
            xml.WriteEmptyTag('rFont', true);
          end;

          if part.Font.Family > 0 then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('val', IntToStr(part.Font.Family), false);
            xml.WriteEmptyTag('family', true);
          end;

          if part.Font.Charset > 0 then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('val', IntToStr(part.Font.Charset), false);
            xml.WriteEmptyTag('charset', true);
          end;

          if not trim(part.Font.Scheme).IsEmpty then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('val', part.Font.Scheme, false);
            xml.WriteEmptyTag('scheme', true);
          end;

          xml.WriteEndTagNode();// rPr

          xml.Attributes.Clear();
          if part.Text.StartsWith(' ') or part.Text.EndsWith(' ') then
            xml.Attributes.Add('xml:space', 'preserve', false);
          xml.WriteTag('t', part.Text);

          xml.WriteEndTagNode(); // r
        end;
      end;

      xml.WriteEndTagNode(); // si
    end;

    xml.WriteEndTagNode(); // sst
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateSharedStrings

//Создаёт app.xml
//INPUT
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateDocPropsApp(Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns',    SCHEMA_DOC + '/extended-properties');
    xml.Attributes.Add('xmlns:vt', SCHEMA_DOC + '/docPropsVTypes', false);
    xml.WriteTagNode('Properties', true, true, false);

    xml.Attributes.Clear();
    xml.WriteTag('TotalTime', '0', true, false, false);
    xml.WriteTag('Application', TZWorkBook.Application, true, false, true);
    xml.WriteEndTagNode(); //Properties
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateDocPropsApp

//Создаёт app.xml
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//    Stream: TStream                   - поток для записи
//    TextConverter: TAnsiToCPConverter - конвертер из локальной кодировки в нужную
//    CodePageName: string              - название кодовой страници
//    BOM: ansistring                   - BOM
//RETURN
//      integer
function ZEXLSXCreateDocPropsCore(var XMLSS: TZWorkBook; Stream: TStream; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring): integer;
var xml: TZsspXMLWriterH; creationDate: string;
begin
  result := 0;
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.TabLength := 1;
    xml.TextConverter := TextConverter;
    xml.TabSymbol := ' ';
    xml.WriteHeader(CodePageName, BOM);
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns:cp', SCHEMA_PACKAGE + '/metadata/core-properties');
    xml.Attributes.Add('xmlns:dc', 'http://purl.org/dc/elements/1.1/', false);
    xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    xml.Attributes.Add('xmlns:dcterms', 'http://purl.org/dc/terms/', false);
    xml.Attributes.Add('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance', false);
    xml.WriteTagNode('cp:coreProperties', true, true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    creationDate := ZEDateTimeToStr(XMLSS.DocumentProperties.Created) + 'Z';
    xml.WriteTag('dcterms:created', creationDate, true, false, false);
    xml.WriteTag('dcterms:modified', creationDate, true, false, false);

    xml.Attributes.Clear();
    xml.WriteTag('cp:revision', '1', true, false, false);

    xml.WriteEndTagNode(); //cp:coreProperties
  finally
    xml.Free();
  end;
end; //ZEXLSXCreateDocPropsCore

//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZWorkBook                 - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of integer; const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName: string; BOM: ansistring = ''): integer; overload;
var
  _pages: TIntegerDynArray;      //номера страниц
  _names: TStringDynArray;      //названия страниц
  kol, i{, ii}: integer;
  Stream: TStream;
  _WriteHelper: TZEXLSXWriteHelper;
  path_xl, path_sheets, path_relsmain, path_relsw, path_docprops: string;
  s: string;
  SharedStrings: TSharedStringBank;
  //iDrawingsCount: Integer;
  //path_draw, path_draw_rel, path_media: string;
  //_drawing: TZEDrawing;
  //_pic: TZEPicture;
begin
  Result := 0;
  Stream := nil;
  _WriteHelper := nil;
  kol := 0;
  SharedStrings := TSharedStringBank.Create();
  try
    if (not TDirectory.Exists(PathName)) then begin
      result := 3;
      exit;
    end;

    if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then begin
      result := 2;
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

    //iDrawingsCount := XMLSS.DrawingCount();
    // sheets of workbook
    for i := 0 to kol - 1 do begin
      Stream := TFileStream.Create(path_sheets + 'sheet' + IntToStr(i + 1) + '.xml', fmCreate);
      try
        ZEXLSXCreateSheet(XMLSS, Stream, _pages[i], SharedStrings, TextConverter, CodePageName, BOM, _WriteHelper);
      finally
        FreeAndNil(Stream);
      end;

      if (_WriteHelper.HyperLinksCount > 0) then begin
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
    end; //for i

    //iDrawingsCount := XMLSS.DrawingCount();
//    if iDrawingsCount <> 0 then begin
//      path_draw := path_xl + 'drawings' + PathDelim;
//      if (not DirectoryExists(path_draw)) then
//        ForceDirectories(path_draw);
//
//      path_draw_rel := path_draw + '_rels' + PathDelim;
//      if (not DirectoryExists(path_draw_rel)) then
//        ForceDirectories(path_draw_rel);
//
//      path_media := path_xl + 'media' + PathDelim;
//      if (not DirectoryExists(path_media)) then
//        ForceDirectories(path_media);
//
//      for i := 0 to iDrawingsCount - 1 do begin
//        _drawing := XMLSS.GetDrawing(i);
//        // drawings/drawingN.xml
//        Stream := TFileStream.Create(path_draw + 'drawing' + IntToStr(_drawing.Id) + '.xml', fmCreate);
//        try
//          ZEXLSXCreateDrawing(XMLSS, Stream, _drawing, TextConverter, CodePageName, BOM);
//        finally
//          FreeAndNil(Stream);
//        end;
//
//        // drawings/_rels/drawingN.xml.rels
//        Stream := TFileStream.Create(path_draw_rel + 'drawing' + IntToStr(i + 1) + '.xml.rels', fmCreate);
//        try
//          ZEXLSXCreateDrawingRels(XMLSS, Stream, _drawing, TextConverter, CodePageName, BOM);
//        finally
//          FreeAndNil(Stream);
//        end;
//
//        // media/imageN.png
//        for ii := 0 to _drawing.PictureStore.Count - 1 do begin
//          _pic := _drawing.PictureStore[ii];
//          if not Assigned(_pic.DataStream) then Continue;
//          Stream := TFileStream.Create(path_media + _pic.Name, fmCreate);
//          try
//            _pic.DataStream.Position := 0;
//            Stream.CopyFrom(_pic.DataStream, _pic.DataStream.Size);
//          finally
//            FreeAndNil(Stream);
//          end;
//        end;
//      end;
//    end;

    // sharedStrings.xml
    Stream := TFileStream.Create(path_xl + 'sharedStrings.xml', fmCreate);
    try
      ZEXLSXCreateSharedStrings(XMLSS, Stream, SharedStrings, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    //workbook.xml - list of shhets
    Stream := TFileStream.Create(path_xl + 'workbook.xml', fmCreate);
    try
      ZEXLSXCreateWorkBook(XMLSS, Stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
    finally
      FreeAndNil(Stream);
    end;

    //[Content_Types].xml
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
    SharedStrings.Free();
  end;
end; //SaveXmlssToXLSXPath

//SaveXmlssToXLSXPath
//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZWorkBook                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string; const SheetsNumbers: array of integer; const SheetsNames: array of string): integer; overload;
begin
  result := SaveXmlssToXLSXPath(XMLSS, PathName, SheetsNumbers, SheetsNames, nil, 'UTF-8', '');
end; //SaveXmlssToXLSXPath

//SaveXmlssToXLSXPath
//Сохраняет незапакованный документ в формате Office Open XML (OOXML)
//INPUT
//  var XMLSS: TZWorkBook                   - хранилище
//      PathName: string                  - путь к директории для сохранения (должна заканчиватся разделителем директории)
//RETURN
//      integer
function SaveXmlssToXLSXPath(var XMLSS: TZWorkBook; PathName: string): integer; overload;
begin
  result := SaveXmlssToXLSXPath(XMLSS, PathName, [], []);
end; //SaveXmlssToXLSXPath


//Сохраняет документ в формате Open Office XML (xlsx)
//INPUT
//  var XMLSS: TZWorkBook                   - хранилище
//      FileName: string                  - имя файла для сохранения
//  const SheetsNumbers:array of integer  - массив номеров страниц в нужной последовательности
//  const SheetsNames: array of string    - массив названий страниц
//                                          (количество элементов в двух массивах должны совпадать)
//      TextConverter: TAnsiToCPConverter - конвертер
//      CodePageName: string              - имя кодировки
//      BOM: ansistring                   - Byte Order Mark
//RETURN
//      integer
function SaveXmlssToXLSX(var XMLSS: TZWorkBook; zipStream: TStream; const SheetsNumbers: array of integer;
                         const SheetsNames: array of string; TextConverter: TAnsiToCPConverter; CodePageName:
                         string; BOM: ansistring = ''): integer;
var
  _pages: TIntegerDynArray; // numbers of sheets
  _names: TStringDynArray;  // names of sheets
  kol, i: integer;
  zip: TZipFile;
  stream: TStream;
  writeHelper: TZEXLSXWriteHelper;
  SharedStrings: TSharedStringBank;
begin
  Result := 0;
  zip := TZipFile.Create();
  try
    writeHelper := TZEXLSXWriteHelper.Create();
    try
      SharedStrings := TSharedStringBank.Create();
      try
        if (not ZECheckTablesTitle(XMLSS, SheetsNumbers, SheetsNames, _pages, _names, kol)) then
          exit(2);

        zip.Open(zipStream, zmReadWrite);

        // styles
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateStyles(XMLSS, stream, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'xl/styles.xml');
        finally
          stream.Free();
        end;

        // _rels/.rels
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateRelsMain(stream, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, '_rels/.rels');
        finally
          stream.Free();
        end;

        // xl/_rels/workbook.xml.rels
        kol := 0;
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateRelsWorkBook(kol, stream, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'xl/_rels/workbook.xml.rels');
        finally
          stream.Free();
        end;

        // sheets of workbook
        for i := 0 to kol - 1 do begin
          if XMLSS.Sheets[_pages[i]].RowCount > 60000 then begin
            stream := TTempFileStream.Create();
            try
              ZEXLSXCreateSheet(XMLSS, stream, _pages[i], SharedStrings, TextConverter, CodePageName, BOM, writeHelper);
              stream.Position := 0;
              zip.Add(stream, 'xl/worksheets/sheet' + IntToStr(i + 1) + '.xml');
            finally
              stream.Free();
            end;
          end
          else begin
            stream := TMemoryStream.Create();
            try
              ZEXLSXCreateSheet(XMLSS, stream, _pages[i], SharedStrings, TextConverter, CodePageName, BOM, writeHelper);
              stream.Position := 0;
              zip.Add(stream, 'xl/worksheets/sheet' + IntToStr(i + 1) + '.xml');
            finally
              stream.Free();
            end;
          end;

          if (writeHelper.HyperLinksCount > 0) then begin
            writeHelper.AddSheetHyperlink(i);
            stream := TMemoryStream.Create();
            try
              writeHelper.CreateSheetRels(stream, TextConverter, CodePageName, BOM);
              stream.Position := 0;
              zip.Add(stream, 'xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels');
            finally
              stream.Free();
            end;
          end;
        end; //for i

        // sharedStrings.xml
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateSharedStrings(XMLSS, stream, SharedStrings, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'xl/sharedStrings.xml');
        finally
          stream.Free();
        end;

        for i := 0 to XMLSS.Sheets.Count - 1 do begin
          if not XMLSS.Sheets[i].Drawing.IsEmpty then begin
            // drawings/drawingN.xml
            stream := TMemoryStream.Create();
            try
              ZEXLSXCreateDrawing(XMLSS.Sheets[i], stream, TextConverter, CodePageName, BOM);
              stream.Position := 0;
              zip.Add(stream, 'xl/drawings/drawing' + IntToStr(i+1) + '.xml');
            finally
              stream.Free();
            end;

            // drawings/_rels/drawingN.xml.rels
            stream := TMemoryStream.Create();
            try
              ZEXLSXCreateDrawingRels(XMLSS.Sheets[i], stream, TextConverter, CodePageName, BOM);
              stream.Position := 0;
              zip.Add(stream, 'xl/drawings/_rels/drawing' + IntToStr(i+1) + '.xml.rels');
            finally
              stream.Free();
            end;
          end;
        end;

        // media/imageN.png
        for I := 0 to High(XMLSS.MediaList) do begin
          zip.Add(XMLSS.MediaList[i].Content, 'xl/media/' + XMLSS.MediaList[i].FileName);
        end;

        //workbook.xml - sheets count
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateWorkBook(XMLSS, stream, _pages, _names, kol, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'xl/workbook.xml');
        finally
          stream.Free();
        end;

        //[Content_Types].xml
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateContentTypes(XMLSS, stream, kol, 0, nil, TextConverter, CodePageName, BOM, writeHelper);
          stream.Position := 0;
          zip.Add(stream, '[Content_Types].xml');
        finally
          stream.Free();
        end;

        // docProps/app.xml
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateDocPropsApp(stream, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'docProps/app.xml');
        finally
          stream.Free();
        end;

        // docProps/core.xml
        stream := TMemoryStream.Create();
        try
          ZEXLSXCreateDocPropsCore(XMLSS, stream, TextConverter, CodePageName, BOM);
          stream.Position := 0;
          zip.Add(stream, 'docProps/core.xml');
        finally
          stream.Free();
        end;
      finally
        SharedStrings.Free;
      end;
    finally
      writeHelper.Free();
    end;
  finally
    zip.Free();
    ZESClearArrays(_pages, _names);
  end;
end; //SaveXmlssToXSLX

{ TZEXMLSSHelper }

procedure TZEXMLSSHelper.LoadFromFile(fileName: string);
var stream: TFileStream;
begin
    stream := TFileStream.Create(fileName, fmOpenRead or fmShareDenyNone);
    try
        LoadFromStream(stream);
    finally
        stream.Free();
    end;
end;

procedure TZEXMLSSHelper.LoadFromStream(stream: TStream);
begin
    ReadXLSXFile(self, stream);
end;

procedure TZEXMLSSHelper.SaveToFile(fileName: string);
var stream: TFileStream;
begin
    stream := TFileStream.Create(fileName, fmCreate or fmOpenReadWrite);
    try
        SaveToStream(stream);
    finally
        stream.Free();
    end;
end;

procedure TZEXMLSSHelper.SaveToStream(stream: TStream);
begin
    SaveXmlssToXLSX(self, stream, [], [], nil, 'UTF-8');
end;

end.
