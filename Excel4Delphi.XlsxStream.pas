unit Excel4Delphi.XlsxStream;

// Типы файлов: Реализовать класс для управления
// Стили: Реализовать отдельный класс для чтения стилей
// Листы: Реаоизовать отдельный класс для чтения данных листа


interface

uses
{$IFDEF MSWINDOWS}
  Winapi.Windows,
{$ENDIF}
  System.SysUtils,
  System.Classes,
  System.Types,
{$IFDEF FMX}
  FMX.Graphics,
{$ELSE}
  Vcl.Graphics,
{$ENDIF}
  System.UITypes,
  System.Zip,
  System.IOUtils,
  System.Generics.Collections,
  Excel4Delphi.Formula,
  Excel4Delphi.Xml,
  Excel4Delphi,
  Excel4Delphi.Common;

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
    rtDrawing    = 9,

    rtWorkBook
  );

type
  TContentTypeRec=record
    ftype: TRelationType;
    name : string;
    rel  : string;
  end;

//  TContentType = class
//  private
//    function GetContentType(): TRelationType;
//  public
//    PartName, ContentType: string;
//    constructor Create();
//    property ContentType: TRelationType read GetContentType;
//  end;

type
  TZXLSXFileItem = record
    name: string;     //путь к файлу
    nameArx: string;
    original: string; //исходная строка
    ftype: TRelationType;   //тип контента
  end;

type
  TZXLSXRelations = record
    id: string;           //rID
    ftype: TRelationType; //тип ссылки
    target: string;       //ссылка на файла
    fileid: integer;      //ссылка на запись
    name: string;         //имя листа
    state: string;        //cостояние, возможные значения см. https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/74cb1d22-b931-4bf8-997d-17517e2416e9
    sheetid: integer;     //номер листа
  end;

//  TZXLSXDiffBorderItemStyle = class(TPersistent)
//  private
//    FUseStyle: boolean;             //заменять ли стиль
//    FUseColor: boolean;             //заменять ли цвет
//    FColor: TColor;                 //цвет линий
//    FLineStyle: TZBorderType;       //стиль линий
//    FWeight: byte;
//  protected
//  public
//    constructor Create();
//    procedure Clear();
//    procedure Assign(Source: TPersistent); override;
//    property UseStyle: boolean read FUseStyle write FUseStyle;
//    property UseColor: boolean read FUseColor write FUseColor;
//    property Color: TColor read FColor write FColor;
//    property LineStyle: TZBorderType read FLineStyle write FLineStyle;
//    property Weight: byte read FWeight write FWeight;
//  end;

//  TZXLSXDiffBorder = class(TPersistent)
//  private
//    FBorder: array [0..5] of TZXLSXDiffBorderItemStyle;
//    procedure SetBorder(Num: TZBordersPos; Const Value: TZXLSXDiffBorderItemStyle);
//    function GetBorder(Num: TZBordersPos): TZXLSXDiffBorderItemStyle;
//  public
//    constructor Create(); virtual;
//    destructor Destroy(); override;
//    procedure Clear();
//    procedure Assign(Source: TPersistent);override;
//    property Border[Num: TZBordersPos]: TZXLSXDiffBorderItemStyle read GetBorder write SetBorder; default;
//  end;

//type TZXLSXDiffFormattingItem = class(TPersistent)
//  private
//    FUseFont: boolean;              //заменять ли шрифт
//    FUseFontColor: boolean;         //заменять ли цвет шрифта
//    FUseFontStyles: boolean;        //заменять ли стиль шрифта
//    FFontColor: TColor;             //цвет шрифта
//    FFontStyles: TFontStyles;       //стиль шрифта
//    FUseBorder: boolean;            //заменять ли рамку
//    FBorders: TZXLSXDiffBorder;     //Что менять в рамке
//    FUseFill: boolean;              //заменять ли заливку
//    FUseCellPattern: boolean;       //Заменять ли тип заливки
//    FCellPattern: TZCellPattern;    //тип заливки
//    FUseBGColor: boolean;           //заменять ли цвет заливки
//    FBGColor: TColor;               //цвет заливки
//    FUsePatternColor: boolean;      //Заменять ли цвет шаблона заливки
//    FPatternColor: TColor;          //Цвет шаблона заливки
//  protected
//  public
//    constructor Create();
//    destructor Destroy(); override;
//    procedure Clear();
//    procedure Assign(Source: TPersistent); override;
//    property UseFont: boolean read FUseFont write FUseFont;
//    property UseFontColor: boolean read FUseFontColor write FUseFontColor;
//    property UseFontStyles: boolean read FUseFontStyles write FUseFontStyles;
//    property FontColor: TColor read FFontColor write FFontColor;
//    property FontStyles: TFontStyles read FFontStyles write FFontStyles;
//    property UseBorder: boolean read FUseBorder write FUseBorder;
//    property Borders: TZXLSXDiffBorder read FBorders write FBorders;
//    property UseFill: boolean read FUseFill write FUseFill;
//    property UseCellPattern: boolean read FUseCellPattern write FUseCellPattern;
//    property CellPattern: TZCellPattern read FCellPattern write FCellPattern;
//    property UseBGColor: boolean read FUseBGColor write FUseBGColor;
//    property BGColor: TColor read FBGColor write FBGColor;
//    property UsePatternColor: boolean read FUsePatternColor write FUsePatternColor;
//    property PatternColor: TColor read FPatternColor write FPatternColor;
//  end;

type
  TXlsxReader = class
  private
    MaximumDigitWidth: double;
    FWorkBook: TZWorkBook;
    //FStyleReader: TZlsxStyleReader;
    FFiles: TList<TZXLSXFileItem>;
    FRelations: TList<TZXLSXRelations>;
    FSharedStrings: TList<string>;
  protected
    procedure ReadDocPropsApp(stream: TStream);
    procedure ReadDocPropsCore(stream: TStream);
    procedure ReadDrawingRels(stream: TStream; sheet: TZSheet);
    procedure ReadDrawing(stream: TStream; sheet: TZSheet);
    procedure ReadTheme(stream: TStream);
    procedure ReadContentTypes(stream: TStream);
    procedure ReadSharedStrings(stream: TStream);
    procedure ReadStyles(stream: TStream);
    procedure ReadComments(stream: TStream);
    procedure ReadRelationships(stream: TStream);
    procedure ReadWorkSheet(stream: TStream; sheet: TZSheet);
    procedure ReadWorkBook(stream: TStream);
  public
    constructor Create(workBook: TZWorkBook);
    destructor Destroy(); override;
    property WorkBook: TZWorkBook read FWorkBook;
    procedure LoadFromStream(AStream: TStream);
  end;

  TXlsxWriter = class
  private
    FSharedStrings: TList<string>;
    FWorkBook: TZWorkBook;
  protected
    procedure WriteDocPropsApp(stream: TStream);
    procedure WriteDocPropsCore(stream: TStream);
    procedure WriteDrawingRels(stream: TStream; sheet: TZSheet);
    procedure WriteDrawing(stream: TStream; sheet: TZSheet);
    procedure WriteTheme(stream: TStream);
    procedure WriteContentTypes(stream: TStream);
    procedure WriteSharedStrings(stream: TStream);
    procedure WriteStyles(stream: TStream);
    procedure WriteComments(stream: TStream);
    procedure WriteRelationships(stream: TStream; sheet: TZSheet);
    procedure WriteRelationshipsMain(stream: TStream);
    procedure WriteWorkSheet(stream: TStream; sheet: TZSheet);
    procedure WriteWorkBook(stream: TStream);
  public
    constructor Create(workBook: TZWorkBook);
    destructor Destroy(); override;
    property WorkBook: TZWorkBook read FWorkBook;
    procedure SaveToStream(AStream: TStream);
  end;

//type TZlsxStyleReader = class
//  constructor Create();
//end;

implementation

uses
  System.AnsiStrings,
  System.StrUtils,
  System.Math,
  System.NetEncoding,
  Excel4Delphi.NumberFormat,
  Excel4Delphi.Stream;

const
  SCHEMA_DOC         = 'http://schemas.openxmlformats.org/officeDocument/2006';
  SCHEMA_DOC_REL     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  SCHEMA_PACKAGE     = 'http://schemas.openxmlformats.org/package/2006';
  SCHEMA_PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  SCHEMA_SHEET_MAIN  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const
CONTENT_TYPES: array[0..10] of TContentTypeRec = (
 (ftype: TRelationType.rtWorkSheet;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'),

 (ftype: TRelationType.rtStyles;
    name:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'),

 (ftype: TRelationType.rtWorkBook;
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
 {
type
TZStyleReader = class
  FReader: TXlsxReader;
  FStream: TStream; // styles.xml
public
  constructor Create(reader: TXlsxReader; stream: TStream);
  destructor Destroy(); override;
end;

TZSheetReader = class
  FReader: TXlsxReader;
  FSheet: TZSheet;
  FStream: TStream;
public
  constructor Create(reader: TXlsxReader; sheet: TZSheet; stream: TStream);
  destructor Destroy(); override;
end;

  }
function XLSXBoolToStr(value: boolean): string;
begin
  if (value) then
    result := 'true'
  else
    result := 'false';
end;

function GetMaximumDigitWidth(fontName: string; fontSize: double): double;
const
  numbers = '0123456789';
var
  bitmap: Vcl.Graphics.TBitmap;
  number: string;
begin
  //А.А.Валуев Расчитываем ширину самого широкого числа.
  Result := 0;
  bitmap := Vcl.Graphics.TBitmap.Create;
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

{ TXlsxWriter }

constructor TXlsxWriter.Create(workBook: TZWorkBook);
begin
  FSharedStrings := TList<string>.Create();
  FWorkBook := workBook;
end;

destructor TXlsxWriter.Destroy;
begin
  FSharedStrings.Free();
  inherited;
end;

procedure TXlsxWriter.SaveToStream(AStream: TStream);
var
  i: integer;
  zip: TZipFile;
  stream: TStream;
begin
  zip := TZipFile.Create();
  try
    zip.Open(AStream, zmReadWrite);

    // styles
    stream := TMemoryStream.Create();
    try
      WriteStyles(stream);
      stream.Position := 0;
      zip.Add(stream, 'xl/styles.xml');
    finally
      stream.Free();
    end;

    // xl/_rels/workbook.xml.rels
    stream := TMemoryStream.Create();
    try
      // todo:
      //ZEXLSXCreateRelsWorkBook(kol, stream, TextConverter, CodePageName, BOM);
      stream.Position := 0;
      zip.Add(stream, 'xl/_rels/workbook.xml.rels');
    finally
      stream.Free();
    end;

    // sheets of workbook
    for I := 0 to FWorkBook.Sheets.Count-1 do begin
      // todo: big file ???
      stream := TMemoryStream.Create();
      try
        WriteWorkSheet(stream, FWorkBook.Sheets[i]);
        stream.Position := 0;
        zip.Add(stream, 'xl/worksheets/sheet' + IntToStr(i + 1) + '.xml');
      finally
        stream.Free();
      end;

      stream := TMemoryStream.Create();
      try
        WriteRelationships(stream, FWorkBook.Sheets[i]);
        stream.Position := 0;
        zip.Add(stream, 'xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels');
      finally
        stream.Free();
      end;

      if not FWorkBook.Sheets[i].Drawing.IsEmpty then begin
        // drawings/drawingN.xml
        stream := TMemoryStream.Create();
        try
          WriteDrawing(stream, FWorkBook.Sheets[i]);
          stream.Position := 0;
          zip.Add(stream, 'xl/drawings/drawing' + IntToStr(i+1) + '.xml');
        finally
          stream.Free();
        end;

        // drawings/_rels/drawingN.xml.rels
        stream := TMemoryStream.Create();
        try
          WriteDrawingRels(stream, FWorkBook.Sheets[i]);
          stream.Position := 0;
          zip.Add(stream, 'xl/drawings/_rels/drawing' + IntToStr(i+1) + '.xml.rels');
        finally
          stream.Free();
        end;
      end;
    end;

    // sharedStrings.xml
    stream := TMemoryStream.Create();
    try
      WriteSharedStrings(stream);
      stream.Position := 0;
      zip.Add(stream, 'xl/sharedStrings.xml');
    finally
      stream.Free();
    end;

    // media/imageN.png
    for I := 0 to High(FWorkBook.MediaList) do
      zip.Add(FWorkBook.MediaList[i].Content,
        'xl/media/' + FWorkBook.MediaList[i].FileName);

    //workbook.xml - sheets count
    stream := TMemoryStream.Create();
    try
      WriteWorkBook(stream);
      stream.Position := 0;
      zip.Add(stream, 'xl/workbook.xml');
    finally
      stream.Free();
    end;

    // docProps/app.xml
    stream := TMemoryStream.Create();
    try
      WriteDocPropsApp(stream);
      stream.Position := 0;
      zip.Add(stream, 'docProps/app.xml');
    finally
      stream.Free();
    end;

    // docProps/core.xml
    stream := TMemoryStream.Create();
    try
      WriteDocPropsCore(stream);
      stream.Position := 0;
      zip.Add(stream, 'docProps/core.xml');
    finally
      stream.Free();
    end;

    // _rels/.rels
    stream := TMemoryStream.Create();
    try
      WriteRelationshipsMain(stream);
      stream.Position := 0;
      zip.Add(stream, '_rels/.rels');
    finally
      stream.Free();
    end;

    //[Content_Types].xml
    stream := TMemoryStream.Create();
    try
      WriteContentTypes(stream);
      stream.Position := 0;
      zip.Add(stream, '[Content_Types].xml');
    finally
      stream.Free();
    end;
  finally
    zip.Free();
  end;
end;

procedure TXlsxWriter.WriteComments(stream: TStream);
begin

end;

procedure TXlsxWriter.WriteContentTypes(stream: TStream);
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

  procedure makeOwerride(const part, content: string);
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('PartName', part);
    xml.Attributes.Add('ContentType', content, false);
    xml.WriteEmptyTag('Override', true);
  end;

  procedure _WriteTypeDefault(extension, contentType: string);
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('Extension', extension);
    xml.Attributes.Add('ContentType', contentType, false);
    xml.WriteEmptyTag('Default', true);
  end;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE + '/content-types');
    xml.WriteTagNode('Types', true, true, true);

    _WriteTypeDefault('rels', 'application/vnd.openxmlformats-package.relationships+xml');
    _WriteTypeDefault('xml',  'application/xml');
    _WriteTypeDefault('png',  'image/png');
    _WriteTypeDefault('jpeg', 'image/jpeg');
    _WriteTypeDefault('wmf',  'image/x-wmf');

    //Страницы
    //_WriteOverride('/_rels/.rels', 3);
    //_WriteOverride('/xl/_rels/workbook.xml.rels', 3);
    for var i := 0 to FWorkBook.Sheets.Count - 1 do begin
      _WriteOverride('/xl/worksheets/sheet' + IntToStr(i + 1) + '.xml', 0);
//      if (WriteHelper.IsSheetHaveHyperlinks(i)) then
//        _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(i + 1) + '.xml.rels', 3);
    end;
    //комментарии
//    for i := 0 to CommentCount - 1 do begin
//      _WriteOverride('/xl/worksheets/_rels/sheet' + IntToStr(PagesComments[i] + 1) + '.xml.rels', 3);
//      _WriteOverride('/xl/comments' + IntToStr(PagesComments[i] + 1) + '.xml', 6);
//    end;

    for var i := 0 to FWorkBook.Sheets.Count - 1 do begin
      if Assigned(FWorkBook.Sheets[i].Drawing) and (FWorkBook.Sheets[i].Drawing.Count > 0) then begin
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

    makeOwerride('/xl/workbook.xml',      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
    makeOwerride('/xl/styles.xml',        'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
    makeOwerride('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
    makeOwerride('/docProps/app.xml',     'application/vnd.openxmlformats-officedocument.extended-properties+xml');
    makeOwerride('/docProps/core.xml',    'application/vnd.openxmlformats-package.core-properties+xml');

    xml.WriteEndTagNode(); //Types
  finally
    xml.Free();
  end;
end;

procedure TXlsxWriter.WriteDocPropsApp(stream: TStream);
var xml: TZsspXMLWriterH;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
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
end;

procedure TXlsxWriter.WriteDocPropsCore(stream: TStream);
var xml: TZsspXMLWriterH;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns:cp',        SCHEMA_PACKAGE + '/metadata/core-properties');
    xml.Attributes.Add('xmlns:dc',       'http://purl.org/dc/elements/1.1/', false);
    xml.Attributes.Add('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/', false);
    xml.Attributes.Add('xmlns:dcterms',  'http://purl.org/dc/terms/', false);
    xml.Attributes.Add('xmlns:xsi',      'http://www.w3.org/2001/XMLSchema-instance', false);
    xml.WriteTagNode('cp:coreProperties', true, true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('xsi:type', 'dcterms:W3CDTF');
    xml.WriteTag('dcterms:created',  ZEDateTimeToStr(FWorkBook.DocumentProperties.Created) + 'Z', true, false, false);
    xml.WriteTag('dcterms:modified', ZEDateTimeToStr(FWorkBook.DocumentProperties.Created) + 'Z', true, false, false);

    xml.Attributes.Clear();
    xml.WriteTag('cp:revision', '1', true, false, false);

    xml.WriteEndTagNode(); //cp:coreProperties
  finally
    xml.Free();
  end;
end;

procedure TXlsxWriter.WriteDrawing(stream: TStream; sheet: TZSheet);
var xml: TZsspXMLWriterH;
  pic: TZEPicture;
  i: integer;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    //xml.NewLine := false;
    xml.WriteHeader();
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

procedure TXlsxWriter.WriteDrawingRels(stream: TStream; sheet: TZSheet);
var xml: TZsspXMLWriterH;
  i: integer;
  dic: TDictionary<integer, string>;
  pair: TPair<integer, string>;
begin
  dic := TDictionary<integer, string>.Create(); // убрать и заменить на список
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL, false);
    xml.WriteTagNode('Relationships', false, false, false);

    for i := 0 to sheet.Drawing.Count - 1 do begin
      dic.AddOrSetValue(sheet.Drawing[i].RelId, sheet.Drawing[i].Name);
    end;

    for pair in dic do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('Id',     'rId' + IntToStr(pair.Key));
      xml.Attributes.Add('Type',   SCHEMA_DOC_REL + '/image');
      xml.Attributes.Add('Target', '../media/' + pair.Value);
      xml.WriteEmptyTag('Relationship', false, true);
    end;
    xml.WriteEndTagNode(); // Relationships
  finally
    xml.Free();
    dic.Free();
  end;
end;

procedure TXlsxWriter.WriteRelationships(stream: TStream; sheet: TZSheet);
begin

end;

procedure TXlsxWriter.WriteRelationshipsMain(stream: TStream);
var xml: TZsspXMLWriterH;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    xml.Attributes.Add('xmlns', SCHEMA_PACKAGE_REL);
    xml.WriteTagNode('Relationships', true, true, false);

    xml.Attributes.Clear();
    xml.Attributes.Add('Id',     'rId1');
    xml.Attributes.Add('Type',   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', false);
    xml.Attributes.Add('Target', 'docProps/app.xml', false);
    xml.WriteEmptyTag('Relationship', true, true);

    xml.Attributes.Clear();
    xml.Attributes.Add('Id',     'rId2');
    xml.Attributes.Add('Type',   'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', false);
    xml.Attributes.Add('Target', 'docProps/core.xml', false);
    xml.WriteEmptyTag('Relationship', true, true);

    xml.Attributes.Clear();
    xml.Attributes.Add('Id',     'rId3');
    xml.Attributes.Add('Type',   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', false);
    xml.Attributes.Add('Target', 'xl/workbook.xml', false);
    xml.WriteEmptyTag('Relationship', true, true);

    xml.WriteEndTagNode(); //Relationships
  finally
    xml.Free();
  end;
end;

procedure TXlsxWriter.WriteSharedStrings(stream: TStream);
var xml: TZsspXMLWriterH; i: integer; str: string;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    // todo: prepare

    xml.WriteHeader();
    xml.Attributes.Clear();
    xml.Attributes.Add('count', FSharedStrings.Count.ToString);
    xml.Attributes.Add('uniqueCount', FSharedStrings.Count.ToString, false);
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN, false);
    xml.WriteTagNode('sst', true, true, false);

    {- Write out the content of Shared Strings: <si><t>Value</t></si> }
    for i := 0 to Pred(FSharedStrings.Count) do begin
      xml.Attributes.Clear();
      xml.WriteTagNode('si', false, false, false);
      str := FSharedStrings[i];
      xml.Attributes.Clear();
      if str.StartsWith(' ') or str.EndsWith(' ') then
        //А.А.Валуев Чтобы ведущие и последние пробелы не терялись,
        //добавляем атрибут xml:space="preserve".
        xml.Attributes.Add('xml:space', 'preserve', false);
      xml.WriteTag('t', str);
      xml.WriteEndTagNode();
    end;

    xml.WriteEndTagNode(); //Relationships
  finally
    xml.Free();
  end;
end;

procedure TXlsxWriter.WriteStyles(stream: TStream);
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
      _style := FWorkBook.Styles[StyleNum];
      if (_style.NumberFormatId > 0) and (_style.NumberFormatId < 164) then
        Exit(_style.NumberFormatId);

      //If cell type is datetime and cell style is empty then need write default NumFmtId = 14.
      if ((Trim(_style.NumberFormat) = '') or (UpperCase(_style.NumberFormat) = 'GENERAL')) then begin
        if (_is_dateTime[StyleNum + 1]) then
          Result := 14
        else begin
          b := false;
          _cs := _currSheet;
          for i := _cs to FWorkBook.Sheets.Count - 1 do begin
            _sheet := FWorkBook.Sheets[i];
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
    kol := FWorkBook.Styles.Count + 1;
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
        if (_isFontsEqual(FWorkBook.Styles[i], FWorkBook.Styles[j])) then
          _FontIndex[j + 1] := i;
    end; //if

    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(_fontCount));
    xml.WriteTagNode('fonts', true, true, true);

    for i := 0 to _StylesCount do
    if (_FontIndex[i] = - 2) then begin
      fnt := FWorkBook.Styles[i - 1].Font;
      xml.Attributes.Clear();
      xml.WriteTagNode('font', true, true, true);

      xml.Attributes.Clear();
      xml.Attributes.Add('val', fnt.Name);
      xml.WriteEmptyTag('name', true);

      xml.Attributes.Clear();
      xml.Attributes.Add('val', IntToStr(fnt.Charset));
      xml.WriteEmptyTag('charset', true);

      xml.Attributes.Clear();
      xml.Attributes.Add('val', FloatToStr(fnt.Size, TFormatSettings.Invariant));
      xml.WriteEmptyTag('sz', true);

      if (fnt.Color <> clWindowText) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('rgb', '00' + ColorToHTMLHex(fnt.Color));
        xml.WriteEmptyTag('color', true);
      end;

      if (fsBold in fnt.Style) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('val', 'true');
        xml.WriteEmptyTag('b', true);
      end;

      if (fsItalic in fnt.Style) then begin
        xml.Attributes.Clear();
        xml.Attributes.Add('val', 'true');
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

      //<vertAlign val="superscript"/>
      if FWorkBook.Styles[i - 1].Superscript then begin
        xml.Attributes.Clear();

        xml.Attributes.Add('val', 'superscript');
        xml.WriteEmptyTag('vertAlign', true);
      end

      //<vertAlign val="subscript"/>

      else if FWorkBook.Styles[i - 1].Subscript then begin

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
        if (_isFillsEqual(FWorkBook.Styles[i], FWorkBook.Styles[j])) then
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

      case FWorkBook.Styles[i].CellPattern of
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

      b := (FWorkBook.Styles[i].PatternColor <> clWindow) or (FWorkBook.Styles[i].BGColor <> clWindow);
      xml.Attributes.Clear();
      if b and (FWorkBook.Styles[i].CellPattern = ZPNone) then
        xml.Attributes.Add('patternType', 'solid')
      else
        xml.Attributes.Add('patternType', s);

      if (b) then
        xml.WriteTagNode('patternFill', true, true, false)
      else
        xml.WriteEmptyTag('patternFill', true, false);

      _reverse := not (FWorkBook.Styles[i].CellPattern in [ZPNone, ZPSolid]);

      if (FWorkBook.Styles[i].BGColor <> clWindow) then
      begin
        xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := FWorkBook.Styles[i].PatternColor
        else
          _tmpColor := FWorkBook.Styles[i].BGColor;
        xml.Attributes.Add('rgb', 'FF' + ColorToHTMLHex(_tmpColor));
        xml.WriteEmptyTag('fgColor', true);
      end;

      if (FWorkBook.Styles[i].PatternColor <> clWindow) then
      begin
        xml.Attributes.Clear();
        if (_reverse) then
          _tmpColor := FWorkBook.Styles[i].BGColor
        else
          _tmpColor := FWorkBook.Styles[i].PatternColor;
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
    _border := FWorkBook.Styles[StyleNum].Border[BorderNum];
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
        if (FWorkBook.Styles[i].Border.isEqual(FWorkBook.Styles[j].Border)) then
          _BorderIndex[j + 1] := i;
    end; //if

    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(_borderCount));
    xml.WriteTagNode('borders', true, true, true);

    for i := -1 to _StylesCount - 1 do
    if (_BorderIndex[i + 1] = -2) then begin
      xml.Attributes.Clear();
      s := 'false';
      if (FWorkBook.Styles[i].Border[bpDiagonalLeft].Weight > 0) and (FWorkBook.Styles[i].Border[bpDiagonalLeft].LineStyle <> ZENone) then
        s := 'true';
      xml.Attributes.Add('diagonalDown', s);
      s := 'false';
      if (FWorkBook.Styles[i].Border[bpDiagonalRight].Weight > 0) and (FWorkBook.Styles[i].Border[bpDiagonalRight].LineStyle <> ZENone) then
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
    _style := FWorkBook.Styles[NumStyle];
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
    xml.Attributes.Add('hidden', XLSXBoolToStr(FWorkBook.Styles[NumStyle].Protect));
    xml.Attributes.Add('locked', XLSXBoolToStr(FWorkBook.Styles[NumStyle].HideFormula));
    xml.WriteEmptyTag('protection', true);

    xml.WriteEndTagNode(); //xf
  end; //_WriteXF

  //<cellStyleXfs> ... </cellStyleXfs> / <cellXfs> ... </cellXfs>
  procedure WriteCellStyleXfs(const TagName: string; isxfId: boolean);
  var i: integer;
  begin
    xml.Attributes.Clear();
    xml.Attributes.Add('count', IntToStr(FWorkBook.Styles.Count + 1));
    xml.WriteTagNode(TagName, true, true, true);
    for i := -1 to FWorkBook.Styles.Count - 1 do  begin
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
  _FmtParser := TNumFormatParser.Create();
  _DateParser := TZDateTimeODSFormatParser.Create();
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    _StylesCount := FWorkBook.Styles.Count;

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
end;

procedure TXlsxWriter.WriteTheme(stream: TStream);
begin

end;

procedure TXlsxWriter.WriteWorkBook(stream: TStream);
var xml: TZsspXMLWriterH; i: integer;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();

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

    xml.Attributes.Add('activeTab', '0');
    xml.Attributes.Add('firstSheet', '0', false);
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
    for i := 0 to FWorkBook.Sheets.Count-1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('name', FWorkBook.Sheets[i].Title, false);
      xml.Attributes.Add('sheetId', IntToStr(i + 1), false);

      case FWorkBook.Sheets[i].Visible of
        svHidden:     xml.Attributes.Add('state', 'hidden', false);
        svVeryHidden: xml.Attributes.Add('state', 'veryhidden', false);
      else
        xml.Attributes.Add('state', 'visible', false);
      end;
      xml.Attributes.Add('r:id', 'rId' + IntToStr(i + 2), false);
      xml.WriteEmptyTag('sheet', true);
    end;
    xml.WriteEndTagNode(); //sheets

    // definedNames
    xml.Attributes.clear();
    xml.WriteTagNode('definedNames', true, true, true);
    for i := 0 to High(FWorkBook.FDefinedNames) do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('localSheetId', IntToStr(FWorkBook.FDefinedNames[i].LocalSheetId), false);
      xml.Attributes.Add('name', FWorkBook.FDefinedNames[i].Name, false);
      xml.WriteTag('definedName', FWorkBook.FDefinedNames[i].Body);
    end;
    xml.WriteEndTagNode(); //definedNames

    xml.Attributes.Clear();
    xml.Attributes.Add('iterateCount', '100');
    xml.Attributes.Add('refMode', 'A1', false);
    xml.Attributes.Add('iterate', 'false', false);
    xml.Attributes.Add('iterateDelta', '0.001', false);
    xml.WriteEmptyTag('calcPr', true);

    xml.WriteEndTagNode(); //workbook
  finally
    xml.Free();
  end;
end;

procedure TXlsxWriter.WriteWorkSheet(stream: TStream; sheet: TZSheet);
var xml: TZsspXMLWriterH;

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
    xml.Attributes.Add('fitToPage', sheet.FitToPage.ToString()); // todo: check it!!!
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
    MaximumDigitWidth := 0;// GetMaximumDigitWidth(XMLSS.Styles[0].Font.Name, XMLSS.Styles[0].Font.Size); //Если совсем нет стилей, пусть будет ошибка.
    xml.Attributes.Clear();
    xml.WriteTagNode('cols', true, true, true);
    for i := 0 to sheet.ColCount - 1 do begin
      xml.Attributes.Clear();
      xml.Attributes.Add('collapsed', 'false', false);
      xml.Attributes.Add('hidden', sheet.Columns[i].Hidden.ToString, false);
      xml.Attributes.Add('max', IntToStr(i + 1), false);
      xml.Attributes.Add('min', IntToStr(i + 1), false);
      s := '0';
      ProcessedColumn := sheet.Columns[i];
      if ((ProcessedColumn.StyleID >= -1) and (ProcessedColumn.StyleID < FWorkBook.Styles.Count)) then
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
    s: string;
    _r: TRect;
    strIndex: integer;
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
        xml.Attributes.Clear();
//        if (not WriteHelper.isHaveComments) then
//          if (sheet.Cell[j, i].Comment > '') then
//            WriteHelper.isHaveComments := true;
        b := (sheet.Cell[j, i].Data > '') or
             (sheet.Cell[j, i].Formula > '');
        s := TZEFormula.GetColAddres(j) + IntToStr(i + 1);

//        if (sheet.Cell[j, i].HRef <> '') then
//          WriteHelper.AddHyperLink(s, sheet.Cell[j, i].HRef, sheet.Cell[j, i].HRefScreenTip, 'External');

        xml.Attributes.Add('r', s);

        if (sheet.Cell[j, i].CellStyle >= -1) and (sheet.Cell[j, i].CellStyle < FWorkBook.Styles.Count) then
          s := IntToStr(sheet.Cell[j, i].CellStyle + 1)
        else
          s := '0';
        xml.Attributes.Add('s', s, false);

        case sheet.Cell[j, i].CellType of
          ZENumber:   s := 'n';
          ZEDateTime: s := 'd'; //??
          ZEBoolean:  s := 'b';
          ZEString:
          begin
            //А.А.Валуев Общие строки пишем только, если в строке есть
            //определённые символы. Хотя можно писать и всё подряд.
            if sheet.Cell[j, i].Data.StartsWith(' ')
                or sheet.Cell[j, i].Data.EndsWith(' ')
                or (sheet.Cell[j, i].Data.IndexOfAny([#10, #13]) >= 0) then
            begin
              //А.А.Валуев С помощью словаря пытаемся находить дубликаты строк.
//              if SharedStringsDictionary.ContainsKey(sheet.Cell[j, i].Data) then
//                strIndex := SharedStringsDictionary[sheet.Cell[j, i].Data]
//              else
//              begin
//                strIndex := Length(SharedStrings);
//                Insert(sheet.Cell[j, i].Data, SharedStrings, strIndex);
//                SharedStringsDictionary.Add(sheet.Cell[j, i].Data, strIndex);
//              end;
              s := 's';
            end
            else
              s := 'str';
          end;
          ZEError: s := 'e';
        end;

        // если тип ячейки ZEGeneral, то атрибут опускаем
        if  (sheet.Cell[j, i].CellType <> ZEGeneral)
        and (sheet.Cell[j, i].CellType <> ZEDateTime) then
          xml.Attributes.Add('t', s, false);

        if (b) then begin
          xml.WriteTagNode('c', true, true, false);
          if (sheet.Cell[j, i].Formula > '') then begin
            xml.Attributes.Clear();
            xml.Attributes.Add('aca', 'false');
            xml.WriteTag('f', sheet.Cell[j, i].Formula, true, false, true);
          end;
          if (sheet.Cell[j, i].Data > '') then begin
            xml.Attributes.Clear();
            if s = 's' then
              xml.WriteTag('v', strIndex.ToString, true, false, true)
            else
              xml.WriteTag('v', sheet.Cell[j, i].Data, true, false, true);
          end;
          xml.WriteEndTagNode();
        end else
          xml.WriteEmptyTag('c', true);
      end;
      xml.WriteEndTagNode(); //row
    end; //for i

    xml.WriteEndTagNode(); //sheetData

    // autoFilter
    if not trim(sheet.AutoFilter).IsEmpty then begin
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

    //WriteHelper.WriteHyperLinksTag(xml);
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

    xml.Attributes.Add('pageOrder', 'downThenOver', false);
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
      rId := 1;
      //rId := WriteHelper.AddDrawing('../drawings/drawing' + IntToStr(sheet.SheetIndex + 1) + '.xml');
      xml.Attributes.Clear();
      xml.Attributes.Add('r:id', 'rId' + IntToStr(rId));
      xml.WriteEmptyTag('drawing');
    end;
  end;
begin
  xml := TZsspXMLWriterH.Create(Stream);
  try
    xml.WriteHeader();
    xml.Attributes.Clear();
    xml.Attributes.Add('xmlns', SCHEMA_SHEET_MAIN);
    xml.Attributes.Add('xmlns:r', SCHEMA_DOC_REL);
    xml.WriteTagNode('worksheet', true, true, false);

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
end;

{ TXlsxReader }

constructor TXlsxReader.Create(workBook: TZWorkBook);
begin
  FWorkBook := workBook;
  MaximumDigitWidth := 0;
  FFiles := TList<TZXLSXFileItem>.Create();
  FRelations := TList<TZXLSXRelations>.Create();
  FSharedStrings := TList<string>.Create();
end;

destructor TXlsxReader.Destroy;
begin
  FFiles.Free();
  FRelations.Free();
  inherited;
end;

procedure TXlsxReader.LoadFromStream(AStream: TStream);
var
  i, j: integer;
  s: string;
  zip: TZipFile;
  zipHdr: TZipHeader;
  buff: TBytes;
  stream: TStream;
  fileRec: TZXLSXFileItem;
  sheet: TZSheet;
begin
  zip := TZipFile.Create();
  zip.Encoding := TEncoding.GetEncoding(437);

  FWorkBook.Styles.Clear();
  FWorkBook.Sheets.Count := 0;
  try
    zip.Open(AStream, zmRead);

    zip.Read('[Content_Types].xml', stream, zipHdr);
    try
      ReadContentTypes(stream);
    finally
      FreeAndNil(stream);
    end;

    // todo: check it by type
    s := '/_rels/.rels';
    if zip.IndexOf(s.Substring(1)) > -1 then begin
      fileRec.original := s;
      fileRec.name := s;
      fileRec.ftype := TRelationType.rtDoc;
      FFiles.Add(fileRec);
    end;

    s := '/xl/_rels/workbook.xml.rels';
    if zip.IndexOf(s.Substring(1)) > -1 then begin
      fileRec.original := s;
      fileRec.name := s;
      fileRec.ftype := TRelationType.rtDoc;
      FFiles.Add(fileRec);
    end;

    for i := 0 to FFiles.Count - 1 do begin
      if (FFiles[i].ftype = TRelationType.rtDoc) then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          ReadRelationships(stream);
        finally
          FreeAndNil(stream);
        end;
      end;
    end;

    // sharedstrings.xml
    for i:= 0 to FFiles.Count - 1 do begin
      if (FFiles[i].ftype = TRelationType.rtCoreProp) then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          ReadSharedStrings(stream);
        finally
          FreeAndNil(stream);
        end;
        break;
      end;
    end;

    // theme
    for i := 0 to FFiles.Count - 1 do begin
      if (FFiles[i].ftype = TRelationType.rtVmlDrawing) then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          ReadTheme(stream);
        finally
          FreeAndNil(stream);
        end;
        break;
      end;
    end;

    // styles
    for i := 0 to FFiles.Count - 1 do begin
      if (FFiles[i].ftype = TRelationType.rtStyles) then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          ReadStyles(stream);
        finally
          FreeAndNil(stream);
        end;
      end;
    end;

    // workbook
    for i := 0 to FFiles.Count - 1 do begin
      if FFiles[i].ftype = TRelationType.rtWorkBook then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          ReadWorkBook(stream);
        finally
          FreeAndNil(stream);
        end;
      end;
    end;

    // worksheets
    for i := 0 to FFiles.Count - 1 do begin
      if FFiles[i].ftype = TRelationType.rtWorkSheet then begin
        zip.Read(FFiles[i].original.Substring(1), stream, zipHdr);
        try
          sheet := FWorkBook.Sheets.Add('');
          ReadWorkSheet(stream, sheet);
        finally
          FreeAndNil(stream);
        end;
      end;
    end;

    // drawings
    for I := 0 to FWorkBook.Sheets.Count-1 do begin
      sheet := FWorkBook.Sheets[i];
      if sheet.DrawingRid > 0 then begin
        // load images
        s := 'xl/drawings/drawing'+IntToStr(i+1)+'.xml';
        zip.Read(s, stream, zipHdr);
        try
          ReadDrawing(stream, sheet);
        finally
          stream.Free();
        end;

        // read drawing rels
        s := 'xl/drawings/_rels/drawing'+IntToStr(i+1)+'.xml.rels';
        zip.Read(s, stream, zipHdr);
        try
          ReadDrawingRels(stream, sheet);
        finally
          stream.Free();
        end;

        // read img file
        for j := 0 to sheet.Drawing.Count-1 do begin
          s := sheet.Drawing[j].Name;
          zip.Read('xl/media/' + s, buff);
          // only unique content
          FWorkBook.AddMediaContent(s, buff, true);
        end;
      end;
    end;
  finally
    zip.Free();
  end;
end;

procedure TXlsxReader.ReadComments(stream: TStream);
begin

end;

procedure TXlsxReader.ReadContentTypes(stream: TStream);
var
  xml: TZsspXMLReaderH;
  contType: string;
  rec: TContentTypeRec;
  newrec: TZXLSXFileItem;
begin
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    xml.BeginReadStream(Stream);

    FFiles.Clear();
    while not xml.Eof() do begin
      xml.ReadTag();
      if xml.IsTagClosedByName('Override') then begin
        contType := xml.Attributes.ItemsByName['ContentType'];
        for rec in CONTENT_TYPES do begin
          if contType = rec.name then begin
            newrec.name     := xml.Attributes.ItemsByName['PartName'];
            newrec.original := xml.Attributes.ItemsByName['PartName'];
            newrec.ftype    := rec.ftype;
            FFiles.Add(newrec);
            break;
          end;
        end;
      end;
    end;
  finally
    xml.Free();
  end;
end;

procedure TXlsxReader.ReadDocPropsApp(stream: TStream);
begin

end;

procedure TXlsxReader.ReadDocPropsCore(stream: TStream);
begin

end;

procedure TXlsxReader.ReadDrawing(stream: TStream; sheet: TZSheet);
begin

end;

procedure TXlsxReader.ReadDrawingRels(stream: TStream; sheet: TZSheet);
begin

end;

procedure TXlsxReader.ReadRelationships(stream: TStream);
begin

end;

procedure TXlsxReader.ReadSharedStrings(stream: TStream);
var
  xml: TZsspXMLReaderH;
  s: string;
  k: integer;
begin
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    xml.BeginReadStream(Stream);
    while not xml.Eof() do begin
      xml.ReadTag();
      if xml.IsTagStartByName('si') then begin
        s := '';
        k := 0;
        while xml.ReadToEndTagByName('si') do begin
          if xml.IsTagEndByName('t') then begin
            if (k > 1) then
              s := s + sLineBreak;
            s := s + xml.TextBeforeTag;
          end;
          if xml.IsTagEndByName('r') then
            inc(k);
        end;
        FSharedStrings.Add(s);
      end;
    end;
  finally
    xml.Free();
  end;
end;

procedure TXlsxReader.ReadStyles(stream: TStream);
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
  FontArray: TArray<TZEXLSXFont>;
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
  //_Style: TZStyle;
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
//      _dfIndex := ReadHelper.DiffFormatting.Count;
//
//      SetLength(_dfFonts, _dfIndex + 1);
//      _dfFonts[_dfIndex].ColorType := 0;
//      _dfFonts[_dfIndex].LumFactor := 0;
//
//      SetLength(_dfFills, _dfIndex + 1);
//      ZEXLSXClearPatternFill(_dfFills[_dfIndex]);
//
//      ReadHelper.DiffFormatting.Add();
//      _df := ReadHelper.DiffFormatting[_dfIndex];
//      while xml.ReadToEndTagByName('dxf') do begin
//        if xml.IsTagStartByName('font') then
//          _ReadDFFont()
//        else
//        if xml.IsTagStartByName('fill') then
//          _ReadDFFill()
//        else
//        if xml.IsTagStartByName('border') then
//          _ReadDFBorder();
//      end; //while
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
      if ((t >= 0) and (t < LEngth(FWorkBook.FTheme.ThemeColors))) then
        AColor := FWorkBook.FTheme.ThemeColors[t];
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
    //if (XLSXStyle.numFmtId >= 0) then
      //XMLSSStyle.NumberFormat := ReadHelper.NumberFormats.GetFormat(XLSXStyle.numFmtId);
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
      //else
//      if xml.IsTagStartByName('numFmts') then
//        ReadHelper.NumberFormats.ReadNumFmts(xml);
    end; //while

    //тут незабыть применить номера цветов, если были введены

    _CheckIndexedColors();

    //
    for i := 0 to FillCount - 1 do begin
      XLSXApplyColor(FillArray[i].bgcolor, FillArray[i].bgColorType, FillArray[i].lumFactorBG);
      XLSXApplyColor(FillArray[i].patterncolor, FillArray[i].patternColorType, FillArray[i].lumFactorPattern);
    end; //for

    //{tut}

    FWorkBook.Styles.Count := CellXfsCount;
//    ReadHelper.NumberFormats.StyleFMTCount := CellXfsCount;
//    for i := 0 to CellXfsCount - 1 do begin
//      t := CellXfsArray[i].xfId;
//      ReadHelper.NumberFormats.StyleFMTID[i] := CellXfsArray[i].numFmtId;
//
//      _Style := XMLSS.Styles[i];
//      if ((t >= 0) and (t < CellStyleCount)) then
//        _ApplyStyle(_Style, CellStyleArray[t]);
//      //else
//        _ApplyStyle(_Style, CellXfsArray[i]);
//    end;
//
//    //Применение цветов к DF
//    for i := 0 to ReadHelper.DiffFormatting.Count - 1 do begin
//      if (ReadHelper.DiffFormatting[i].UseFontColor) then begin
//        XLSXApplyColor(_dfFonts[i].Color, _dfFonts[i].ColorType, _dfFonts[i].LumFactor);
//        ReadHelper.DiffFormatting[i].FontColor := _dfFonts[i].Color;
//      end;
//      if (ReadHelper.DiffFormatting[i].UseBGColor) then begin
//        XLSXApplyColor(_dfFills[i].bgcolor, _dfFills[i].bgColorType, _dfFills[i].lumFactorBG);
//        ReadHelper.DiffFormatting[i].BGColor := _dfFills[i].bgcolor;
//      end;
//      if (ReadHelper.DiffFormatting[i].UsePatternColor) then begin
//        XLSXApplyColor(_dfFills[i].patterncolor, _dfFills[i].patternColorType, _dfFills[i].lumFactorPattern);
//        ReadHelper.DiffFormatting[i].PatternColor := _dfFills[i].patterncolor;
//      end;
//    end;
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
end;

procedure TXlsxReader.ReadTheme(stream: TStream);
var
  xml: TZsspXMLReaderH;
  flag: boolean;
  procedure _addFillColor(const _rgb: string);
  begin
    SetLength(FWorkBook.FTheme.ThemeColors, Length(FWorkBook.FTheme.ThemeColors)+1);
    FWorkBook.FTheme.ThemeColors[Length(FWorkBook.FTheme.ThemeColors)-1] := HTMLHexToColor(_rgb);
  end;
begin
  xml := TZsspXMLReaderH.Create();
  flag := false;
  try
    xml.AttributesMatch := false;
    xml.BeginReadStream(Stream);

    while not xml.Eof() do begin
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
    end;
  finally
    xml.Free();
  end;
end;

procedure TXlsxReader.ReadWorkBook(stream: TStream);
var
  xml: TZsspXMLReaderH;
  s: string;
  i, t, dn: integer;
  rel: TZXLSXRelations;
begin
  xml := TZsspXMLReaderH.Create();
  try
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    dn := 0;
    while (not xml.Eof()) do begin
      xml.ReadTag();

      if xml.IsTagStartByName('definedName') then begin
        SetLength(FWorkBook.FDefinedNames, dn + 1);
        FWorkBook.FDefinedNames[dn].LocalSheetId := StrToIntDef(xml.Attributes.ItemsByName['localSheetId'], 0);
        FWorkBook.FDefinedNames[dn].Name := xml.Attributes.ItemsByName['name'];
        inc(dn);
      end
      else if xml.IsTagEndByName('definedName') and (Length(FWorkBook.FDefinedNames) = dn) then
        FWorkBook.FDefinedNames[dn - 1].Body := xml.TagValue
      else if xml.IsTagClosedByName('sheet') then begin
        s := xml.Attributes.ItemsByName['r:id'];
        for i := 0 to FRelations.Count - 1 do begin
          rel := FRelations[i];
          if (rel.id = s) then begin
            rel.name := ZEReplaceEntity(xml.Attributes.ItemsByName['name']);
            s := xml.Attributes.ItemsByName['sheetId'];
            rel.sheetid := -1;
            if (TryStrToInt(s, t)) then
              rel.sheetid := t;
            rel.state := xml.Attributes.ItemsByName['state'];
            FRelations[i] := rel;
            break;
          end;
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
  finally
    xml.Free();
  end;
end;

procedure TXlsxReader.ReadWorkSheet(stream: TStream; sheet: TZSheet);
var
  xml: TZsspXMLReaderH;
  //currentPage: integer;
  currentRow: integer;
  currentCol: integer;
  currentCell: TZCell;
  str: string;
  tempReal: real;
  tempInt: integer;
  tempDate: TDateTime;
  tempFloat: Double;

  procedure CheckRow(const RowCount: integer);
  begin
    if (sheet.RowCount < RowCount) then
      sheet.RowCount := RowCount;
  end;

  procedure CheckCol(const ColCount: integer);
  begin
    if (sheet.ColCount < ColCount) then
      sheet.ColCount := ColCount
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
        currentCell := sheet.Cell[currentCol, currentRow];
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
//          if (ReadHelper.NumberFormats.IsDateFormat(currentCell.CellStyle)) then
//            if (ZEIsTryStrToFloat(v, tempFloat)) then begin
//              currentCell.CellType := ZEDateTime;
//              v := ZEDateTimeToStr(tempFloat);
//            end;
        end else if (_type = 's') then begin
          currentCell.CellType := ZEString;
          if (TryStrToInt(v, t)) then
            if ((t >= 0) and (t < FSharedStrings.Count)) then
              v := FSharedStrings[t];
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
        sheet.Rows[currentRow].Hidden := ZETryStrToBoolean(xml.Attributes.ItemsByName['hidden'], false);

        str := xml.Attributes.ItemsByName['ht']; //в поинтах
        if (str > '') then begin
          tempReal := ZETryStrToFloat(str, 10);
          sheet.Rows[currentRow].Height := tempReal;
          //tempReal := tempReal / 2.835; //???
          //currentSheet.Rows[currentRow].HeightMM := tempReal;
        end
        else
          sheet.Rows[currentRow].Height := sheet.DefaultRowHeight;

        str := xml.Attributes.ItemsByName['outlineLevel'];
        sheet.Rows[currentRow].OutlineLevel := StrToIntDef(str, 0);

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
    end;
    sheet.ColCount := maxCol;
  end; //_ReadSheetData

  procedure _ReadAutoFilter();
  begin
    sheet.AutoFilter := xml.Attributes.ItemsByName['ref'];
  end;

  //Чтение объединённых ячеек
  procedure _ReadMerge();
  var
    i, t, num: integer;
    x1, x2, y1, y2: integer;
    s1, s2: string;
    b: boolean;
    function _GetCoords(var x, y: integer): boolean;
    begin
      result := true;
      x := TZEFormula.GetColIndex(s1);
      if (x < 0) then
        result := false;
      if (not TryStrToInt(s2, y)) then
        result := false
      else
        dec(y);
      b := result;
    end; //_GetCoords

  begin
    x1 := 0;
    y1 := 0;
    x2 := 0;
    y2 := 0;
    while xml.ReadToEndTagByName('mergeCells') do begin
      if xml.IsTagStartOrClosedByName('mergeCell') then begin
        str := xml.Attributes.ItemsByName['ref'];
        t := length(str);
        if (t > 0) then begin
          str := str + ':';
          s1 := '';
          s2 := '';
          b := true;
          num := 0;
          for i := 1 to t + 1 do
          case str[i] of
            'A'..'Z', 'a'..'z': s1 := s1 + str[i];
            '0'..'9': s2 := s2 + str[i];
            ':':
              begin
                inc(num);
                if (num > 2) then begin
                  b := false;
                  break;
                end;
                if (num = 1) then begin
                  if (not _GetCoords(x1, y1)) then
                    break;
                end else begin
                  if (not _GetCoords(x2, y2)) then
                    break;
                end;
                s1 := '';
                s2 := '';
              end;
            else begin
              b := false;
              break;
            end;
          end;

          if (b) then begin
            CheckRow(y1 + 1);
            CheckRow(y2 + 1);
            CheckCol(x1 + 1);
            CheckCol(x2 + 1);
            sheet.MergeCells.AddRectXY(x1, y1, x2, y2);
          end;
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
        sheet.Columns[j-1].AutoFitWidth := colInf[i].bestFit;
        sheet.Columns[j-1].Hidden := colInf[i].hidden;
        sheet.Columns[j-1].WidthPix := colInf[i].width;
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
  var st, s: string;
    i, l, _maxC, _maxR, c, r: integer;
  begin
    c := 0;
    r := 0;
    st := xml.Attributes.ItemsByName['ref'];
    l := Length(st);
    if (l > 0) then begin
      st := st + ':';
      inc(l);
      s := '';
      _maxC := -1;
      _maxR := -1;
      for i := 1 to l do
      if (st[i] = ':') then begin
        if TZEFormula.GetCellCoords(s, c, r) then begin;
          if (c > _maxC) then
            _maxC := c;
          if (r > _maxR) then
            _maxR := r;
        end else
          break;
        s := '';
      end else
        s := s + st[i];
      if (_maxC > 0) then
        CheckCol(_maxC);
      if (_maxR > 0) then
        CheckRow(_maxR);
    end;
  end; //_GetDimension()

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
            sheet.Cell[_c, _r].HRefScreenTip := xml.Attributes.ItemsByName['tooltip'];
            str := xml.Attributes.ItemsByName['r:id'];
            //по r:id подставить ссылку
            for i := 0 to FRelations.Count - 1 do
              if ((FRelations[i].id = str) and (FRelations[i].ftype = TRelationType.rtHyperlink)) then begin
                sheet.Cell[_c, _r].Href := FRelations[i].target;
                break;
              end;
          end;
        //доп. атрибуты:
        //  display - ??
        //  id - id <> r:id??
        //  location - ??
      end;
    end; //while
  end; //_ReadHyperLinks();

  procedure _ReadSheetPr();
  begin
    while xml.ReadToEndTagByName('sheetPr') do begin
      if xml.TagName = 'tabColor' then
        sheet.TabColor := ARGBToColor(xml.Attributes.ItemsByName['rgb']);

      if xml.TagName = 'pageSetUpPr' then
        sheet.FitToPage := ZEStrToBoolean(xml.Attributes.ItemsByName['fitToPage']);

      if xml.TagName = 'outlinePr' then begin
        sheet.ApplyStyles := ZEStrToBoolean(xml.Attributes.ItemsByName['applyStyles']);
        sheet.SummaryBelow := xml.Attributes.ItemsByName['summaryBelow'] <> '0';
        sheet.SummaryRight := xml.Attributes.ItemsByName['summaryRight'] <> '0';
      end;
    end;
  end; //_ReadSheetPr();

  procedure _ReadRowBreaks();
  begin
    sheet.RowBreaks := [];
    while xml.ReadToEndTagByName('rowBreaks') do begin
      if xml.TagName = 'brk' then
        sheet.RowBreaks := sheet.RowBreaks
            + [ StrToIntDef(xml.Attributes.ItemsByName['id'], 0) ];
    end;
  end;

  procedure _ReadColBreaks();
  begin
    sheet.ColBreaks := [];
    while xml.ReadToEndTagByName('colBreaks') do begin
      if xml.TagName = 'brk' then
        sheet.ColBreaks := sheet.ColBreaks
            + [ StrToIntDef(xml.Attributes.ItemsByName['id'], 0) ];
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
        sheet.Selected := sheet.SheetIndex = 0;// s = '1';
        sheet.ViewMode := zvmNormal;
        if xml.Attributes.ItemsByName['view'] = 'pageBreakPreview' then
            sheet.ViewMode := zvmPageBreakPreview;
        sheet.ShowZeros := ZETryStrToBoolean(xml.Attributes.ItemsByName['showZeros'], true);
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

        sheet.SheetOptions.SplitVerticalValue := vValue;
        sheet.SheetOptions.SplitHorizontalValue := hValue;

        sheet.SheetOptions.SplitHorizontalMode := ZSplitNone;
        sheet.SheetOptions.SplitVerticalMode := ZSplitNone;
        if (hValue <> 0) then
          sheet.SheetOptions.SplitHorizontalMode := SplitMode;
        if (vValue <> 0) then
          sheet.SheetOptions.SplitVerticalMode := SplitMode;

        if (sheet.SheetOptions.SplitHorizontalMode = ZSplitSplit) then
          sheet.SheetOptions.SplitHorizontalValue := PointToPixel(hValue/20);
        if (sheet.SheetOptions.SplitVerticalMode = ZSplitSplit) then
          sheet.SheetOptions.SplitVerticalValue := PointToPixel(vValue/20);

      end; //if
    end; //while
  end; //_ReadSheetViews()
  {
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
          sheet.ConditionalFormatting.Add();
          _CF := sheet.ConditionalFormatting[sheet.ConditionalFormatting.Count - 1];
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

//          if (ZEXLSX_getCFCondition(_type, _operator, _CFCondition, _CFOperator)) then
//            _TryApplyCF();
        end; //if
      end; //while
    finally
      SetLength(_formulas, 0);
      FreeAndNil(_tmpStyle);
    end;
  end; //_ReadConditionFormatting
  }
  procedure _ReadHeaderFooter();
  begin
    sheet.SheetOptions.IsDifferentFirst   := ZEStrToBoolean(xml.Attributes['differentFirst']);
    sheet.SheetOptions.IsDifferentOddEven := ZEStrToBoolean(xml.Attributes['differentOddEven']);
    while xml.ReadToEndTagByName('headerFooter') do begin
      if xml.IsTagEndByName('oddHeader') then
        sheet.SheetOptions.Header := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('oddFooter') then
        sheet.SheetOptions.Footer := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('evenHeader') then
        sheet.SheetOptions.EvenHeader := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('evenFooter') then
        sheet.SheetOptions.EvenFooter := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('firstHeader') then
        sheet.SheetOptions.FirstPageHeader := ClenuapXmlTagValue(xml.TextBeforeTag)
      else if xml.IsTagEndByName('firstFooter') then
        sheet.SheetOptions.FirstPageFooter := ClenuapXmlTagValue(xml.TextBeforeTag);
    end;
  end;
begin
  xml := TZsspXMLReaderH.Create();
  try
    xml.AttributesMatch := false;
    if (xml.BeginReadStream(Stream) <> 0) then
      exit;

    //currentPage := FWorkBook.Sheets.Count;
    currentRow := 0;

//    if assigned(SheetRelations) then
//    begin
//      sheet.Title := SheetRelations^.name;
//      if SameText(SheetRelations^.state, 'hidden') then
//        sheet.Visible := svHidden
//      else if SameText(SheetRelations^.state, 'veryhidden') then
//        sheet.Visible := svVeryHidden
//      else
//        sheet.Visible := svVisible;
//    end
//    else
      sheet.Title := '';

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
         sheet.DrawingRid := StrtoIntDef(xml.Attributes.ItemsByName['r:id'].Substring(3), 0);
      end else
      if xml.IsTagClosedByName('pageMargins') then begin
        str := xml.Attributes.ItemsByName['bottom'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.MarginBottom := round(tempReal);
        str := xml.Attributes.ItemsByName['footer'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.FooterMargins.Height := abs(round(tempReal));
        str := xml.Attributes.ItemsByName['header'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.HeaderMargins.Height := abs(round(tempReal));
        str := xml.Attributes.ItemsByName['left'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.MarginLeft := round(tempReal);
        str := xml.Attributes.ItemsByName['right'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.MarginRight := round(tempReal);
        str := xml.Attributes.ItemsByName['top'];
        if (_StrToMM(str, tempReal)) then
          sheet.SheetOptions.MarginTop := round(tempReal);
      end else
      //Настройки страницы
      if xml.IsTagClosedByName('pageSetup') then begin
        //str := xml.Attributes.ItemsByName['blackAndWhite'];
        //str := xml.Attributes.ItemsByName['cellComments'];
        //str := xml.Attributes.ItemsByName['copies'];
        //str := xml.Attributes.ItemsByName['draft'];
        //str := xml.Attributes.ItemsByName['errors'];
        str := xml.Attributes.ItemsByName['firstPageNumber'];
        if (str > '') then
          if (TryStrToInt(str, tempInt)) then
            sheet.SheetOptions.StartPageNumber := tempInt;

        str := xml.Attributes.ItemsByName['fitToHeight'];
        if (str > '') then
          if (TryStrToInt(str, tempInt)) then
            sheet.SheetOptions.FitToHeight := tempInt;

        str := xml.Attributes.ItemsByName['fitToWidth'];
        if (str > '') then
          if (TryStrToInt(str, tempInt)) then
            sheet.SheetOptions.FitToWidth := tempInt;

        //str := xml.Attributes.ItemsByName['horizontalDpi'];
        //str := xml.Attributes.ItemsByName['id'];
        str := xml.Attributes.ItemsByName['orientation'];
        if (str > '') then begin
          sheet.SheetOptions.PortraitOrientation := false;
          if (str = 'portrait') then
            sheet.SheetOptions.PortraitOrientation := true;
        end;

        //str := xml.Attributes.ItemsByName['pageOrder'];

        str := xml.Attributes.ItemsByName['paperSize'];
        if (str > '') then
          if (TryStrToInt(str, tempInt)) then
            sheet.SheetOptions.PaperSize := tempInt;
        //str := xml.Attributes.ItemsByName['paperHeight']; //если утановлены paperHeight и Width, то paperSize игнорируется
        //str := xml.Attributes.ItemsByName['paperWidth'];

        str := xml.Attributes.ItemsByName['scale'];
        sheet.SheetOptions.ScaleToPercent := StrToIntDef(str, 100);
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
          sheet.SheetOptions.CenterHorizontal := ZEStrToBoolean(str);

        str := xml.Attributes.ItemsByName['verticalCentered'];
        if (str > '') then
          sheet.SheetOptions.CenterVertical := ZEStrToBoolean(str);
      end
      else
      if xml.IsTagClosedByName('sheetFormatPr') then
      begin
        str := xml.Attributes.ItemsByName['defaultColWidth'];
        if (str > '') then
          sheet.DefaultColWidth := ZETryStrToFloat(str, sheet.DefaultColWidth);
        str := xml.Attributes.ItemsByName['defaultRowHeight'];
        if (str > '') then
          sheet.DefaultRowHeight := ZETryStrToFloat(str, sheet.DefaultRowHeight);
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
        //_ReadConditionFormatting()
      else
      if xml.IsTagStartByName('headerFooter') then
        _ReadHeaderFooter();
    end; //while
  finally
    xml.Free();
  end;
end;

end.
