unit Excel4Delphi;

interface

uses
  System.Classes, System.SysUtils,
  {$IFDEF FMX}
  FMX.Graphics,
  {$ELSE}
  Vcl.Graphics,
  {$ENDIF}
  {$IFDEF MSWINDOWS}
  Winapi.Windows,
  {$ENDIF}
  System.UITypes, System.Math, RegularExpressions, System.Generics.Collections,
  System.Generics.Defaults, Excel4Delphi.Xml;

var
  ZE_XLSX_APPLICATION: string;

// 1 (topographical point) = 0.3528 mm
const
  _PointToMM: Real = 0.3528;

{$IFDEF FMX}
  DEFAULT_CHARSET = 1;
{$ENDIF}

type
  TObjectList = TObjectList<TObject>;

type
  /// <summary>
  /// Data type of cell
  /// </summary>
  TZCellType = (ZENumber, ZEDateTime, ZEBoolean, ZEString, ZEError, ZEGeneral);

  /// <summary>
  /// Style lines of the cell border
  /// </summary>
  TZBorderType = (ZENone, ZEContinuous, ZEHair, ZEDot, ZEDash, ZEDashDot, ZEDashDotDot, ZESlantDashDot, ZEDouble);

  /// <summary>
  /// Horizontal alignment
  /// </summary>
  TZHorizontalAlignment = (ZHAutomatic, ZHLeft, ZHCenter, ZHRight, ZHFill, ZHJustify, ZHCenterAcrossSelection,
    ZHDistributed, ZHJustifyDistributed);

  /// <summary>
  /// Vertical alignment
  /// </summary>
  TZVerticalAlignment = (ZVAutomatic, ZVTop, ZVBottom, ZVCenter, ZVJustify, ZVDistributed, ZVJustifyDistributed);

  /// <summary>
  /// Fill pattern of the cell
  /// </summary>
  TZCellPattern = (ZPNone, ZPSolid, ZPGray75, ZPGray50, ZPGray25, ZPGray125, ZPGray0625, ZPHorzStripe, ZPVertStripe,
    ZPReverseDiagStripe, ZPDiagStripe, ZPDiagCross, ZPThickDiagCross, ZPThinHorzStripe, ZPThinVertStripe,
    ZPThinReverseDiagStripe, ZPThinDiagStripe, ZPThinHorzCross, ZPThinDiagCross);

  /// <summary>
  /// Borders position.
  /// </summary>
  TZBordersPos = (bpLeft, bpTop, bpRight, bpBottom, bpDiagonalLeft, bpDiagonalRight);

  /// <summary>
  /// Vertical/horizontal split/freeze mode.
  /// </summary>
  TZSplitMode = (ZSplitNone, ZSplitFrozen, ZSplitSplit);

  /// <summary>
  /// View mode.
  /// </summary>
  TZViewMode = (zvmNormal, zvmPageBreakPreview);

  /// <summary>
  /// An angle of the rotation (direction) for a text within a cell.
  /// Nominative range is -90 .. +90, extended is -180 .. +180 (in degrees)
  /// </summary>
  TZCellTextRotate = -180 .. +359;

  TMediaRec = record
    FileName: string;
    Content: TBytes;
  end;

  TZSheet = class;
  TZStyle = class;
  TZFont = class;

  TZExcelColor = record
    RGB: TColor; // rgb
    Indexed: Byte; // 0 - not exists (deprecated but used)
    Theme: Byte; // 0 - not exists
    Tint: Double; // 0 - not exists
    procedure Load(const ARgb, AIndexed, ATheme, ATint: string);
    class operator Equal(const ALeft, ARight: TZExcelColor): Boolean;
    class operator NotEqual(const ALeft, ARight: TZExcelColor): Boolean;
  end;

  TRichString = class(TPersistent)
  private
    FText: string;
    FFont: TZFont;
    // FSCheme: string; // todo: add to TZFont or use "minor" for default
  public
    property Text: string read FText write FText;
    property Font: TZFont read FFont write FFont;

    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function GetHashCode(): Integer; override;
    function Equals(Obj: TObject): Boolean; override;
  end;

  TRichText = class(TPersistent)
  private
    FList: TList<TRichString>;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    property List: TList<TRichString> read FList;
    function GetHashCode(): Integer; override;
    function Equals(Obj: TObject): Boolean; override;
    // function ToHtml(): string;
    // function ToString(): string; override;
  end;

  TRichTextComparer = class(TInterfacedObject, IEqualityComparer<TRichText>)
  public
    function Equals(const left, right: TRichText): Boolean; reintroduce;
    function GetHashCode(const Value: TRichText): Integer; reintroduce;
  end;

  /// <summary>
  /// Cell of spreadsheet.
  /// </summary>
  TZCell = class(TPersistent)
  private
    FFormula: string;
    FData: string;
    FHref: string;
    FHRefScreenTip: string;
    FComment: string;
    FCommentAuthor: string;
    FAlwaysShowComment: Boolean; // default = false;
    FShowComment: Boolean; // default = false
    FCellType: TZCellType;
    FCellStyle: Integer;
    FRichText: TRichText;
    FSheet: TZSheet;
    procedure ApplyStyleValue(proc: TProc<TZStyle>);
    function GetDataAsDouble: Double;
    procedure SetDataAsDouble(const Value: Double);
    procedure SetDataAsInteger(const Value: Integer);
    function GetDataAsInteger: Integer;
    function GetDataAsDateTime(): TDateTime;
    procedure SetDataAsDateTime(const Value: TDateTime);
    procedure SetDataAsString(const Value: string);
    function GetStyle(): TZStyle;
    function GetBgColor: TColor;
    function GetBorderColor(num: TZBordersPos): TColor;
    function GetBorderStyle(num: TZBordersPos): TZBorderType;
    function GetBorderWidht(num: TZBordersPos): Byte;
    function GetFontColor: TColor;
    function GetFontSize: Double;
    function GetFontStyle: TFontStyles;
    function GetHorizontalAlignment: TZHorizontalAlignment;
    function GetNumberFormat: string;
    function GetRotate: TZCellTextRotate;
    function GetVerticalAlignment: TZVerticalAlignment;
    function GetVerticalText: Boolean;
    function GetWrapText: Boolean;
    procedure SetBgColor(const Value: TColor);
    procedure SetBorderColor(num: TZBordersPos; const Value: TColor);
    procedure SetBorderStyle(num: TZBordersPos; const Value: TZBorderType);
    procedure SetBorderWidht(num: TZBordersPos; const Value: Byte);
    procedure SetFontColor(const Value: TColor);
    procedure SetFontSize(const Value: Double);
    procedure SetFontStyle(const Value: TFontStyles);
    procedure SetHorizontalAlignment(const Value: TZHorizontalAlignment);
    procedure SetNumberFormat(const Value: string);
    procedure SetRotate(const Value: TZCellTextRotate);
    procedure SetVerticalAlignment(const Value: TZVerticalAlignment);
    procedure SetVerticalText(const Value: Boolean);
    procedure SetWrapText(const Value: Boolean);
  public
    constructor Create(ASheet: TZSheet); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Clear cell data.
    /// </summary>
    procedure Clear();
    /// <summary>
    /// Current sheet.
    /// </summary>
    property Sheet: TZSheet read FSheet;
    /// <summary>
    /// Current cell style.
    /// </summary>
    property Style: TZStyle read GetStyle;
    /// <summary>
    /// Always show comment. <br />False by default.
    /// </summary>
    property AlwaysShowComment: Boolean read FAlwaysShowComment write FAlwaysShowComment default false;
    /// <summary>
    /// Comment text.
    /// </summary>
    property Comment: string read FComment write FComment;
    /// <summary>
    /// Author of comment.
    /// </summary>
    property CommentAuthor: string read FCommentAuthor write FCommentAuthor;
    /// <summary>
    /// Cell style number. <br />-1 by default.
    /// </summary>
    property CellStyle: Integer read FCellStyle write FCellStyle default -1;
    /// <summary>
    /// Cell type. <br />ZEString by default.
    /// </summary>
    property CellType: TZCellType read FCellType write FCellType default ZEString;
    /// <summary>
    /// Specifies the value of this cell to show on screen.
    /// </summary>
    property Data: string read FData write FData;
    /// <summary>
    /// Formula in R1C1 style.
    /// </summary>
    property Formula: string read FFormula write FFormula;
    /// <summary>
    /// Specifies the URL to link this cell.
    /// </summary>
    property HRef: string read FHref write FHref;
    /// <summary>
    /// Displays the caption of URL to show on screen.
    /// </summary>
    property HRefScreenTip: string read FHRefScreenTip write FHRefScreenTip;
    /// <summary>
    /// Show comment. <br />False by default.
    /// </summary>
    property ShowComment: Boolean read FShowComment write FShowComment default false;
    /// <summary>
    /// Rich formated text.
    /// </summary>
    property RichText: TRichText read FRichText;
    /// <summary>
    /// Present cell data as double value
    /// </summary>
    property AsDouble: Double read GetDataAsDouble write SetDataAsDouble;
    /// <summary>
    /// Present cell data as integer value
    /// </summary>
    property AsInteger: Integer read GetDataAsInteger write SetDataAsInteger;
    /// <summary>
    /// Present cell data as TDateTime value
    /// </summary>
    property AsDateTime: TDateTime read GetDataAsDateTime write SetDataAsDateTime;
    /// <summary>
    /// Present cell data as string value
    /// </summary>
    property AsString: string read FData write SetDataAsString;
    /// <summary>
    /// Vertical content alignment
    /// </summary>
    property VerticalAlignment: TZVerticalAlignment read GetVerticalAlignment write SetVerticalAlignment;
    /// <summary>
    /// Horisontal content alignment
    /// </summary>
    property HorizontalAlignment: TZHorizontalAlignment read GetHorizontalAlignment write SetHorizontalAlignment;
    /// <summary>
    /// Background cell color
    /// </summary>
    property BgColor: TColor read GetBgColor write SetBgColor;
    /// <summary>
    /// Font color
    /// </summary>
    property FontColor: TColor read GetFontColor write SetFontColor;
    /// <summary>
    /// Font size
    /// </summary>
    property FontSize: Double read GetFontSize write SetFontSize;
    /// <summary>
    /// Font style
    /// </summary>
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
    /// <summary>
    /// Border style
    /// </summary>
    property BorderStyle[num: TZBordersPos]: TZBorderType read GetBorderStyle write SetBorderStyle;
    /// <summary>
    /// Border width
    /// </summary>
    property BorderWidht[num: TZBordersPos]: Byte read GetBorderWidht write SetBorderWidht;
    /// <summary>
    /// Border color
    /// </summary>
    property BorderColor[num: TZBordersPos]: TColor read GetBorderColor write SetBorderColor;
    /// <summary>
    /// Word wrap
    /// </summary>
    property WrapText: Boolean read GetWrapText write SetWrapText;
    /// <summary>
    /// Vertical text
    /// </summary>
    property VerticalText: Boolean read GetVerticalText write SetVerticalText;
    /// <summary>
    /// Text rotation
    /// </summary>
    property Rotate: TZCellTextRotate read GetRotate write SetRotate;
    /// <summary>
    /// Number format
    /// </summary>
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    procedure SetBorderAround(borderWidth: Byte; BorderColor: TColor = TColorRec.Black; BorderStyle: TZBorderType = TZBorderType.ZEContinuous);
  end;

  /// <summary>
  /// Border's style.
  /// </summary>
  TZBorderStyle = class(TPersistent)
  private
    FLineStyle: TZBorderType;
    FWeight: Byte;
    FColor: TZExcelColor;
    procedure SetLineStyle(const Value: TZBorderType);
    procedure SetWeight(const Value: Byte);
    procedure SetColor(const Value: TColor);
    function GetColor: TColor;
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <returns>
    /// True, when border equal to Source.
    /// </returns>
    function IsEqual(Source: TPersistent): Boolean; virtual;
    /// <summary>
    /// Line style.
    /// </summary>
    property LineStyle: TZBorderType read FLineStyle write SetLineStyle default ZENone;
    /// <summary>
    /// Specifies the thickness of this border (0-3).
    /// </summary>
    property Weight: Byte read FWeight write SetWeight default 0;
    /// <summary>
    /// Specifies the color of this border.
    /// </summary>
    property Color: TColor read GetColor write SetColor;
    /// <summary>
    /// Specifies the inner color strucure of this border.
    /// </summary>
    property ExcelColor: TZExcelColor read FColor write FColor;
  end;

  /// <summary>
  /// Borders of cell.
  /// </summary>
  TZBorder = class(TPersistent)
  private
    FBorder: array [0 .. 5] of TZBorderStyle;
    procedure SetBorder(num: TZBordersPos; Const Value: TZBorderStyle);
    function GetBorder(num: TZBordersPos): TZBorderStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Set border by border position.
    /// </summary>
    property Border[num: TZBordersPos]: TZBorderStyle read GetBorder write SetBorder; default;
    /// <returns>
    /// True when borders equal Source.
    /// </returns>
    function IsEqual(Source: TPersistent): Boolean; virtual;
    /// <summary>
    /// Left border.
    /// </summary>
    property left: TZBorderStyle index bpLeft read GetBorder write SetBorder;
    /// <summary>
    /// Top border.
    /// </summary>
    property Top: TZBorderStyle index bpTop read GetBorder write SetBorder;
    /// <summary>
    /// Right border.
    /// </summary>
    property right: TZBorderStyle index bpRight read GetBorder write SetBorder;
    /// <summary>
    /// Bottom border.
    /// </summary>
    property Bottom: TZBorderStyle index bpBottom read GetBorder write SetBorder;
    /// <summary>
    /// Diagonal from upper left to lower right.
    /// </summary>
    property DiagonalLeft: TZBorderStyle index bpDiagonalLeft read GetBorder write SetBorder;
    /// <summary>
    /// Diagonal from lower left to upper right.
    /// </summary>
    property DiagonalRight: TZBorderStyle index bpDiagonalRight read GetBorder write SetBorder;
  end;

  /// <summary>
  /// Specifies how text is aligned within the cell.
  /// </summary>
  TZAlignment = class(TPersistent)
  private
    FHorizontal: TZHorizontalAlignment;
    FVertical: TZVerticalAlignment;
    FIndent: Integer;
    FRotate: TZCellTextRotate;
    FWrapText: Boolean;
    FShrinkToFit: Boolean;
    FVerticalText: Boolean;
    procedure SetHorizontal(const Value: TZHorizontalAlignment);
    procedure SetIndent(const Value: Integer);
    procedure SetRotate(const Value: TZCellTextRotate);
    procedure SetShrinkToFit(const Value: Boolean);
    procedure SetVertical(const Value: TZVerticalAlignment);
    procedure SetVerticalText(const Value: Boolean);
    procedure SetWrapText(const Value: Boolean);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <returns>
    /// True when all properties equal to Source's properties.
    /// </returns>
    function IsEqual(Source: TPersistent): Boolean; virtual;
    /// <summary>
    /// Specifies how text is aligned by horizontally within the cell.
    /// </summary>
    property Horizontal: TZHorizontalAlignment read FHorizontal write SetHorizontal default ZHAutomatic;
    /// <summary>
    /// Specifies how far the cell's text is indented.
    /// </summary>
    property Indent: Integer read FIndent write SetIndent default 0;
    /// <summary>
    /// Specifies the rotation of the text within the cell (from -90 to 90).
    /// </summary>
    property Rotate: TZCellTextRotate read FRotate write SetRotate;
    /// <summary>
    /// If True then the text size will shrunk so to all of the text fits within the cell.
    /// </summary>
    property ShrinkToFit: Boolean read FShrinkToFit write SetShrinkToFit default false;
    /// <summary>
    /// Specifies how text is aligned by vertically within the cell.
    /// </summary>
    property Vertical: TZVerticalAlignment read FVertical write SetVertical default ZVAutomatic;
    /// <summary>
    /// If True each letter is drawn horizontally, one above the other.
    /// </summary>
    property VerticalText: Boolean read FVerticalText write SetVerticalText default false;
    /// <summary>
    /// Specifies whether the text in cell should wrap at the cell boundary.
    /// </summary>
    property WrapText: Boolean read FWrapText write SetWrapText default false;
  end;

  /// <summary>
  /// Excel Font - https://docs.microsoft.com/ru-ru/office/vba/api/excel.font(object)
  /// </summary>
  TZFont = class(TPersistent)
  private
    // todo: double underLine
    FColor: TZExcelColor;
    FSize: Double;
    FCharset: TFontCharset;
    FName: TFontName;
    FStyle: TFontStyles;
    function GetColor: TColor;
    procedure SetColor(const Value: TColor);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    procedure AssignTo(Dest: TPersistent); override;
    function GetHashCode(): Integer; override;
    property Color: TColor read GetColor write SetColor;
    property Size: Double read FSize write FSize;
    property Charset: TFontCharset read FCharset write FCharset;
    property Name: TFontName read FName write FName;
    property Style: TFontStyles read FStyle write FStyle;
    property ExcelColor: TZExcelColor read FColor write FColor;
  end;

  /// <summary>
  /// Cell style.
  /// </summary>
  TZStyle = class(TPersistent)
  private
    FBorder: TZBorder;
    FAlignment: TZAlignment;
    FFont: TZFont;
    FBGColor: TZExcelColor;
    FPatternColor: TColor;
    FCellPattern: TZCellPattern;
    FNumberFormatId: Integer;
    FNumberFormat: string;
    FProtect: Boolean;
    FHideFormula: Boolean;
    FSuperscript: Boolean;
    FSubscript: Boolean;
    procedure SetFont(const Value: TZFont);
    procedure SetBorder(const Value: TZBorder);
    procedure SetAlignment(const Value: TZAlignment);
    procedure SetBgColor(const Value: TColor);
    procedure SetPatternColor(const Value: TColor);
    procedure SetCellPattern(const Value: TZCellPattern);
    procedure SetSuperscript(const Value: Boolean);
    procedure SetSubscript(const Value: Boolean);
    function GetBgColor: TColor;
  protected
    procedure SetNumberFormat(const Value: string); virtual;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// True when style equal Source.
    /// </summary>
    function IsEqual(Source: TPersistent): Boolean; virtual;
    /// <summary>
    /// Cell font.
    /// </summary>
    property Font: TZFont read FFont write SetFont;
    /// <summary>
    /// Cell borders.
    /// </summary>
    property Border: TZBorder read FBorder write SetBorder;
    /// <summary>
    /// Specifies how text is aligned within the cell.
    /// </summary>
    property Alignment: TZAlignment read FAlignment write SetAlignment;
    /// <summary>
    /// Background color of the cell. <br />clWindow by default.
    /// </summary>
    property BgColor: TColor read GetBgColor write SetBgColor;
    /// <summary>
    /// Background color of the cell. <br />clWindow by default.
    /// </summary>
    property BGExcelColor: TZExcelColor read FBGColor write FBGColor;
    /// <summary>
    /// Color of fill pattern. <br />clWindow by default.
    /// </summary>
    property PatternColor: TColor read FPatternColor write SetPatternColor default TColorRec.cWINDOW;
    /// <summary>
    /// Indicates whether or not this cell is protected. <br />True by default.
    /// </summary>
    property Protect: Boolean read FProtect write FProtect default true;
    /// <summary>
    /// Indicates whether or not this cell's formula should be hidden when sheet protection is enabled. <br />False by default.
    /// </summary>
    property HideFormula: Boolean read FHideFormula write FHideFormula default false;
    /// <summary>
    /// Indicates whether or not the text is slightly above the baseline. <br />False by default.
    /// </summary>
    property Superscript: Boolean read FSuperscript write SetSuperscript default false;
    /// <summary>
    /// Indicates whether or not the text is slightly below the baseline. <br />False by default.
    /// </summary>
    property Subscript: Boolean read FSubscript write SetSubscript default false;
    /// <summary>
    /// Fill pattern of the cell. <br />ZPNone by default.
    /// </summary>
    property CellPattern: TZCellPattern read FCellPattern write SetCellPattern default ZPNone;
    /// <summary>
    /// Defines the number format that should be in cells referencing this style.
    /// </summary>
    property NumberFormat: string read FNumberFormat write SetNumberFormat;
    property NumberFormatId: Integer read FNumberFormatId write FNumberFormatId default -1;
  end;

  /// <summary>
  /// Contains styles of the document.
  /// </summary>
  TZStyles = class(TPersistent)
  private
    FDefaultStyle: TZStyle;
    FStyles: array of TZStyle;
    FCount: Integer;
    procedure SetDefaultStyle(const Value: TZStyle);
    function GetStyle(num: Integer): TZStyle;
    procedure SetStyle(num: Integer; const Value: TZStyle);
    procedure SetCount(const Value: Integer);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Add a Style.
    /// </summary>
    /// <param name="Style">
    /// Style.
    /// </param>
    /// <param name="CheckMatch">
    /// Checks the coincidence of this style with introduced in earlier styles.
    /// <br />Return number of added (or, if CheckMatch = True, previously introduced) style.
    /// </param>
    function Add(const Style: TZStyle; CheckMatch: Boolean = false): Integer;
    /// <summary>
    /// Delete all styles.
    /// </summary>
    procedure Clear(); virtual;
    /// <summary>
    /// Delete style num, styles with a larger number are shifting.
    /// </summary>
    /// <param name="num">
    /// Style number
    /// </param>
    /// <returns>
    /// Return: 0 - if successfully deleted, -1 - can not delete style
    /// </returns>
    function DeleteStyle(num: Integer): Integer; virtual;
    /// <summary>
    /// Find number to match the Style introduced in earlier styles.
    /// </summary>
    /// <param name="Style">
    /// Style.
    /// </param>
    /// <returns>
    /// Return: -2 - if style not found, -1 - if style = DefaultStyle, 0..Count-1 - Style
    /// </returns>
    function Find(const Style: TZStyle): Integer;
    /// <summary>
    /// Style num (-1 - DefaultStyle).
    /// </summary>
    property Items[num: Integer]: TZStyle read GetStyle write SetStyle; default;
    /// <summary>
    /// Count styles in the document.
    /// </summary>
    property Count: Integer read FCount write SetCount;
    /// <summary>
    /// Default style ( = Items[-1]).
    /// </summary>
    property DefaultStyle: TZStyle read FDefaultStyle write SetDefaultStyle;
  end;

  TZMergeArea = record
    left, Top, right, Bottom: Integer;
    constructor Create(ALeft, ATop, ARight, ABottom: Integer);
  end;

  /// <summary>
  /// Merged cells
  /// </summary>
  TZMergeCells = class
  private
    FSheet: TZSheet;
    FCount: Integer;
    FMergeArea: TArray<TZMergeArea>;
    function GetItem(num: Integer): TZMergeArea;
    procedure SetItem(num: Integer; const rect: TZMergeArea);
  public
    constructor Create(ASheet: TZSheet); virtual;
    destructor Destroy(); override;
    /// <summary>
    /// Adds a merged cell enclosed into rectangle.
    /// </summary>
    function AddRect(Rct: TZMergeArea): Byte;
    /// <summary>
    /// Adds a merged cell enclosed into rectangle
    /// </summary>
    function AddRectXY(left, Top, right, Bottom: Integer): Byte;
    /// <summary>
    /// Delete merged cell num. <br />Return True if the cell is successfully deleted.
    /// </summary>
    function DeleteItem(num: Integer): Boolean;
    /// <summary>
    /// Returns the number of merged cell, in which the cell [ACol, ARow] is top left.
    /// If returns a negative value - there is no such area.
    /// </summary>
    function InLeftTopCorner(ACol, ARow: Integer): Integer;
    /// <summary>
    /// Returns the number of merged cell that includes cell [ACol, ARow].
    /// If returns a negative value - cell [ACol, ARow] is not contained in the Merged area.
    /// </summary>
    function InMergeRange(ACol, ARow: Integer): Integer;
    /// <summary>
    /// True if region cross with merged cells.
    /// </summary>
    function IsCrossWithArea(AID, AC1, AR1, AC2, AR2: Integer): Boolean;
    /// <summary>
    /// Rows count in merged region.
    /// </summary>
    function MergedRows(ACol, ARow: Integer): Integer;
    /// <summary>
    /// Columns count in merged region.
    /// </summary>
    function MergedCols(ACol, ARow: Integer): Integer;
    /// <summary>
    /// Removes all merged cells.
    /// </summary>
    procedure Clear();
    /// <summary>
    /// Merged regions count.
    /// </summary>
    property Count: Integer read FCount;
    /// <summary>
    /// Current merged region.
    /// </summary>
    property Items[num: Integer]: TZMergeArea read GetItem write SetItem; default;
  end;

  TZCellColumn = array of TZCell;

  TZWorkBook = class;

  /// <summary>
  /// Common options for columns and rows. Ancestor for TZColOptions and TZRowOptions.
  /// </summary>
  TZRowColOptions = class(TPersistent)
  private
    FSheet: TZSheet;
    FHidden: Boolean;
    FStyleID: Integer;
    FSize: Real;
    FAuto: Boolean;
    FBreaked: Boolean;
    FOutlineLevel: Integer;
  protected
    function GetAuto(): Boolean;
    procedure SetAuto(Value: Boolean);
    function GetSizePoint(): Real;
    procedure SetSizePoint(Value: Real);
    function GetSizeMM(): Real;
    procedure SetSizeMM(Value: Real);
    function GetSizePix(): Integer; virtual; abstract;
    procedure SetSizePix(Value: Integer); virtual; abstract;
  public
    constructor Create(ASheet: TZSheet); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// True specifies that column or row is hidden. <br />False (not hidden) by default.
    /// </summary>
    property Hidden: Boolean read FHidden write FHidden default false;
    /// <summary>
    /// Specifies a style for column or row. <br />-1 by default.
    /// </summary>
    property StyleID: Integer read FStyleID write FStyleID default -1;
    /// <summary>
    /// Page break after column or row. <br />False (no break) by default.
    /// </summary>
    property Breaked: Boolean read FBreaked write FBreaked default false;

    property OutlineLevel: Integer read FOutlineLevel write FOutlineLevel;
  end;

  /// <summary>
  /// Column options.
  /// </summary>
  TZColOptions = class(TZRowColOptions)
  protected
    function GetSizePix(): Integer; override;
    procedure SetSizePix(Value: Integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    /// <summary>
    /// If True, it means that this column should be autosized.
    /// </summary>
    property AutoFitWidth: Boolean read GetAuto write SetAuto;
    /// <summary>
    /// Column width in points.
    /// </summary>
    property Width: Real read GetSizePoint write SetSizePoint;
    /// <summary>
    /// Column width in mm.
    /// </summary>
    property WidthMM: Real read GetSizeMM write SetSizeMM;
    /// <summary>
    /// Column width in pixels.
    /// </summary>
    property WidthPix: Integer read GetSizePix write SetSizePix;
  end;

  /// <summary>
  /// Row options.
  /// </summary>
  TZRowOptions = class(TZRowColOptions)
  protected
    function GetSizePix(): Integer; override;
    procedure SetSizePix(Value: Integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    /// <summary>
    /// If True, it means that this row should be autosized.
    /// </summary>
    property AutoFitHeight: Boolean read GetAuto write SetAuto;
    /// <summary>
    /// Row height in points (1 point = 1/72" = 0.3528 mm).
    /// </summary>
    property Height: Real read GetSizePoint write SetSizePoint;
    /// <summary>
    /// Row height in mm.
    /// </summary>
    property HeightMM: Real read GetSizeMM write SetSizeMM;
    /// <summary>
    /// Row height in pixels.
    /// </summary>
    property HeightPix: Integer read GetSizePix write SetSizePix;
  end;

  /// <summary>
  /// Repeating columns or rows when printing the sheet
  /// </summary>
  TZSheetPrintTitles = class(TPersistent)
  private
    FOwner: TZSheet;
    FColumns: Boolean;
    FActive: Boolean;
    FTill: word;
    FFrom: word;
    procedure SetActive(const Value: Boolean);
    procedure SetFrom(const Value: word);
    procedure SetTill(const Value: word);
    function Valid(const AFrom, ATill: word): Boolean;
    procedure RequireValid(const AFrom, ATill: word);
  public
    procedure Assign(Source: TPersistent); override;
    constructor Create(const owner: TZSheet; const ForColumns: Boolean);
    function ToString: string; override;
    property From: word read FFrom write SetFrom;
    property Till: word read FTill write SetTill;
    property Active: Boolean read FActive write SetActive;
  end;

  /// <summary>
  /// Footer or header margins.
  /// </summary>
  TZHeaderFooterMargins = class(TPersistent)
  private
    FMarginTopBottom: word; // Bottom or top margin
    FMarginLeft: word;
    FMarginRight: word;
    FHeight: word;
    FUseAutoFitHeight: Boolean; // If true then in LibreOffice Calc
                                  // in window "Page style settings" on tabs "Header" or "Footer"
                                  // check box "AutoFit height" will be checked.
  public
    constructor Create();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    /// <summary>
    /// Bottom / top margin of the footer / header in mm. <br />13 by default.
    /// </summary>
    property MarginTopBottom: word read FMarginTopBottom write FMarginTopBottom default 13;
    /// <summary>
    /// Left margin in mm. <br />0 by default.
    /// </summary>
    property MarginLeft: word read FMarginLeft write FMarginLeft default 0;
    /// <summary>
    /// Right margin in mm. <br />0 by default.
    /// </summary>
    property MarginRight: word read FMarginRight write FMarginRight default 0;
    /// <summary>
    /// Height of footer / header in mm. <br />7 by default.
    /// </summary>
    property Height: word read FHeight write FHeight default 7;
    /// <summary>
    /// Automatically fit height. <br />true by default.
    /// </summary>
    property UseAutoFitHeight: Boolean read FUseAutoFitHeight write FUseAutoFitHeight default true;
  end;

  /// <summary>
  /// Inherited from TPersistent. Sheet options.
  /// </summary>
  TZSheetOptions = class(TPersistent)
  private
    FActiveCol: word;
    FActiveRow: word;
    FMarginBottom: word;
    FMarginLeft: word;
    FMarginTop: word;
    FMarginRight: word;
    FHeaderMargins: TZHeaderFooterMargins;
    FFooterMargins: TZHeaderFooterMargins;
    FPortraitOrientation: Boolean;
    FCenterHorizontal: Boolean;
    FCenterVertical: Boolean;
    FStartPageNumber: Integer;
    FDifferentOddEven: Boolean;
    FDifferentFirst: Boolean;

    FHeader: string; // Header for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FFooter: string; // Footer for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FEvenHeader: string; // Header for even pages. Used only if IsEvenFooterEqual = false
    FEvenFooter: string; // Footer for even pages. Used only if IsEvenFooterEqual = false
    FFirstPageHeader: string;
    FFirstPageFooter: string;

    FHeaderBGColor: TColor;
    FFooterBGColor: TColor;

    FScaleToPercent: Integer; // Document must be scaled to percentage value (100 - no scale)
    FScaleToPages: Integer; // Document scaled to fit a number of pages (1 - no scale)

    FPaperSize: Byte;
    FPaperWidth: Integer;
    FPaperHeight: Integer;

    FFitToHeight: Integer; // Number of vertical pages to fit on
    FFitToWidth: Integer; // Number of horizontal pages to fit on

    FSplitVerticalMode: TZSplitMode;
    FSplitHorizontalMode: TZSplitMode;
    FSplitVerticalValue: Integer; // Вроде можно вводить отрицательные
    FSplitHorizontalValue: Integer; // Измеряться будут:
                                        // в пикселях, если SplitMode = ZSplitSplit
                                        // в кол-ве строк/столбцов, если SplitMode = ZSplitFrozen
                                        // Если SplitMode = ZSplitNone, то фиксация столбцов/ячеек не работает
    function GetHeaderMargin(): word;
    procedure SetHeaderMargin(Value: word);
    function GetFooterMargin(): word;
    procedure SetFooterMargin(Value: word);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Column number with active cell. <br />0 by default.
    /// </summary>
    property ActiveCol: word read FActiveCol write FActiveCol default 0;
    /// <summary>
    /// Row number with active cell. <br />0 by default.
    /// </summary>
    property ActiveRow: word read FActiveRow write FActiveRow default 0;
    /// <summary>
    /// Specifies the bottom margin on the page in millimeters. <br />25 mm by default.
    /// </summary>
    property MarginBottom: word read FMarginBottom write FMarginBottom default 25;
    /// <summary>
    /// Specifies the left margin on the page in millimeters 19mm (1.9sm) by default.
    /// </summary>
    property MarginLeft: word read FMarginLeft write FMarginLeft default 19;
    /// <summary>
    /// Specifies the top margin on the page in millimeters. <br />25 mm by default.
    /// </summary>
    property MarginTop: word read FMarginTop write FMarginTop default 25;
    /// <summary>
    /// Specifies the right margin on the page in millimeters. <br />20 mm by default.
    /// </summary>
    property MarginRight: word read FMarginRight write FMarginRight default 20;
    /// <summary>
    /// Paper Size (Paper size table). <br />9 (A4) by default.
    /// </summary>
    property PaperSize: Byte read FPaperSize write FPaperSize default 9;
    /// <summary>
    /// Paper width in mm. Used only when PaperSize = 0!
    /// </summary>
    property PaperWidth: Integer read FPaperWidth write FPaperWidth default 0;
    /// <summary>
    /// Paper height in mm. Used only when PaperSize = 0!
    /// </summary>
    property PaperHeight: Integer read FPaperHeight write FPaperHeight default 0;
    property FitToHeight: Integer read FFitToHeight write FFitToHeight default -1;
    property FitToWidth: Integer read FFitToWidth write FFitToWidth default -1;
    /// <summary>
    /// Specifies the orientation of the page (True - Portrait, False - Landscape). <br />True by default.
    /// </summary>
    property PortraitOrientation: Boolean read FPortraitOrientation write FPortraitOrientation default true;
    /// <summary>
    /// If True, the document should be centered horizontally on the page. <br />False by default.
    /// </summary>
    property CenterHorizontal: Boolean read FCenterHorizontal write FCenterHorizontal default false;
    /// <summary>
    /// If True, the document should be centered vertically on the page. <br />False by default.
    /// </summary>
    property CenterVertical: Boolean read FCenterVertical write FCenterVertical default false;
    /// <summary>
    /// Specifies the starting page number for print. <br />1 by default.
    /// </summary>
    property StartPageNumber: Integer read FStartPageNumber write FStartPageNumber default 1;
    /// <summary>
    /// The size of header in millimeters. <br />13 mm by default. <br />
    /// </summary>
    /// <remarks>
    /// deprecated 'Use HeaderMargins.Height!'
    /// </remarks>
    property HeaderMargin: word read GetHeaderMargin write SetHeaderMargin default 13; // deprecated!
    /// <summary>
    /// The size of footer in millimeters. <br />13 mm by default. <br />
    /// </summary>
    /// <remarks>
    /// deprecated 'Use FooterMargins.Height!'
    /// </remarks>
    property FooterMargin: word read GetFooterMargin write SetFooterMargin default 13; // deprecated!
    /// <summary>
    /// Sizes and margins for header in mm.
    /// </summary>
    property HeaderMargins: TZHeaderFooterMargins read FHeaderMargins;
    /// <summary>
    /// Sizes and margins for footer in mm.
    /// </summary>
    property FooterMargins: TZHeaderFooterMargins read FFooterMargins;
    /// <summary>
    ///
    /// </summary>
    property IsDifferentFirst: Boolean read FDifferentFirst write FDifferentFirst;
    /// <summary>
    ///
    /// </summary>
    property IsDifferentOddEven: Boolean read FDifferentOddEven write FDifferentOddEven;
    property Header: string read FHeader write FHeader;
    property Footer: string read FFooter write FFooter;
    property EvenHeader: string read FEvenHeader write FEvenHeader;
    property EvenFooter: string read FEvenFooter write FEvenFooter;
    property FirstPageHeader: string read FFirstPageHeader write FFirstPageHeader;
    property FirstPageFooter: string read FFirstPageFooter write FFirstPageFooter;
    /// <summary>
    /// Background color for header. <br />clWindow by default.
    /// </summary>
    property HeaderBGColor: TColor read FHeaderBGColor write FHeaderBGColor default TColorRec.cWINDOW;
    /// <summary>
    /// Background color for footer. <br />clWindow by default.
    /// </summary>
    property FooterBGColor: TColor read FFooterBGColor write FFooterBGColor default TColorRec.cWINDOW;
    /// <summary>
    /// Document must be scaled to percentage value (100 - no scale). <br />100 by default.
    /// </summary>
    property ScaleToPercent: Integer read FScaleToPercent write FScaleToPercent default 100;
    /// <summary>
    /// Document scaled to fit a number of pages (1 - no scale). <br />1 by default.
    /// </summary>
    property ScaleToPages: Integer read FScaleToPages write FScaleToPages default 1;
    /// <summary>
    /// Vertical columns split/freeze mode.
    /// Does the same thing as LibreOffice Calc commands "Window - Freeze"/ "Window - Split" <br />ZSplitNone by default.
    /// </summary>
    property SplitVerticalMode: TZSplitMode read FSplitVerticalMode write FSplitVerticalMode default ZSplitNone;
    /// <summary>
    /// Horizontal rows split/freeze mode. <br />ZSplitNone by default.
    /// </summary>
    property SplitHorizontalMode: TZSplitMode read FSplitHorizontalMode write FSplitHorizontalMode default ZSplitNone;
    /// <summary>
    /// If SplitVerticalMode = ZSplitFrozen then count of column for freeze.
    /// If SplitVerticalMode = ZSplitSplit then size in pixels. <br />0 by default.
    /// </summary>
    property SplitVerticalValue: Integer read FSplitVerticalValue write FSplitVerticalValue;
    /// <summary>
    /// If SplitHorizontalMode = ZSplitFrozen then count of rows for freeze.
    /// If SplitHorizontalMode = ZSplitSplit then size if pixels. <br />0 by default.
    /// </summary>
    property SplitHorizontalValue: Integer read FSplitHorizontalValue write FSplitHorizontalValue;
  end;

  /// <summary>
  /// Conditions
  /// </summary>
  TZCondition = (ZCFIsTrueFormula, ZCFCellContentIsBetween, ZCFCellContentIsNotBetween, ZCFCellContentOperator,
    ZCFNumberValue, ZCFString, ZCFBoolTrue, ZCFBoolFalse, ZCFFormula, ZCFContainsText, ZCFNotContainsText,
    ZCFBeginsWithText, ZCFEndsWithText, ZCFCellIsEmpty, ZCFDuplicate, ZCFUnique, ZCFAboveAverage, ZCFBellowAverage,
    ZCFAboveEqualAverage, ZCFBelowEqualAverage, ZCFTopElements, ZCFBottomElements, ZCFTopPercent, ZCFBottomPercent,
    ZCFIsError, ZCFIsNoError);

  // Оператор для условного форматирования
  TZConditionalOperator = (ZCFOpGT, // > (Greater Than)
    ZCFOpLT, // < (Less Than)
    ZCFOpGTE, // >= (Greater or Equal)
    ZCFOpLTE, // <= (Less or Equal)
    ZCFOpEqual, // = (Equal)
    ZCFOpNotEqual // <> (Not Equal)
    );

  /// <summary>
  /// Conditional formatting
  /// </summary>
  TZConditionalStyleItem = class(TPersistent)
  private
    FCondition: TZCondition; // условие
    FConditionOperator: TZConditionalOperator; // Оператор
    FValue1: String;
    FValue2: String;
    FApplyStyleID: Integer; // номер применяемого стиля

                                                // Базовая ячейка (только для формул):
    FBaseCellPageIndex: Integer; // Номер страницы для адреса базовой ячейки
                                                // -1 - текущая страница
    FBaseCellRowIndex: Integer; // Номер строки
    FBaseCellColumnIndex: Integer; // Номер столбца
  public
    constructor Create(); virtual;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    property ApplyStyleID: Integer read FApplyStyleID write FApplyStyleID;
    property BaseCellColumnIndex: Integer read FBaseCellColumnIndex write FBaseCellColumnIndex;
    property BaseCellPageIndex: Integer read FBaseCellPageIndex write FBaseCellPageIndex;
    property BaseCellRowIndex: Integer read FBaseCellRowIndex write FBaseCellRowIndex;
    property Condition: TZCondition read FCondition write FCondition;
    property ConditionOperator: TZConditionalOperator read FConditionOperator write FConditionOperator;
    property Value1: String read FValue1 write FValue1;
    property Value2: String read FValue2 write FValue2;
  end;

  /// <summary>
  /// Conditional formatting area
  /// </summary>
  TZConditionalAreaItem = class(TPersistent)
  private
    FRow: Integer;
    FColumn: Integer;
    FWidth: Integer;
    FHeight: Integer;
    procedure SetRow(Value: Integer);
    procedure SetColumn(Value: Integer);
    procedure SetWidth(Value: Integer);
    procedure SetHeight(Value: Integer);
  public
    constructor Create(); overload; virtual;
    constructor Create(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer); overload; virtual;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    property Row: Integer read FRow write SetRow;
    property Column: Integer read FColumn write SetColumn;
    property Width: Integer read FWidth write SetWidth;
    property Height: Integer read FHeight write SetHeight;
  end;

  // Области для применения условного форматирования
  TZConditionalAreas = class(TPersistent)
  private
    FCount: Integer;
    FItems: array of TZConditionalAreaItem;
    procedure SetCount(Value: Integer);
    function GetItem(num: Integer): TZConditionalAreaItem;
    procedure SetItem(num: Integer; Value: TZConditionalAreaItem);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalAreaItem; overload;
    function Add(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer): TZConditionalAreaItem; overload;
    procedure Assign(Source: TPersistent); override;
    procedure Delete(num: Integer);
    function IsCellInArea(ColumnNum, RowNum: Integer): Boolean;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    property Count: Integer read FCount write SetCount;
    property Items[num: Integer]: TZConditionalAreaItem read GetItem write SetItem; default;
  end;

  // Условное форматирование: список условий
  TZConditionalStyle = class(TPersistent)
  private
    FCount: Integer; // кол-во условий
    FMaxCount: Integer;
    FAreas: TZConditionalAreas;
    FConditions: array of TZConditionalStyleItem;
    function GetItem(num: Integer): TZConditionalStyleItem;
    procedure SetItem(num: Integer; Value: TZConditionalStyleItem);
    procedure SetCount(Value: Integer);
    procedure SetAreas(Value: TZConditionalAreas);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyleItem; overload;
    function Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem; overload;
    procedure Delete(num: Integer);
    procedure Insert(num: Integer); overload;
    procedure Insert(num: Integer; StyleItem: TZConditionalStyleItem); overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    property Areas: TZConditionalAreas read FAreas write SetAreas;
    property Count: Integer read FCount write SetCount;
    property Items[num: Integer]: TZConditionalStyleItem read GetItem write SetItem; default;
  end;

  TZConditionalFormatting = class(TPersistent)
  private
    FStyles: array of TZConditionalStyle;
    FCount: Integer;
    procedure SetCount(Value: Integer);
    function GetItem(num: Integer): TZConditionalStyle;
    procedure SetItem(num: Integer; Value: TZConditionalStyle);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyle; overload;
    function Add(Style: TZConditionalStyle): TZConditionalStyle; overload;
    function Add(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer): TZConditionalStyle; overload;
    procedure Clear();
    function Delete(num: Integer): Boolean;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): Boolean; virtual;
    property Count: Integer read FCount write SetCount;
    property Items[num: Integer]: TZConditionalStyle read GetItem write SetItem; default;
  end;

  // Transformations that can applied to a chart/image.
  TZETransform = class(TPersistent)
  private
    FRotate: Double;
    FScaleX: Double;
    FScaleY: Double;
    FSkewX: Double;
    FSkewY: Double;
    FTranslateX: Double;
    FTranslateY: Double;
  public
    constructor Create();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean;
    procedure Clear();
  published
    property Rotate: Double read FRotate write FRotate;
    property ScaleX: Double read FScaleX write FScaleX;
    property ScaleY: Double read FScaleY write FScaleY;
    property SkewX: Double read FSkewX write FSkewX;
    property SkewY: Double read FSkewY write FSkewY;
    property TranslateX: Double read FTranslateX write FTranslateX;
    property TranslateY: Double read FTranslateY write FTranslateY;
  end;

  // Common frame ancestor for Charts and Images
  TZECommonFrameAncestor = class(TPersistent)
  private
    FX: Integer;
    FY: Integer;
    FWidth: Integer;
    FHeight: Integer;
    FTransform: TZETransform;
  protected
    procedure SetX(Value: Integer); virtual;
    procedure SetY(Value: Integer); virtual;
    procedure SetWidth(Value: Integer); virtual;
    procedure SetHeight(Value: Integer); virtual;
    procedure SetTransform(const Value: TZETransform); virtual;
  public
    constructor Create(); overload; virtual;
    constructor Create(AX, AY, AWidth, AHeight: Integer); overload; virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
    property X: Integer read FX write SetX default 0;
    property Y: Integer read FY write SetY default 0;
    property Width: Integer read FWidth write SetWidth default 10;
    property Height: Integer read FHeight write SetHeight default 10;
    property Transform: TZETransform read FTransform write SetTransform;
  end;

  // Possible chart types
  TZEChartType = (ZEChartTypeArea, ZEChartTypeBar, ZEChartTypeBubble, ZEChartTypeCircle, ZEChartTypeGantt,
    ZEChartTypeLine, ZEChartTypeRadar, ZEChartTypeRing, ZEChartTypeScatter, ZEChartTypeStock, ZEChartTypeSurface);

  // Specifies the rendering of bars for 3D bar charts
  TZEChartSolidType = (ZEChartSolidTypeCone, ZEChartSolidTypeCuboid, ZEChartSolidTypeCylinder, ZEChartSolidTypePyramid);

  // Type of symbols for a data point in a chart
  TZEChartSymbolType = (ZEChartSymbolTypeNone, // No symbol should be used
    ZEChartSymbolTypeAutomatic, // Auto select from TZEChartSymbol
    ZEChartSymbolTypeNamedSymbol // Use selected from TZEChartSymbol
    // ZEChartSymbolTypeImage          //not for now
    );

  // Symbol to be used for a data point in a chart, used only for chart type = ZEChartSymbolTypeNamedSymbol
  TZEChartSymbol = (ZEChartSymbolArrowDown, ZEChartSymbolArrowUp, ZEChartSymbolArrowRight, ZEChartSymbolArrowLeft,
    ZEChartSymbolAsterisk, ZEChartSymbolCircle, ZEChartSymbolBowTie, ZEChartSymbolDiamond, ZEChartSymbolHorizontalBar,
    ZEChartSymbolHourglass, ZEChartSymbolPlus, ZEChartSymbolStar, ZEChartSymbolSquare, ZEChartSymbolX,
    ZEChartSymbolVerticalBar);

  // Position of Legend in chart
  TZEChartLegendPosition = (ZELegendBottom, // Legend below the plot area
    ZELegendEnd, // Legend on the right side of the plot area
    ZELegendStart, // Legend on the left side of the plot area
    ZELegendTop, // Legend above the plot area
    ZELegendBottomEnd, // Legend in the bottom right corner of the plot area
    ZELegendBottomStart, // Legend in the bottom left corner
    ZELegendTopEnd, // Legend in the top right corner
    ZELegendTopStart // Legend in the top left corner
    );

  // Alignment of a legend
  TZELegendAlign = (ZELegengAlignStart, // Legend aligned at the beginning of a plot area (left or top)
    ZELegendAlignCenter, // Legend aligned at the center of a plot area
    ZELegendAlignEnd // Legend aligned at the end of a plot area (right or bottom)
    );

  // Range item
  TZEChartRangeItem = class(TPersistent)
  private
    FSheetNum: Integer;
    FCol: Integer;
    FRow: Integer;
    FWidth: Integer;
    FHeight: Integer;
  public
    constructor Create(); overload;
    constructor Create(ASheetNum: Integer; ACol, ARow, AWidth, AHeight: Integer); overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
  published
    property SheetNum: Integer read FSheetNum write FSheetNum;
    property Col: Integer read FCol write FCol;
    property Row: Integer read FRow write FRow;
    property Width: Integer read FWidth write FWidth;
    property Height: Integer read FHeight write FHeight;
  end;

  // Store for Range items
  TZEChartRange = class(TPersistent)
  private
    FItems: array of TZEChartRangeItem;
    FCount: Integer;
  protected
    function GetItem(num: Integer): TZEChartRangeItem;
    procedure SetItem(num: Integer; const Value: TZEChartRangeItem);
  public
    constructor Create();
    destructor Destroy(); override;
    function Add(): TZEChartRangeItem; overload;
    function Add(const ItemForClone: TZEChartRangeItem): TZEChartRangeItem; overload;
    function Delete(num: Integer): Boolean;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
    property Items[num: Integer]: TZEChartRangeItem read GetItem write SetItem; default;
    property Count: Integer read FCount;
  end;

  // Title or Subtitle for chart, axis etc
  TZEChartTitleItem = class(TPersistent)
  private
    FText: string;
    FFont: TZFont;
    FRotationAngle: Integer;
    FIsDisplay: Boolean;
  protected
    procedure SetFont(const Value: TZFont);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
    property Text: string read FText write FText;
    property Font: TZFont read FFont write SetFont;
    property RotationAngle: Integer read FRotationAngle write FRotationAngle default 0;
    property IsDisplay: Boolean read FIsDisplay write FIsDisplay default true;
  end;

  // Chart legend
  TZEChartLegend = class(TZEChartTitleItem)
  private
    FPosition: TZEChartLegendPosition;
    FAlign: TZELegendAlign;
  public
    constructor Create(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; override;
    property Position: TZEChartLegendPosition read FPosition write FPosition;
    property Align: TZELegendAlign read FAlign write FAlign;
  end;

  // Chart axis
  TZEChartAxis = class(TZEChartTitleItem)
  private
    FLogarithmic: Boolean;
    FReverseDirection: Boolean;
    FScaleMin: Double;
    FScaleMax: Double;
    FAutoScaleMin: Boolean;
    FAutoScaleMax: Boolean;
  public
    constructor Create(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; override;
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic default false;
    property ReverseDirection: Boolean read FReverseDirection write FReverseDirection default false;
    property ScaleMin: Double read FScaleMin write FScaleMin;
    property ScaleMax: Double read FScaleMax write FScaleMax;
    property AutoScaleMin: Boolean read FAutoScaleMin write FAutoScaleMin;
    property AutoScaleMax: Boolean read FAutoScaleMax write FAutoScaleMax;
  end;

  // Chart/series settings
  TZEChartSettings = class(TPersistent)
  private
    FJapanCandle: Boolean;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
    property JapanCandle: Boolean read FJapanCandle write FJapanCandle;
    // True - japanese candle. Used only for ZEChartTypeStock chart type.
  end;

  // Chart series
  TZEChartSeries = class(TPersistent)
  private
    FChartType: TZEChartType;
    FSeriesName: string;
    FSeriesNameSheet: Integer;
    FSeriesNameRow: Integer;
    FSeriesNameCol: Integer;
    FRanges: TZEChartRange;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
                                    // For each series it's own ChartType
    property ChartType: TZEChartType read FChartType write FChartType;

                                    // If (SeriesNameRow >= 0) and (SeriesNameCol >= 0) then
                                    // text for Series label will be from cell[SeriesNameCol, SeriesNameRow] and
                                    // from property SeriesName otherwise.
    property SeriesName: string read FSeriesName write FSeriesName;

                                    // If SeriesNameSheet < 0 then uses current sheet for Series label
    property SeriesNameSheet: Integer read FSeriesNameSheet write FSeriesNameSheet;
    property SeriesNameRow: Integer read FSeriesNameRow write FSeriesNameRow;
    property SeriesNameCol: Integer read FSeriesNameCol write FSeriesNameCol;
    property Ranges: TZEChartRange read FRanges;
  end;

  // Chart item
  TZEChart = class(TZECommonFrameAncestor)
  private
    FTitle: TZEChartTitleItem;
    FSubtitle: TZEChartTitleItem;
    FFooter: TZEChartTitleItem;
    FLegend: TZEChartLegend;
    FAxisX: TZEChartAxis;
    FAxisY: TZEChartAxis;
    FAxisZ: TZEChartAxis;
    FSecondaryAxisX: TZEChartAxis;
    FSecondaryAxisY: TZEChartAxis;
    FSecondaryAxisZ: TZEChartAxis;
    FDefaultChartType: TZEChartType;
    FView3D: Boolean;
    FViewDeep: Boolean;
  protected
    procedure SetTitle(const Value: TZEChartTitleItem);
    procedure SetSubtitle(const Value: TZEChartTitleItem);
    procedure SetFooter(const Value: TZEChartTitleItem);
    procedure SetLegend(const Value: TZEChartLegend);
    procedure CommonInit();
    procedure SetAxisX(const Value: TZEChartAxis);
    procedure SetAxisY(const Value: TZEChartAxis);
    procedure SetAxisZ(const Value: TZEChartAxis);
    procedure SetSecondaryAxisX(const Value: TZEChartAxis);
    procedure SetSecondaryAxisY(const Value: TZEChartAxis);
    procedure SetSecondaryAxisZ(const Value: TZEChartAxis);
  public
    constructor Create(); overload; override;
    constructor Create(AX, AY, AWidth, AHeight: Integer); overload; override;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; override;
    property AxisX: TZEChartAxis read FAxisX write SetAxisX;
    property AxisY: TZEChartAxis read FAxisY write SetAxisY;
    property AxisZ: TZEChartAxis read FAxisZ write SetAxisZ;
    property SecondaryAxisX: TZEChartAxis read FSecondaryAxisX write SetSecondaryAxisX;
    property SecondaryAxisY: TZEChartAxis read FSecondaryAxisY write SetSecondaryAxisY;
    property SecondaryAxisZ: TZEChartAxis read FSecondaryAxisZ write SetSecondaryAxisZ;
    property DefaultChartType: TZEChartType read FDefaultChartType write FDefaultChartType;
    property Title: TZEChartTitleItem read FTitle write SetTitle;
    property Subtitle: TZEChartTitleItem read FSubtitle write SetSubtitle;
    property Footer: TZEChartTitleItem read FFooter write SetFooter;
    property Legend: TZEChartLegend read FLegend write SetLegend;
    property View3D: Boolean read FView3D write FView3D;
    property ViewDeep: Boolean read FViewDeep write FViewDeep;
  end;

  // Store for charts on a sheet
  TZEChartStore = class(TPersistent)
  private
    FCount: Integer;
    FItems: array of TZEChart;
  protected
    function GetItem(num: Integer): TZEChart;
    procedure SetItem(num: Integer; const Value: TZEChart);
  public
    constructor Create();
    destructor Destroy(); override;
    function Add(): TZEChart; overload;
    function Add(const ItemForClone: TZEChart): TZEChart; overload;
    function Delete(num: Integer): Boolean;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; virtual;
    property Items[num: Integer]: TZEChart read GetItem write SetItem; default;
    property Count: Integer read FCount;
  end;

  TZCellAnchor = (ZACell, ZAAbsolute);

  // Picture item
  TZEPicture = class(TZECommonFrameAncestor)
  private
    FId: Integer;
    FRelId: Integer;

    FFileName: string;
    FTitle: string;
    FDescription: string;
    FCellAnchor: TZCellAnchor;

    FRow: Integer;
    FCol: Integer;
    // FHidden: Boolean;

    FFromCol: Integer;
    FFromColOff: Integer;
    FFromRow: Integer;
    FFromRowOff: Integer;
    FToCol: Integer;
    FToColOff: Integer;
    FToRow: Integer;
    FToRowOff: Integer;
    FFrmOffX: Integer;
    FFrmOffY: Integer;
    FFrmExtCX: Integer;
    FFrmExtCY: Integer;

    FSheet: TZSheet;
    function GetImage: TBytes;
    procedure SetImage(const Value: TBytes);
  protected
    procedure CommonInit();
  public
    constructor Create(ASheet: TZSheet);
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): Boolean; override;
    // through workbook
    property Id: Integer read FId write FId;
    // through worksheet
    property RelId: Integer read FRelId write FRelId;
    property Name: string read FFileName write FFileName;
    property Title: string read FTitle write FTitle;
    property Description: string read FDescription write FDescription;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property CellAnchor: TZCellAnchor read FCellAnchor write FCellAnchor;

    property FromCol: Integer read FFromCol write FFromCol;
    property FromColOff: Integer read FFromColOff write FFromColOff;
    property FromRow: Integer read FFromRow write FFromRow;
    property FromRowOff: Integer read FFromRowOff write FFromRowOff;
    property ToCol: Integer read FToCol write FToCol;
    property ToColOff: Integer read FToColOff write FToColOff;
    property ToRow: Integer read FToRow write FToRow;
    property ToRowOff: Integer read FToRowOff write FToRowOff;
    property FrmOffX: Integer read FFrmOffX write FFrmOffX;
    property FrmOffY: Integer read FFrmOffY write FFrmOffY;
    property FrmExtCX: Integer read FFrmExtCX write FFrmExtCX;
    property FrmExtCY: Integer read FFrmExtCY write FFrmExtCY;

    property Image: TBytes read GetImage write SetImage;
  end;

  { Store pictures for worksheet }
  TZEDrawing = class(TPersistent)
  private
    FId: Integer;
    FItems: TObjectList;
    FSheet: TZSheet;
    function GetIsEmpty(): Boolean;
    function GetCount(): Integer;
    function GetItem(idx: Integer): TZEPicture;
    procedure SetItem(idx: Integer; const Value: TZEPicture);
  public
    constructor Create(ASheet: TZSheet);
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;

    function Add(): TZEPicture; overload;
    function Add(ARow, ACol: Integer; APicture: TBytes): TZEPicture; overload;

    procedure Delete(idx: Integer);
    procedure Clear();

    property Count: Integer read GetCount;
    property Id: Integer read FId write FId;
    property IsEmpty: Boolean read GetIsEmpty;
    property Items[idx: Integer]: TZEPicture read GetItem write SetItem; default;
  end;

  TZRange = class;
  IZRange = interface;

  /// <summary>
  /// Inherited from TPersistent. Contains a sheet of the document.
  /// </summary>
  TZSheet = class(TPersistent)
  private
    FStore: TZWorkBook;
    FCells: array of TZCellColumn;
    FRows: array of TZRowOptions;
    FColumns: array of TZColOptions;
    FAutoFilter: string;
    FTitle: string;
    FRowCount: Integer;
    FColCount: Integer;
    FTabColor: TColor; // цвет закладки
    FFitToPage: Boolean;
    FDefaultRowHeight: Real;
    FDefaultColWidth: Real;
    FMergeCells: TZMergeCells;
    FProtect: Boolean;
    FRightToLeft: Boolean;
    FSheetOptions: TZSheetOptions;
    FSelected: Boolean;
    FPrintRows, FPrintCols: TZSheetPrintTitles;
    FCharts: TZEChartStore;
    FDrawing: TZEDrawing;
    FViewMode: TZViewMode;
    FSummaryBelow: Boolean;
    FSummaryRight: Boolean;
    FApplyStyles: Boolean;
    FDrawingRid: Integer;
    FOutlineLevelRow: Integer;
    FOutlineLevelCol: Integer;
    FRowBreaks: TArray<Integer>;
    FColBreaks: TArray<Integer>;
    FConditionalFormatting: TZConditionalFormatting;
    procedure SetConditionalFormatting(Value: TZConditionalFormatting);
    procedure SetCharts(const Value: TZEChartStore);
    procedure SetColumn(num: Integer; const Value: TZColOptions);
    function GetColumn(num: Integer): TZColOptions;
    procedure SetRow(num: Integer; const Value: TZRowOptions);
    function GetRow(num: Integer): TZRowOptions;
    function GetSheetOptions(): TZSheetOptions;
    procedure SetSheetOptions(Value: TZSheetOptions);
    procedure SetPrintCols(const Value: TZSheetPrintTitles);
    procedure SetPrintRows(const Value: TZSheetPrintTitles);
  protected
    procedure SetColWidth(num: Integer; const Value: Real); virtual;
    function GetColWidth(num: Integer): Real; virtual;
    procedure SetRowHeight(num: Integer; const Value: Real); virtual;
    function GetRowHeight(num: Integer): Real; virtual;
    procedure SetDefaultColWidth(const Value: Real); virtual;
    procedure SetDefaultRowHeight(const Value: Real); virtual;
    function GetCell(ACol, ARow: Integer): TZCell; virtual;
    procedure SetCell(ACol, ARow: Integer; const Value: TZCell); virtual;
    function GetCellRef(ACol: string; ARow: Integer): TZCell; virtual;
    procedure SetCellRef(ACol: string; ARow: Integer; const Value: TZCell); virtual;
    function GetRowCount: Integer; virtual;
    procedure SetRowCount(const Value: Integer); virtual;
    function GetColCount: Integer; virtual;
    procedure SetColCount(const Value: Integer); virtual;
    function GetRange(AC1, AR1, AC2, AR2: Integer): IZRange; virtual;
    // procedure SetRange(AC1,AR1,AC2,AR2: integer; const Value: TZRange); virtual;
    function GetRangeRef(AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: Integer): IZRange; virtual;
    // procedure SetRangeRef(AFrom, ATo: string; const Value: TZRange); virtual;
    function GetSheetIndex(): Integer;
  public
    constructor Create(AStore: TZWorkBook); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    procedure Clear(); virtual;
    procedure InsertRows(ARow, ACount: Integer);
    procedure CopyRows(ARowDst, ARowSrc, ACount: Integer);
    procedure SetCorrectTitle(const Value: string);

    function ColsWidth(AFrom, ATo: Integer): Real;
    function RowsHeight(AFrom, ATo: Integer): Real;

    /// <summary>
    /// Get or set the width (in points) of column num in the sheet.
    /// </summary>
    property ColWidths[num: Integer]: Real read GetColWidth write SetColWidth;
    /// <summary>
    /// Options of column num.
    /// </summary>
    property Columns[num: Integer]: TZColOptions read GetColumn write SetColumn;
    /// <summary>
    /// Specifies various properties of the Row num.
    /// </summary>
    property Rows[num: Integer]: TZRowOptions read GetRow write SetRow;
    /// <summary>
    /// Specifies various properties of the Cells range.
    /// </summary>
    property Range[AC1, AR1, AC2, AR2: Integer]: IZRange read GetRange { write SetRange };
    /// <summary>
    /// Specifies various properties of the Cells range.
    /// </summary>
    property RangeRef[AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: Integer]: IZRange
      read GetRangeRef { write SetRangeRef };
    /// <summary>
    /// Get or set the height (in points) of row num in the sheet.
    /// </summary>
    property RowHeights[num: Integer]: Real read GetRowHeight write SetRowHeight;
    /// <summary>
    /// Default column width.
    /// </summary>
    property DefaultColWidth: Real read FDefaultColWidth write SetDefaultColWidth; // default 48;
    /// <summary>
    /// Default row height.
    /// </summary>
    property DefaultRowHeight: Real read FDefaultRowHeight write SetDefaultRowHeight; // default 12.75;
    /// <summary>
    /// Cell at the intersection of column ACol and row ARow.
    /// </summary>
    property Cell[ACol, ARow: Integer]: TZCell read GetCell write SetCell; default;
    /// <summary>
    /// Cell at the intersection of column by "A1" reference.
    /// </summary>
    property CellRef[ACol: string; ARow: Integer]: TZCell read GetCellRef write SetCellRef;
    property AutoFilter: string read FAutoFilter write FAutoFilter;
    /// <summary>
    /// Indicates whether or not this sheet is protected. <br />False by default.
    /// </summary>
    property Protect: Boolean read FProtect write FProtect default false;
    property TabColor: TColor read FTabColor write FTabColor default TColorRec.cWINDOW;
    property FitToPage: Boolean read FFitToPage write FFitToPage default false;

    property SummaryBelow: Boolean read FSummaryBelow write FSummaryBelow;
    property SummaryRight: Boolean read FSummaryRight write FSummaryRight;
    property ApplyStyles: Boolean read FApplyStyles write FApplyStyles;
    property DrawingRid: Integer read FDrawingRid write FDrawingRid;

    property OutlineLevelRow: Integer read FOutlineLevelRow write FOutlineLevelRow;
    property OutlineLevelCol: Integer read FOutlineLevelCol write FOutlineLevelCol;

    /// <summary>
    /// Sheet title.
    /// </summary>
    property Title: string read FTitle write SetCorrectTitle;
    /// <summary>
    /// Sheet index in the workbook.
    /// </summary>
    property SheetIndex: Integer read GetSheetIndex;
    /// <summary>
    /// Specifies the number of rows in the sheet.
    /// </summary>
    property RowCount: Integer read GetRowCount write SetRowCount;
    /// <summary>
    /// If True, the window displays from right to left else window displays from left to right. <br />False by default.
    /// </summary>
    property RightToLeft: Boolean read FRightToLeft write FRightToLeft default false;
    /// <summary>
    /// Specifies the number of columns in the sheet.
    /// </summary>
    property ColCount: Integer read GetColCount write SetColCount;
    /// <summary>
    /// Merged cells.
    /// </summary>
    property MergeCells: TZMergeCells read FMergeCells write FMergeCells;
    /// <summary>
    /// Specifies various properties of the sheet.
    /// </summary>
    property SheetOptions: TZSheetOptions read GetSheetOptions write SetSheetOptions;
    /// <summary>
    /// Indicates whether or not this sheet is selecteded.
    /// </summary>
    property Selected: Boolean read FSelected write FSelected;
    property WorkBook: TZWorkBook read FStore;
    property RowsToRepeat: TZSheetPrintTitles read FPrintRows write SetPrintRows;
    property ColsToRepeat: TZSheetPrintTitles read FPrintCols write SetPrintCols;
    /// <summary>
    /// Conditional formatting for sheet.
    /// </summary>
    property ConditionalFormatting: TZConditionalFormatting read FConditionalFormatting write SetConditionalFormatting;
    property Charts: TZEChartStore read FCharts write SetCharts;
    property Drawing: TZEDrawing read FDrawing;
    property ViewMode: TZViewMode read FViewMode write FViewMode;
    property RowBreaks: TArray<Integer> read FRowBreaks write FRowBreaks;
    property ColBreaks: TArray<Integer> read FColBreaks write FColBreaks;
  end;

  /// <summary>
  /// Document sheets.
  /// </summary>
  TZSheets = class(TPersistent)
  private
    FStore: TZWorkBook;
    FSheets: array of TZSheet;
    FCount: Integer;
    procedure SetSheetCount(const Value: Integer);
    procedure SetSheet(num: Integer; Const Value: TZSheet);
    function GetSheet(num: Integer): TZSheet;
  public
    constructor Create(AStore: TZWorkBook); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Number of sheets in the document.
    /// </summary>
    property Count: Integer read FCount write SetSheetCount;
    /// <summary>
    /// Document's sheet num.
    /// </summary>
    property Sheet[num: Integer]: TZSheet read GetSheet write SetSheet; default;
    /// <summary>
    /// Add new sheet to the workbook.
    /// </summary>
    function Add(const ATitle: string = ''): TZSheet;
  end;

  IZRange = interface(IInterface)
    function HasStyle: Boolean;
    procedure ApplyStyleValue(proc: TProc<TZStyle>);
    function GetVerticalAlignment(): TZVerticalAlignment;
    procedure SetVerticalAlignment(const Value: TZVerticalAlignment);
    function GetHorizontalAlignment(): TZHorizontalAlignment;
    procedure SetHorizontalAlignment(const Value: TZHorizontalAlignment);
    function GetRotate(): TZCellTextRotate;
    procedure SetRotate(const Value: TZCellTextRotate);
    function GetBgColor(): TColor;
    procedure SetBgColor(const Value: TColor);
    function GetFontColor(): TColor;
    procedure SetFontColor(const Value: TColor);
    function GetFontSize(): Double;
    procedure SetFontSize(const Value: Double);
    function GetFontStyle(): TFontStyles;
    procedure SetFontStyle(const Value: TFontStyles);
    function GetBorderStyle(num: TZBordersPos): TZBorderType;
    procedure SetBorderStyle(num: TZBordersPos; const Value: TZBorderType);
    function GetBorderWidht(num: TZBordersPos): Byte;
    procedure SetBorderWidht(num: TZBordersPos; const Value: Byte);
    function GetBorderColor(num: TZBordersPos): TColor;
    procedure SetBorderColor(num: TZBordersPos; const Value: TColor);
    function GetBordersStyle(): TZBorderType;
    procedure SetBordersStyle(const Value: TZBorderType);
    function GetBordersWidht(): Byte;
    procedure SetBordersWidht(const Value: Byte);
    function GetBordersColor(): TColor;
    procedure SetBordersColor(const Value: TColor);
    function GetWrapText(): Boolean;
    procedure SetWrapText(const Value: Boolean);
    function GetVerticalText(): Boolean;
    procedure SetVerticalText(const Value: Boolean);
    function GetNumberFormat(): string;
    procedure SetNumberFormat(const Value: string);
    //
    property VerticalAlignment: TZVerticalAlignment read GetVerticalAlignment write SetVerticalAlignment;
    property HorizontalAlignment: TZHorizontalAlignment read GetHorizontalAlignment write SetHorizontalAlignment;
    property BgColor: TColor read GetBgColor write SetBgColor;
    property FontColor: TColor read GetFontColor write SetFontColor;
    property FontSize: Double read GetFontSize write SetFontSize;
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
    property BorderStyle[num: TZBordersPos]: TZBorderType read GetBorderStyle write SetBorderStyle;
    property BorderWidht[num: TZBordersPos]: Byte read GetBorderWidht write SetBorderWidht;
    property BorderColor[num: TZBordersPos]: TColor read GetBorderColor write SetBorderColor;
    property BordersStyle: TZBorderType read GetBordersStyle write SetBordersStyle;
    property BordersWidht: Byte read GetBordersWidht write SetBordersWidht;
    property BordersColor: TColor read GetBordersColor write SetBordersColor;
    property WrapText: Boolean read GetWrapText write SetWrapText;
    property VerticalText: Boolean read GetVerticalText write SetVerticalText;
    property Rotate: TZCellTextRotate read GetRotate write SetRotate;
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    procedure SetBorderAround(borderWidth: Byte; BorderColor: TColor = TColorRec.Black; BorderStyle: TZBorderType = TZBorderType.ZEContinuous);
    procedure Merge();
    procedure Clear();
  end;

  TZRange = class(TInterfacedObject, IZRange)
  private
    FSheet: TZSheet;
    FLeft, FTop, FRight, FBottom: Integer;

    function HasStyle: Boolean;
    procedure ApplyStyleValue(proc: TProc<TZStyle>);
    function GetVerticalAlignment(): TZVerticalAlignment;
    procedure SetVerticalAlignment(const Value: TZVerticalAlignment);
    function GetHorizontalAlignment(): TZHorizontalAlignment;
    procedure SetHorizontalAlignment(const Value: TZHorizontalAlignment);
    function GetRotate(): TZCellTextRotate;
    procedure SetRotate(const Value: TZCellTextRotate);
    function GetBgColor(): TColor;
    procedure SetBgColor(const Value: TColor);
    function GetFontColor(): TColor;
    procedure SetFontColor(const Value: TColor);
    function GetFontSize(): Double;
    procedure SetFontSize(const Value: Double);
    function GetFontStyle(): TFontStyles;
    procedure SetFontStyle(const Value: TFontStyles);
    function GetBorderStyle(num: TZBordersPos): TZBorderType;
    procedure SetBorderStyle(num: TZBordersPos; const Value: TZBorderType);
    function GetBorderWidht(num: TZBordersPos): Byte;
    procedure SetBorderWidht(num: TZBordersPos; const Value: Byte);
    function GetBorderColor(num: TZBordersPos): TColor;
    procedure SetBorderColor(num: TZBordersPos; const Value: TColor);
    function GetBordersStyle(): TZBorderType;
    procedure SetBordersStyle(const Value: TZBorderType);
    function GetBordersWidht(): Byte;
    procedure SetBordersWidht(const Value: Byte);
    function GetBordersColor(): TColor;
    procedure SetBordersColor(const Value: TColor);
    function GetWrapText(): Boolean;
    procedure SetWrapText(const Value: Boolean);
    function GetVerticalText(): Boolean;
    procedure SetVerticalText(const Value: Boolean);
    function GetNumberFormat(): string;
    procedure SetNumberFormat(const Value: string);
  protected
  public
    constructor Create(ASheet: TZSheet; ALeft, ATop, ARight, ABottom: Integer); virtual;
    procedure Assign(Source: TZRange);
    destructor Destroy(); override;
    property VerticalAlignment: TZVerticalAlignment read GetVerticalAlignment write SetVerticalAlignment;
    property HorizontalAlignment: TZHorizontalAlignment read GetHorizontalAlignment write SetHorizontalAlignment;
    property BgColor: TColor read GetBgColor write SetBgColor;
    property FontColor: TColor read GetFontColor write SetFontColor;
    property FontSize: Double read GetFontSize write SetFontSize;
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
    property BorderStyle[num: TZBordersPos]: TZBorderType read GetBorderStyle write SetBorderStyle;
    property BorderWidht[num: TZBordersPos]: Byte read GetBorderWidht write SetBorderWidht;
    property BorderColor[num: TZBordersPos]: TColor read GetBorderColor write SetBorderColor;
    property BordersStyle: TZBorderType read GetBordersStyle write SetBordersStyle;
    property BordersWidht: Byte read GetBordersWidht write SetBordersWidht;
    property BordersColor: TColor read GetBordersColor write SetBordersColor;
    property WrapText: Boolean read GetWrapText write SetWrapText;
    property VerticalText: Boolean read GetVerticalText write SetVerticalText;
    property Rotate: TZCellTextRotate read GetRotate write SetRotate;
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    procedure SetBorderAround(borderWidth: Byte; BorderColor: TColor = TColorRec.Black; BorderStyle: TZBorderType = TZBorderType.ZEContinuous);
    procedure Merge();
    procedure Clear();
  end;

  /// <summary>
  /// Document properties
  /// </summary>
  TZEXMLDocumentProperties = class(TPersistent)
  private
    FAuthor: string;
    FLastAuthor: string;
    FCreated: TDateTime;
    FCompany: string;
    FVersion: string; // - should be integer by Spec but hardcoded float in real MS Office apps
    FWindowHeight: word;
    FWindowWidth: word;
    FWindowTopX: Integer;
    FWindowTopY: Integer;
    FModeR1C1: Boolean;
  protected
    procedure SetAuthor(const Value: string);
    procedure SetLastAuthor(const Value: string);
    procedure SetCompany(const Value: string);
    procedure SetVersion(const Value: string);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Author of the document
    /// </summary>
    property Author: string read FAuthor write SetAuthor;
    /// <summary>
    /// Author of last changes in the document
    /// </summary>
    property LastAuthor: string read FLastAuthor write SetLastAuthor;
    /// <summary>
    /// Date and time of document creation
    /// </summary>
    property Created: TDateTime read FCreated write FCreated;
    /// <summary>
    /// Company name
    /// </summary>
    property Company: string read FCompany write SetCompany;
    /// <summary>
    /// Document version
    /// </summary>
    property Version: string read FVersion write SetVersion;
    /// <summary>
    /// Enabled R1C1 style in Excel. <br />False by default
    /// </summary>
    property ModeR1C1: Boolean read FModeR1C1 write FModeR1C1 default false;
    property WindowHeight: word read FWindowHeight write FWindowHeight default 20000;
    property WindowWidth: word read FWindowWidth write FWindowWidth default 20000;
    property WindowTopX: Integer read FWindowTopX write FWindowTopX default 150;
    property WindowTopY: Integer read FWindowTopY write FWindowTopY default 150;
  end;

  TDefinedName = record
    LocalSheetId: Integer;
    Name: string;
    Body: string;
  end;

  /// <summary>
  /// Contains spreadsheet document
  /// </summary>
  TZWorkBook = class(TComponent)
  private
    FSheets: TZSheets;
    FDocumentProperties: TZEXMLDocumentProperties;
    FStyles: TZStyles;
    FHorPixelSize: Real;
    FVertPixelSize: Real;
    FDefaultSheetOptions: TZSheetOptions;
    FMediaList: TArray<TMediaRec>;
    procedure SetHorPixelSize(Value: Real);
    procedure SetVertPixelSize(Value: Real);
    function GetDefaultSheetOptions(): TZSheetOptions;
    procedure SetDefaultSheetOptions(Value: TZSheetOptions);
  public
    FDefinedNames: TArray<TDefinedName>;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    {$IFNDEF FMX}
    procedure GetPixelSize(hdc: THandle);
    {$ENDIF}
    property Sheets: TZSheets read FSheets write FSheets;
    property MediaList: TArray<TMediaRec> read FMediaList write FMediaList;
    function AddMediaContent(AFileName: string; AContent: TBytes; ACheckByName: Boolean): Integer;
    function GetDrawing(num: Integer): TZEDrawing;
    function GetDrawingSheetNum(Value: TZEDrawing): Integer;
    property Styles: TZStyles read FStyles write FStyles;
    property DefaultSheetOptions: TZSheetOptions read GetDefaultSheetOptions write SetDefaultSheetOptions;
    property DocumentProperties: TZEXMLDocumentProperties read FDocumentProperties write FDocumentProperties;
    property HorPixelSize: Real read FHorPixelSize write SetHorPixelSize; // размер пикселя по горизонтали
    property VertPixelSize: Real read FVertPixelSize write SetVertPixelSize; // размер пикселя по вертикали
  end;

/// <summary>
/// Convert TColor to Hex RGB
/// </summary>
function ColorToHTMLHex(Color: TColor): string;

/// <summary>
/// Convert Hex RGB (string) to TColor
/// </summary>
function HTMLHexToColor(Value: string): TColor;

/// <summary>
/// Convert ARGB (string) to TColor
/// </summary>
function ARGBToColor(Value: string): TColor;

/// <summary>
/// Convert pixels to point
/// </summary>
function PixelToPoint(inPixel: Integer; PixelSizeMM: Real = 0.265): Real;

/// <summary>
/// Convert typographical point to pixels.
/// </summary>
function PointToPixel(inPoint: Real; PixelSizeMM: Real = 0.265): Integer;

/// <summary>
/// Convert typographical point to mm.
/// </summary>
function PointToMM(inPoint: Real): Real;

/// <summary>
/// Convert mm to typographical point.
/// </summary>
function MMToPoint(inMM: Real): Real;

/// <summary>
/// Checks is Font1 equal Font2
/// </summary>
function ZEIsFontsEquals(const Font1, Font2: TZFont): Boolean; overload;

/// <summary>
/// Checks is Font1 equal Font2
/// </summary>
function ZEIsFontsEquals(const Font1, Font2: TFont): Boolean; overload;

/// <summary>
/// Convert datetime value to string (YYYY-MM-DDTHH:MM:SS[.mmm]).
/// </summary>
function ZEDateTimeToStr(ATime: TDateTime; Addmms: Boolean = false): string;

/// <summary>
/// Try convert string (YYYY-MM-DDTHH:MM:SS[.mmm]) to datetime
/// </summary>
function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): Boolean;

/// <summary>
/// Convert the number to string min count NullCount
/// </summary>
function IntToStrN(Value: Integer; NullCount: Integer): string;

function IsIdenticalByteArray(Src, Dst: TBytes): Boolean;

implementation

uses
  Excel4Delphi.Formula;

var
  invariantFormatSertting: TFormatSettings;

function IsIdenticalByteArray(Src, Dst: TBytes): Boolean;
var
  i: Integer;
begin
  if Length(Src) <> Length(Dst) then
    exit(false);

  for i := Low(Src) to High(Src) do
  begin
    if Src[i] <> Dst[i] then
      exit(false);
  end;

  Result := true;
end;

function IntToStrN(Value: Integer; NullCount: Integer): string;
var
  t, k: Integer;
begin
  t := Value;
  k := 0;
  if (t = 0) then
    k := 1;
  while t > 0 do
  begin
    Inc(k);
    t := t div 10;
  end;
  Result := IntToStr(Value);
  for t := 1 to (NullCount - k) do
    Result := '0' + Result;
end;

function ZEDateTimeToStr(ATime: TDateTime; Addmms: Boolean = false): string;
var
  HH, MM, SS, MS: word;
begin
  DecodeDate(ATime, HH, MM, SS);
  Result := IntToStrN(HH, 4) + '-' + IntToStrN(MM, 2) + '-' + IntToStrN(SS, 2) + 'T';
  DecodeTime(ATime, HH, MM, SS, MS);
  Result := Result + IntToStrN(HH, 2) + ':' + IntToStrN(MM, 2) + ':' + IntToStrN(SS, 2);
  if (Addmms) then
    Result := Result + '.' + IntToStrN(MS, 3);
end;

function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): Boolean;
var
  a: array [0 .. 10] of word;
  i, l: Integer;
  s, SS: string;
  Count: Integer;
  ch: char;
  datedelimeters: Integer;
  istimesign: Boolean;
  timedelimeters: Integer;
  istimezone: Boolean;
  lastdateindex: Integer;
  tmp: Integer;
  msindex: Integer;
  tzindex: Integer;
  timezonemul: Integer;
  _ms: word;

  function TryAddToArray(const ST: string): Boolean;
  begin
    if (Count > 10) then
    begin
      Result := false;
      exit;
    end;
    Result := TryStrToInt(ST, tmp);
    if (Result) then
    begin
      a[Count] := word(tmp);
      Inc(Count);
    end
  end;

  procedure _CheckDigits();
  var
    _l: Integer;
  begin
    _l := Length(s);
    if (_l > 0) then
    begin
      if (_l > 4) then
      begin // it is not good
        if (istimesign) then
        begin
          // HHMMSS?
          if (_l = 6) then
          begin
            SS := copy(s, 1, 2);
            if (TryAddToArray(SS)) then
            begin
              SS := copy(s, 3, 2);
              if (TryAddToArray(SS)) then
              begin
                SS := copy(s, 5, 2);
                if (not TryAddToArray(SS)) then
                  Result := false;
              end
              else
                Result := false;
            end
            else
              Result := false
          end
          else
            Result := false;
        end
        else
        begin
          // YYYYMMDD?
          if (_l = 8) then
          begin
            SS := copy(s, 1, 4);
            if (not TryAddToArray(SS)) then
              Result := false
            else
            begin
              SS := copy(s, 5, 2);
              if (not TryAddToArray(SS)) then
                Result := false
              else
              begin
                SS := copy(s, 7, 2);
                if (not TryAddToArray(SS)) then
                  Result := false;
              end;
            end;
          end
          else
            Result := false;
        end;
      end
      else if (not TryAddToArray(s)) then
        Result := false;
    end; // if
    if (Count > 10) then
      Result := false;
    s := '';
  end;

  procedure _processDigit();
  begin
    s := s + ch;
  end;

  procedure _processTimeSign();
  begin
    istimesign := true;
    if (Count > 0) then
      lastdateindex := Count;

    _CheckDigits();
  end;

  procedure _processTimeDelimiter();
  begin
    _CheckDigits();
    Inc(timedelimeters)
  end;

  procedure _processDateDelimiter();
  begin
    _CheckDigits();
    if (istimesign) then
    begin
      tzindex := Count;
      istimezone := true;
      timezonemul := -1;
    end
    else
      Inc(datedelimeters);
  end;

  procedure _processMSDelimiter();
  begin
    _CheckDigits();
    msindex := Count;
  end;

  procedure _processTimeZoneSign();
  begin
    _CheckDigits();
    istimezone := true;
  end;

  procedure _processTimeZonePlus();
  begin
    _CheckDigits();
    istimezone := true;
    timezonemul := -1;
  end;

  function _TryGetDateTime(): Boolean;
  var
    _time, _date: TDateTime;
  begin
    // Result := true;
    if (msindex >= 0) then
      _ms := a[msindex];
    if (lastdateindex >= 0) then
    begin
      Result := TryEncodeDate(a[0], a[1], a[2], _date);
      if (Result) then
      begin
        Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, _time);
        if (Result) then
          retDateTime := _date + _time;
      end;
    end
    else
      Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, retDateTime);
  end;

  function _TryGetDate(): Boolean;
  begin
    if (datedelimeters = 0) and (timedelimeters >= 2) then
    begin
      if (msindex >= 0) then
        _ms := a[msindex];
      Result := TryEncodeTime(a[0], a[1], a[2], _ms, retDateTime);
    end
    else if (Count >= 3) then
      Result := TryEncodeDate(a[0], a[1], a[2], retDateTime)
    else
      Result := false;
  end;

begin
  Result := true;
  datedelimeters := 0;
  istimesign := false;
  timedelimeters := 0;
  istimezone := false;
  lastdateindex := -1;
  msindex := -1;
  tzindex := -1;
  timezonemul := 0;
  _ms := 0;
  FillChar(a, sizeof(a), 0);

  l := Length(AStrDateTime);
  s := '';
  Count := 0;
  for i := 1 to l do
  begin
    ch := AStrDateTime[i];
    case (ch) of
      '0' .. '9':
        _processDigit();
      't', 'T':
        _processTimeSign();
      '-':
        _processDateDelimiter();
      ':':
        _processTimeDelimiter();
      '.', ',':
        _processMSDelimiter();
      'z', 'Z':
        _processTimeZoneSign();
      '+':
        _processTimeZonePlus();
    end;
    if (not Result) then
      break
  end;

  if (Result and (s <> '')) then
    _CheckDigits();

  if (Result) then
  begin
    if (istimesign) then
      Result := _TryGetDateTime()
    else
      Result := _TryGetDate();
  end;
end; // TryZEStrToDateTime

function ZEIsFontsEquals(const Font1, Font2: TZFont): Boolean;
begin
  Result := Assigned(Font1) and (Assigned(Font2));
  if (Result) then
  begin
    Result := false;
    if (Font1.Color <> Font2.Color) then
      exit;

    if (Font1.Name <> Font2.Name) then
      exit;

    if (Font1.Size <> Font2.Size) then
      exit;

    if (Font1.Style <> Font2.Style) then
      exit;

    Result := true;
  end;
end;

function ZEIsFontsEquals(const Font1, Font2: TFont): Boolean;
begin
  Result := Assigned(Font1) and (Assigned(Font2));
  if Result then
  begin
    Result := false;
    {$IFDEF FMX}
    if (Font1.Family <> Font2.Family) then
      exit;

    {$ELSE}
    if (Font1.Color <> Font2.Color) then
      exit;

    if (Font1.Name <> Font2.Name) then
      exit;
    {$ENDIF}

    if (Font1.Size <> Font2.Size) then
      exit;

    if (Font1.Style <> Font2.Style) then
      exit;

    Result := true;
  end;
end;

function ColorToHTMLHex(Color: TColor): string;
var
  _RGB: Integer;
begin
  _RGB := TColorRec.ColorToRGB(Color);
  // result := IntToHex(GetRValue(_RGB), 2) + IntToHex(GetGValue(_RGB), 2) + IntToHex(GetBValue(_RGB), 2);
  Result := IntToHex(Byte(_RGB), 2) + IntToHex(Byte(_RGB shr 8), 2) + IntToHex(Byte(_RGB shr 16), 2);
end;

function HTMLHexToColor(Value: string): TColor;
var
  a: array [0 .. 2] of Integer;
  i, n, t: Integer;
begin
  Result := 0;
  if Value > '' then
  begin
    Value := UpperCase(Value);
{$HINTS OFF}
    FillChar(a, sizeof(a), 0);
{$HINTS ON}
    n := 0;
    if Value[1] = '#' then
      Delete(Value, 1, 1);
    // А что, если будут цвета типа "black"?  {tut}
    for i := 1 to Length(Value) do
    begin
      if n > 2 then
        break;
      case Value[i] of
        '0' .. '9':
          t := ord(Value[i]) - 48;
        'A' .. 'F':
          t := 10 + ord(Value[i]) - 65;
      else
        t := 0;
      end;
      a[n] := a[n] * 16 + t;
      if i mod 2 = 0 then
        Inc(n);
    end;
    Result := a[2] shl 16 or a[1] shl 8 or a[0];
  end;
end;

function ARGBToColor(Value: string): TColor;
var
  a: array [0 .. 2] of Integer;
  i, n, t: Integer;
begin
  Result := 0;
  if Value > '' then
  begin
    Value := UpperCase(Value);
{$HINTS OFF}
    FillChar(a, sizeof(a), 0);
{$HINTS ON}
    n := 0;
    if Value[1] = '#' then
      Delete(Value, 1, 1);
    for i := 3 to Length(Value) do
    begin
      if n > 2 then
        break;
      case Value[i] of
        '0' .. '9':
          t := ord(Value[i]) - 48;
        'A' .. 'F':
          t := 10 + ord(Value[i]) - 65;
      else
        t := 0;
      end;
      a[n] := a[n] * 16 + t;
      if i mod 2 = 0 then
        Inc(n);
    end;
    Result := a[2] shl 16 or a[1] shl 8 or a[0];
  end;
end;

function PixelToPoint(inPixel: Integer; PixelSizeMM: Real = 0.265): Real;
begin
  Result := inPixel * PixelSizeMM / _PointToMM;
  // и оставим 2 знака после запятой ^_^
  Result := round(Result * 100) / 100;
end;

function PointToPixel(inPoint: Real; PixelSizeMM: Real = 0.265): Integer;
begin
  Result := round(inPoint * _PointToMM / PixelSizeMM);
end;

function PointToMM(inPoint: Real): Real;
begin
  Result := round(inPoint * _PointToMM * 100) / 100;
end;

function MMToPoint(inMM: Real): Real;
begin
  Result := round(inMM / _PointToMM * 100) / 100;
end;

/// /::::::::::::: TZBorderStyle :::::::::::::::::////

constructor TZBorderStyle.Create();
begin
  FWeight := 0;
  FColor.RGB := TColorRec.Black;
  FColor.Indexed := 0;
  FColor.Theme := 0;
  FColor.Tint := 0;
  FLineStyle := ZENone;
end;

procedure TZBorderStyle.Assign(Source: TPersistent);
begin
  if Source is TZBorderStyle then
  begin
    Weight := (Source as TZBorderStyle).Weight;
    Color := (Source as TZBorderStyle).Color;
    LineStyle := (Source as TZBorderStyle).LineStyle;
  end
  else
    inherited Assign(Source);
end;

function TZBorderStyle.IsEqual(Source: TPersistent): Boolean;
var
  zSource: TZBorderStyle;
begin
  Result := false;
  if not(Source is TZBorderStyle) then
    exit;
  zSource := Source as TZBorderStyle;

  if Self.LineStyle <> zSource.LineStyle then
    exit;

  if (Self.FColor <> zSource.FColor) then
    exit;

  if Self.Weight <> zSource.Weight then
    exit;

  Result := true;
end;

procedure TZBorderStyle.SetLineStyle(const Value: TZBorderType);
begin
  FLineStyle := Value;
end;

procedure TZBorderStyle.SetWeight(const Value: Byte);
begin
  FWeight := min(3, Value);
end;

function TZBorderStyle.GetColor: TColor;
begin
  Result := FColor.RGB;
end;

procedure TZBorderStyle.SetColor(const Value: TColor);
begin
  if FColor.RGB <> Value then
  begin
    FColor.RGB := Value;
    FColor.Indexed := 0;
    FColor.Theme := 0;
    FColor.Tint := 0;
  end;
end;

/// /::::::::::::: TZBorder :::::::::::::::::////

constructor TZBorder.Create();
var
  i: Integer;
begin
  for i := 0 to 5 do
    FBorder[i] := TZBorderStyle.Create();
end;

destructor TZBorder.Destroy();
var
  i: Integer;
begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);
  inherited Destroy;
end;

procedure TZBorder.Assign(Source: TPersistent);
var
  zSource: TZBorder;
  i: TZBordersPos;
begin
  if (Source is TZBorder) then
  begin
    zSource := Source as TZBorder;
    for i := bpLeft to bpDiagonalRight do
      Border[i].Assign(zSource.Border[i]);
  end
  else
    inherited Assign(Source);
end;

function TZBorder.IsEqual(Source: TPersistent): Boolean;
var
  zSource: TZBorder;
  i: TZBordersPos;
begin
  Result := false;
  if not(Source is TZBorder) then
    exit;
  zSource := Source as TZBorder;
  for i := bpLeft to bpDiagonalRight do
    if not FBorder[ord(i)].IsEqual(zSource.Border[i]) then
      exit;
  Result := true;
end;

procedure TZBorder.SetBorder(num: TZBordersPos; Const Value: TZBorderStyle);
begin
  if (num >= bpLeft) and (num <= bpDiagonalRight) then
    Border[num].Assign(Value);
end;

function TZBorder.GetBorder(num: TZBordersPos): TZBorderStyle;
begin
  if (num >= bpLeft) and (num <= bpDiagonalRight) then
    Result := FBorder[ord(num)]
  else
    Result := nil;
end;

/// /::::::::::::: TZAlignment :::::::::::::::::////

constructor TZAlignment.Create();
begin
  FIndent := 0;
  FRotate := 0;
  FHorizontal := ZHAutomatic;
  FShrinkToFit := false;
  FVertical := ZVAutomatic;
  FVerticalText := false;
  FWrapText := false;
end;

procedure TZAlignment.Assign(Source: TPersistent);
var
  zSource: TZAlignment;
begin
  if (Source is TZAlignment) then
  begin
    zSource := Source as TZAlignment;
    FHorizontal := zSource.Horizontal;
    FIndent := zSource.Indent;
    FRotate := zSource.Rotate;
    FShrinkToFit := zSource.ShrinkToFit;
    FVertical := zSource.Vertical;
    FVerticalText := zSource.VerticalText;
    FWrapText := zSource.WrapText;
  end
  else
    inherited Assign(Source);
end;

function TZAlignment.IsEqual(Source: TPersistent): Boolean;
var
  zSource: TZAlignment;
begin
  Result := true;
  if (Source is TZAlignment) then
  begin
    zSource := Source as TZAlignment;
    if Horizontal <> zSource.Horizontal then
      exit(false);
    if Indent <> zSource.Indent then
      exit(false);
    if Rotate <> zSource.Rotate then
      exit(false);
    if ShrinkToFit <> zSource.ShrinkToFit then
      exit(false);
    if Vertical <> zSource.Vertical then
      exit(false);
    if VerticalText <> zSource.VerticalText then
      exit(false);
    if WrapText <> zSource.WrapText then
      exit(false);
  end
  else
    Result := false;
end;

procedure TZAlignment.SetHorizontal(const Value: TZHorizontalAlignment);
begin
  FHorizontal := Value;
end;

procedure TZAlignment.SetIndent(const Value: Integer);
begin
  FIndent := Value;
end;

procedure TZAlignment.SetRotate(const Value: TZCellTextRotate);
begin
  FRotate := Value;
end;

procedure TZAlignment.SetShrinkToFit(const Value: Boolean);
begin
  FShrinkToFit := Value;
end;

procedure TZAlignment.SetVertical(const Value: TZVerticalAlignment);
begin
  FVertical := Value;
end;

procedure TZAlignment.SetVerticalText(const Value: Boolean);
begin
  FVerticalText := Value;
end;

procedure TZAlignment.SetWrapText(const Value: Boolean);
begin
  FWrapText := Value;
end;

/// /::::::::::::: TZFont :::::::::::::::::////

constructor TZFont.Create;
begin
  inherited;
  FColor.RGB := TColorRec.cWindowText;
  FColor.Indexed := 0;
  FColor.Theme := 0;
  FColor.Tint := 0;
  FSize := 8;
  FCharset := DEFAULT_CHARSET;
  FName := 'MS Sans Serif';
  FStyle := [];
end;

destructor TZFont.Destroy;
begin
  inherited;
end;

function TZFont.GetColor: TColor;
begin
  Result := FColor.RGB;
end;

procedure TZFont.SetColor(const Value: TColor);
begin
  if Value <> FColor.RGB then
  begin
    FColor.RGB := Value;
    FColor.Indexed := 0;
    FColor.Theme := 0;
    FColor.Tint := 0;
  end;
end;

function TZFont.GetHashCode: Integer;
var
  ST: Integer;
begin
  ST := 0;
  if TFontStyle.fsBold in FStyle then
    Inc(ST, 1);
  if TFontStyle.fsItalic in FStyle then
    Inc(ST, 2);
  if TFontStyle.fsUnderline in FStyle then
    Inc(ST, 4);
  if TFontStyle.fsStrikeOut in FStyle then
    Inc(ST, 8);

  Result := 17;
  Result := Result * 23 + Integer(FColor.RGB);
  Result := Result * 23 + Integer(FColor.Indexed);
  Result := Result * 23 + Integer(FColor.Theme);
  Result := Result * 23 + trunc(FColor.Tint * 100000.0);
  Result := Result * 23 + trunc(FSize * 1000.0);
  Result := Result * 23 + Integer(FCharset);
  Result := Result * 23 + string(FName).GetHashCode();
  Result := Result * 23 + ST;
end;

procedure TZFont.Assign(Source: TPersistent);
var
  zSource: TZFont;
  srcFont: TFont;
begin
  if Source is TZFont then
  begin
    zSource := Source as TZFont;
    FColor := zSource.ExcelColor;
    FSize := zSource.Size;
    FCharset := zSource.Charset;
    FName := zSource.Name;
    FStyle := zSource.Style;
  end
  else if Source is TFont then
  begin
    srcFont := Source as TFont;
    {$IFDEF FMX}
    FName := srcFont.Family;
    {$ELSE}
    FColor.RGB := srcFont.Color;
    FColor.Indexed := 0;
    FColor.Theme := 0;
    FColor.Tint := 0;
    FCharset := srcFont.Charset;
    FName := srcFont.Name;
    {$ENDIF}
    FSize := srcFont.Size;
    FStyle := srcFont.Style;
  end
  else
    inherited Assign(Source);
end;

procedure TZFont.AssignTo(Dest: TPersistent);
var
  dstFont: TFont;
begin
  if Dest is TZFont then
    TZFont(Dest).Assign(Self)
  else if Dest is TFont then
  begin
    dstFont := Dest as TFont;
    // А.А.Валуев Свойства, которых нет в TZFont сбрасываем на значения по умолчанию.
    dstFont.Size := Round(FSize);
    dstFont.Style := FStyle;
    {$IFDEF FMX}
    dstFont.Family := FName;
    {$ELSE}
    dstFont.Color := FColor.RGB;
    dstFont.Pitch := fpDefault;
    dstFont.Orientation := 0;
    dstFont.Charset := FCharset;
    dstFont.Name := FName;
    {$ENDIF}
  end
  else
    inherited AssignTo(Dest);
end;

/// /::::::::::::: TZStyle :::::::::::::::::////
// about default font in Excel - http://support.microsoft.com/kb/214123
constructor TZStyle.Create();
begin
  FFont := TZFont.Create();
  FFont.Size := 10;
  FFont.Name := 'Arial';
  FFont.Color := TColorRec.Black;
  FBorder := TZBorder.Create();
  FAlignment := TZAlignment.Create();
  FBGColor.RGB := TColorRec.cWindow;
  FPatternColor := TColorRec.cWindow;
  FCellPattern := ZPNone;
  FNumberFormat := '';
  FNumberFormatId := -1;
  FProtect := true;
  FHideFormula := false;
  FSuperscript := false;
  FSubscript := false;
end;

destructor TZStyle.Destroy();
begin
  FreeAndNil(FFont);
  FreeAndNil(FBorder);
  FreeAndNil(FAlignment);
  inherited Destroy();
end;

procedure TZStyle.Assign(Source: TPersistent);
var
  zSource: TZStyle;
begin
  if Source is TZStyle then
  begin
    zSource := Source as TZStyle;
    FFont.Assign(zSource.Font);
    FBorder.Assign(zSource.Border);
    FAlignment.Assign(zSource.Alignment);
    FBGColor := zSource.FBGColor;
    FPatternColor := zSource.PatternColor;
    FCellPattern := zSource.CellPattern;
    FNumberFormat := zSource.NumberFormat;
    FNumberFormatId := zSource.NumberFormatId;
    FProtect := zSource.Protect;
    FHideFormula := zSource.HideFormula;
    FSuperscript := zSource.Superscript;
    FSubscript := zSource.Subscript;
  end
  else
    inherited Assign(Source);
end;

function TZStyle.IsEqual(Source: TPersistent): Boolean;
var
  zSource: TZStyle;
begin
  Result := false;
  if not(Source is TZStyle) then
    exit;
  zSource := Source as TZStyle;

  if not Border.IsEqual(zSource.Border) then
    exit;
  if not Self.Alignment.IsEqual(zSource.Alignment) then
    exit;
  if BgColor <> zSource.BgColor then
    exit;
  if PatternColor <> zSource.PatternColor then
    exit;
  if CellPattern <> zSource.CellPattern then
    exit;
  if NumberFormat <> zSource.NumberFormat then
    exit;
  if NumberFormatId <> zSource.NumberFormatId then
    exit;
  if Protect <> zSource.Protect then
    exit;
  if HideFormula <> zSource.HideFormula then
    exit;
  if Superscript <> zSource.Superscript then
    exit;
  if Subscript <> zSource.Subscript then
    exit;
  Result := ZEIsFontsEquals(FFont, zSource.Font);
end;

procedure TZStyle.SetFont(const Value: TZFont);
begin
  FFont.Assign(Value);
end;

procedure TZStyle.SetBorder(const Value: TZBorder);
begin
  FBorder.Assign(Value);
end;

procedure TZStyle.SetAlignment(const Value: TZAlignment);
begin
  FAlignment.Assign(Value);
end;

function TZStyle.GetBgColor: TColor;
begin
  Result := FBGColor.RGB;
end;

procedure TZStyle.SetBgColor(const Value: TColor);
begin
  if FBGColor.RGB <> Value then
  begin
    FBGColor.RGB := Value;
    FBGColor.Indexed := 0;
    FBGColor.Theme := 0;
    FBGColor.Tint := 0;
  end;
end;

procedure TZStyle.SetPatternColor(const Value: TColor);
begin
  FPatternColor := Value;
end;

procedure TZStyle.SetCellPattern(const Value: TZCellPattern);
begin
  FCellPattern := Value;
end;

procedure TZStyle.SetSuperscript(const Value: Boolean);
begin
  if FSuperscript <> Value then
  begin
    FSuperscript := Value;
    if FSuperscript then
      FSubscript := false;
  end;
end;

procedure TZStyle.SetSubscript(const Value: Boolean);
begin
  if FSubscript <> Value then
  begin
    FSubscript := Value;
    if FSubscript then
      FSuperscript := false;
  end;
end;

procedure TZStyle.SetNumberFormat(const Value: string);
begin
  FNumberFormat := Value;
  FNumberFormatId := -1;
end;

/// /::::::::::::: TZStyles :::::::::::::::::////

constructor TZStyles.Create();
begin
  FDefaultStyle := TZStyle.Create();
  FCount := 0;
end;

destructor TZStyles.Destroy();
var
  i: Integer;
begin
  FreeAndNil(FDefaultStyle);
  for i := 0 to FCount - 1 do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited Destroy();
end;

function TZStyles.Find(const Style: TZStyle): Integer;
var
  i: Integer;
begin
  Result := -2;
  if DefaultStyle.IsEqual(Style) then
    exit(-1);

  for i := 0 to Count - 1 do
    if Items[i].IsEqual(Style) then
    begin
      exit(i);
    end;
end;

function TZStyles.Add(const Style: TZStyle; CheckMatch: Boolean = false): Integer;
begin
  Result := -2;
  if CheckMatch then
    Result := Find(Style);
  if Result = -2 then
  begin
    Count := Count + 1;
    Items[Count - 1].Assign(Style);
    Result := Count - 1;
  end;
end;

procedure TZStyles.Assign(Source: TPersistent);
var
  srcStyles: TZStyles;
  i: Integer;
begin
  if Source is TZStyles then
  begin
    srcStyles := Source as TZStyles;
    FDefaultStyle.Assign(srcStyles.DefaultStyle);
    Count := srcStyles.Count;
    for i := 0 to Count - 1 do
      FStyles[i].Assign(srcStyles[i]);
  end
  else
    inherited Assign(Source);
end;

procedure TZStyles.SetDefaultStyle(const Value: TZStyle);
begin
  if Assigned(Value) then
    FDefaultStyle.Assign(Value);
end;

procedure TZStyles.SetCount(const Value: Integer);
var
  i: Integer;
begin
  if FCount < Value then
  begin
    SetLength(FStyles, Value);
    for i := FCount to Value - 1 do
    begin
      FStyles[i] := TZStyle.Create;
      FStyles[i].Assign(FDefaultStyle);
    end;
  end
  else
  begin
    for i := Value to FCount - 1 do
      FreeAndNil(FStyles[i]);
    SetLength(FStyles, Value);
  end;
  FCount := Value;
end;

function TZStyles.GetStyle(num: Integer): TZStyle;
begin
  if (num >= 0) and (num < Count) then
    Result := FStyles[num]
  else
    Result := DefaultStyle;
end;

procedure TZStyles.SetStyle(num: Integer; const Value: TZStyle);
begin
  if (num >= 0) and (num < Count) then
    FStyles[num].Assign(Value)
  else if num = -1 then
    DefaultStyle.Assign(Value);
end;

function TZStyles.DeleteStyle(num: Integer): Integer;
begin
  if (num >= 0) and (num < Count) then
  begin
    FreeAndNil(FStyles[num]);
    System.Move(FStyles[num + 1], FStyles[num], (Count - num - 1) * sizeof(FStyles[num]));
    FStyles[Count - 1] := nil;
    Dec(FCount);
    SetLength(FStyles, FCount);
    // при удалении глянуть на ячейки - изменить num на 0, остальные сдвинуть.
    Result := 0;
  end
  else
    Result := -1;
end;

procedure TZStyles.Clear();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FCount := 0;
end;

/// /::::::::::::: TZCell :::::::::::::::::////

constructor TZCell.Create(ASheet: TZSheet);
begin
  FSheet := ASheet;
  FFormula := '';
  FData := '';
  FHref := '';
  FComment := '';
  FCommentAuthor := '';
  FHRefScreenTip := '';
  FCellType := ZEString;
  FCellStyle := -1; // по дефолту
  FAlwaysShowComment := false;
  FShowComment := false;
  FRichText := TRichText.Create();
end;

destructor TZCell.Destroy;
begin
  FRichText.Free();
  inherited;
end;

procedure TZCell.Assign(Source: TPersistent);
var
  zSource: TZCell;
begin
  if Source is TZCell then
  begin
    zSource := Source as TZCell;
    FFormula := zSource.Formula;
    FData := zSource.Data;
    FHref := zSource.HRef;
    FComment := zSource.Comment;
    FCommentAuthor := zSource.CommentAuthor;
    FCellStyle := zSource.CellStyle;
    FCellType := zSource.CellType;
    FAlwaysShowComment := zSource.AlwaysShowComment;
    FShowComment := zSource.ShowComment;
    FRichText.Assign(zSource.FRichText);
  end
  else
    inherited Assign(Source);
end;

procedure TZCell.Clear();
begin
  FFormula := '';
  FData := '';
  FHref := '';
  FComment := '';
  FCommentAuthor := '';
  FHRefScreenTip := '';
  FCellType := ZEString;
  FCellStyle := -1;
  FAlwaysShowComment := false;
  FShowComment := false;
end;

procedure TZCell.ApplyStyleValue(proc: TProc<TZStyle>);
var
  Style: TZStyle;
begin
  Style := TZStyle.Create();
  try
    Style.Assign(Self.Style);
    proc(Style);
    FCellStyle := FSheet.FStore.Styles.Add(Style, true);
  finally
    Style.Free();
  end;
end;

procedure TZCell.SetBorderAround(borderWidth: Byte; BorderColor: TColor; BorderStyle: TZBorderType);
var
  Style: TZStyle;
  bp: TZBordersPos;
begin
  Style := TZStyle.Create();
  try
    Style.Assign(Self.Style);
    for bp := TZBordersPos.bpLeft to TZBordersPos.bpBottom do
    begin
      Style.FBorder[bp].LineStyle := BorderStyle;
      Style.FBorder[bp].Weight := borderWidth;
      Style.FBorder[bp].Color := BorderColor;
    end;
    FCellStyle := FSheet.FStore.Styles.Add(Style, true);
  finally
    Style.Free();
  end;
end;

function TZCell.GetStyle: TZStyle;
begin
  Result := nil;
  if FCellStyle > -1 then
    Result := FSheet.FStore.FStyles[FCellStyle]
  else if FSheet.FStore.FStyles.Count > 0 then
    Result := FSheet.FStore.FStyles[0];
end;

function TZCell.GetBgColor: TColor;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.BgColor;
end;

function TZCell.GetBorderColor(num: TZBordersPos): TColor;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.Border[num].Color;
end;

function TZCell.GetBorderStyle(num: TZBordersPos): TZBorderType;
begin
  Result := TZBorderType.ZENone;
  if FCellStyle > -1 then
    Result := Style.Border[num].LineStyle;
end;

function TZCell.GetBorderWidht(num: TZBordersPos): Byte;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.Border[num].Weight;
end;

function TZCell.GetFontColor: TColor;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.Font.Color;
end;

function TZCell.GetFontSize: Double;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.Font.Size;
end;

function TZCell.GetFontStyle: TFontStyles;
begin
  Result := [];
  if FCellStyle > -1 then
    Result := Style.Font.Style;
end;

function TZCell.GetHorizontalAlignment: TZHorizontalAlignment;
begin
  Result := TZHorizontalAlignment.ZHAutomatic;
  if FCellStyle > -1 then
    Result := Style.Alignment.Horizontal;
end;

function TZCell.GetNumberFormat: string;
begin
  Result := '';
  if FCellStyle > -1 then
    Result := Style.NumberFormat;
end;

function TZCell.GetRotate: TZCellTextRotate;
begin
  Result := 0;
  if FCellStyle > -1 then
    Result := Style.Alignment.Rotate;
end;

function TZCell.GetVerticalAlignment: TZVerticalAlignment;
begin
  Result := TZVerticalAlignment.ZVAutomatic;
  if FCellStyle > -1 then
    Result := Style.Alignment.Vertical;
end;

function TZCell.GetVerticalText: Boolean;
begin
  Result := false;
  if FCellStyle > -1 then
    Result := Style.Alignment.VerticalText;
end;

function TZCell.GetWrapText: Boolean;
begin
  Result := false;
  if FCellStyle > -1 then
    Result := Style.Alignment.WrapText;
end;

procedure TZCell.SetFontColor(const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Color := Value;
    end);
end;

procedure TZCell.SetFontSize(const Value: Double);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Size := Value;
    end);
end;

procedure TZCell.SetFontStyle(const Value: TFontStyles);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Style := Value;
    end);
end;

procedure TZCell.SetBgColor(const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.BgColor := Value;
    end);
end;

procedure TZCell.SetBorderColor(num: TZBordersPos; const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].Color := Value;
    end);
end;

procedure TZCell.SetBorderStyle(num: TZBordersPos; const Value: TZBorderType);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].LineStyle := Value;
    end);
end;

procedure TZCell.SetBorderWidht(num: TZBordersPos; const Value: Byte);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].Weight := Value;
    end);
end;

procedure TZCell.SetHorizontalAlignment(const Value: TZHorizontalAlignment);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Horizontal := Value;
    end);
end;

procedure TZCell.SetNumberFormat(const Value: string);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.NumberFormat := Value;
    end);
end;

procedure TZCell.SetRotate(const Value: TZCellTextRotate);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Rotate := Value;
    end);
end;

procedure TZCell.SetVerticalAlignment(const Value: TZVerticalAlignment);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Vertical := Value;
    end);
end;

procedure TZCell.SetVerticalText(const Value: Boolean);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.VerticalText := Value;
    end);
end;

procedure TZCell.SetWrapText(const Value: Boolean);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.WrapText := Value;
    end);
end;

function TZCell.GetDataAsInteger: Integer;
begin
  Result := StrToInt(Data);
end;

procedure TZCell.SetDataAsDateTime(const Value: TDateTime);
begin
  FCellType := ZEDateTime;
  // FData := ZEDateTimeToStr(Value, true);
  FData := FloatToStr(Value).Replace(',', '.');
end;

function TZCell.GetDataAsDouble: Double;
var
  err: Integer;
  b: Boolean;
  _dt: TDateTime;
begin
  Val(FData, Result, err); // need old-school to ignore regional settings
  if err > 0 then
  begin
    b := true;
    // If datetime ...
    if CellType = ZEDateTime then
      b := not TryZEStrToDateTime(FData, _dt);

    if b then
      Raise EConvertError.Create('ZxCell: Cannot cast data to number')
    else
      Result := _dt;
  end;
end;

function TZCell.GetDataAsDateTime(): TDateTime;
var
  b: Boolean;
begin
  b := false;
  if FData = '' then
    Result := 0
  else if not TryZEStrToDateTime(FData, Result) then
  begin
    // If cell type is number then try convert "float" to datetime
    if CellType = ZENumber then
      Result := AsDouble
    else
      b := true;
  end;
  if (b) then
    Raise EConvertError.Create('ZxCell: Cannot cast data to DateTime');;
end;

procedure TZCell.SetDataAsInteger(const Value: Integer);
begin
  Data := Trim(IntToStr(Value));
  CellType := ZENumber;
// Val adds the prepending space, maybe some I2S implementation would adds too
// and Excel dislikes it. Better safe than sorry.
end;

procedure TZCell.SetDataAsDouble(const Value: Double);
begin
  CellType := ZENumber;
  FData := FloatToStr(Value, invariantFormatSertting).ToUpper;
end;

procedure TZCell.SetDataAsString(const Value: string);
begin
  FData := Value;
  CellType := ZEString;
end;

/// /::::::::::::: TZMergeCells :::::::::::::::::////

constructor TZMergeCells.Create(ASheet: TZSheet);
begin
  FSheet := ASheet;
  FCount := 0;
end;

destructor TZMergeCells.Destroy();
begin
  Clear();
  inherited Destroy();
end;

procedure TZMergeCells.Clear();
begin
  SetLength(FMergeArea, 0);
  FCount := 0;
end;

function TZMergeCells.GetItem(num: Integer): TZMergeArea;
begin
  if (num >= 0) and (num < Count) then
    Result := FMergeArea[num]
  else
  begin
    Result.left := 0;
    Result.Top := 0;
    Result.right := 0;
    Result.Bottom := 0;
  end;
end;

procedure TZMergeCells.SetItem(num: Integer; const rect: TZMergeArea);
begin
  FMergeArea[num] := rect;
end;

function TZMergeCells.AddRect(Rct: TZMergeArea): Byte;
var
  i: Integer;
  function IsCross(rct1, rct2: TZMergeArea): Boolean;
  begin
    Result := (((rct1.left >= rct2.left) and (rct1.left <= rct2.right)) or
      ((rct1.right >= rct2.left) and (rct1.right <= rct2.right))) and
      (((rct1.Top >= rct2.Top) and (rct1.Top <= rct2.Bottom)) or
      ((rct1.Bottom >= rct2.Top) and (rct1.Bottom <= rct2.Bottom)));
  end;

begin
  if Rct.left > Rct.right then
  begin
    i := Rct.left;
    Rct.left := Rct.right;
    Rct.right := i;
  end;

  if Rct.Top > Rct.Bottom then
  begin
    i := Rct.Top;
    Rct.Top := Rct.Bottom;
    Rct.Bottom := i;
  end;

  if (Rct.left < 0) or (Rct.Top < 0) then
  begin
    Result := 1;
    exit;
  end;

  if (Rct.right - Rct.left = 0) and (Rct.Bottom - Rct.Top = 0) then
  begin
    Result := 3;
    exit;
  end;

  for i := 0 to Count - 1 do
    if IsCross(FMergeArea[i], Rct) or IsCross(Rct, FMergeArea[i]) then
    begin
      Result := 2;
      exit;
    end;

  // если надо, увеличиваем кол-во строк/столбцов в хранилище
  if Assigned(FSheet) then
  begin
    if Rct.right > FSheet.ColCount - 1 then
      FSheet.ColCount := Rct.right { + 1 };
    if Rct.Bottom > FSheet.RowCount - 1 then
      FSheet.RowCount := Rct.Bottom { + 1 };
  end;

  Inc(FCount);
  SetLength(FMergeArea, FCount);
  FMergeArea[FCount - 1] := Rct;
  Result := 0;
end;

function TZMergeCells.InLeftTopCorner(ACol, ARow: Integer): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to FCount - 1 do
    if (ACol = FMergeArea[i].left) and (ARow = FMergeArea[i].Top) then
    begin
      Result := i;
      break;
    end;
end;

function TZMergeCells.InMergeRange(ACol, ARow: Integer): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to FCount - 1 do
    if (ACol >= FMergeArea[i].left) and (ACol <= FMergeArea[i].right) and (ARow >= FMergeArea[i].Top) and
      (ARow <= FMergeArea[i].Bottom) then
    begin
      Result := i;
      break;
    end;
end;

function TZMergeCells.IsCrossWithArea(AID, AC1, AR1, AC2, AR2: Integer): Boolean;
begin
  Result := (((Items[AID].left >= AC1) and (Items[AID].left <= AC2)) or
    ((Items[AID].right >= AC1) and (Items[AID].right <= AC2))) and
    (((Items[AID].Top >= AR1) and (Items[AID].Top <= AR2)) or ((Items[AID].Bottom >= AR1) and
    (Items[AID].Bottom <= AR2)));
end;

function TZMergeCells.MergedCols(ACol, ARow: Integer): Integer;
var
  i: Integer;
begin
  Result := 1;
  for i := 0 to FCount - 1 do
  begin
    if (ACol >= FMergeArea[i].left) and (ACol <= FMergeArea[i].right) and (ARow >= FMergeArea[i].Top) and
      (ARow <= FMergeArea[i].Bottom) then
    begin
      Result := (FMergeArea[i].right - FMergeArea[i].left) + 1;
      break;
    end;
  end;
end;

function TZMergeCells.MergedRows(ACol, ARow: Integer): Integer;
var
  i: Integer;
begin
  Result := 1;
  for i := 0 to FCount - 1 do
  begin
    if (ACol >= FMergeArea[i].left) and (ACol <= FMergeArea[i].right) and (ARow >= FMergeArea[i].Top) and
      (ARow <= FMergeArea[i].Bottom) then
    begin
      Result := (FMergeArea[i].Bottom - FMergeArea[i].Top) + 1;
      break;
    end;
  end;
end;

function TZMergeCells.DeleteItem(num: Integer): Boolean;
var
  i: Integer;
begin
  if (num > Count - 1) or (num < 0) then
  begin
    Result := false;
    exit;
  end;

  for i := num to Count - 2 do
    FMergeArea[i] := FMergeArea[i + 1];

  Dec(FCount);
  SetLength(FMergeArea, FCount);
  Result := true;
end;

function TZMergeCells.AddRectXY(left, Top, right, Bottom: Integer): Byte;
var
  Rct: TZMergeArea;
begin
  Rct.left := left;
  Rct.Top := Top;
  Rct.right := right;
  Rct.Bottom := Bottom;
  Result := AddRect(Rct);
end;

/// /::::::::::::: TZRowColOptions :::::::::::::::::////

constructor TZRowColOptions.Create(ASheet: TZSheet);
begin
  inherited Create();
  FSheet := ASheet;
  FHidden := false;
  FAuto := true;
  FStyleID := -1;
  FBreaked := false;
  FOutlineLevel := 0;
end;

procedure TZRowColOptions.Assign(Source: TPersistent);
begin
  if Source is TZRowColOptions then
  begin
    Hidden := (Source as TZRowColOptions).Hidden;
    StyleID := (Source as TZRowColOptions).StyleID;
    FSize := (Source as TZRowColOptions).FSize;
    FAuto := (Source as TZRowColOptions).FAuto;
    FBreaked := (Source as TZRowColOptions).Breaked;
    FOutlineLevel := (Source as TZRowColOptions).FOutlineLevel;
  end
  else
    inherited Assign(Source);
end;

function TZRowColOptions.GetAuto(): Boolean;
begin
  Result := FAuto;
end;

procedure TZRowColOptions.SetAuto(Value: Boolean);
begin
  FAuto := Value;
end;

function TZRowColOptions.GetSizePoint(): Real;
begin
  Result := FSize;
end;

procedure TZRowColOptions.SetSizePoint(Value: Real);
begin
  if Value >= 0 then
    FSize := Value;
end;

function TZRowColOptions.GetSizeMM(): Real;
begin
  Result := PointToMM(FSize);
end;

procedure TZRowColOptions.SetSizeMM(Value: Real);
begin
  FSize := MMToPoint(Value);
end;

/// /::::::::::::: TZRowOptions :::::::::::::::::////

constructor TZRowOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 48;
end;

function TZRowOptions.GetSizePix(): Integer;
var
  t: Real;
begin
  t := 0.265;
  if Assigned(FSheet) and Assigned(FSheet.FStore) then
    t := FSheet.FStore.HorPixelSize;
  Result := PointToPixel(FSize, t);
end;

procedure TZRowOptions.SetSizePix(Value: Integer);
var
  t: Real;
begin
  if Value < 0 then
    exit;
  t := 0.265;
  if Assigned(FSheet) and Assigned(FSheet.FStore) then
    t := FSheet.FStore.HorPixelSize;
  FSize := PixelToPoint(Value, t);
end;

/// /::::::::::::: TZColOptions :::::::::::::::::////

constructor TZColOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 12.75;
end;

function TZColOptions.GetSizePix(): Integer;
var
  t: Real;
begin
  t := 0.265;
  if Assigned(FSheet) and Assigned(FSheet.FStore) then
    t := FSheet.FStore.VertPixelSize;
  Result := PointToPixel(FSize, t);
end;

procedure TZColOptions.SetSizePix(Value: Integer);
var
  t: Real;
begin
  if Value < 0 then
    exit;
  t := 0.265;
  if Assigned(FSheet) and Assigned(FSheet.FStore) then
    t := FSheet.FStore.VertPixelSize;
  FSize := PixelToPoint(Value, t);
end;

/// /::::::::::::: TZHeaderFooterMargins :::::::::::::::::////

constructor TZHeaderFooterMargins.Create();
begin
  FMarginTopBottom := 13;
  FMarginLeft := 0;
  FMarginRight := 0;
  FHeight := 7;
  FUseAutoFitHeight := true;
end;

procedure TZHeaderFooterMargins.Assign(Source: TPersistent);
var
  t: TZHeaderFooterMargins;
begin
  if Source is TZHeaderFooterMargins then
  begin
    t := Source as TZHeaderFooterMargins;
    FMarginTopBottom := t.MarginTopBottom;
    FMarginLeft := t.MarginLeft;
    FMarginRight := t.MarginRight;
    FHeight := t.Height;
    FUseAutoFitHeight := t.UseAutoFitHeight;
  end
  else
    inherited Assign(Source);
end;

function TZHeaderFooterMargins.IsEqual(Source: TPersistent): Boolean;
var
  t: TZHeaderFooterMargins;
begin
  Result := false;
  if Assigned(Source) and (Source is TZHeaderFooterMargins) then
  begin
    t := Source as TZHeaderFooterMargins;
    Result := (FMarginTopBottom = t.MarginTopBottom) and (FMarginLeft = t.MarginLeft) and (FMarginRight = t.MarginRight)
      and (FHeight = t.FHeight) and (FUseAutoFitHeight = t.UseAutoFitHeight);
  end;
end;

/// /::::::::::::: TZSheetOptions :::::::::::::::::////

constructor TZSheetOptions.Create();
begin
  inherited;
  FHeaderMargins := TZHeaderFooterMargins.Create();
  FFooterMargins := TZHeaderFooterMargins.Create();
  FActiveCol := 0;
  FActiveRow := 0;
  FMarginBottom := 25;
  FMarginLeft := 20;
  FMarginTop := 25;
  FMarginRight := 20;
  FPortraitOrientation := true;
  FCenterHorizontal := false;
  FCenterVertical := false;
  FStartPageNumber := 1;
  HeaderMargin := 13;
  FooterMargin := 13;
  FPaperSize := 9;
  FFitToHeight := -1;
  FFitToWidth := -1;
  FDifferentFirst := false;
  FDifferentOddEven := false;
  FHeader := '';
  FFooter := '';
  FEvenHeader := '';
  FEvenFooter := '';
  FFirstPageHeader := '';
  FFirstPageFooter := '';
  FHeaderBGColor := TColorRec.cWindow;
  FFooterBGColor := TColorRec.cWindow;
  FSplitVerticalMode := ZSplitNone;
  FSplitHorizontalMode := ZSplitNone;
  FSplitVerticalValue := 0;
  FSplitHorizontalValue := 0;
  FPaperWidth := 0;
  FPaperHeight := 0;
  FScaleToPercent := 100;
  FScaleToPages := 1;
end;

destructor TZSheetOptions.Destroy();
begin
  FreeAndNil(FHeaderMargins);
  FreeAndNil(FFooterMargins);
  inherited;
end;

function TZSheetOptions.GetHeaderMargin(): word;
begin
  Result := FHeaderMargins.Height;
end;

procedure TZSheetOptions.SetHeaderMargin(Value: word);
begin
  FHeaderMargins.Height := Value;
end;

function TZSheetOptions.GetFooterMargin(): word;
begin
  Result := FFooterMargins.Height;
end;

procedure TZSheetOptions.SetFooterMargin(Value: word);
begin
  FFooterMargins.Height := Value;
end;

procedure TZSheetOptions.Assign(Source: TPersistent);
var
  t: TZSheetOptions;
begin
  if Source is TZSheetOptions then
  begin
    t := Source as TZSheetOptions;
    ActiveCol := t.ActiveCol;
    ActiveRow := t.ActiveRow;
    MarginBottom := t.MarginBottom;
    MarginLeft := t.MarginLeft;
    MarginTop := t.MarginTop;
    MarginRight := t.MarginRight;
    PortraitOrientation := t.PortraitOrientation;
    CenterHorizontal := t.CenterHorizontal;
    CenterVertical := t.CenterVertical;
    StartPageNumber := t.StartPageNumber;
    PaperSize := t.PaperSize;
    FitToHeight := t.FitToHeight;
    FitToWidth := t.FitToWidth;
    SplitVerticalMode := t.SplitVerticalMode;
    SplitHorizontalMode := t.SplitHorizontalMode;
    SplitVerticalValue := t.SplitVerticalValue;
    SplitHorizontalValue := t.SplitHorizontalValue;
    Footer := t.Footer;
    Header := t.Header;
    EvenHeader := t.EvenHeader;
    EvenFooter := t.EvenFooter;
    FirstPageHeader := t.FirstPageHeader;
    FirstPageFooter := t.FirstPageFooter;
    HeaderBGColor := t.HeaderBGColor;
    FooterBGColor := t.FooterBGColor;
    IsDifferentFirst := t.IsDifferentFirst;
    IsDifferentOddEven := t.IsDifferentOddEven;
    ScaleToPercent := t.ScaleToPercent;
    ScaleToPages := t.ScaleToPages;
    HeaderMargins.Assign(t.HeaderMargins);
    FooterMargins.Assign(t.FooterMargins);
  end
  else
    inherited Assign(Source);
end;

/// /::::::::::::: TZSheet :::::::::::::::::////

constructor TZSheet.Create(AStore: TZWorkBook);
var
  i, j: Integer;
begin
  FStore := AStore;
  FRowCount := 0;
  FColCount := 0;
  FDefaultRowHeight := 12.75; // 16;
  FDefaultColWidth := 48; // 60;
  FMergeCells := TZMergeCells.Create(Self);
  SetLength(FCells, FColCount);
  FTabColor := TColorRec.cWindow;
  FProtect := false;
  FRightToLeft := false;
  FSelected := false;
  FSummaryBelow := true;
  FSummaryRight := true;
  FApplyStyles := false;
  FOutlineLevelRow := 0;
  FOutlineLevelCol := 0;
  FDrawingRid := 0;
  SetLength(FRows, FRowCount);
  SetLength(FColumns, FColCount);

  for i := 0 to FColCount - 1 do
  begin
    SetLength(FCells[i], FRowCount);
    for j := 0 to FRowCount - 1 do
      FCells[i, j] := TZCell.Create(Self);
    FColumns[i] := TZColOptions.Create(Self);
    FColumns[i].Width := DefaultColWidth;
  end;

  for i := 0 to FRowCount - 1 do
  begin
    FRows[i] := TZRowOptions.Create(Self);
    FRows[i].Height := DefaultRowHeight;
  end;

  FSheetOptions := TZSheetOptions.Create();
  if Assigned(FStore) and Assigned(FStore.DefaultSheetOptions) then
    FSheetOptions.Assign(FStore.DefaultSheetOptions);

  FPrintRows := TZSheetPrintTitles.Create(Self, false);
  FPrintCols := TZSheetPrintTitles.Create(Self, true);
  FConditionalFormatting := TZConditionalFormatting.Create();
  FCharts := TZEChartStore.Create();
  FDrawing := TZEDrawing.Create(Self);
end;

destructor TZSheet.Destroy();
begin
  try
    FreeAndNil(FMergeCells);
    FreeAndNil(FSheetOptions);
    FPrintRows.Free;
    FPrintCols.Free;
    Clear();
    FCells := nil;
    FRows := nil;
    FColumns := nil;
    FreeAndNil(FConditionalFormatting);
    FreeAndNil(FCharts);
    FreeAndNil(FDrawing);
  finally
    inherited Destroy();
  end;
end;

procedure TZSheet.Assign(Source: TPersistent);
var
  zSource: TZSheet;
  i, j: Integer;
begin
  if Source is TZSheet then
  begin
    zSource := Source as TZSheet;
    /// /////////////////////////////
    RowCount := zSource.RowCount;
    ColCount := zSource.ColCount;
    TabColor := zSource.TabColor;
    FitToPage := zSource.FitToPage;
    Title := zSource.Title;
    Protect := zSource.Protect;
    RightToLeft := zSource.RightToLeft;
    DefaultRowHeight := zSource.DefaultRowHeight;
    DefaultColWidth := zSource.DefaultColWidth;
    FViewMode := zSource.FViewMode;
    FRowBreaks := zSource.FRowBreaks;
    FColBreaks := zSource.FColBreaks;
    FDrawingRid := zSource.FDrawingRid;

    FSummaryBelow := zSource.FSummaryBelow;
    FSummaryRight := zSource.FSummaryRight;
    FApplyStyles := zSource.FApplyStyles;
    FOutlineLevelRow := zSource.FOutlineLevelRow;
    FOutlineLevelCol := zSource.FOutlineLevelCol;

    for i := 0 to RowCount - 1 do
      Rows[i] := zSource.Rows[i];

    for i := 0 to ColCount - 1 do
    begin
      Columns[i] := zSource.Columns[i];
      for j := 0 to RowCount - 1 do
        Cell[i, j] := zSource.Cell[i, j];
    end;

    FSelected := zSource.Selected;
    SheetOptions.Assign(zSource.SheetOptions);

    MergeCells.Clear();
    for i := 0 to zSource.MergeCells.Count - 1 do
      MergeCells.AddRect(zSource.MergeCells.GetItem(i));

    ConditionalFormatting.Assign(zSource.ConditionalFormatting);
    // На этой строке перезаписывается указатель на объект TZEDrawing,
    // но старый объект не удаляется, т.е. происходит утечка памяти.
    // FDrawing := TZEDrawing.Create();
    FDrawing.Assign(zSource.FDrawing);
    RowsToRepeat.Assign(zSource.RowsToRepeat);
    ColsToRepeat.Assign(zSource.ColsToRepeat);
  end
  else
    inherited Assign(Source);
end;

procedure TZSheet.SetConditionalFormatting(Value: TZConditionalFormatting);
begin
  if Assigned(Value) then
    FConditionalFormatting.Assign(Value);
end;

procedure TZSheet.SetCharts(const Value: TZEChartStore);
begin
  if Assigned(Value) then
    FCharts.Assign(Value);
end;

function TZSheet.GetSheetIndex: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to WorkBook.FSheets.Count - 1 do
    if WorkBook.FSheets[i] = Self then
      exit(i);
end;

function TZSheet.GetSheetOptions(): TZSheetOptions;
begin
  Result := FSheetOptions;
end;

procedure TZSheet.InsertRows(ARow, ACount: Integer);
var
  r, c: Integer;
begin
  // resize
  SetRowCount(FRowCount + ACount);

  // append and reloc cells
  for r := Length(FRows) - 1 downto ARow do
  begin
    // reloc rows
    if (r - ACount) < ARow then
    begin
      FRows[r] := TZRowOptions.Create(Self);
      FRows[r].Height := DefaultRowHeight;
    end
    else
    begin
      FRows[r] := FRows[r - ACount];
    end;

    // reloc cells
    for c := 0 to ColCount - 1 do
    begin
      if (r - ACount) < ARow then
      begin
        FCells[c][r].Clear();
      end
      else
      begin
        FCells[c][r].Assign(FCells[c][r - ACount]);
      end;
    end;
  end;

  // reloc merged areas
  for r := 0 to MergeCells.Count - 1 do
  begin
    if MergeCells[r].Top >= ARow then
    begin
      MergeCells.Items[r] := TZMergeArea.Create(MergeCells.Items[r].left, MergeCells.Items[r].Top + ACount,
        MergeCells.Items[r].right, MergeCells.Items[r].Bottom + ACount);
    end;
  end;
end;

function TZSheet.ColsWidth(AFrom, ATo: Integer): Real;
begin
  Result := 0;
  while (AFrom < ColCount) and (AFrom < ATo) do
  begin
    Result := Result + ColWidths[AFrom];
    Inc(AFrom);
  end;
end;

function TZSheet.RowsHeight(AFrom, ATo: Integer): Real;
begin
  Result := 0;
  while (AFrom < RowCount) and (AFrom < ATo) do
  begin
    Result := Result + RowHeights[AFrom];
    Inc(AFrom);
  end;
end;

procedure TZSheet.CopyRows(ARowDst, ARowSrc, ACount: Integer);
var
  r, c, delta: Integer;
begin
  // copy row and cell info
  for r := 0 to ACount - 1 do
  begin
    FRows[ARowDst + r].Assign(FRows[ARowSrc + r]);
    for c := 0 to FColCount - 1 do
    begin
      FCells[c][ARowDst + r].Assign(FCells[c][ARowSrc + r]);
    end;
  end;

  delta := ARowDst - ARowSrc;

  // reloc merged areas
  for r := 0 to MergeCells.Count - 1 do
  begin
    if (MergeCells[r].Top >= ARowSrc) and (MergeCells[r].Bottom < ARowSrc + ACount) then
    begin
      MergeCells.AddRect(TZMergeArea.Create(MergeCells.Items[r].left, MergeCells.Items[r].Top + delta,
        MergeCells.Items[r].right, MergeCells.Items[r].Bottom + delta));
    end;
  end;
end;

procedure TZSheet.SetSheetOptions(Value: TZSheetOptions);
begin
  if Assigned(Value) then
    FSheetOptions.Assign(Value);
end;

procedure TZSheet.SetCorrectTitle(const Value: string);
var
  i: Integer;
  suffix, newTitle: string;
  sheetNames: TDictionary<string, string>;
  regEx: TRegEx;
begin
  newTitle := Trim(Value);
  if newTitle.Length > 31 then
    newTitle := newTitle.Substring(0, 31);
  regEx := TRegEx.Create('[\\/\*\[\]\?:]');
  newTitle := regEx.Replace(newTitle, ' ');
  if newTitle.Trim.IsEmpty then
    newTitle := 'Лист1';

  if Assigned(WorkBook) then
  begin
    sheetNames := TDictionary<string, string>.Create;
    try
      for i := 0 to Self.WorkBook.Sheets.Count - 1 do
        if Self <> WorkBook.Sheets[i] then
          sheetNames.AddOrSetValue(Self.WorkBook.Sheets[i].Title.ToLower, '');
      i := 1;
      suffix := '';
      repeat
        if i > 1 then
          suffix := ' (' + IntToStr(i) + ')';
        if newTitle.Length + suffix.Length > 31 then
          newTitle := newTitle.Substring(0, 31 - suffix.Length);
        Inc(i);
      until not sheetNames.ContainsKey(newTitle.ToLower + suffix);
      newTitle := newTitle + suffix;
    finally
      sheetNames.Free;
    end;
  end;
  Self.FTitle := newTitle;
end;

procedure TZSheet.SetColumn(num: Integer; const Value: TZColOptions);
begin
  if (num >= 0) and (num < FColCount) then
    FColumns[num].Assign(Value);
end;

function TZSheet.GetColumn(num: Integer): TZColOptions;
begin
  if (num >= 0) and (num < FColCount) then
    Result := FColumns[num]
  else
    Result := nil;
end;

{
procedure TZSheet.SetRange(AC1,AR1,AC2,AR2: integer; const Value: TZRange);
begin

end;

procedure TZSheet.SetRangeRef(AFrom, ATo: string; const Value: TZRange);
begin

end;
}

function TZSheet.GetRange(AC1, AR1, AC2, AR2: Integer): IZRange;
begin
  Result := TZRange.Create(Self, AC1, AR1, AC2, AR2);
end;

function TZSheet.GetRangeRef(AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: Integer): IZRange;
var
  AC1, AR1, AC2, AR2: Integer;
begin
  AC1 := ZEGetColByA1(AFromCol);
  AR1 := AFromRow;
  AC2 := ZEGetColByA1(AToCol);
  AR2 := AToRow;
  Result := TZRange.Create(Self, AC1, AR1, AC2, AR2);
end;

procedure TZSheet.SetRow(num: Integer; const Value: TZRowOptions);
begin
  if (num >= 0) and (num < FRowCount) then
    FRows[num].Assign(Value);
end;

function TZSheet.GetRow(num: Integer): TZRowOptions;
begin
  if (num >= 0) and (num < FRowCount) then
    Result := FRows[num]
  else
    Result := nil;
end;

procedure TZSheet.SetColWidth(num: Integer; const Value: Real);
begin
  if (num < ColCount) and (num >= 0) and (Value >= 0) then
    if FColumns[num] <> nil then
      FColumns[num].Width := Value;
end;

function TZSheet.GetColWidth(num: Integer): Real;
begin
  Result := 0;
  if (num < ColCount) and (num >= 0) then
    if Assigned(FColumns[num]) then
      Result := FColumns[num].Width
end;

procedure TZSheet.SetRowHeight(num: Integer; const Value: Real);
begin
  if (num < RowCount) and (num >= 0) and (Value >= 0) then
    if Assigned(FRows[num]) then
      FRows[num].Height := Value;
end;

function TZSheet.GetRowHeight(num: Integer): Real;
begin
  Result := 0;
  if (num < RowCount) and (num >= 0) then
    if Assigned(FRows[num]) then
      Result := FRows[num].Height
end;

procedure TZSheet.SetDefaultColWidth(const Value: Real);
begin
  if Value >= 0 then
    FDefaultColWidth := round(Value * 100) / 100;
end;

procedure TZSheet.SetDefaultRowHeight(const Value: Real);
begin
  if Value >= 0 then
    FDefaultRowHeight := round(Value * 100) / 100;
end;

procedure TZSheet.SetPrintCols(const Value: TZSheetPrintTitles);
begin
  FPrintCols.Assign(Value);
end;

procedure TZSheet.SetPrintRows(const Value: TZSheetPrintTitles);
begin
  FPrintRows.Assign(Value);
end;

procedure TZSheet.Clear();
var
  i, j: Integer;
begin
  for i := 0 to FColCount - 1 do
  begin
    for j := 0 to FRowCount - 1 do
      FreeAndNil(FCells[i][j]);
    SetLength(FCells[i], 0);
    FColumns[i].Free;
    FCells[i] := nil;
  end;

  for i := 0 to FRowCount - 1 do
    FRows[i].Free;

  SetLength(FCells, 0);
  FRowCount := 0;
  FColCount := 0;
  SetLength(FRows, 0);
  SetLength(FColumns, 0);
end;

procedure TZSheet.SetCell(ACol, ARow: Integer; const Value: TZCell);
begin
  if (ACol >= 0) and (ACol < FColCount) and (ARow >= 0) and (ARow < FRowCount) then
    FCells[ACol, ARow].Assign(Value);
end;

procedure TZSheet.SetCellRef(ACol: string; ARow: Integer; const Value: TZCell);
begin
  SetCell(ZEGetColByA1(ACol), ARow, Value);
end;

function TZSheet.GetCell(ACol, ARow: Integer): TZCell;
begin
  Result := nil;
  if (ACol >= 0) and (ACol < FColCount) and (ARow >= 0) and (ARow < FRowCount) then
    Result := FCells[ACol, ARow];
end;

function TZSheet.GetCellRef(ACol: string; ARow: Integer): TZCell;
begin
  Result := GetCell(ZEGetColByA1(ACol), ARow);
end;

procedure TZSheet.SetColCount(const Value: Integer);
var
  i, j: Integer;
begin
  if Value < 0 then
    exit;

  if FColCount > Value then
  begin // todo Repeatable columns may be affected
    for i := Value to FColCount - 1 do
    begin
      for j := 0 to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      SetLength(FCells[i], 0);
      FreeAndNil(FColumns[i]);
    end;

    SetLength(FCells, Value);
    SetLength(FColumns, Value);
  end
  else
  begin
    SetLength(FCells, Value);
    SetLength(FColumns, Value);

    for i := FColCount to Value - 1 do
    begin
      SetLength(FCells[i], FRowCount);
      FColumns[i] := TZColOptions.Create(Self);
      FColumns[i].Width := DefaultColWidth;
      for j := 0 to FRowCount - 1 do
        FCells[i][j] := TZCell.Create(Self);
    end;
  end;
  FColCount := Value;
end;

function TZSheet.GetColCount: Integer;
begin
  Result := FColCount;
end;

procedure TZSheet.SetRowCount(const Value: Integer);
var
  i, j: Integer;
begin
  if Value < 0 then
    exit;

  if FRowCount > Value then
  begin // todo Repeatable rows may be affected
    for i := 0 to FColCount - 1 do
    begin
      for j := Value to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      SetLength(FCells[i], Value);
    end;

    for i := Value to FRowCount - 1 do
      FreeAndNil(FRows[i]);
    SetLength(FRows, Value);
  end
  else
  begin
    for i := 0 to FColCount - 1 do
    begin
      SetLength(FCells[i], Value);
      for j := FRowCount to Value - 1 do
        FCells[i][j] := TZCell.Create(Self);
    end;

    SetLength(FRows, Value);
    for i := FRowCount to Value - 1 do
    begin
      FRows[i] := TZRowOptions.Create(Self);
      FRows[i].Height := DefaultRowHeight;
    end;
  end;
  FRowCount := Value;
end;

function TZSheet.GetRowCount: Integer;
begin
  Result := FRowCount;
end;

/// /::::::::::::: TZSheets :::::::::::::::::////
constructor TZSheets.Create(AStore: TZWorkBook);
begin
  FStore := AStore;
  FCount := 0;
  SetLength(FSheets, 0);
end;

destructor TZSheets.Destroy();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FSheets[i]);
  SetLength(FSheets, 0);
  FSheets := nil;
  FStore := nil;
  inherited Destroy();
end;

function TZSheets.Add(const ATitle: string = ''): TZSheet;
begin
  Result := TZSheet.Create(FStore);
  Result.Title := ATitle;
  SetLength(FSheets, Length(FSheets) + 1);
  FSheets[High(FSheets)] := Result;
  Inc(FCount);
end;

procedure TZSheets.Assign(Source: TPersistent);
var
  t: TZSheets;
  i: Integer;
begin
  if Source is TZSheets then
  begin
    t := Source as TZSheets;
    Count := t.Count;
    for i := 0 to Count - 1 do
      Sheet[i].Assign(t.Sheet[i]);
  end
  else
    inherited Assign(Source);
end;

procedure TZSheets.SetSheetCount(const Value: Integer);
var
  i: Integer;
begin
  if Count < Value then
  begin
    SetLength(FSheets, Value);
    for i := Count to Value - 1 do
      FSheets[i] := TZSheet.Create(FStore);
  end
  else
  begin
    for i := Value to Count - 1 do
      FreeAndNil(FSheets[i]);
    SetLength(FSheets, Value);
  end;
  FCount := Value;
end;

procedure TZSheets.SetSheet(num: Integer; Const Value: TZSheet);
begin
  if (num < Count) and (num >= 0) then
    FSheets[num].Assign(Value);
end;

function TZSheets.GetSheet(num: Integer): TZSheet;
begin
  if (num < Count) and (num >= 0) then
    Result := FSheets[num]
  else
    Result := nil;
end;

/// /::::::::::::: TZEXMLDocumentProperties :::::::::::::::::////

constructor TZEXMLDocumentProperties.Create();
begin
  FAuthor := 'none';
  FLastAuthor := 'none';
  FCompany := 'none';
  FVersion := '11.9999';
  FCreated := Now();
  FWindowHeight := 20000;
  FWindowWidth := 20000;
  FWindowTopX := 150;
  FWindowTopY := 150;
  FModeR1C1 := false;
end;

procedure TZEXMLDocumentProperties.SetAuthor(const Value: string);
begin
  FAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetLastAuthor(const Value: string);
begin
  FLastAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetCompany(const Value: string);
begin
  FCompany := Value;
end;

procedure TZEXMLDocumentProperties.SetVersion(const Value: string);
begin
  FVersion := Value;
end;

procedure TZEXMLDocumentProperties.Assign(Source: TPersistent);
var
  props: TZEXMLDocumentProperties;
begin
  if Source is TZEXMLDocumentProperties then
  begin
    props := Source as TZEXMLDocumentProperties;
    Author := props.Author;
    LastAuthor := props.LastAuthor;
    Created := props.Created;
    Company := props.Company;
    Version := props.Version;
    WindowHeight := props.WindowHeight;
    WindowWidth := props.WindowWidth;
    WindowTopX := props.WindowTopX;
    WindowTopY := props.WindowTopY;
    ModeR1C1 := props.ModeR1C1;
  end
  else
    inherited Assign(Source);
end;

/// /::::::::::::: TZWorkBook :::::::::::::::::////

constructor TZWorkBook.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDocumentProperties := TZEXMLDocumentProperties.Create;
  FStyles := TZStyles.Create();
  FSheets := TZSheets.Create(Self);
  FHorPixelSize := 0.265;
  FVertPixelSize := 0.265;
  FDefaultSheetOptions := TZSheetOptions.Create();
end;

destructor TZWorkBook.Destroy();
begin
  FreeAndNil(FDefaultSheetOptions);
  FreeAndNil(FDocumentProperties);
  FreeAndNil(FStyles);
  FreeAndNil(FSheets);
  inherited Destroy();
end;

function TZWorkBook.AddMediaContent(AFileName: string; AContent: TBytes; ACheckByName: Boolean): Integer;
var
  i: Integer;
begin
  if ACheckByName then
  begin
    for i := 0 to High(FMediaList) do
    begin
      if FMediaList[i].FileName.ToUpper = AFileName.ToUpper then
        exit(i);
    end;
  end
  else
  begin
    for i := 0 to High(FMediaList) do
    begin
      if IsIdenticalByteArray(FMediaList[i].Content, AContent) then
        exit(i);
    end;
  end;

  SetLength(FMediaList, Length(FMediaList) + 1);
  FMediaList[High(FMediaList)].FileName := AFileName;
  FMediaList[High(FMediaList)].Content := AContent;
  Result := High(FMediaList);
end;

procedure TZWorkBook.Assign(Source: TPersistent);
var
  t: TZWorkBook;
begin
  if Source is TZWorkBook then
  begin
    t := Source as TZWorkBook;
    FMediaList := t.FMediaList;
    FDefinedNames := t.FDefinedNames;
    Styles.Assign(t.Styles);
    Sheets.Assign(t.Sheets);
  end
  else if Source is TZStyles then
    Styles.Assign(Source as TZStyles)
  else if Source is TZSheets then
    Sheets.Assign(Source as TZSheets)
  else
    inherited Assign(Source);
end;

procedure TZWorkBook.SetHorPixelSize(Value: Real);
begin
  if Value > 0 then
    FHorPixelSize := Value;
end;

procedure TZWorkBook.SetVertPixelSize(Value: Real);
begin
  if Value > 0 then
    FVertPixelSize := Value;
end;

{$IFNDEF FMX}
procedure TZWorkBook.GetPixelSize(hdc: THandle);
begin
  // горизонтальный размер пикселя в миллиметрах
  HorPixelSize := GetDeviceCaps(hdc, HORZSIZE) / GetDeviceCaps(hdc, HORZRES);
  // вертикальный размер пикселя в миллиметрах
  VertPixelSize := GetDeviceCaps(hdc, VERTSIZE) / GetDeviceCaps(hdc, VERTRES);
end;
{$ENDIF}

function TZWorkBook.GetDefaultSheetOptions(): TZSheetOptions;
begin
  Result := FDefaultSheetOptions;
end;

procedure TZWorkBook.SetDefaultSheetOptions(Value: TZSheetOptions);
begin
  if Assigned(Value) then
    FDefaultSheetOptions.Assign(Value);
end;

function TZWorkBook.GetDrawing(num: Integer): TZEDrawing;
var
  i, n: Integer;
begin
  Result := nil;
  n := 0;
  for i := 0 to Sheets.Count - 1 do
  begin
    if not Sheets[i].Drawing.IsEmpty then
    begin
      if n = num then
      begin
        Result := Sheets[i].Drawing;
        exit;
      end;
      Inc(n);
    end;
  end;
end;

function TZWorkBook.GetDrawingSheetNum(Value: TZEDrawing): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Sheets.Count - 1 do
  begin
    if Value = Sheets[i].Drawing then
    begin
      Result := i;
      exit;
    end;
  end;
end;

{ TZSheetPrintTitles }

procedure TZSheetPrintTitles.Assign(Source: TPersistent);
var
  f, t: word;
  a: Boolean;
begin
  if Source is TZSheetPrintTitles then
  begin
    f := TZSheetPrintTitles(Source).From;
    t := TZSheetPrintTitles(Source).Till;
    a := TZSheetPrintTitles(Source).Active;

    if a then
      RequireValid(f, t);

    FFrom := f;
    FTill := t;
    FActive := a;
  end
  else
    inherited;
end;

constructor TZSheetPrintTitles.Create(const owner: TZSheet; const ForColumns: Boolean);
begin
  if nil = owner then
    raise Exception.Create(Self.ClassName + ' requires an existing worksheet for owners.');
  Self.FOwner := owner;
  Self.FColumns := ForColumns;
end;

procedure TZSheetPrintTitles.SetActive(const Value: Boolean);
begin
  if Value then
    RequireValid(FFrom, FTill);
  FActive := Value
end;

procedure TZSheetPrintTitles.SetFrom(const Value: word);
begin
  if Active then
    RequireValid(Value, FTill);
  FFrom := Value;
end;

procedure TZSheetPrintTitles.SetTill(const Value: word);
begin
  if Active then
    RequireValid(FFrom, Value);
  FTill := Value;
end;

function TZSheetPrintTitles.ToString: string;
var
  c: char;
begin
  If Active then
  begin
    if FColumns then
      c := 'C'
    else
      c := 'R';
    Result := c + IntToStr(From + 1) + ':' + c + IntToStr(Till + 1);
  end
  else
    Result := '';
end;

procedure TZSheetPrintTitles.RequireValid(const AFrom, ATill: word);
begin
  if not Valid(AFrom, ATill) then
    raise Exception.Create('Invalid printable titles for the worksheet.');
end;

function TZSheetPrintTitles.Valid(const AFrom, ATill: word): Boolean;
var
  UpperLimit: word;
begin
  Result := false;
  if AFrom > ATill then
    exit;

  if FColumns then
    UpperLimit := FOwner.ColCount
  else
    UpperLimit := FOwner.RowCount;

  if ATill >= UpperLimit then
    exit;

  Result := true;
end;

/// /::::::::::::: TZConditionalStyleItem :::::::::::::::::////

constructor TZConditionalStyleItem.Create();
begin
  Clear();
end;

procedure TZConditionalStyleItem.Clear();
begin
  FCondition := ZCFCellContentOperator;
  FConditionOperator := ZCFOpEqual;
  FApplyStyleID := -1;
  FValue1 := '';
  FValue2 := '';
  FBaseCellPageIndex := -1;
  FBaseCellRowIndex := 0;
  FBaseCellColumnIndex := 0;
end;

procedure TZConditionalStyleItem.Assign(Source: TPersistent);
var
  t: TZConditionalStyleItem;
begin
  if (Source is TZConditionalStyleItem) then
  begin
    t := (Source as TZConditionalStyleItem);
    Condition := t.Condition;
    ConditionOperator := t.ConditionOperator;
    Value1 := t.Value1;
    Value2 := t.Value2;
    ApplyStyleID := t.ApplyStyleID;
    BaseCellColumnIndex := t.BaseCellColumnIndex;
    BaseCellPageIndex := t.BaseCellPageIndex;
    BaseCellRowIndex := t.BaseCellRowIndex;
  end
  else
    inherited Assign(Source);
end;

function TZConditionalStyleItem.IsEqual(Source: TPersistent): Boolean;
var
  t: TZConditionalStyleItem;
begin
  Result := false;
  if (Source is TZConditionalStyleItem) then
  begin
    t := (Source as TZConditionalStyleItem);
    if (Condition <> t.Condition) then
      exit;
    if (ConditionOperator <> t.ConditionOperator) then
      exit;
    if (ApplyStyleID <> t.ApplyStyleID) then
      exit;
    if (BaseCellColumnIndex <> t.BaseCellColumnIndex) then
      exit;
    if (BaseCellPageIndex <> t.BaseCellPageIndex) then
      exit;
    if (BaseCellRowIndex <> t.BaseCellRowIndex) then
      exit;
    if (Value1 <> t.Value1) then
      exit;
    if (Value2 <> t.Value2) then
      exit;
    Result := true;
  end;
end;

/// /::::::::::::: TZConditionalStyle :::::::::::::::::////

constructor TZConditionalStyle.Create();
var
  i: Integer;
begin
  FCount := 0;
  FMaxCount := 3;
  SetLength(FConditions, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FConditions[i] := TZConditionalStyleItem.Create();
  FAreas := TZConditionalAreas.Create();
end;

destructor TZConditionalStyle.Destroy();
var
  i: Integer;
begin
  for i := 0 to FMaxCount - 1 do
    if (Assigned(FConditions[i])) then
      FreeAndNil(FConditions[i]);
  SetLength(FConditions, 0);
  FConditions := nil;
  FreeAndNil(FAreas);
  inherited;
end;

procedure TZConditionalStyle.Assign(Source: TPersistent);
var
  t: TZConditionalStyle;
  i: Integer;
begin
  if (Source is TZConditionalStyle) then
  begin
    t := Source as TZConditionalStyle;
    Count := t.Count;
    FAreas.Assign(t.Areas);
    for i := 0 to Count - 1 do
      FConditions[i].Assign(t.Items[i]);
  end
  else
    inherited Assign(Source);
end;

function TZConditionalStyle.IsEqual(Source: TPersistent): Boolean;
var
  t: TZConditionalStyle;
  i: Integer;
begin
  Result := false;
  if (Source is TZConditionalStyle) then
  begin
    t := Source as TZConditionalStyle;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FConditions[i].IsEqual(t.Items[i])) then
        exit;
    Result := true;
  end;
end;

function TZConditionalStyle.GetItem(num: Integer): TZConditionalStyleItem;
begin
  Result := nil;
  if (num >= 0) and (num < Count) then
    Result := FConditions[num];
end;

procedure TZConditionalStyle.SetItem(num: Integer; Value: TZConditionalStyleItem);
begin
  if (num >= 0) and (num < Count) then
    FConditions[num].Assign(Value);
end;

procedure TZConditionalStyle.SetCount(Value: Integer);
var
  i: Integer;
begin
  // TODO: нужно ли ограничение на максимальное кол-во?
  if (Value >= 0) then
  begin
    if (Value < FCount) then
    begin
      for i := Value to FCount - 1 do
        FConditions[i].Clear();
    end
    else if (Value > FMaxCount) then
    begin
      SetLength(FConditions, Value);
      for i := FMaxCount to Value - 1 do
        FConditions[i] := TZConditionalStyleItem.Create();
      FMaxCount := Value;
    end;
    FCount := Value;
  end;
end;

procedure TZConditionalStyle.SetAreas(Value: TZConditionalAreas);
begin
  if (Assigned(Value)) then
    FAreas.Assign(Value);
end;

function TZConditionalStyle.Add(): TZConditionalStyleItem;
begin
  Count := Count + 1;
  Result := FConditions[Count - 1];
end;

function TZConditionalStyle.Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem;
begin
  Result := Add();
  if (Assigned(StyleItem)) then
    Result.Assign(StyleItem);
end;

procedure TZConditionalStyle.Delete(num: Integer);
var
  i: Integer;
  t: TZConditionalStyleItem;
begin
  if (num >= 0) and (num < Count) then
  begin
    t := FConditions[num];
    for i := num to Count - 2 do
      FConditions[i] := FConditions[i + 1];
    if (Count > 0) then
      FConditions[Count - 1] := t;
    Count := Count - 1;
  end;
end;

procedure TZConditionalStyle.Insert(num: Integer);
begin
  Insert(num, nil);
end;

procedure TZConditionalStyle.Insert(num: Integer; StyleItem: TZConditionalStyleItem);
var
  i: Integer;
  t: TZConditionalStyleItem;
begin
  if (num >= 0) and (num < Count) then
  begin
    Add();
    t := FConditions[Count - 1];
    for i := Count - 1 downto num + 1 do
      FConditions[i] := FConditions[i - 1];
    FConditions[num] := t;
    if (Assigned(StyleItem)) then
      FConditions[num].Assign(StyleItem);
  end;
end;

/// /::::::::::::: TZConditionalAreaItem :::::::::::::::::////

constructor TZConditionalAreaItem.Create();
begin
  Create(0, 0, 1, 1);
end;

constructor TZConditionalAreaItem.Create(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer);
begin
  Row := RowNum;
  Column := ColumnNum;
  Width := AreaWidth;
  Height := AreaHeight;
end;

procedure TZConditionalAreaItem.SetRow(Value: Integer);
begin
  if Value >= 0 then
    FRow := Value;
end;

procedure TZConditionalAreaItem.SetColumn(Value: Integer);
begin
  if Value >= 0 then
    FColumn := Value;
end;

procedure TZConditionalAreaItem.SetWidth(Value: Integer);
begin
  if Value >= 0 then
    FWidth := Value;
end;

procedure TZConditionalAreaItem.SetHeight(Value: Integer);
begin
  if Value >= 0 then
    FHeight := Value;
end;

procedure TZConditionalAreaItem.Assign(Source: TPersistent);
var
  t: TZConditionalAreaItem;
begin
  if (Source is TZConditionalAreaItem) then
  begin
    t := Source as TZConditionalAreaItem;
    Row := t.Row;
    Column := t.Column;
    Height := t.Height;
    Width := t.Width;
  end
  else
    inherited Assign(Source);
end;

function TZConditionalAreaItem.IsEqual(Source: TPersistent): Boolean;
var
  t: TZConditionalAreaItem;
begin
  Result := false;
  if (Source is TZConditionalAreaItem) then
  begin
    t := Source as TZConditionalAreaItem;
    if (FRow <> t.Row) then
      exit;
    if (FColumn <> t.Column) then
      exit;
    if (FWidth <> t.Width) then
      exit;
    if (FHeight <> t.Height) then
      exit;
    Result := true;
  end;
end;

/// /::::::::::::: TZConditionalAreas :::::::::::::::::////

constructor TZConditionalAreas.Create();
begin
  FCount := 1;
  SetLength(FItems, FCount);
  FItems[0] := TZConditionalAreaItem.Create();
end;

destructor TZConditionalAreas.Destroy();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    if (Assigned(FItems[i])) then
      FreeAndNil(FItems[i]);
  SetLength(FItems, 0);
  FItems := nil;
  inherited;
end;

procedure TZConditionalAreas.SetCount(Value: Integer);
var
  i: Integer;
begin
  if ((Value >= 0) and (Value <> Count)) then
  begin
    if (Value < Count) then
    begin
      for i := Value to Count - 1 do
        if (Assigned(FItems[i])) then
          FreeAndNil(FItems[i]);
      SetLength(FItems, Value);
    end
    else if (Value > Count) then
    begin
      SetLength(FItems, Value);
      for i := Count to Value - 1 do
        FItems[i] := TZConditionalAreaItem.Create();
    end;
    FCount := Value;
  end;
end;

function TZConditionalAreas.GetItem(num: Integer): TZConditionalAreaItem;
begin
  Result := nil;
  if ((num >= 0) and (num < FCount)) then
    Result := FItems[num];
end;

procedure TZConditionalAreas.SetItem(num: Integer; Value: TZConditionalAreaItem);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FItems[num].Assign(Value);
end;

function TZConditionalAreas.Add(): TZConditionalAreaItem;
begin
  SetCount(FCount + 1);
  Result := FItems[FCount - 1];
end;

function TZConditionalAreas.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer): TZConditionalAreaItem;
begin
  Result := Add();
  Result.Row := RowNum;
  Result.Column := ColumnNum;
  Result.Width := AreaWidth;
  Result.Height := AreaHeight;
end;

procedure TZConditionalAreas.Assign(Source: TPersistent);
var
  t: TZConditionalAreas;
  i: Integer;
begin
  if (Source is TZConditionalAreas) then
  begin
    t := Source as TZConditionalAreas;
    Count := t.Count;
    for i := 0 to Count - 1 do
      FItems[i].Assign(t.Items[i]);
  end
  else
    inherited Assign(Source);
end;

procedure TZConditionalAreas.Delete(num: Integer);
var
  t: TZConditionalAreaItem;
  i: Integer;
begin
  if ((num >= 0) and (num < Count)) then
  begin
    t := FItems[num];
    for i := num to Count - 2 do
      FItems[i] := FItems[i + 1];
    FItems[Count - 1] := t;
    Count := Count - 1;
  end;
end;

// Определяет, находится ли ячейка в области
// INPUT
// ColumnNum: integer  - номер столбца ячейки
// RowNum: integer     - номер строки ячейки
// RETURN
// boolean - true - ячейка входит в область
function TZConditionalAreas.IsCellInArea(ColumnNum, RowNum: Integer): Boolean;
var
  i, X, Y, xx, yy: Integer;
begin
  Result := false;
  for i := 0 to Count - 1 do
  begin
    X := FItems[i].Column;
    Y := FItems[i].Row;
    xx := X + FItems[i].Width;
    yy := Y + FItems[i].Height;
    if ((ColumnNum >= X) and (ColumnNum < xx) and (RowNum >= Y) and (RowNum < yy)) then
    begin
      Result := true;
      break;
    end;
  end;
end;

function TZConditionalAreas.IsEqual(Source: TPersistent): Boolean;
var
  t: TZConditionalAreas;
  i: Integer;
begin
  Result := false;
  if (Source is TZConditionalAreas) then
  begin
    t := Source as TZConditionalAreas;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FItems[i].IsEqual(t.Items[i])) then
        exit;
    Result := true;
  end;
end;

/// /::::::::::::: TConditionalFormatting :::::::::::::::::////

constructor TZConditionalFormatting.Create();
begin
  FCount := 0;
  SetLength(FStyles, 0);
end;

destructor TZConditionalFormatting.Destroy();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    if (Assigned(FStyles[i])) then
      FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited;
end;

procedure TZConditionalFormatting.SetCount(Value: Integer);
var
  i: Integer;
begin
  if Value >= 0 then
  begin
    if Value < FCount then
    begin
      for i := Value to FCount - 1 do
        FreeAndNil(FStyles[i]);
      SetLength(FStyles, Value);
    end
    else if Value > FCount then
    begin
      SetLength(FStyles, Value);
      for i := FCount to Value - 1 do
        FStyles[i] := TZConditionalStyle.Create();
    end;
    FCount := Value;
  end;
end;

function TZConditionalFormatting.Add(): TZConditionalStyle;
begin
  Result := Add(nil);
end;

function TZConditionalFormatting.GetItem(num: Integer): TZConditionalStyle;
begin
  Result := nil;
  if (num >= 0) and (num < Count) then
    Result := FStyles[num];
end;

procedure TZConditionalFormatting.SetItem(num: Integer; Value: TZConditionalStyle);
begin
  if (num >= 0) and (num < Count) then
    if Assigned(Value) then
      FStyles[num].Assign(Value);
end;

function TZConditionalFormatting.Add(Style: TZConditionalStyle): TZConditionalStyle;
begin
  Count := Count + 1;
  Result := FStyles[Count - 1];
  if Assigned(Style) then
    Result.Assign(Style);
end;

// Добавить условное форматирование с областью
// INPUT
// ColumnNum: integer  - номер колонки
// RowNum: integer     - номер строки
// AreaWidth: integer  - ширина области
// AreaHeight: integer - высота области
// RETURN
// TZConditionalStyle - добавленный стиль
function TZConditionalFormatting.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: Integer): TZConditionalStyle;
var
  t: TZConditionalAreaItem;
begin
  Result := Add(nil);
  t := Result.Areas[0];
  t.Row := RowNum;
  t.Column := ColumnNum;
  t.Width := AreaWidth;
  t.Height := AreaHeight;
end;

procedure TZConditionalFormatting.Clear();
begin
  SetCount(0);
end;

// Delete condition formatting item
// INPUT
// num: integer - number of CF item
// RETURN
// boolean - true - item deleted
function TZConditionalFormatting.Delete(num: Integer): Boolean;
var
  i: Integer;
begin
  Result := false;
  if (num >= 0) and (num < Count) then
  begin
    FreeAndNil(FStyles[num]);
    for i := num to FCount - 2 do
      FStyles[num] := FStyles[num + 1];
    Dec(FCount);
  end;
end;

procedure TZConditionalFormatting.Assign(Source: TPersistent);
var
  t: TZConditionalFormatting;
  i: Integer;
begin
  if Source is TZConditionalFormatting then
  begin
    t := Source as TZConditionalFormatting;
    FCount := t.Count;
    for i := 0 to FCount - 1 do
      FStyles[i].Assign(t.Items[i]);
  end
  else
    inherited Assign(Source);
end;

function TZConditionalFormatting.IsEqual(Source: TPersistent): Boolean;
var
  t: TZConditionalFormatting;
  i: Integer;
begin
  Result := false;
  if Source is TZConditionalFormatting then
  begin
    t := Source as TZConditionalFormatting;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FStyles[i].IsEqual(t.Items[i])) then
        exit;
    Result := true;
  end;
end;

/// /::::::::::::: TZECommonFrameAncestor :::::::::::::::::////

constructor TZECommonFrameAncestor.Create();
begin
  FX := 0;
  FY := 0;
  FWidth := 10;
  FHeight := 10;
  FTransform := TZETransform.Create();
end;

constructor TZECommonFrameAncestor.Create(AX, AY, AWidth, AHeight: Integer);
begin
  FX := AX;
  FY := AY;
  FWidth := AWidth;
  FHeight := AHeight;
  FTransform := TZETransform.Create();
end;

destructor TZECommonFrameAncestor.Destroy();
begin
  FreeAndNil(FTransform);
  inherited;
end;

function TZECommonFrameAncestor.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZECommonFrameAncestor;
begin
  Result := false;
  if Assigned(Source) and (Source is TZECommonFrameAncestor) then
  begin
    tmp := Source as TZECommonFrameAncestor;
    Result := (FX = tmp.X) and (FY = tmp.Y) and (FWidth = tmp.Width) and (FHeight = tmp.Height);
    if (Result) then
      Result := FTransform.IsEqual(tmp.Transform);
  end;
end;

procedure TZECommonFrameAncestor.Assign(Source: TPersistent);
var
  b: Boolean;
  tmp: TZECommonFrameAncestor;
begin
  b := Assigned(Source);
  if b then
    b := Source is TZECommonFrameAncestor;

  if b then
  begin
    tmp := Source as TZECommonFrameAncestor;
    FX := tmp.X;
    FY := tmp.Y;
    FWidth := tmp.Width;
    FHeight := tmp.Height;
    FTransform.Assign(tmp.Transform);
  end
  else
    inherited Assign(Source);
end;

procedure TZECommonFrameAncestor.SetHeight(Value: Integer);
begin
  FHeight := Value;
end;

procedure TZECommonFrameAncestor.SetTransform(const Value: TZETransform);
begin
  if Assigned(Value) then
    FTransform.Assign(Value);
end;

procedure TZECommonFrameAncestor.SetWidth(Value: Integer);
begin
  FWidth := Value;
end;

procedure TZECommonFrameAncestor.SetX(Value: Integer);
begin
  FX := Value;
end;

procedure TZECommonFrameAncestor.SetY(Value: Integer);
begin
  FY := Value;
end;

/// /::::::::::::: TZETransform :::::::::::::::::////

constructor TZETransform.Create();
begin
  Clear();
end;

procedure TZETransform.Assign(Source: TPersistent);
var
  b: Boolean;
  tmp: TZETransform;
begin
  b := Assigned(Source);
  if b then
    b := Source is TZETransform;

  if b then
  begin
    tmp := Source as TZETransform;
    FRotate := tmp.Rotate;
    FScaleX := tmp.ScaleX;
    FScaleY := tmp.ScaleY;
    FSkewX := tmp.SkewX;
    FSkewY := tmp.SkewY;
    FTranslateX := tmp.TranslateX;
    FTranslateY := tmp.TranslateY;
  end
  else
    inherited Assign(Source);
end;

procedure TZETransform.Clear();
begin
  FRotate := 0;
  FScaleX := 1;
  FScaleY := 1;
  FSkewX := 0;
  FSkewY := 0;
  FTranslateX := 0;
  FTranslateY := 0;
end;

function TZETransform.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZETransform;
begin
  Result := Assigned(Source);
  if (Result) then
    Result := Source is TZETransform;

  if (Result) then
  begin
    tmp := Source as TZETransform;
    Result := (FRotate = tmp.Rotate) and (FScaleX = tmp.ScaleX) and (FScaleY = tmp.ScaleY) and (FSkewX = tmp.SkewX) and
      (FSkewY = tmp.SkewY) and (FTranslateX = tmp.TranslateX) and (FTranslateY = tmp.TranslateY);
  end;
end;

/// /::::::::::::: TZEChartRangeItem :::::::::::::::::////

constructor TZEChartRangeItem.Create();
begin
  FSheetNum := -1;
  FRow := 0;
  FCol := 0;
  FWidth := 1;
  FHeight := 1;
end;

constructor TZEChartRangeItem.Create(ASheetNum: Integer; ACol, ARow, AWidth, AHeight: Integer);
begin
  FSheetNum := ASheetNum;
  FRow := ARow;
  FCol := ACol;
  FWidth := AWidth;
  FHeight := AHeight;
end;

procedure TZEChartRangeItem.Assign(Source: TPersistent);
var
  tmp: TZEChartRangeItem;
  b: Boolean;
begin
  b := Assigned(Source);
  if b then
  begin
    b := Source is TZEChartRangeItem;
    if b then
    begin
      tmp := Source as TZEChartRangeItem;
      FSheetNum := tmp.SheetNum;
      FRow := tmp.Row;
      FCol := tmp.Col;
      FWidth := tmp.Width;
      FHeight := tmp.Height;
    end;
  end;

  if not b then
    inherited Assign(Source);
end;

function TZEChartRangeItem.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartRangeItem;
begin
  Result := Assigned(Source);
  if Result then
  begin
    Result := Source is TZEChartRangeItem;
    if Result then
    begin
      tmp := Source as TZEChartRangeItem;
      Result := (FSheetNum = tmp.SheetNum) and (FCol = tmp.Col) and (FRow = tmp.Row) and (FHeight = tmp.Height) and
        (FWidth = tmp.Width);
    end;
  end;
end;

/// /::::::::::::: TZEChartRange :::::::::::::::::////

constructor TZEChartRange.Create();
begin
  FCount := 0;
end;

destructor TZEChartRange.Destroy();
begin
  Clear();
  inherited;
end;

function TZEChartRange.GetItem(num: Integer): TZEChartRangeItem;
begin
  Result := nil;
  if (num >= 0) and (num < FCount) then
    Result := FItems[num];
end;

procedure TZEChartRange.SetItem(num: Integer; const Value: TZEChartRangeItem);
begin
  if (num >= 0) and (num < FCount) then
    FItems[num].Assign(Value);
end;

function TZEChartRange.Add(): TZEChartRangeItem;
begin
  SetLength(FItems, FCount + 1);
  Result := TZEChartRangeItem.Create();
  FItems[FCount] := Result;
  Inc(FCount);
end;

function TZEChartRange.Add(const ItemForClone: TZEChartRangeItem): TZEChartRangeItem;
begin
  Result := Add();
  if Assigned(ItemForClone) then
    Result.Assign(ItemForClone);
end;

function TZEChartRange.Delete(num: Integer): Boolean;
var
  i: Integer;
begin
  Result := (num >= 0) and (num < FCount);
  if Result then
  begin
    FreeAndNil(FItems[num]);
    for i := num to FCount - 2 do
      FItems[i] := FItems[i + 1];
    Dec(FCount);
    SetLength(FItems, FCount);
  end;
end;

procedure TZEChartRange.Clear();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FItems[i]);
  FCount := 0;
  SetLength(FItems, 0);
end;

procedure TZEChartRange.Assign(Source: TPersistent);
var
  tmp: TZEChartRange;
  b: Boolean;
  i: Integer;
begin
  b := Assigned(Source);
  if b then
  begin
    b := Source is TZEChartRange;
    if b then
    begin
      tmp := Source as TZEChartRange;
      if FCount > tmp.Count then
      begin
        for i := tmp.Count to FCount - 1 do
          FreeAndNil(FItems[i]);
        FCount := tmp.Count;
        SetLength(FItems, FCount);
      end
      else if FCount < tmp.Count then
      begin
        SetLength(FItems, tmp.Count);
        for i := FCount to tmp.Count - 1 do
          FItems[i] := TZEChartRangeItem.Create();
        FCount := tmp.Count;
        SetLength(FItems, FCount);
      end;

      for i := 0 to FCount - 1 do
        FItems[i].Assign(tmp.Items[i]);
    end;
  end;

  if not b then
    inherited Assign(Source);
end;

function TZEChartRange.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartRange;
  i: Integer;
begin
  Result := Assigned(Source);
  if Result then
  begin
    Result := Source is TZEChartRange;
    if Result then
    begin
      tmp := Source as TZEChartRange;
      Result := FCount = tmp.Count;
      if Result then
        for i := 0 to FCount - 1 do
          if not FItems[i].IsEqual(tmp.Items[i]) then
          begin
            Result := false;
            break;
          end;
    end;
  end;
end;

/// /::::::::::::: TZEChartTitleItem :::::::::::::::::////

constructor TZEChartTitleItem.Create();
begin
  FFont := TZFont.Create();
  FText := '';
  FRotationAngle := 0;
  FIsDisplay := true;
end;

destructor TZEChartTitleItem.Destroy();
begin
  FreeAndNil(FFont);
  inherited;
end;

procedure TZEChartTitleItem.Assign(Source: TPersistent);
var
  tmp: TZEChartTitleItem;
  b: Boolean;
begin
  b := Assigned(Source);
  if (b) then
  begin
    b := Source is TZEChartTitleItem;
    if (b) then
    begin
      tmp := Source as TZEChartTitleItem;

      FFont.Assign(tmp.Font);
      FText := tmp.Text;
      FRotationAngle := tmp.RotationAngle;
      FIsDisplay := tmp.IsDisplay;
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartTitleItem.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartTitleItem;
begin
  Result := Assigned(Source);
  if (Result) then
  begin
    Result := Source is TZEChartTitleItem;
    if (Result) then
    begin
      tmp := Source as TZEChartTitleItem;
      Result := false;
      if (FIsDisplay = tmp.IsDisplay) then
        if (FRotationAngle = tmp.RotationAngle) then
          if (FText = tmp.Text) then
            Result := ZEIsFontsEquals(FFont, tmp.Font);
    end;
  end;
end;

procedure TZEChartTitleItem.SetFont(const Value: TZFont);
begin
  if (Assigned(Value)) then
    FFont.Assign(Value);
end;

/// /::::::::::::: TZEChartLegend :::::::::::::::::////

constructor TZEChartLegend.Create();
begin
  inherited;
  FPosition := ZELegendStart;
  FAlign := ZELegendAlignCenter;
end;

procedure TZEChartLegend.Assign(Source: TPersistent);
var
  tmp: TZEChartLegend;
begin
  inherited Assign(Source);

  if Assigned(Source) and (Source is TZEChartLegend) then
  begin
    tmp := Source as TZEChartLegend;
    FAlign := tmp.Align;
    FPosition := tmp.Position;
  end;
end;

function TZEChartLegend.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartLegend;
begin
  Result := inherited IsEqual(Source);
  if (Result) then
  begin
    Result := false;
    if (Source is TZEChartLegend) then
    begin
      tmp := Source as TZEChartLegend;
      Result := (FPosition = tmp.Position) and (FAlign = tmp.Align);
    end;
  end;
end;

/// /::::::::::::: TZEChartAxis :::::::::::::::::////

constructor TZEChartAxis.Create();
begin
  inherited;
  FLogarithmic := false;
  FReverseDirection := false;
  FScaleMin := 0;
  FScaleMax := 20000;
  FAutoScaleMin := true;
  FAutoScaleMax := true;
end;

procedure TZEChartAxis.Assign(Source: TPersistent);
var
  tmp: TZEChartAxis;
begin
  inherited Assign(Source);
  if (Source is TZEChartAxis) then
  begin
    tmp := Source as TZEChartAxis;
    FLogarithmic := tmp.Logarithmic;
    FReverseDirection := tmp.ReverseDirection;
    FScaleMin := tmp.ScaleMin;
    FScaleMax := tmp.ScaleMax;
    FAutoScaleMin := tmp.AutoScaleMin;
    FAutoScaleMax := tmp.AutoScaleMax;
  end
end;

function TZEChartAxis.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartAxis;
begin
  Result := inherited IsEqual(Source);
  if (Result) then
  begin
    Result := Source is TZEChartAxis;
    if (Result) then
    begin
      tmp := Source as TZEChartAxis;
      Result := (FLogarithmic = tmp.FLogarithmic) and (FReverseDirection = tmp.FReverseDirection) and
        (FAutoScaleMin = tmp.AutoScaleMin) and (FAutoScaleMax = tmp.AutoScaleMax);
      if (Result) then
      begin
        // TODO: for comparsion double values need check observational errors?
        if (not FAutoScaleMin) then
          Result := FScaleMin <> tmp.ScaleMin;
        if (Result and (not FAutoScaleMin)) then
          Result := FScaleMax <> tmp.ScaleMax;
      end;
    end;
  end;
end;

/// /::::::::::::: TZEChartSettings :::::::::::::::::////

constructor TZEChartSettings.Create();
begin
  FJapanCandle := true;
end;

destructor TZEChartSettings.Destroy();
begin
  inherited;
end;

procedure TZEChartSettings.Assign(Source: TPersistent);
var
  tmp: TZEChartSettings;
  b: Boolean;
begin
  b := Source <> nil;
  if (b) then
  begin
    b := Source is TZEChartSettings;
    if (b) then
    begin
      tmp := Source as TZEChartSettings;
      FJapanCandle := tmp.JapanCandle;
    end;
  end;
  if (not b) then
    inherited Assign(Source);
end;

function TZEChartSettings.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartSettings;
begin
  Result := Assigned(Source);
  if (Result) then
  begin
    Result := Source is TZEChartSettings;
    if (Result) then
    begin
      tmp := Source as TZEChartSettings;
      Result := FJapanCandle = tmp.JapanCandle;
    end;
  end;
end;

/// /::::::::::::: TZEChartSeries :::::::::::::::::////

constructor TZEChartSeries.Create();
begin
  FChartType := ZEChartTypeBar;
  FSeriesName := '';
  FSeriesNameSheet := -1;
  FSeriesNameRow := -1;
  FSeriesNameCol := -1;
  FRanges := TZEChartRange.Create();
end;

destructor TZEChartSeries.Destroy();
begin
  FreeAndNil(FRanges);
  inherited;
end;

procedure TZEChartSeries.Assign(Source: TPersistent);
var
  tmp: TZEChartSeries;
  b: Boolean;
begin
  b := Assigned(Source);

  if (b) then
  begin
    b := Source is TZEChartSeries;
    if (b) then
    begin
      tmp := TZEChartSeries.Create();
      FChartType := tmp.ChartType;
      FSeriesName := tmp.SeriesName;
      FSeriesNameSheet := tmp.SeriesNameSheet;
      FSeriesNameRow := tmp.SeriesNameRow;
      FSeriesNameCol := tmp.SeriesNameCol;
      FRanges.Assign(tmp.Ranges);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartSeries.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartSeries;
begin
  Result := Source <> nil;
  if (Result) then
    Result := Source is TZEChartSeries;
  if (Result) then
  begin
    tmp := TZEChartSeries.Create();
    Result := (FChartType = tmp.ChartType) and (FSeriesNameSheet = tmp.SeriesNameSheet);
    if (Result) then
    begin
      if (((FSeriesNameRow < 0) or (FSeriesNameCol < 0)) and (tmp.SeriesNameRow < 0) or (tmp.SeriesNameCol < 0)) then
        Result := tmp.SeriesName = FSeriesName
      else
        Result := (tmp.SeriesNameRow = FSeriesNameRow) and (tmp.SeriesNameCol = FSeriesNameCol);
    end;
    if (Result) then
      Result := FRanges.IsEqual(tmp.Ranges);
  end;
end;

/// /::::::::::::: TZEChart :::::::::::::::::////

procedure TZEChart.CommonInit();
begin
  FTitle := TZEChartTitleItem.Create();
  FSubtitle := TZEChartTitleItem.Create();
  FLegend := TZEChartLegend.Create();
  FFooter := TZEChartTitleItem.Create();
  FAxisX := TZEChartAxis.Create();
  FAxisY := TZEChartAxis.Create();
  FAxisZ := TZEChartAxis.Create();
  FSecondaryAxisX := TZEChartAxis.Create();
  FSecondaryAxisY := TZEChartAxis.Create();
  FSecondaryAxisZ := TZEChartAxis.Create();
  FDefaultChartType := ZEChartTypeBar;
  FView3D := false;
  FViewDeep := false;
end;

constructor TZEChart.Create(AX, AY, AWidth, AHeight: Integer);
begin
  inherited;
  CommonInit();
  X := AX;
  Y := AY;
  Width := AWidth;
  Height := AHeight;
end;

constructor TZEChart.Create();
begin
  inherited;
  CommonInit();
end;

destructor TZEChart.Destroy();
begin
  FreeAndNil(FSubtitle);
  FreeAndNil(FTitle);
  FreeAndNil(FLegend);
  FreeAndNil(FFooter);
  FreeAndNil(FAxisX);
  FreeAndNil(FAxisY);
  FreeAndNil(FAxisZ);
  FreeAndNil(FSecondaryAxisX);
  FreeAndNil(FSecondaryAxisY);
  FreeAndNil(FSecondaryAxisZ);

  inherited;
end;

function TZEChart.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChart;
begin
  Result := inherited IsEqual(Source);
  if Result and (Source is TZEChart) then
  begin
    tmp := Source as TZEChart;
    Result := inherited IsEqual(Source);

    if (Result) then
      Result := FSubtitle.IsEqual(tmp.Subtitle) and FTitle.IsEqual(tmp.Title) and FLegend.IsEqual(tmp.Legend) and
        FAxisX.IsEqual(tmp.AxisX) and FAxisY.IsEqual(tmp.AxisY) and FAxisZ.IsEqual(tmp.AxisZ) and
        FSecondaryAxisX.IsEqual(tmp.SecondaryAxisX) and FSecondaryAxisY.IsEqual(tmp.SecondaryAxisY) and
        FSecondaryAxisZ.IsEqual(tmp.SecondaryAxisZ) and (FView3D = tmp.View3D) and (FViewDeep = tmp.ViewDeep);
  end;
end;

procedure TZEChart.SetAxisX(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FAxisX.Assign(Value);
end;

procedure TZEChart.SetAxisY(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FAxisY.Assign(Value);
end;

procedure TZEChart.SetAxisZ(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FAxisZ.Assign(Value);
end;

procedure TZEChart.SetSecondaryAxisX(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FSecondaryAxisX.Assign(Value);
end;

procedure TZEChart.SetSecondaryAxisY(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FSecondaryAxisY.Assign(Value);
end;

procedure TZEChart.SetSecondaryAxisZ(const Value: TZEChartAxis);
begin
  if (Assigned(Value)) then
    FSecondaryAxisZ.Assign(Value);
end;

procedure TZEChart.SetFooter(const Value: TZEChartTitleItem);
begin
  if (Assigned(Value)) then
    FFooter.Assign(Value);
end;

procedure TZEChart.SetLegend(const Value: TZEChartLegend);
begin
  if (Assigned(Value)) then
    FLegend.Assign(Value);
end;

procedure TZEChart.SetSubtitle(const Value: TZEChartTitleItem);
begin
  if (Assigned(Value)) then
    FSubtitle.Assign(Value);
end;

procedure TZEChart.SetTitle(const Value: TZEChartTitleItem);
begin
  if (Assigned(Value)) then
    FTitle.Assign(Value);
end;

procedure TZEChart.Assign(Source: TPersistent);
var
  tmp: TZEChart;
begin
  inherited;
  if Assigned(Source) and (Source is TZEChart) then
  begin
    tmp := Source as TZEChart;
    FSubtitle.Assign(tmp.Subtitle);
    FTitle.Assign(tmp.Title);
    FLegend.Assign(tmp.Legend);
    FAxisX.Assign(tmp.AxisX);
    FAxisY.Assign(tmp.AxisY);
    FAxisZ.Assign(tmp.AxisZ);
    FSecondaryAxisX.Assign(tmp.SecondaryAxisX);
    FSecondaryAxisY.Assign(tmp.SecondaryAxisY);
    FSecondaryAxisZ.Assign(tmp.SecondaryAxisZ);
    FView3D := tmp.View3D;
    FViewDeep := tmp.ViewDeep;
  end;
end;

/// /::::::::::::: TZEChartStore :::::::::::::::::////

constructor TZEChartStore.Create();
begin
  FCount := 0;
  SetLength(FItems, 0);
end;

destructor TZEChartStore.Destroy();
begin
  Clear();
  inherited;
end;

function TZEChartStore.Add(): TZEChart;
begin
  Result := TZEChart.Create();
  SetLength(FItems, FCount + 1);
  FItems[FCount] := Result;
  Inc(FCount);
end;

function TZEChartStore.Add(const ItemForClone: TZEChart): TZEChart;
begin
  Result := Add();
  Result.Assign(ItemForClone);
end;

procedure TZEChartStore.Assign(Source: TPersistent);
var
  tmp: TZEChartStore;
  b: Boolean;
  i: Integer;
begin
  b := Assigned(Source);
  if (b) then
  begin
    b := Source is TZEChartStore;
    if (b) then
    begin
      tmp := Source as TZEChartStore;

      if (FCount > tmp.Count) then
      begin
        for i := tmp.Count to FCount - 1 do
          FreeAndNil(FItems[i]);
        SetLength(FItems, tmp.Count);
      end
      else if (FCount < tmp.Count) then
      begin
        SetLength(FItems, tmp.Count);
        for i := FCount to tmp.Count - 1 do
          FItems[i] := TZEChart.Create();
      end;
      FCount := tmp.Count;

      for i := 0 to FCount - 1 do
        FItems[i].Assign(tmp[i]);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartStore.Delete(num: Integer): Boolean;
var
  i: Integer;
begin
  Result := (num >= 0) and (num < FCount);
  if (Result) then
  begin
    FreeAndNil(FItems[num]);
    for i := num to FCount - 2 do
      FItems[i] := FItems[i + 1];
    Dec(FCount);
    SetLength(FItems, FCount);
  end;
end;

procedure TZEChartStore.Clear();
var
  i: Integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FItems[i]);
  FCount := 0;
  SetLength(FItems, 0);
end;

function TZEChartStore.GetItem(num: Integer): TZEChart;
begin
  Result := nil;
  if ((num >= 0) and (num < FCount)) then
    Result := FItems[num];
end;

procedure TZEChartStore.SetItem(num: Integer; const Value: TZEChart);
begin
  if ((num >= 0) and (num < FCount)) then
    FItems[num].Assign(Value);
end;

function TZEChartStore.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEChartStore;
  i: Integer;
begin
  Result := Assigned(Source);
  if (Result) then
  begin
    Result := Source is TZEChartStore;
    if (Result) then
    begin
      tmp := Source as TZEChartStore;
      Result := FCount = tmp.Count;
      if (Result) then
        for i := 0 to FCount - 1 do
          if (not FItems[i].IsEqual(tmp[i])) then
          begin
            Result := false;
            break;
          end;
    end;
  end;
end;

/// /::::::::::::: TZEPicture :::::::::::::::::////

procedure TZEPicture.Assign(Source: TPersistent);
var
  tmp: TZEPicture;
begin
  inherited;
  if Assigned(Source) and (Source is TZEPicture) then
  begin
    tmp := Source as TZEPicture;

    FId := tmp.FId;
    FRelId := tmp.FRelId;
    FTitle := tmp.Title;
    FDescription := tmp.Description;
    FCellAnchor := tmp.FCellAnchor;
    FFileName := tmp.FFileName;
    FRow := tmp.Row;
    FCol := tmp.Col;
    FFromCol := tmp.FFromCol;
    FFromColOff := tmp.FFromColOff;
    FFromRow := tmp.FFromRow;
    FFromRowOff := tmp.FFromRowOff;
    FToCol := tmp.FToCol;
    FToColOff := tmp.FToColOff;
    FToRow := tmp.FToRow;
    FToRowOff := tmp.FToRowOff;
    FFrmOffX := tmp.FFrmOffX;
    FFrmOffY := tmp.FFrmOffY;
    FFrmExtCX := tmp.FFrmExtCX;
    FFrmExtCY := tmp.FFrmExtCY;
    FSheet := tmp.FSheet;
  end;
end;

procedure TZEPicture.CommonInit();
begin
  FId := 0;
  FRelId := 0;
  FTitle := '';
  FDescription := '';
  FRow := 0;
  FCol := 0;
  FCellAnchor := ZACell;
end;

constructor TZEPicture.Create(ASheet: TZSheet);
begin
  inherited Create;
  FSheet := ASheet;
  CommonInit();
end;

destructor TZEPicture.Destroy;
begin
  FSheet := nil;
  inherited;
end;

function TZEPicture.GetImage: TBytes;
begin
  Result := FSheet.WorkBook.MediaList[Self.RelId - 1].Content;
end;

procedure TZEPicture.SetImage(const Value: TBytes);
begin
  // FSheet.WorkBook.MediaList
end;

function TZEPicture.IsEqual(const Source: TPersistent): Boolean;
var
  tmp: TZEPicture;
begin
  Result := inherited IsEqual(Source);
  if (Result) and (Source is TZEPicture) then
  begin
    tmp := Source as TZEPicture;
    Result := (FId = tmp.Id) and (FTitle = tmp.Title) and (FDescription = tmp.Description)
          // and (FHidden = tmp.Hidden);
          // TODO: compare filename/streams
  end;
end;

{ TZEDrawing }

procedure TZEDrawing.Clear;
begin
  FItems.Clear();
end;

constructor TZEDrawing.Create(ASheet: TZSheet);
begin
  inherited Create;
  FItems := TObjectList.Create(true);
  FSheet := ASheet;
end;

procedure TZEDrawing.Delete(idx: Integer);
begin
  FItems.Delete(idx);
end;

destructor TZEDrawing.Destroy();
begin
  FItems.Clear();
  FItems.Free();
  inherited;
end;

procedure TZEDrawing.Assign(Source: TPersistent);
var
  tmp: TZEDrawing;
  b: Boolean;
  i: Integer;
begin
  b := Assigned(Source);
  if b then
  begin
    b := Source is TZEDrawing;
    if b then
    begin
      tmp := Source as TZEDrawing;
      Self.FId := tmp.FId;
      for i := 0 to tmp.FItems.Count - 1 do
        Add.Assign(TZEPicture(tmp.FItems[i]));
    end;
  end;

  if not b then
    inherited Assign(Source);
end;

function TZEDrawing.Add(ARow, ACol: Integer; APicture: TBytes): TZEPicture;
begin
  Result := TZEPicture.Create(FSheet);
  FItems.Add(Result);
  Result.Image := APicture;
  Result.RelId := FItems.Count;
  Result.Row := ARow;
  Result.Col := ACol;
  Result.CellAnchor := ZACell;
end;

function TZEDrawing.Add: TZEPicture;
begin
  Result := TZEPicture.Create(FSheet);
  Result.CellAnchor := ZACell;
  FItems.Add(Result);
end;

function TZEDrawing.GetCount: Integer;
begin
  Result := FItems.Count;
end;

function TZEDrawing.GetIsEmpty(): Boolean;
begin
  Result := FItems.Count = 0;
end;

function TZEDrawing.GetItem(idx: Integer): TZEPicture;
begin
  Result := TZEPicture(FItems[idx]);
end;

procedure TZEDrawing.SetItem(idx: Integer; const Value: TZEPicture);
begin
  TZEPicture(FItems[idx]).Assign(Value);
end;

{ TZRange }
procedure TZRange.Assign(Source: TZRange);
var
  Src: TZRange;
  r, c, Id: Integer;
  Style: TZStyle;
  rect: TZMergeArea;
begin
  inherited;
  if Source is TZRange then
  begin
    Src := TZRange(Source);

    // resize columns or rows if need
    if FSheet.RowCount <= Max(Src.FBottom, FBottom) then
      FSheet.RowCount := Max(Src.FBottom, FBottom);

    if FSheet.ColCount <= Max(Src.FRight, FRight) then
      FSheet.ColCount := Max(Src.FRight, FRight);

    Style := TZStyle.Create();
    try
      // copy cells and styles
      for c := 0 to Src.FRight - Src.FLeft do
      begin
        for r := 0 to Src.FBottom - Src.FTop do
        begin
          FSheet.Cell[FLeft + c, FTop + r].Assign(Src.FSheet.Cell[Src.FLeft + c, Src.FTop + r]);
          Id := FSheet.Cell[FLeft + c, FTop + r].CellStyle;
          if Id > -1 then
          begin
            // style must copy from source sheet
            Style.Assign(Src.FSheet.WorkBook.Styles[Id]);
            Id := FSheet.WorkBook.Styles.Add(Style, true);
            FSheet.Cell[FLeft + c, FTop + r].CellStyle := Id;
          end;
        end;
      end;
    finally
      Style.Free();
    end;

    // remove cross merges
    for Id := FSheet.MergeCells.Count - 1 downto 0 do
    begin
      if FSheet.MergeCells.IsCrossWithArea(Id, FLeft, FTop, FLeft + (Src.FRight - Src.FLeft),
        FTop + (Src.FBottom - Src.FTop)) then
        FSheet.MergeCells.DeleteItem(Id);
    end;

    // copy and reloc merges
    for Id := 0 to Src.FSheet.MergeCells.Count - 1 do
    begin
      if Src.FSheet.MergeCells.IsCrossWithArea(Id, Src.FLeft, Src.FTop, Src.FRight, Src.FBottom) then
      begin
        rect := Src.FSheet.MergeCells[Id];
        FSheet.MergeCells.AddRectXY(FLeft + (rect.left - Src.FLeft), FTop + (rect.Top - Src.FTop),
          FRight + (rect.right - Src.FRight), FBottom + (rect.Bottom - Src.FBottom));
      end;
    end;
  end;
end;

constructor TZRange.Create(ASheet: TZSheet; ALeft, ATop, ARight, ABottom: Integer);
begin
  FSheet := ASheet;
  FLeft := ALeft;
  FTop := ATop;
  FRight := ARight;
  FBottom := ABottom;
end;

destructor TZRange.Destroy;
begin
  FSheet := nil;
  FLeft := 0;
  FTop := 0;
  FRight := 0;
  FBottom := 0;
  inherited;
end;

function TZRange.HasStyle: Boolean;
begin
  Result := FSheet.Cell[FLeft, FTop].FCellStyle > -1;
end;

procedure TZRange.Merge();
var
  i: Integer;
begin
  for i := FSheet.MergeCells.Count - 1 downto 0 do
  begin
    if FSheet.MergeCells.IsCrossWithArea(i, FLeft, FTop, FRight, FBottom) then
      FSheet.MergeCells.DeleteItem(i);
  end;
  FSheet.MergeCells.AddRectXY(FLeft, FTop, FRight, FBottom);
end;

procedure TZRange.SetBorderAround(borderWidth: Byte; BorderColor: TColor = TColorRec.Black; BorderStyle: TZBorderType = TZBorderType.ZEContinuous);
var
  Row, Col: Integer;
  Style: TZStyle;
begin
  for Row := FTop to FBottom do
  begin
    for Col := FLeft to FRight do
    begin
      if (Col = FLeft) or (Row = FTop) or (Col = FRight) or (Row = FBottom) then
      begin
        Style := TZStyle.Create();
        try
          Style.Assign(FSheet.Cell[Col, Row].Style);
          if Col = FLeft then
          begin
            Style.Border[bpLeft].LineStyle := BorderStyle;
            Style.Border[bpLeft].Weight := borderWidth;
            Style.Border[bpLeft].Color := BorderColor;
          end;
          if Row = FTop then
          begin
            Style.Border[bpTop].LineStyle := BorderStyle;
            Style.Border[bpTop].Weight := borderWidth;
            Style.Border[bpTop].Color := BorderColor;
          end;
          if Col = FRight then
          begin
            Style.Border[bpRight].LineStyle := BorderStyle;
            Style.Border[bpRight].Weight := borderWidth;
            Style.Border[bpRight].Color := BorderColor;
          end;
          if Row = FBottom then
          begin
            Style.Border[bpBottom].LineStyle := BorderStyle;
            Style.Border[bpBottom].Weight := borderWidth;
            Style.Border[bpBottom].Color := BorderColor;
          end;
          FSheet.Cell[Col, Row].CellStyle := FSheet.FStore.Styles.Add(Style, true);
        finally
          Style.Free();
        end;
      end;
    end;
  end;
end;

procedure TZRange.ApplyStyleValue(proc: TProc<TZStyle>);
var
  Col, Row, Id: Integer;
  Style: TZStyle;
begin
  for Col := FLeft to FRight do
  begin
    for Row := FTop to FBottom do
    begin
      Style := TZStyle.Create();
      try
        Id := FSheet.Cell[Col, Row].CellStyle;
        if Id > -1 then
          Style.Assign(FSheet.FStore.Styles[Id])
        else if (FSheet.FStore.Styles.Count > 0) then
          Style.Assign(FSheet.FStore.Styles[0]);

        proc(Style);

        Id := FSheet.FStore.Styles.Add(Style, true);
        FSheet.Cell[Col, Row].CellStyle := Id;
      finally
        Style.Free();
      end;
    end;
  end;
end;

procedure TZRange.Clear();
var
  Col, Row: Integer;
begin
  for Col := FLeft to FRight do
  begin
    for Row := FTop to FBottom do
    begin
      FSheet.Cell[Col, Row].Data := '';
      FSheet.Cell[Col, Row].Formula := '';
      FSheet.Cell[Col, Row].Comment := '';
      FSheet.Cell[Col, Row].CommentAuthor := '';
    end;
  end;
end;

function TZRange.GetBgColor(): TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].BgColor;
end;

function TZRange.GetBorderColor(num: TZBordersPos): TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[num].Color;
end;

function TZRange.GetBordersStyle(): TZBorderType;
begin
  Result := TZBorderType.ZENone;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[bpLeft].LineStyle;
end;

function TZRange.GetBordersWidht(): Byte;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[bpLeft].Weight;
end;

function TZRange.GetBordersColor(): TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[bpLeft].Color;
end;

function TZRange.GetBorderStyle(num: TZBordersPos): TZBorderType;
begin
  Result := TZBorderType.ZENone;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[num].LineStyle;
end;

function TZRange.GetBorderWidht(num: TZBordersPos): Byte;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Border[num].Weight;
end;

function TZRange.GetFontColor(): TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Font.Color;
end;

function TZRange.GetFontSize(): Double;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Font.Size;
end;

function TZRange.GetFontStyle(): TFontStyles;
begin
  Result := [];
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Font.Style;
end;

function TZRange.GetHorizontalAlignment(): TZHorizontalAlignment;
begin
  Result := TZHorizontalAlignment.ZHAutomatic;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Alignment.Horizontal;
end;

function TZRange.GetNumberFormat(): string;
begin
  Result := '';
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].NumberFormat;
end;

function TZRange.GetRotate(): TZCellTextRotate;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Alignment.Rotate;
end;

function TZRange.GetVerticalAlignment(): TZVerticalAlignment;
begin
  Result := TZVerticalAlignment.ZVAutomatic;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Alignment.Vertical;
end;

function TZRange.GetVerticalText(): Boolean;
begin
  Result := false;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Alignment.VerticalText;
end;

function TZRange.GetWrapText(): Boolean;
begin
  Result := false;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FLeft, FTop].FCellStyle].Alignment.WrapText;
end;

procedure TZRange.SetBgColor(const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.BgColor := Value;
    end);
end;

procedure TZRange.SetBorderColor(num: TZBordersPos; const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].Color := Value;
    end);
end;

procedure TZRange.SetBorderStyle(num: TZBordersPos; const Value: TZBorderType);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].LineStyle := Value;
    end);
end;

procedure TZRange.SetBorderWidht(num: TZBordersPos; const Value: Byte);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Border[num].Weight := Value;
    end);
end;

procedure TZRange.SetBordersColor(const Value: TColor);
var
  Style: TZStyle;
  Col, Row: Integer;
begin
  Style := TZStyle.Create();
  try
    for Row := FTop to FBottom do
    begin
      for Col := FLeft to FRight do
      begin
        Style.Assign(FSheet[Col, Row].Style);
        Style.Border[bpLeft].Color := Value;
        Style.Border[bpTop].Color := Value;
        Style.Border[bpRight].Color := Value;
        Style.Border[bpBottom].Color := Value;
        FSheet[Col, Row].CellStyle := FSheet.FStore.Styles.Add(Style, true);
      end;
    end;
  finally
    Style.Free();
  end;
end;

procedure TZRange.SetBordersStyle(const Value: TZBorderType);
var
  Style: TZStyle;
  Col, Row: Integer;
begin
  Style := TZStyle.Create();
  try
    for Row := FTop to FBottom do
    begin
      for Col := FLeft to FRight do
      begin
        Style.Assign(FSheet[Col, Row].Style);
        Style.Border[bpLeft].LineStyle := Value;
        Style.Border[bpTop].LineStyle := Value;
        Style.Border[bpRight].LineStyle := Value;
        Style.Border[bpBottom].LineStyle := Value;
        FSheet[Col, Row].CellStyle := FSheet.FStore.Styles.Add(Style, true);
      end;
    end;
  finally
    Style.Free();
  end;
end;

procedure TZRange.SetBordersWidht(const Value: Byte);
var
  Style: TZStyle;
  Col, Row: Integer;
begin
  Style := TZStyle.Create();
  try
    for Row := FTop to FBottom do
    begin
      for Col := FLeft to FRight do
      begin
        Style.Assign(FSheet[Col, Row].Style);
        Style.Border[bpLeft].Weight := Value;
        Style.Border[bpTop].Weight := Value;
        Style.Border[bpRight].Weight := Value;
        Style.Border[bpBottom].Weight := Value;
        FSheet[Col, Row].CellStyle := FSheet.FStore.Styles.Add(Style, true);
      end;
    end;
  finally
    Style.Free();
  end;
end;

procedure TZRange.SetFontColor(const Value: TColor);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Color := Value;
    end);
end;

procedure TZRange.SetFontSize(const Value: Double);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Size := Value;
    end);
end;

procedure TZRange.SetFontStyle(const Value: TFontStyles);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Font.Style := Value;
    end);
end;

procedure TZRange.SetHorizontalAlignment(const Value: TZHorizontalAlignment);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Horizontal := Value;
    end);
end;

procedure TZRange.SetNumberFormat(const Value: string);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.NumberFormat := Value;
    end);
end;

procedure TZRange.SetRotate(const Value: TZCellTextRotate);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Rotate := Value;
    end);
end;

procedure TZRange.SetVerticalAlignment(const Value: TZVerticalAlignment);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.Vertical := Value;
    end);
end;

procedure TZRange.SetVerticalText(const Value: Boolean);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.VerticalText := Value;
    end);
end;

procedure TZRange.SetWrapText(const Value: Boolean);
begin
  ApplyStyleValue(
    procedure(Style: TZStyle)
    begin
      Style.Alignment.WrapText := Value;
    end);
end;

{ TRichText }

constructor TRichText.Create();
begin
  FList := TList<TRichString>.Create();
end;

destructor TRichText.Destroy();
var
  i: Integer;
begin
  for i := 0 to FList.Count - 1 do
    FList[i].Free();
  FList.Clear();
  FList.Free();
end;

function TRichText.Equals(Obj: TObject): Boolean;
var
  i: Integer;
begin
  if Assigned(Obj) and (Obj is TRichText) then
  begin
    if Self.FList.Count <> ((Obj as TRichText).List.Count) then
      exit(false);

    for i := 0 to Self.FList.Count - 1 do
      if not Self.FList[i].Equals((Obj as TRichText).FList[i]) then
        exit(false);
  end;
  Result := true;
end;

procedure TRichText.Assign(Source: TPersistent);
var
  i: Integer;
begin
  if Source is TRichText then
  begin
    FList.Clear();
    for i := 0 to TRichText(Source).FList.Count - 1 do
      FList.Add(TRichText(Source).FList[i]);
  end;
end;

function TRichText.GetHashCode(): Integer;
var
  i: Integer;
begin

  Result := 17;
  for i := 0 to FList.Count - 1 do
    Result := Result * 23 + FList[i].GetHashCode();
end;

{ TRichString }

destructor TRichString.Destroy();
begin
  if Assigned(FFont) then
    FFont.Free();
end;

function TRichString.Equals(Obj: TObject): Boolean;
begin
  if Assigned(Obj) and (Obj is TRichString) then
  begin
    if Self.FText <> (Obj as TRichString).FText then
      exit(false);
    if Assigned(Self.FFont) <> Assigned((Obj as TRichString).FFont) then
      exit(false);
    if not ZEIsFontsEquals(Self.FFont, (Obj as TRichString).FFont) then
      exit(false)
  end;
  Result := true;
end;

procedure TRichString.Assign(Source: TPersistent);
begin
  if Source is TRichString then
  begin
    FText := TRichString(Source).FText;
    if Assigned(TRichString(Source).FFont) then
      FFont.Assign(TRichString(Source).FFont);
  end;
end;

function TRichString.GetHashCode(): Integer;
begin
  Result := 17;
  Result := Result * 23 + FText.GetHashCode();
  if Assigned(FFont) then
    Result := Result * 23 + FFont.GetHashCode();
end;

{ TRichTextComparer }

function TRichTextComparer.Equals(const left, right: TRichText): Boolean;
begin
  Result := Assigned(left) and Assigned(right) and left.Equals(right);
end;

function TRichTextComparer.GetHashCode(const Value: TRichText): Integer;
begin
  Result := Value.GetHashCode();
end;

{ TZExcelColor }

class operator TZExcelColor.Equal(const ALeft, ARight: TZExcelColor): Boolean;
begin
  Result := (ALeft.RGB = ARight.RGB) and (ALeft.Indexed = ARight.Indexed) and (ALeft.Theme = ARight.Theme) and
    SameValue(ALeft.Tint, ARight.Tint);
end;

procedure TZExcelColor.Load(const ARgb, AIndexed, ATheme, ATint: string);
begin
  Indexed := StrToIntDef(ARgb, 0);
  Theme := StrToIntDef(ATheme, 0);
  Tint := StrToFloatDef(ATint, 0, TFormatSettings.Invariant);
  if ARgb.IsEmpty then
    RGB := 0
  else
    RGB := ARGBToColor(ARgb);
end;

class operator TZExcelColor.NotEqual(const ALeft, ARight: TZExcelColor): Boolean;
begin
  Result := not((ALeft.RGB = ARight.RGB) and (ALeft.Indexed = ARight.Indexed) and (ALeft.Theme = ARight.Theme) and
    SameValue(ALeft.Tint, ARight.Tint));
end;

{ TZMegeCell }

constructor TZMergeArea.Create(ALeft, ATop, ARight, ABottom: Integer);
begin
  Top := ATop;
  left := ALeft;
  right := ARight;
  Bottom := ABottom;
end;

initialization

invariantFormatSertting := TFormatSettings.Create();
invariantFormatSertting.DecimalSeparator := '.';

end.
