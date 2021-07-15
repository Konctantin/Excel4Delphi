unit Excel4Delphi.Xml;

interface

uses
  System.SysUtils, System.Classes;

const
  BOMUTF8 = #239#187#191; // EF BB BF
  BOMUTF16BE = #254#255; // FE FF
  BOMUTF16LE = #255#254; // FF FE
  BOMUTF32BE = #0#0#254#255; // 00 00 FE FF
  BOMUTF32LE = #255#254#0#0; // FF FE 00 00

type
  /// <summary>
  /// Xml tag type.
  /// </summary>
  TXmlTagType = (
    // Unknown.
    xttUnknown = 0,
    /// <?...?>
    xttDeclare = 1,
    /// <![CDATA[..]]>
    xttCData = 2,
    /// <!--..-->
    xttComment = 3,
    /// <...>
    xttStart = 4,
    /// <.../>
    xttClosed = 5,
    // </...>
    xttEnd = 6);

  /// <summary>
  /// Text converter from local encoding to needed encoding
  /// </summary>
  TAnsiToCPConverter = function(const AnsiText: ansistring): ansistring;
  TReadCPCharObj = procedure(var RetChar: ansistring; var _eof: Boolean) of object;
  TReadCPChar = procedure(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
  /// <summary>
  /// Text converter from reading encoding to local encoding
  /// </summary>
  TCPToAnsiConverter = TAnsiToCPConverter;
  /// <summary>
  /// Attribute-value
  /// </summary>
  TZAttrArray = array [0 .. 1] of ansistring;

  TagsProp = record
    Name: ansistring;
    CloseTagNewLine: Boolean;
  end;

  /// <summary>
  /// Builder attributes for tags.
  /// </summary>
  TZAttributes = class(TPersistent)
  private
    FCount: integer;
    FMaxCount: integer; // реальное кол-во элементов
    FItems: array of TZAttrArray;
    function GetAttrS(const Att: ansistring): ansistring;
    procedure SetAttrS(const Att: ansistring; const Value: ansistring);
    function GetAttrI(num: integer): ansistring;
    procedure SetAttrI(num: integer; const Value: ansistring);
    function GetAttr(num: integer): TZAttrArray;
    procedure SetAttr(num: integer; const Value: TZAttrArray);
  protected
    procedure ResizeItemsArray(NewSize: integer);
  public
    constructor Create();
    destructor Destroy(); override;
    /// <summary>
    /// Adds attribute
    /// </summary>
    /// <param name="AttrName">
    /// Attribute Name
    /// </param>
    /// <param name="TestMatch">
    /// Checks for attribute AttrName entry, if attribute was entered then changes it's value, else new attribute is added.
    /// </param>
    procedure Add(const AttrName: ansistring; const Value: ansistring; TestMatch: Boolean = true); overload;
    /// <summary>
    /// Adds attributes
    /// </summary>
    /// <param name="Attr">
    /// Attributes
    /// </param>
    /// <param name="TestMatch">
    /// Checks for attribute AttrName entry, if attribute was entered then changes it's value, else new attribute is added.
    /// </param>
    procedure Add(const Attr: TZAttrArray; TestMatch: Boolean = true); overload;
    /// <summary>
    /// Adds attributes
    /// </summary>
    /// <param name="Att">
    /// Attributes
    /// </param>
    /// <param name="TestMatch">
    /// Checks for attribute AttrName entry, if attribute was entered then changes it's value, else new attribute is added.
    /// </param>
    procedure Add(Att: array of TZAttrArray; TestMatch: Boolean = true); overload;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Delete all atributes.
    /// </summary>
    procedure Clear();
    /// <summary>
    /// Delete attribute with number Index. Attributes with larger number are shifting left.
    /// </summary>
    procedure DeleteItem(Index: integer);
    /// <summary>
    /// Adds attribute AttrName with value Value on position Index.
    /// </summary>
    procedure Insert(Index: integer; const AttrName: ansistring; const Value: ansistring;
      TestMatch: Boolean = true); overload;
    /// <summary>
    /// Adds attribute AttrName with value Value on position Index.
    /// </summary>
    procedure Insert(Index: integer; const Attr: TZAttrArray; TestMatch: Boolean = true); overload;
    /// <summary>
    /// Return string with attributes.
    /// </summary>
    /// <param name="quote">
    /// quotation mark for attribute.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity
    /// </param>
    /// <param name="addempty">
    /// true then do not add attributes with empty value.
    /// </param>
    function ToString(quote: ansichar; CheckEntity: Boolean; addempty: Boolean): ansistring; reintroduce;
      overload; virtual;
    /// <summary>
    /// Return string with attributes.
    /// </summary>
    /// <param name="quote">
    /// quotation mark for attribute.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity
    /// </param>
    function ToString(quote: ansichar; CheckEntity: Boolean): ansistring; reintroduce; overload; virtual;
    /// <summary>
    /// Return string with attributes.
    /// </summary>
    /// <param name="quote">
    /// quotation mark for attribute.
    /// </param>
    function ToString(quote: ansichar): ansistring; reintroduce; overload; virtual;
    /// <summary>
    /// Return string with attributes.
    /// </summary>
    /// <param name="CheckEntity">
    /// If true then corrects entity
    /// </param>
    function ToString(CheckEntity: Boolean): ansistring; reintroduce; overload; virtual;
    /// <summary>
    /// Return string with attributes.
    /// </summary>
    function ToString(): ansistring; reintroduce; overload; virtual;
    function IsContainsAttribute(const AttrName: string; CaseSensitivity: Boolean = true): Boolean;
    /// <summary>
    /// Number of attributes. (RO)
    /// </summary>
    property Count: integer read FCount;
    /// <summary>
    /// Access to the attribute-value by number num.
    /// </summary>
    property Items[num: integer]: TZAttrArray read GetAttr write SetAttr;
    /// <summary>
    /// Access to the attribute value by name Att.
    /// </summary>
    property ItemsByName[const Att: ansistring]: ansistring read GetAttrS write SetAttrS; default;
    /// <summary>
    /// Access to the attribute value by number num.
    /// </summary>
    property ItemsByNum[num: integer]: ansistring read GetAttrI write SetAttrI;
  end;

  /// <summary>
  /// Class-writer can write xml to string, stream or file in Windows-1251, UTF-8, UTF-16 (BE and LE) encodings.
  /// Use TextConverter property for write in other encodings.
  /// </summary>
  TZsspXMLWriter = class
  private
    FAttributeQuote: ansichar; // какие кавычки используем для атрибутов
    FAttributes: TZAttributes;
    FTags: array of TagsProp; // "Стек тэгов"
    FTagCount: integer; // кол-во тэгов в "стеке тэгов"
    FMaxTagCount: integer; // Реальный размер FTags
    FBuffer: ansistring;
    FMaxBufferLength: integer;
    FInProcess: Boolean;
    FStream: TStream;
    FTextConverter: TAnsiToCPConverter;
    FNewLine: Boolean;
    FUnixNLSeparator: Boolean;
    FNLSeparator: ansistring;
    FTab: ansistring;
    FTabSymbol: ansistring;
    FTabLength: integer;
    function GetTag(num: integer): ansistring;
    function GetTabSymbol(): ansichar;
    procedure SetAttributeQuote(Value: ansichar);
    procedure SetMaxBufferLength(Value: integer);
    procedure SetNewLine(Value: Boolean);
    procedure SetTabLength(Value: integer);
    procedure SetTabSymbol(Value: ansichar);
    procedure SetTextConverter(Value: TAnsiToCPConverter);
    procedure SetUnixNLSeparator(Value: Boolean);
    procedure SetAttributes(Value: TZAttributes);
  protected
    procedure AddText(const text: ansistring; UseConverter: Boolean = true);
    procedure AddNode(const TagName: ansistring; CloseTagNewLine: Boolean);
    function GetTab(num: integer = 0): ansistring;
    procedure _AddTag(const _begin: ansistring; text: ansistring; const _end: ansistring; StartNewLine: Boolean;
      _tab: integer = 0);
    procedure ResizeTagArray(NewSize: integer);
  public
    constructor Create(Stream: TStream);
    destructor Destroy(); override;
    /// <summary>
    /// End writing to string/file/stream. Sets InProcess False.
    /// </summary>
    procedure EndSaveTo();
    /// <summary>
    /// Empties the Buffer. All characters from Buffer have been written to a file.
    /// </summary>
    procedure FlushBuffer();
    /// <summary>
    /// Write CDATA. CDATA - text. CorrectCDATA - if true then replaces ']]&gt;' on ']]&amp;gt;'.
    /// StartNewLine - if true then starts from new line (ignored when NewLine = false).
    /// </summary>
    /// <param name="CDATA">
    /// text
    /// </param>
    /// <param name="CorrectCDATA">
    /// if true then replaces ']]&gt;' on ']]&amp;gt;'
    /// </param>
    /// <param name="StartNewLine">
    /// if true then starts from new line (ignored when NewLine = false)
    /// </param>
    procedure WriteCDATA(CDATA: ansistring; CorrectCDATA: Boolean; StartNewLine: Boolean = true); overload;
    // <![CDATA[ bla-bla-bla <><><>...]]>
    /// <summary>
    /// Write CDATA. CDATA - text. CorrectCDATA - if true then replaces ']]&gt;' on ']]&amp;gt;'.
    /// StartNewLine - if true then starts from new line (ignored when NewLine = false).
    /// </summary>
    /// <param name="CDATA">
    /// text
    /// </param>
    procedure WriteCDATA(const CDATA: ansistring); overload;
    /// <summary>
    /// Write comment.
    /// </summary>
    /// <param name="Comment">
    /// text of comment.
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    procedure WriteComment(const Comment: ansistring; StartNewLine: Boolean = true);
    procedure WriteEmptyTag(const TagName: ansistring; SAttributes: TZAttributes; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload; // <tag a="a"... />
    procedure WriteEmptyTag(const TagName: ansistring; AttrArray: array of TZAttrArray; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write empty tag.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteEmptyTag(const TagName: ansistring; StartNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write empty tag.
    /// </summary>
    procedure WriteEmptyTag(const TagName: ansistring; SAttributes: TZAttributes); overload;
    /// <summary>
    /// Write empty tag.
    /// </summary>
    procedure WriteEmptyTag(const TagName: ansistring; AttrArray: array of TZAttrArray); overload;
    /// <summary>
    /// Write empty tag.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    procedure WriteEmptyTag(const TagName: ansistring); overload;
    /// <summary>
    /// Write end of node.
    /// </summary>
    procedure WriteEndTagNode(); overload;
    /// <summary>
    /// Write end of node.
    /// </summary>
    procedure WriteEndTagNode(isForce: Boolean; CloseTagNewLine: Boolean); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    /// <param name="SAttributes">
    /// attributes
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false)
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring; SAttributes: TZAttributes; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    /// <param name="SAttributes">
    /// attributes
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false)
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring; AttrArray: array of TZAttrArray;
      StartNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false)
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    /// <param name="SAttributes">
    /// attributes
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring; SAttributes: TZAttributes); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    /// <param name="SAttributes">
    /// attributes
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring; AttrArray: array of TZAttrArray); overload;
    /// <summary>
    /// Write instruction
    /// </summary>
    /// <param name="InstructionName">
    /// name of instruction.
    /// </param>
    procedure WriteInstruction(const InstructionName: ansistring); overload;
    /// <summary>
    /// <para>
    /// Write not processed text. Text
    /// </para>
    /// <note type="warning">
    /// can break XML!
    /// </note>
    /// </summary>
    /// <param name="Text">
    /// text
    /// </param>
    /// <param name="UseConverter">
    /// used TextConverter
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    procedure WriteRaw(text: ansistring; UseConverter: Boolean; StartNewLine: Boolean = true);
    procedure WriteTag(const TagName: ansistring; const text: ansistring; SAttributes: TZAttributes;
      StartNewLine: Boolean; CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteTag(const TagName: ansistring; const text: ansistring; AttrArray: array of TZAttrArray;
      StartNewLine: Boolean; CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write tag (&lt;tag ...&gt;text&lt;/tag&gt;).
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="Text">
    /// text
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    /// <param name="CloseTagNewLine">
    /// if true then closed tag starts from new line.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteTag(const TagName: ansistring; const text: ansistring; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write tag (&lt;tag ...&gt;text&lt;/tag&gt;).
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="Text">
    /// text
    /// </param>
    /// <param name="SAttributes">
    /// Attributes
    /// </param>
    procedure WriteTag(const TagName: ansistring; const text: ansistring; SAttributes: TZAttributes); overload;
    /// <summary>
    /// Write tag (&lt;tag ...&gt;text&lt;/tag&gt;).
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="Text">
    /// text
    /// </param>
    /// <param name="SAttributes">
    /// Attributes
    /// </param>
    procedure WriteTag(const TagName: ansistring; const text: ansistring; AttrArray: array of TZAttrArray); overload;
    /// <summary>
    /// Write tag (&lt;tag ...&gt;text&lt;/tag&gt;).
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="Text">
    /// text
    /// </param>
    procedure WriteTag(const TagName: ansistring; const text: ansistring); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="SAttributes">
    /// Attributes
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    /// <param name="CloseTagNewLine">
    /// if true then closed tag starts from new line.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity. Attributes get from Attributes.
    /// </param>
    procedure WriteTagNode(const TagName: ansistring; SAttributes: TZAttributes; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="AttrArray">
    /// Attributes
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    /// <param name="CloseTagNewLine">
    /// if true then closed tag starts from new line.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity. Attributes get from Attributes.
    /// </param>
    procedure WriteTagNode(const TagName: ansistring; AttrArray: array of TZAttrArray; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="StartNewLine">
    /// if true then start from new line (ignored when NewLine = false).
    /// </param>
    /// <param name="CloseTagNewLine">
    /// if true then closed tag starts from new line.
    /// </param>
    /// <param name="CheckEntity">
    /// if true then corrects entity.
    /// </param>
    procedure WriteTagNode(const TagName: ansistring; StartNewLine: Boolean; CloseTagNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="SAttributes">
    /// Attributes
    /// </param>
    procedure WriteTagNode(const TagName: ansistring; SAttributes: TZAttributes); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    /// <param name="AttrArray">
    /// Attributes
    /// </param>
    procedure WriteTagNode(const TagName: ansistring; AttrArray: array of TZAttrArray); overload;
    /// <summary>
    /// Write node.
    /// </summary>
    /// <param name="TagName">
    /// name of tag
    /// </param>
    procedure WriteTagNode(const TagName: ansistring); overload;
    /// <summary>
    /// Tag's attributes.
    /// </summary>
    property Attributes: TZAttributes read FAttributes write SetAttributes;
    /// <summary>
    /// Quotation mark for attribute values ( ' (prime) or " (double prime)). <br />" (double prime) by default.
    /// </summary>
    property AttributeQuote: ansichar read FAttributeQuote write SetAttributeQuote;
    /// <summary>
    /// Buffer. (read only)
    /// </summary>
    property Buffer: ansistring read FBuffer;
    /// <summary>
    /// Return True if writing in process. (read only)
    /// </summary>
    property InProcess: Boolean read FInProcess;
    /// <summary>
    /// Buffer length (&gt;0). If InProcess = True then property read only. <br />4096 bytes by default.
    /// </summary>
    property MaxBufferLength: integer read FMaxBufferLength write SetMaxBufferLength;
    /// <summary>
    /// If True then tag starts with new line. If InProcess = True then property read only. <br />True by default
    /// </summary>
    property NewLine: Boolean read FNewLine write SetNewLine;
    /// <summary>
    /// Number of tab-symbols before tag. If InProcess = True then property read only. <br />0 by default.
    /// </summary>
    property TabLength: integer read FTabLength write SetTabLength;
    /// <summary>
    /// Tab symbol (#32 (space) or #9 (tab)). If InProcess = True then property read only. <br />#32 by default.
    /// </summary>
    property TabSymbol: ansichar read GetTabSymbol write SetTabSymbol;
    /// <summary>
    /// Number of open tags before current tag. (read only)
    /// </summary>
    property TagCount: integer read FTagCount;
    /// <summary>
    /// Return open tag number num. (read only)
    /// </summary>
    property Tags[num: integer]: ansistring read GetTag;
    /// <summary>
    /// Set text converter. If InProcess = True then property read only.
    /// </summary>
    Property TextConverter: TAnsiToCPConverter read FTextConverter write SetTextConverter;
    /// <summary>
    /// If True then newline character is #10 (LF) else newline "character" is #13#10 (CR+LF).
    /// </summary>
    property UnixNLSeparator: Boolean read FUnixNLSeparator write SetUnixNLSeparator;
  end;

  /// <summary>
  /// Class-reader. Read XML from string, stream or file in CP866, Windows-1251, UTF-8 and UTF-16 (BE and LE) encodings.
  /// </summary>
  TZsspXMLReader = class
  private
    FAttributes: TZAttributes; // Атрибуты
    FStream: TStream;
    FTags: array of ansistring; // "Стек тэгов"
    FTagCount: integer; // кол-во тэгов в "стеке тэгов"
    FMaxTagCount: integer; // Реальный размер FTags
    FBuffer: ansistring; // буфер
    FMaxBufferLength: integer; // Размер буфера
    FInProcess: Boolean; // в процессе чтения
    FSourceType: byte; // 1 - файл, 2 - поток, 3 - строка
    FPFirst: integer; // Начало буфера
    FPLast: integer; // Конец буфера
    FTextBeforeTag: ansistring; // Текст до тэга
    FErrorCode: integer; // код ошибки (0 - всё нормально)
             // and    1  =    1 - значение параметра без кавычек (<TAG param=value>)
             // and    2  =    2 - <TAG ... param = > или <TAG ... param = = ...> (значение параметра без кавычек)
             // and    4  =    4 - <TAG ... param = value = ...> (значение параметра без кавычек)
             // and    8  =    8 - <TAG ... param =/...> (значение параметра без кавычек)
             // and   16  =   16 - <TAG ... param = value"...> (нет открывающей кавычки)
             // and   32  =   32 - <TAG ... param = <*va<*lue<*... > <TAG ... param = <*"value"... > (* - 0 или более)
             // and   64  =   64 - <TAG ... param = value/ ... >
             // and  128  =  128 - <!Unknown... > Странный комментарий
             // and  256  =  256 - <=...>
             // and  512  =  512 - <tag!*...> - символ '!' в неправильном месте
             // and 1024  = 1024 - <tag ... param = ?...>
             // and 2048  = 2048 - < tag ... > пробельный символ после <
             // and 4096  = 4096 - < tag ... param param ... > ('=' отсутствует)
             // and 8192  = 8192 - </|?tag ... /> ('/' или '?')
             // and 16384 = 16384 - <tag ... ?>
             // and 32768 = 32768 - '<' в неподходящем месте
             // and 65536 = 65536 - "'" или '"' в неподходящем месте
             // and 131072 = 131072 - неожиданный конец тега
             // and 262144 = 262144 - перед закрывающим тэгом нету открывающего
             // and 524288 = 524288 - EOF и кол-во тэгов > 0

    FRawTextTag: ansistring; // Текст тэга
// FRawTextTagNotDecoded: ansistring; //Текст тэга не декодированный
    FTagName: ansistring; // Имя тэга (инструкции/комментария)
    FValue: ansistring; // Текст CDATA или комментария
    FTagType: TXmlTagType;
    FCharReader: TReadCPChar; // Читает "символ"
    FCharConverter: TCPToAnsiConverter; // конвертер
    FStreamEnd: Boolean;
    FIgnoreCase: Boolean;
    FQuotesEqual: Boolean;
    FAttributesMatch: Boolean;
    function GetTag(num: integer): ansistring;
    procedure SetMaxBufferLength(Value: integer);
    procedure SetAttributes(Value: TZAttributes);
    procedure AddTag(const Value: ansistring);
    procedure DeleteClosedTag();
    procedure DeleteTag();
    procedure SetIgnoreCase(Value: Boolean);
    procedure SetQuotesEqual(Value: Boolean);
    procedure SetAttributesMatch(Value: Boolean);
  protected
    procedure Clear();
    procedure ClearAll();
    procedure RecognizeEncoding(var txt: ansistring); // Попытка распознания кодировки
    procedure ReadBuffer();
    procedure GetOneChar(var OneChar: ansistring; var err: Boolean);
    procedure ResizeTagArray(NewSize: integer);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    /// <summary>
    /// Begin to read XML from Stream.
    /// </summary>
    /// <returns>
    /// 0 if no errors.
    /// </returns>
    function BeginReadStream(Stream: TStream): integer;
    /// <summary>
    /// Read tag.
    /// </summary>
    function ReadTag(): Boolean;
    /// <summary>
    /// End reading.
    /// </summary>
    procedure EndRead();
    /// <returns>
    /// True when current file (stream or string) position is the end.
    /// </returns>
    function Eof(): Boolean; virtual;
    /// <summary>
    /// Attributes of tag.
    /// </summary>
    property Attributes: TZAttributes read FAttributes write SetAttributes;
    property AttributesMatch: Boolean read FAttributesMatch write SetAttributesMatch;
    /// <summary>
    /// Return True when processed XML. (RO)
    /// </summary>
    property InProcess: Boolean read FInProcess;
    /// <summary>
    /// Raw text. (RO)
    /// </summary>
    property RawTextTag: ansistring read FRawTextTag;
    /// <summary>
    /// Return error code. 0 - OK, 1 - value without quotes, etc. (RO)
    /// </summary>
    property ErrorCode: integer read FErrorCode;
    /// <summary>
    /// If True then ignore case sensitive. If InProcess = True then property read only. <br />False by default.
    /// </summary>
    property IgnoreCase: Boolean read FIgnoreCase write SetIgnoreCase;
    /// <summary>
    /// Return name of readed tag/instruction/comment. (RO)
    /// </summary>
    property TagName: ansistring read FTagName;
    /// <summary>
    /// Return CDATA or comment text. (RO)
    /// </summary>
    property TagValue: ansistring read FValue;
    /// <summary>
    /// Return type of tag.
    /// </summary>
    property TagType: TXmlTagType read FTagType;
    /// <summary>
    /// Number of opened tags before current tag. (RO)
    /// </summary>
    property TagCount: integer read FTagCount;
    /// <summary>
    /// Return opened tag number num. (RO)
    /// </summary>
    property Tags[num: integer]: ansistring read GetTag;
    property TextBeforeTag: ansistring read FTextBeforeTag;
    /// <summary>
    /// Buffer size (&gt;=512). If InProcess = True then property read only. <br />4096 bytes by default.
    /// </summary>
    property MaxBufferLength: integer read FMaxBufferLength write SetMaxBufferLength;
    property QuotesEqual: Boolean read FQuotesEqual write SetQuotesEqual;
    // признак того, что двойные и одинарные кавычки
                                                                          // равны. Если установлено в false, то в <tagname attr1="asda' attr2='adsas">
                                                                          // будет один атрибут "attr1" cо значением "asda' attr2='adsas",
                                                                          // в противном случае распознается 2 атрибута: "attr1"="asda" и "attr2"="adsas".
                                                                          // По умолчанию false.
  end;

  TZAttrArrayH = array [0 .. 1] of string;

// Класс-обёртка-костыль для атрибутов над TZAttributes
  TZAttributesH = class(TPersistent)
  private
    FAttributes: TZAttributes;
    function GetAttrCount(): integer;
    function GetAttrS(Att: string): string;
    procedure SetAttrS(Att: string; const Value: string);
    function GetAttrI(num: integer): string;
    procedure SetAttrI(num: integer; const Value: string);
    function GetAttr(num: integer): TZAttrArrayH;
    procedure SetAttr(num: integer; const Value: TZAttrArrayH);
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Add(const AttrName: string; const Value: string; TestMatch: Boolean = true); overload;
    procedure Add(const Attr: TZAttrArrayH; TestMatch: Boolean = true); overload;
    procedure Add(Att: array of TZAttrArrayH; TestMatch: Boolean = true); overload;
    procedure Assign(Source: TPersistent); override;
    procedure Clear();
    procedure DeleteItem(Index: integer);
    procedure Insert(Index: integer; const AttrName: string; const Value: string; TestMatch: Boolean = true); overload;
    procedure Insert(Index: integer; const Attr: TZAttrArrayH; TestMatch: Boolean = true); overload;
    function ToString(quote: char; CheckEntity: Boolean; addempty: Boolean): string; reintroduce; overload; virtual;
    function ToString(quote: char; CheckEntity: Boolean): string; reintroduce; overload; virtual;
    function ToString(quote: char): string; reintroduce; overload; virtual;
    function ToString(CheckEntity: Boolean): string; reintroduce; overload; virtual;
    function ToString(): string; overload; override;
    function IsContainsAttribute(const AttrName: string; CaseSensitivity: Boolean = true): Boolean;
    property Count: integer read GetAttrCount;
    property Items[num: integer]: TZAttrArrayH read GetAttr write SetAttr;
    property ItemsByName[Att: string]: string read GetAttrS write SetAttrS; default;
    property ItemsByNum[num: integer]: string read GetAttrI write SetAttrI;
  end;

  // пишет XML
  TZsspXMLWriterH = class
  private
    FAttributes: TZAttributesH;
    FXMLWriter: TZsspXMLWriter;
    function GetXMLBuffer(): string;
    function GetAttributeQuote(): char;
    function GetInProcess(): Boolean;
    function GetMaxBufferLength(): integer;
    function GetNewLine(): Boolean;
    function GetTabLength(): integer;
    function GetTagCount(): integer;
    function GetTextConverter(): TCPToAnsiConverter;
    function GetUnixNLSeparator(): Boolean;
    function GetTag(num: integer): string;
    function GetTabSymbol(): char;
    procedure SetAttributeQuote(Value: char);
    procedure SetMaxBufferLength(Value: integer);
    procedure SetNewLine(Value: Boolean);
    procedure SetTabLength(Value: integer);
    procedure SetTabSymbol(Value: char);
    procedure SetTextConverter(Value: TAnsiToCPConverter);
    procedure SetUnixNLSeparator(Value: Boolean);
    procedure SetAttributes(Value: TZAttributesH);
  protected
  public
    constructor Create(Stream: TStream);
    destructor Destroy(); override;
    procedure EndSaveTo();
    procedure FlushBuffer();
    procedure WriteCDATA(CDATA: string; CorrectCDATA: Boolean; StartNewLine: Boolean = true); overload;
    // <![CDATA[ bla-bla-bla <><><>...]]>
    procedure WriteCDATA(CDATA: string); overload;
    procedure WriteComment(Comment: string; StartNewLine: Boolean = true); // <!-- bla-bla-bla -->
    procedure WriteEmptyTag(TagName: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload; // <tag a="a"... />
    procedure WriteEmptyTag(TagName: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteEmptyTag(TagName: string; StartNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteEmptyTag(TagName: string; SAttributes: TZAttributesH); overload;
    procedure WriteEmptyTag(TagName: string; AttrArray: array of TZAttrArrayH); overload;
    procedure WriteEmptyTag(TagName: string); overload;
    procedure WriteEndTagNode(); overload;
    procedure WriteEndTagNode(isForce: Boolean; CloseTagNewLine: Boolean); overload;
    procedure WriteInstruction(InstructionName: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteInstruction(InstructionName: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteInstruction(InstructionName: string; StartNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteInstruction(InstructionName: string; SAttributes: TZAttributesH); overload;
    procedure WriteInstruction(InstructionName: string; AttrArray: array of TZAttrArrayH); overload;
    procedure WriteInstruction(InstructionName: string); overload;
    procedure WriteRaw(text: string; UseConverter: Boolean; StartNewLine: Boolean = true); overload;
    procedure WriteRaw(text: ansistring; UseConverter: Boolean; StartNewLine: Boolean = true); overload;
    procedure WriteTag(TagName: string; text: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteTag(TagName: string; text: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteTag(TagName: string; text: string; StartNewLine: Boolean; CloseTagNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteTag(TagName: string; text: string; SAttributes: TZAttributesH); overload;
    procedure WriteTag(TagName: string; text: string; AttrArray: array of TZAttrArrayH); overload;
    procedure WriteTag(TagName: string; text: string); overload;
    procedure WriteTagNode(TagName: string; SAttributes: TZAttributesH; StartNewLine: Boolean; CloseTagNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteTagNode(TagName: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
      CloseTagNewLine: Boolean; CheckEntity: Boolean = true); overload;
    procedure WriteTagNode(TagName: string; StartNewLine: Boolean; CloseTagNewLine: Boolean;
      CheckEntity: Boolean = true); overload;
    procedure WriteTagNode(TagName: string; SAttributes: TZAttributesH); overload;
    procedure WriteTagNode(TagName: string; AttrArray: array of TZAttrArrayH); overload;
    procedure WriteTagNode(TagName: string); overload;
    procedure WriteHeader(CodePageName: string; BOM: ansistring);
    property Attributes: TZAttributesH read FAttributes write SetAttributes;
    property AttributeQuote: char read GetAttributeQuote write SetAttributeQuote;
    property Buffer: string read GetXMLBuffer;
    property InProcess: Boolean read GetInProcess;
    property MaxBufferLength: integer read GetMaxBufferLength write SetMaxBufferLength;
    property NewLine: Boolean read GetNewLine write SetNewLine;
    property TabLength: integer read GetTabLength write SetTabLength;
    property TabSymbol: char read GetTabSymbol write SetTabSymbol;
    property TagCount: integer read GetTagCount;
    property Tags[num: integer]: string read GetTag;
    Property TextConverter: TAnsiToCPConverter read GetTextConverter write SetTextConverter;
    property UnixNLSeparator: Boolean read GetUnixNLSeparator write SetUnixNLSeparator;
  end;

  // читает XML
  TZsspXMLReaderH = class
  private
    FAttributes: TZAttributesH; // Атрибуты
    FXMLReader: TZsspXMLReader;
    function GetAttributes(): TZAttributesH;
    function GetInProcess(): Boolean;
    function GetRawTextTag(): string;
    function GetErrorCode(): integer;
    function GetIgnoreCase(): Boolean;
    function GetValue(): string;
    function GetTagType(): TXmlTagType;
    function GetIsTagStart(): Boolean;
    function GetIsTagClosed(): Boolean;
    function GetIsTagEnd(): Boolean;
    function GetTagCount(): integer;
    function GetTextBeforeTag(): string;
    function GetTagName(): string;
    function GetTag(num: integer): string;
    procedure SetMaxBufferLength(Value: integer);
    function GetMaxBufferLength(): integer;
    procedure SetAttributes(Value: TZAttributesH);
    procedure SetIgnoreCase(Value: Boolean);
    procedure SetQuotesEqual(Value: Boolean);
    function GetQuotesEqual(): Boolean;
    procedure SetAttributesMatch(Value: Boolean);
    function GetAttributesMatch(): Boolean;
    function GetIsTagStartOrClosed: Boolean;
    function GetIsTagOfData: Boolean;
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function BeginReadStream(Stream: TStream): integer;
    function ReadTag(): Boolean;
    procedure EndRead();
    function Eof(): Boolean; virtual;
    property Attributes: TZAttributesH read GetAttributes write SetAttributes;
    property AttributesMatch: Boolean read GetAttributesMatch write SetAttributesMatch;
    property InProcess: Boolean read GetInProcess;
    property RawTextTag: string read GetRawTextTag;
    property ErrorCode: integer read GetErrorCode;
    property IgnoreCase: Boolean read GetIgnoreCase write SetIgnoreCase;
    property TagName: string read GetTagName;
    property TagValue: string read GetValue;
    property TagType: TXmlTagType read GetTagType;
    property IsTagStart: Boolean read GetIsTagStart;
    property IsTagClosed: Boolean read GetIsTagClosed;
    property IsTagStartOrClosed: Boolean read GetIsTagStartOrClosed;
    property IsTagOfData: Boolean read GetIsTagOfData;
    property IsTagEnd: Boolean read GetIsTagEnd;
    property TagCount: integer read GetTagCount;
    property Tags[num: integer]: string read GetTag;
    property TextBeforeTag: string read GetTextBeforeTag;
    property MaxBufferLength: integer read GetMaxBufferLength write SetMaxBufferLength;
    property QuotesEqual: Boolean read GetQuotesEqual write SetQuotesEqual;
    function IsTagEndByName(TagName: string): Boolean;
    function IsTagStartByName(TagName: string): Boolean;
    function IsTagClosedByName(TagName: string): Boolean;
    function IsTagStartOrClosedByName(TagName: string): Boolean;
    function ReadToEndTagByName(TagName: string): Boolean;
  end;

// конец для Delphi >=2009

// Читатали
procedure ReadCharUTF8(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
procedure ReadCharUTF16LE(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
procedure ReadCharUTF16BE(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
procedure ReadCharUTF32(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
procedure ReadCharOneByte(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);

/// ///////// Конвертеры
function conv_UTF8ToLocal(const text: ansistring): ansistring;
function conv_UTF16LEToLocal(const text: ansistring): ansistring;
function conv_UTF16BEToLocal(const text: ansistring): ansistring;
function conv_UTF32LEToLocal(const text: ansistring): ansistring;
function conv_UTF32BEToLocal(const text: ansistring): ansistring;
function conv_WIN1251ToLocal(const text: ansistring): ansistring;
function conv_CP866ToLocal(const text: ansistring): ansistring;

// заменяет в строке спецсимволы
function CheckStrEntity(const st: ansistring; checkamp: Boolean = true): ansistring; overload;
function CheckStrEntity(const st: string; checkamp: Boolean = true): string; overload;
function ClenuapXmlTagValue(const str: string): string;

// проверяем на корректность сущность (не факт, что валидную), в случае
// чего заменяем '&' на '&amp;'
procedure Correct_Entity(const _St: ansistring; num: integer; var _result: ansistring); overload;
procedure Correct_Entity(const _St: string; num: integer; var _result: string); overload;

// Добавляет аттрибут
function ToAttribute(const AttrName: ansistring; const Value: ansistring): TZAttrArray; overload;
function ToAttribute(const AttrName: string; const Value: string): TZAttrArrayH; overload;

// Распознаёт кодировку XML и HTML текста
function RecognizeEncodingXML(startpos: integer; var txt: ansistring; out cpfromtext: integer; out cpname: ansistring;
  out ftype: integer): Boolean; overload;

// Распознаёт BOM (Byte Order Mark) текста
function RecognizeBOM(var txt: ansistring): integer;

// Распознаёт кодировку XML и HTML текста вместе с BOM
function RecognizeEncodingXML(var txt: ansistring; out BOM: integer; out cpfromtext: integer; out cpname: ansistring;
  out ftype: integer): Boolean; overload;

implementation

/// / читатели

function DUAnsiPos(const Substr: ansistring; const S: ansistring): integer;
var
  i, j: integer;
  kol_str: integer;
  kol_sub: integer;
  b: Boolean;
begin
  kol_str := Length(S);
  kol_sub := Length(Substr);
  result := 0;
  if (kol_sub = 0) then
    exit;
  for i := 1 to kol_str - kol_sub + 1 do
    if (S[i] = Substr[1]) then
    begin
      b := true;
      for j := 2 to kol_sub do
        if (S[i + j - 1] <> Substr[j]) then
        begin
          b := false;
          break;
        end;
      if (b) then
      begin
        result := i;
        break;
      end;
    end;
end;

procedure ReadCharUTF8(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
var
  S: ansistring;
  t, i: integer;
  kol: integer;

begin
  _eof := false;
  text := '';
  S := '';
  if Assigned(ReadCPChar) then
  begin
    ReadCPChar(S, _eof);
    text := text + S;
    if _eof then
      exit;
    if Length(S) > 0 then
    begin
      t := ord(S[1]);
      if t > 127 then
      begin
        {
        //вроде как возможны 5-ти байтные, но в стандарте нету
        if t and 248 = 248 then
          kol := 4 else } { tut }
        kol := 0;
        if t and 240 = 240 then
          kol := 3
        else if t and 224 = 224 then
          kol := 2
        else if t and 192 = 192 then
          kol := 1;
        for i := 1 to kol do
        begin
          // все последующие симовлы должны быть вида 10xxxxxx {tut}
          ReadCPChar(S, _eof);
          text := text + S;
          if _eof then
            exit;
        end;
      end; // if
    end;
  end
  else
    _eof := true;
end;

procedure ReadCharUTF16LE(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
var
  i, num: integer;
  S: ansistring;

begin
  _eof := false;
  text := '';
  S := '';
  if Assigned(ReadCPChar) then
  begin
    ReadCPChar(S, _eof);
    text := text + S;
    if _eof then
      exit;
    num := ord(S[1]) shl 8;
    ReadCPChar(S, _eof);
    text := text + S;
    if _eof then
      exit;
    num := num + ord(S[1]);
    if num >= $D800 then
      for i := 1 to 2 do
      begin
        ReadCPChar(S, _eof);
        text := text + S;
        if _eof then
          exit;
      end;
  end
  else
    _eof := true;
end;

procedure ReadCharUTF16BE(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
var
  i, num: integer;
  S: ansistring;

begin
  _eof := false;
  text := '';
  S := '';
  if Assigned(ReadCPChar) then
  begin
    ReadCPChar(S, _eof);
    text := text + S;
    if _eof then
      exit;
    num := ord(S[1]);
    ReadCPChar(S, _eof);
    text := text + S;
    if _eof then
      exit;
    // {$HINTS OFF}
    num := num + (ord(S[1]) shl 8);
    // {$HINTS ON}
    if num >= $D800 then
      for i := 1 to 2 do
      begin
        ReadCPChar(S, _eof);
        text := text + S;
        if _eof then
          exit;
      end;
  end
  else
    _eof := true;
end;

procedure ReadCharUTF32(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
var
  S: ansistring;
  i: integer;

begin
  _eof := false;
  text := '';
  S := '';
  if Assigned(ReadCPChar) then
    for i := 1 to 4 do
    begin
      ReadCPChar(S, _eof);
      text := text + S;
      if _eof then
        exit;
    end
  else
    _eof := true;
end;

// для однобайтных кодировок
procedure ReadCharOneByte(const ReadCPChar: TReadCPCharObj; var text: ansistring; var _eof: Boolean);
begin
  _eof := false;
  text := '';
  if Assigned(ReadCPChar) then
    ReadCPChar(text, _eof)
  else
    _eof := true;
end;

/// ///////// Конвертеры

function CP866ToWin1251(const cp866: ansistring): ansistring;
var
  i, n: integer;
  ch: byte;
begin
  n := Length(cp866);
  setlength(result, n);
  for i := 1 to n do
  begin
    ch := ord(cp866[i]);
    if (ch >= 128) and (ch <= 175) then
      inc(ch, 64)
    else if ch = 240 then
      ch := 168
    else if ch = 241 then
      ch := 184
    else if ch = 252 then
      ch := 185
    else if (ch >= 224) and (ch <= 239) then
      inc(ch, 16);
    result[i] := ansichar(ch);
  end;
end;

function conv_UTF8ToLocal(const text: ansistring): ansistring;
begin
  result := text;
end;

function conv_UTF16LEToLocal(const text: ansistring): ansistring;
var
  ws: WideString;
  w: word;
  i, l: integer;
begin
  ws := '';
  i := 2;
  { tut }// что делать, если длина не кратна 2?
  l := Length(text);
  while i <= l do
  begin
    w := ord(text[i]) shl 8;
    w := w + ord(text[i - 1]);
    ws := ws + WideChar(w);
    inc(i, 2);
  end;
  result := UTF8Encode(PWideChar(ws));
end;

function conv_UTF16BEToLocal(const text: ansistring): ansistring;
var
  ws: WideString;
  w: word;
  i, l: integer;
begin
  ws := '';
  i := 2;
  { tut }// что делать, если длина не кратна 2?
  l := Length(text);
  while i <= l do
  begin
    w := ord(text[i - 1]) shl 8;
    w := w + ord(text[i]);
    ws := ws + WideChar(w);
    inc(i, 2);
  end;
  result := UTF8Encode(PWideChar(ws));
end;

function conv_UTF32LEToLocal(const text: ansistring): ansistring;
begin
  result := text;
end;

function conv_UTF32BEToLocal(const text: ansistring): ansistring;
begin
  result := text;
end;

function conv_WIN1251ToLocal(const text: ansistring): ansistring;
begin
  result := text;
end;

function conv_CP866ToLocal(const text: ansistring): ansistring;
begin
  result := CP866ToWin1251(text);
end;
/// ////////////////////////

// Добавляет аттрибут
function ToAttribute(const AttrName: ansistring; const Value: ansistring): TZAttrArray;
begin
  result[0] := AttrName;
  result[1] := Value;
end;

function ToAttribute(const AttrName: string; const Value: string): TZAttrArrayH;
begin
  result[0] := AttrName;
  result[1] := Value;
end;

// проверяем на корректность сущность (не факт, что валидную), в случае
// чего заменяем '&' на '&amp;'
procedure Correct_Entity(const _St: ansistring; num: integer; var _result: ansistring);
var
  b: Boolean;
  i, l: integer;

begin
  b := true;
  l := Length(_St);
  for i := num + 1 to Length(_St) do
    case _St[i] of
      ' ', #13, #10, #9, '<', '>', '''', '"', '&':
        begin
          b := false;
          break;
        end;
      ';':
        break;
    end;
  if num >= l then
    b := false;
  if b then
    _result := _result + _St[num]
  else
    _result := _result + '&amp;';
end;

procedure Correct_Entity(const _St: string; num: integer; var _result: string);
var
  b: Boolean;
  i, l: integer;

begin
  b := true;
  l := Length(_St);
  for i := num + 1 to Length(_St) do
    case _St[i] of
      ' ', #13, #10, #9, '<', '>', '''', '"', '&':
        begin
          b := false;
          break;
        end;
      ';':
        break;
    end;
  if num >= l then
    b := false;
  if b then
    _result := _result + _St[num]
  else
    _result := _result + '&amp;';
end;

// заменяет в строке спецсимволы
// INPUT
// St: ansistring - исходная строка
// checkamp: boolean - true - заменять & сразу на &amp;
// false - если сущность - оставлять
// RETURN
// ansistring - Обработанная строка
function CheckStrEntity(const st: ansistring; checkamp: Boolean = true): ansistring;
var
  i, kol: integer;

begin
  result := '';
  kol := Length(st);
  for i := 1 to kol do
  begin
    case st[i] of
      '<':
        result := result + '&lt;';
      '&':
        begin
          if checkamp then
            result := result + '&amp;'
          else
            Correct_Entity(st, i, result);
        end;
      '>':
        result := result + '&gt;';
      '''':
        result := result + '&apos;';
      '"':
        result := result + '&quot;';
    else
      result := result + st[i];
    end;
  end;
end;

function ClenuapXmlTagValue(const str: string): string;
begin
  result := str.Replace('&lt;', '<').Replace('&amp;', '&').Replace('&gt;', '>').Replace('&apos;', '''')
    .Replace('&quot;', '"');
end;

// заменяет в строке спецсимволы
// INPUT
// St: string - исходная строка
// checkamp: boolean - true - заменять & сразу на &amp;
// false - если сущность - оставлять
// RETURN
// ansistring - Обработанная строка
function CheckStrEntity(const st: string; checkamp: Boolean = true): string;
var
  i, kol: integer;

begin
  result := '';
  kol := Length(st);
  for i := 1 to kol do
  begin
    case st[i] of
      '<':
        result := result + '&lt;';
      '&':
        begin
          if checkamp then
            result := result + '&amp;'
          else
            Correct_Entity(st, i, result);
        end;
      '>':
        result := result + '&gt;';
      '''':
        result := result + '&apos;';
      '"':
        result := result + '&quot;';
    else
      result := result + st[i];
    end;
  end;
end; // CheckStrEntity

// Возвращает номер кодировки по его названию (UPCASE не забываем!)
// INPUT
// txt: ansistring - имя кодировки в ВЕРХНЕМ РЕГИСТРЕ
// RETURN: integer - номер кодировки
function CPFromName(txt: ansistring): integer;
begin
  result := 0;
  // если нету BOMа, то UTF-16/32 считаем как UTF-16/32BE
  if (txt = 'UTF-8') or (txt = 'UTF8') then
    result := 1
  else if txt = 'UTF-16' then
    result := 2
  else
 // if txt = 'UTF-16LE' then result := 3 else
    if txt = 'UTF-32' then
      result := 4
    else
 // if txt = 'UTF-32LE' then result := 5 else
      if txt = 'WINDOWS-1251' then
        result := 6
      else if txt = 'CP866' then
        result := 7;
end;

// Распознаёт кодировку XML и HTML текста
// (во всяком случае попытается)
// INPUT
// startpos: integer   - стартовая позиция
// txt: ansistring         - текст для распознания
// OUTPUT
// cpfromtext: integer - номер расспознаной кодировки из текста
// 0 - неопределена
// 1 - UTF-8
// 2 - UTF-16BE
// 3 - UTF-16LE
// 4 - UTF-32BE
// 5 - UTF-32LE
// 6 - Windows-1251
// 7 - CP866
// cpname: ansistring      - название кодировки из текста
// ftype: integer      - тип файла:
// 0 - непонятно
// 1 - xml
// 2 - html
// RETURN: boolean         - true - кодировка скорее всего точно опознана
// false - есть сомнение насчёт правильности распознания
function RecognizeEncodingXML(startpos: integer; var txt: ansistring; out cpfromtext: integer; out cpname: ansistring;
  out ftype: integer): Boolean; overload;
var
  i, ll: integer;
  kol16BE: integer;
  kol16LE: integer;
  kol32BE: integer;
  kol32LE: integer;
  S: ansistring;
  _kol: integer;
  _l, _f: integer;

  function checkCPFromText(_name: ansistring; _b: Boolean): Boolean;
  var
    n: integer;

  begin
    result := _b;
    n := CPFromName(_name);
    // если кодировка раньше не была определена
    if cpfromtext = 0 then
      cpfromtext := n
    else
    // если кодировка была определена ранее, а название кодировки не то -
    // оставляем номер ранее определённой кодировкой и ставим признак
    // недоверяи
      if cpfromtext <> n then
        result := false;
  end;

begin
  cpfromtext := 0;
  ftype := 0;
  ll := Length(txt);
  if (ll <= 0) then
  begin
    result := false;
    exit;
  end;
  result := true;
  kol16BE := 0;
  kol16LE := 0;
  kol32BE := 0;
  kol32LE := 0;
  cpname := '';
  if startpos < 0 then
    startpos := 1;
  if startpos > ll then
    startpos := ll;

  // сначала проверяем на UTF-16 и UTF-32
  // если каждый второй символ #0 - значит UTF-16
  // если 2 подряд #0, а 4-ый <>#0, значит UTF-32
  i := startpos;
  while (i <= ll - 1) do
  begin
    if (txt[i] = #0) and (txt[i + 1] <> #0) then
      inc(kol16BE);
    if (txt[i] <> #0) and (txt[i + 1] = #0) then
      inc(kol16LE);
    inc(i, 2);
  end;
  i := startpos;
  while (i <= ll - 3) do
  begin
    if (txt[i] = #0) and (txt[i + 1] = #0) and (txt[i + 3] <> #0) then
      inc(kol32BE);
    if (txt[i + 2] = #0) and (txt[i + 3] = #0) and (txt[i] <> #0) then
      inc(kol32LE);
    inc(i, 4);
  end;
  // перебор возможных значений
  // точно UTF16BE
  if (kol16BE > 0) and (kol16LE = 0) and (kol32BE = 0) and (kol32LE = 0) then
    cpfromtext := 2
  else
  // точно UTF16LE
    if (kol16BE = 0) and (kol16LE > 0) and (kol32BE = 0) and (kol32LE = 0) then
      cpfromtext := 3
    else
  // точно UTF32BE
      if (kol16BE = 0) and (kol16LE = 0) and (kol32BE > 0) and (kol32LE = 0) then
        cpfromtext := 4
      else
  // точно UTF32LE
        if (kol16BE = 0) and (kol16LE = 0) and (kol32BE = 0) and (kol32LE > 0) then
          cpfromtext := 5
        else
  // точно не UTF-16/32
          if (kol16BE = 0) and (kol16LE = 0) and (kol32BE = 0) and (kol32LE = 0) then
          begin
            cpfromtext := 0;
          end
          else
          begin
            result := false; // точно не понятно что за кодировка
    // выбираем по макс. кол-ву
            i := kol16BE;
            cpfromtext := 2;
            if i < kol16LE then
            begin
              i := kol16LE;
              cpfromtext := 3;
            end;
            if i < kol32BE * 2 then
            begin
              i := kol32BE * 2;
              cpfromtext := 4;
            end;
            if i < kol32LE * 2 then
              cpfromtext := 5;
          end;
  // удаляем из текста все символы #0 - получаем (скорее всего) ASCII текст,
  // текст приводим к верхнему регистру,
  // в нем находим строку <?xml version="1.0" encoding="some_encoding"?> (xml)
  // или <META bla-bla-bla CHARSET=some_encoding"> (html)
  S := '';
  for i := startpos to ll do
    if txt[i] <> #0 then
      S := S + UpCase(txt[i]);
  // XML?
  // todo: надо как-то определить, не php-ли это
  _l := DUAnsiPos(UTF8Encode('?>'), S);
  if _l <> 0 then
  begin
    ftype := 1;
    _f := DUAnsiPos(UTF8Encode('ENCODING'), S);
    if (_f < _l) and (_f > 0) then
    begin
      _kol := 0;
      for i := _f + 8 to _l do
        case S[i] of
          '"', '''':
            if _kol = 0 then
              inc(_kol)
            else
              break;
        else
          if _kol = 1 then
            cpname := cpname + S[i];
        end;
      result := checkCPFromText(cpname, result);
    end
    else
    begin
      // если это XML и нету ENCODING - значит UTF-8 (вроде бы так)
      if cpfromtext = 0 then
        cpfromtext := 1
      else if cpfromtext <> 1 then
        result := false;
    end;
  end;
  // HTML?
  if ftype = 0 then
  begin
    _f := DUAnsiPos(UTF8Encode('CHARSET'), S);
    if _f > 0 then
    begin
      _l := DUAnsiPos(UTF8Encode('>'), S); // tut
      while (_l < _f) and (_l > 0) do
      begin
        S[_l] := '"';
        _l := DUAnsiPos(UTF8Encode('>'), S);
      end;
      // наверное это HTML
      if (_l > _f) then
      begin
        ftype := 2;
        _kol := 0;
        for i := _f + 7 to _l - 1 do
        begin
          case S[i] of
            '=':
              inc(_kol);
            '>', '<', '"', '''':
              break;
          else
            if (_kol > 0) and not(S[i] in [' ', #13, #10, #9]) then
              cpname := cpname + S[i];
          end;
        end;
        result := checkCPFromText(cpname, result);
      end;
    end
    else
      result := false;
  end;
end;

// Распознаёт BOM (Byte Order Mark) текста
// Предполагаем, что в самом начале текста стоит BOM,
// тогда это будет уникод:
// UTF-8:    EF BB BF
// UTF-16BE: FE FF     (00 SS)
// UTF-16LE: FF FE     (SS 00) - windows
// UTF-32BE: 00 00 FE FF
// UTF-32LE: FF FE 00 00
// INPUT
// var txt: ansistring - текст для распознания BOM-а
// RETURN: integer       - номер распознанной кодировки по BOMу
function RecognizeBOM(var txt: ansistring): integer;
var
  ltxt: integer;

begin
  result := 0;
  // BOM:
  // 0 - неопределён
  // 1 - UTF-8
  // 2 - UTF-16BE
  // 3 - UTF-16LE
  // 4 - UTF-32BE
  // 5 - UTF-32LE
  ltxt := Length(txt);
  if ltxt >= 2 then
  begin
    // если есть BOM:
    case txt[1] of
      #239: // UTF-8?
        begin
          if ltxt >= 3 then
            if (txt[2] = #187) and (txt[3] = #191) then
              result := 1;
        end;
      #254: // UTF-16BE?
        begin
          if txt[2] = #255 then
            result := 2;
        end;
      #255: // UTF-16LE/UTF-32LE?
        begin
          if txt[2] = #254 then
          begin
            result := 3;
            if ltxt >= 4 then
            begin
                 // не учитываем, что текст в UTF-16LE может начинаться с 00 00
              if (txt[3] = #0) and (txt[4] = #0) then
                result := 5;
            end;
          end;
        end;
      #0: // UTF-32BE?
        begin
          if ltxt >= 4 then
            if (txt[2] = #0) and (txt[3] = #254) and (txt[4] = #255) then
              result := 4;
        end;
    end;
  end;
end;

// Распознаёт кодировку XML и HTML текста вместе с BOM
// (во всяком случае попытается)
// INPUT
// txt: ansistring         - текст для распознания
// OUTPUT
// BOM: integer        - номер распознаной кодировки из BOM
// cpfromtext: integer - номер расспознаной кодировки из текста
// 0 - неопределена
// 1 - UTF-8
// 2 - UTF-16BE
// 3 - UTF-16LE
// 4 - UTF-32BE
// 5 - UTF-32LE
// 6 - Windows-1251
// 7 - CP866
// cpname: ansistring      - название кодировки из текста
// ftype: integer      - тип файла:
// 0 - непонятно
// 1 - xml
// 2 - html
// RETURN: boolean         - true - кодировка скорее всего точно опознана
// false - есть сомнение насчёт правильности распознания
function RecognizeEncodingXML(var txt: ansistring; out BOM: integer; out cpfromtext: integer; out cpname: ansistring;
  out ftype: integer): Boolean; overload;
var
  t: integer;

begin
  BOM := RecognizeBOM(txt);
  t := 1;
  case BOM of
    1:
      t := 4; // utf-8
    2, 3:
      t := 3; // utf-16
    4, 5:
      t := 5; // utf-32
  end;
  result := RecognizeEncodingXML(t, txt, cpfromtext, cpname, ftype);
end;

/// /::::::::::::: TZAttributes :::::::::::::::::////

constructor TZAttributes.Create();
begin
  inherited;
  FMaxCount := 15;
  setlength(FItems, FMaxCount);
end;

destructor TZAttributes.Destroy();
begin
  setlength(FItems, 0);
  FItems := nil;
  inherited;
end;

procedure TZAttributes.ResizeItemsArray(NewSize: integer);
var
  delta: integer;

begin
  if (NewSize >= FMaxCount) then
  begin
    delta := NewSize;
    if (NewSize < 50) then
      delta := delta * 4
    else if (NewSize < 100) then
      delta := delta * 2
    else
      delta := delta + 20;
    setlength(FItems, delta);
  end
  else if (NewSize > 50) then
  begin
    if (FMaxCount - NewSize > 200) then
    begin
      delta := NewSize + 100;
      setlength(FItems, delta);
    end;
  end;
end; // ResizeItemsArray

procedure TZAttributes.Clear();
begin
  FCount := 0;
  ResizeItemsArray(0);
end;

function TZAttributes.GetAttrS(const Att: ansistring): ansistring;
var
  i: integer;

begin
  result := '';
  for i := 0 to FCount - 1 do
    if FItems[i][0] = Att then
    begin
      result := FItems[i][1];
      break;
    end;
end;

procedure TZAttributes.SetAttrS(const Att: ansistring; const Value: ansistring);
begin
  Add(Att, Value, true);
end;

function TZAttributes.GetAttrI(num: integer): ansistring;
begin
  result := '';
  if (num >= 0) and (num < FCount) then
    result := FItems[num][1];
end;

procedure TZAttributes.SetAttrI(num: integer; const Value: ansistring);
begin
  if (num >= 0) and (num < FCount) then
    FItems[num][1] := Value;
end;

function TZAttributes.GetAttr(num: integer): TZAttrArray;
begin
  result[0] := '';
  result[1] := '';
  if (num >= 0) and (num < FCount) then
    result := FItems[num];
end;

procedure TZAttributes.SetAttr(num: integer; const Value: TZAttrArray);
begin
  if (num >= 0) and (num < FCount) then
    FItems[num] := Value;
end;

procedure TZAttributes.DeleteItem(Index: integer);
var
  i: integer;

begin
  if (Index >= 0) and (Index < Count) then
  begin
    for i := Index to Count - 2 do
      FItems[i] := FItems[i + 1];
    dec(FCount);
    ResizeItemsArray(FCount);
  end;
end;

procedure TZAttributes.Insert(Index: integer; const AttrName: ansistring; const Value: ansistring;
  TestMatch: Boolean = true);
var
  i: integer;

begin
  if TestMatch then
  begin
    for i := 0 to FCount - 1 do
      if FItems[i][0] = AttrName then
      begin
        FItems[i][1] := Value;
        exit;
      end;
  end;
  if (Index >= 0) and (Index < Count) then
  begin
    inc(FCount);
    ResizeItemsArray(FCount);
    for i := FCount - 2 downto Index do
      FItems[i + 1] := FItems[i];
    FItems[Index][0] := AttrName;
    FItems[Index][1] := Value;
  end
  else if Count = 0 then
    Add(AttrName, Value, TestMatch);
end;

procedure TZAttributes.Insert(Index: integer; const Attr: TZAttrArray; TestMatch: Boolean = true);
begin
  Insert(Index, Attr[0], Attr[1], TestMatch);
end;

procedure TZAttributes.Add(const AttrName: ansistring; const Value: ansistring; TestMatch: Boolean = true);
var
  i: integer;

begin
  if Length(AttrName) = 0 then
    exit;
  if TestMatch then
  begin
    for i := 0 to FCount - 1 do
      if FItems[i][0] = AttrName then
      begin
        FItems[i][1] := Value;
        exit;
      end;
  end;
  ResizeItemsArray(FCount + 1);
  FItems[FCount][0] := AttrName;
  FItems[FCount][1] := Value;
  inc(FCount);
end;

procedure TZAttributes.Add(const Attr: TZAttrArray; TestMatch: Boolean = true);
begin
  Add(Attr[0], Attr[1], TestMatch);
end;

procedure TZAttributes.Add(Att: array of TZAttrArray; TestMatch: Boolean = true);
var
  i: integer;
begin
  for i := low(Att) to High(Att) do
    Add(Att[i], TestMatch);
end;

function TZAttributes.ToString(quote: ansichar; CheckEntity: Boolean; addempty: Boolean): ansistring;
var
  i: integer;
begin
  if (quote <> '"') and (quote <> '''') then
    quote := '"';
  result := '';
  if CheckEntity then
  begin
    // название атрибута, надеюсь, будет без спецсимволов ^__^
    for i := 0 to Count - 1 do
      if (Length(FItems[i][1]) > 0) or (addempty) then
        result := result + ' ' + FItems[i][0] + '=' + quote + CheckStrEntity(FItems[i][1]) + quote;
  end
  else
  begin
    for i := 0 to Count - 1 do
      if (Length(FItems[i][1]) > 0) or (addempty) then
        result := result + ' ' + FItems[i][0] + '=' + quote + FItems[i][1] + quote;
  end;
end;

function TZAttributes.ToString(quote: ansichar; CheckEntity: Boolean): ansistring;
begin
  result := ToString(quote, CheckEntity, true);
end;

function TZAttributes.ToString(quote: ansichar): ansistring;
begin
  result := ToString(quote, true);
end;

function TZAttributes.ToString(CheckEntity: Boolean): ansistring;
begin
  result := ToString('"', CheckEntity);
end;

function TZAttributes.ToString(): ansistring;
begin
  result := ToString('"', true);
end;

function TZAttributes.IsContainsAttribute(const AttrName: string; CaseSensitivity: Boolean = true): Boolean;
var
  i: integer;
  S: string;
begin
  result := false;
  if (not CaseSensitivity) then
    S := UpperCase(AttrName)
  else
    S := AttrName;

  for i := 0 to FCount - 1 do
  begin
    if (CaseSensitivity) then
    begin
      if (string(FItems[i][1]) = S) then
        result := true;
    end
    else if (UpperCase(string(FItems[i][1])) = S) then
      result := true;

    if (result) then
      break;
  end;
end;

procedure TZAttributes.Assign(Source: TPersistent);
var
  t: TZAttributes;
  i: integer;
begin
  if Source is TZAttributes then
  begin
    t := Source as TZAttributes;
    FCount := t.Count;
    ResizeItemsArray(FCount + 1);
    for i := 0 to t.Count - 1 do
      FItems[i] := t.Items[i];
  end
  else
    inherited Assign(Source);
end;

/// /::::::::::::: TZsspXMLWriter :::::::::::::::::////

constructor TZsspXMLWriter.Create(Stream: TStream);
begin
  if not Assigned(Stream) then
    raise EArgumentException.Create('Stream is null');

  inherited Create();
  FStream := Stream;
  FAttributeQuote := '"';
  FBuffer := '';
  FMaxBufferLength := 4096;
  FInProcess := true;
  FNewLine := true;
  FUnixNLSeparator := false;
  FTabLength := 0;
  FTab := '';
  FTabSymbol := ' ';
  FNLSeparator := #13#10;
  FAttributes := TZAttributes.Create();
  FMaxTagCount := 20;
  setlength(FTags, FMaxTagCount);
end;

destructor TZsspXMLWriter.Destroy();
begin
  // если забыли завершить запись:
  if InProcess then
    EndSaveTo();
  setlength(FTags, 0);
  FTags := nil;
  FreeAndNil(FAttributes);
  inherited;
end;

// Изменяет размер массива с тэгами
// INPUT
// NewSize: integer - новый размер
procedure TZsspXMLWriter.ResizeTagArray(NewSize: integer);
var
  delta: integer;
begin
  delta := 0;
  if (NewSize >= FMaxTagCount) then
  begin
    delta := NewSize;
    if (NewSize < 50) then
      delta := delta * 4
    else if (NewSize < 100) then
      delta := delta * 2
    else
      delta := delta + 20;
  end
  else if (NewSize > 50) then
  begin
    if (FMaxTagCount - NewSize > 200) then
      delta := NewSize + 100;
  end;
  if (delta > 0) then
    setlength(FTags, delta);
end; // ResizeTagArray

procedure TZsspXMLWriter.SetAttributes(Value: TZAttributes);
begin
  if Value <> nil then
    FAttributes.Assign(Value);
end;

function TZsspXMLWriter.GetTag(num: integer): ansistring;
begin
  if (num >= 0) and (num < TagCount) then
    result := FTags[num].Name
  else
    result := '';
end;

procedure TZsspXMLWriter.SetAttributeQuote(Value: ansichar);
begin
  if (Value = '''') or (Value = '"') then
    FAttributeQuote := Value;
end;

// Устанавливает максимальную длинну буфера
procedure TZsspXMLWriter.SetMaxBufferLength(Value: integer);
begin
  if Value > 0 then
    if not InProcess then
      FMaxBufferLength := Value;
end;

// Устанавливает длину табуляции
procedure TZsspXMLWriter.SetTabLength(Value: integer);
var
  i: integer;
begin
  if Value > 0 then
    if not InProcess then
    begin
      FTabLength := Value;
      FTab := '';
      for i := 1 to FTabLength do
        FTab := FTab + FTabSymbol;
    end;
end;

procedure TZsspXMLWriter.SetNewLine(Value: Boolean);
begin
  if not InProcess then
    FNewLine := Value;
end;

procedure TZsspXMLWriter.SetUnixNLSeparator(Value: Boolean);
begin
  if not InProcess then
  begin
    FUnixNLSeparator := Value;
    if Value then
      FNLSeparator := #10
    else
      FNLSeparator := #13#10;
  end;
end;

procedure TZsspXMLWriter._AddTag(const _begin: ansistring; text: ansistring; const _end: ansistring;
  StartNewLine: Boolean; _tab: integer = 0);
begin
  if not FInProcess then
    exit;
  text := _begin + text + _end;
  if StartNewLine and NewLine then
    text := FNLSeparator + GetTab(_tab) + text;
  AddText(text);
end;

procedure TZsspXMLWriter.WriteCDATA(CDATA: ansistring; CorrectCDATA: Boolean; StartNewLine: Boolean = true);
var
  p: integer;
begin
  if CorrectCDATA then
  begin
    p := DUAnsiPos(']]>', CDATA);
    while p <> 0 do
    begin
      delete(CDATA, p, 3);
      Insert(']]&gt;', CDATA, p);
      p := DUAnsiPos(']]>', CDATA);
    end;
  end;
  _AddTag('<![CDATA[', CDATA, ']]>', StartNewLine);
end;

procedure TZsspXMLWriter.WriteCDATA(const CDATA: ansistring);
begin
  WriteCDATA(CDATA, true, true);
end;

procedure TZsspXMLWriter.WriteComment(const Comment: ansistring; StartNewLine: Boolean = true);
begin
  _AddTag('<!-- ', Comment, ' -->', StartNewLine);
end;

procedure TZsspXMLWriter.WriteRaw(text: ansistring; UseConverter: Boolean; StartNewLine: Boolean = true);
begin
  if not FInProcess then
    exit;
  if StartNewLine and NewLine then
    text := FNLSeparator + text;
  AddText(text, UseConverter);
end;

// Закончить запись
procedure TZsspXMLWriter.EndSaveTo();
begin
  while TagCount > 0 do
    WriteEndTagNode();
  if NewLine then
    AddText(FNLSeparator, true);
  FlushBuffer();
  // FStream.Free();
  FStream := nil;
  FInProcess := false;
end;

// записываем буфер в поток и очищаем буфер
procedure TZsspXMLWriter.FlushBuffer();
begin
  if not FInProcess then
    exit;
  if FStream <> nil then
    FStream.WriteBuffer(Pointer(FBuffer)^, Length(FBuffer));
  FBuffer := '';
end;

// Установить конвертер текста
// Если началась запись - конвертер менять нельзя
procedure TZsspXMLWriter.SetTextConverter(Value: TAnsiToCPConverter);
begin
  if not InProcess then
    FTextConverter := Value;
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring; AttrArray: array of TZAttrArray;
  StartNewLine: Boolean; CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  t: TZAttributes;
begin
  t := TZAttributes.Create();
  try
    t.Add(AttrArray, true);
    WriteTag(TagName, text, t, StartNewLine, CloseTagNewLine, CheckEntity);
  finally
    FreeAndNil(t);
  end;
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring; SAttributes: TZAttributes;
  StartNewLine: Boolean; CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  S: ansistring;
begin
  if not FInProcess then
    exit;
  WriteTagNode(TagName, SAttributes, StartNewLine, CloseTagNewLine, CheckEntity);
  S := text;
  if CheckEntity then
    S := CheckStrEntity(text);
  WriteRaw(S, true, CloseTagNewLine);
  WriteEndTagNode();
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
begin
  WriteTag(TagName, text, Attributes, StartNewLine, CloseTagNewLine, CheckEntity);
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring; SAttributes: TZAttributes);
begin
  WriteTag(TagName, text, SAttributes, true, false, true);
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring; AttrArray: array of TZAttrArray);
var
  t: TZAttributes;
begin
  t := TZAttributes.Create();
  try
    t.Add(AttrArray, true);
    WriteTag(TagName, text, t, true, false, true);
  finally
    FreeAndNil(t);
  end;
end;

procedure TZsspXMLWriter.WriteTag(const TagName: ansistring; const text: ansistring);
begin
  WriteTag(TagName, text, Attributes, true, false, true);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring; SAttributes: TZAttributes);
begin
  WriteTagNode(TagName, SAttributes, true, false, true);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring; AttrArray: array of TZAttrArray);
begin
  WriteTagNode(TagName, AttrArray, true, false, true);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring);
begin
  WriteTagNode(TagName, Attributes, true, false, true);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring; StartNewLine: Boolean; CloseTagNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  WriteTagNode(TagName, Attributes, StartNewLine, CloseTagNewLine, CheckEntity);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring; SAttributes: TZAttributes; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  S: ansistring;
begin
  if not FInProcess then
    exit;
  S := '';
  if FNewLine and StartNewLine then
  begin
    S := S + FNLSeparator;
    S := S + GetTab();
  end;
  S := S + '<' + TagName;
  if SAttributes <> nil then
    if SAttributes.Count > 0 then
      S := S + SAttributes.ToString(FAttributeQuote, CheckEntity);
  S := S + '>';
  AddText(S);
  AddNode(TagName, CloseTagNewLine);
end;

procedure TZsspXMLWriter.WriteTagNode(const TagName: ansistring; AttrArray: array of TZAttrArray; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  arttrib: TZAttributes;
begin
  arttrib := TZAttributes.Create();
  try
    arttrib.Add(AttrArray, true);
    WriteTagNode(TagName, arttrib, StartNewLine, CloseTagNewLine, CheckEntity);
  finally
    arttrib.Free();
  end;
end;

procedure TZsspXMLWriter.WriteEndTagNode();
begin
  if not FInProcess then
    exit;
  if TagCount > 0 then
  begin
    _AddTag('</', FTags[TagCount - 1].Name, '>', FTags[TagCount - 1].CloseTagNewLine, -1);
    dec(FTagCount);
    ResizeTagArray(FTagCount);
  end;
end;

procedure TZsspXMLWriter.WriteEndTagNode(isForce: Boolean; CloseTagNewLine: Boolean);
var
  b: Boolean;
begin
  if not FInProcess then
    exit;
  if TagCount > 0 then
  begin
    b := FTags[TagCount - 1].CloseTagNewLine;
    if (isForce) then
      b := CloseTagNewLine;
    _AddTag('</', FTags[TagCount - 1].Name, '>', b, -1);
    dec(FTagCount);
    ResizeTagArray(FTagCount);
  end;
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring; SAttributes: TZAttributes; StartNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  _AddTag('<' + TagName, SAttributes.ToString(AttributeQuote, CheckEntity), '/>', StartNewLine)
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring; AttrArray: array of TZAttrArray;
  StartNewLine: Boolean; CheckEntity: Boolean = true);
var
  t: TZAttributes;
begin
  t := TZAttributes.Create();
  try
    t.Add(AttrArray, true);
    WriteEmptyTag(TagName, t, StartNewLine, CheckEntity);
  finally
    FreeAndNil(t);
  end;
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring; StartNewLine: Boolean; CheckEntity: Boolean = true);
begin
  WriteEmptyTag(TagName, Attributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring; SAttributes: TZAttributes);
begin
  WriteEmptyTag(TagName, SAttributes, true, true);
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring; AttrArray: array of TZAttrArray);
begin
  WriteEmptyTag(TagName, AttrArray, true, true);
end;

procedure TZsspXMLWriter.WriteEmptyTag(const TagName: ansistring);
begin
  WriteEmptyTag(TagName, Attributes, true, true);
end;

procedure TZsspXMLWriter.AddText(const text: ansistring; UseConverter: Boolean = true);
var
  b: Boolean;
begin
  b := false;
  if UseConverter then
    if Assigned(TextConverter) then
      b := true;
  if (b) then
    FBuffer := FBuffer + TextConverter(text)
  else
    FBuffer := FBuffer + text;
  if Length(FBuffer) >= MaxBufferLength then
    FlushBuffer;
end;

function TZsspXMLWriter.GetTab(num: integer = 0): ansistring;
var
  i: integer;
begin
  result := '';
  for i := 1 to TagCount + num do
    result := result + FTab;
end;

// Добавляет в "Стек тэгов" новый тэг
// Input:
// Tag: ansistring                - тэг
// CloseTagNewLine: boolean   - Начинать ли закрывающий тэг с новой строки
procedure TZsspXMLWriter.AddNode(const TagName: ansistring; CloseTagNewLine: Boolean);
begin
  ResizeTagArray(FTagCount + 1);
  FTags[FTagCount].Name := TagName;
  FTags[FTagCount].CloseTagNewLine := CloseTagNewLine;
  inc(FTagCount);
end;

function TZsspXMLWriter.GetTabSymbol(): ansichar;
begin
  result := #0;
  if Length(FTabSymbol) >= 1 then
    result := FTabSymbol[1];
end;

// Устанавливает символ табуляции (только #32 и #9 - пробел и таб)
procedure TZsspXMLWriter.SetTabSymbol(Value: ansichar);
begin
  if not InProcess then
    if (Value = #9) or (Value = #32) then
    begin
      FTabSymbol := Value;
      SetTabLength(TabLength);
    end;
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring; SAttributes: TZAttributes;
  StartNewLine: Boolean; CheckEntity: Boolean = true);
begin
  _AddTag('<?' + InstructionName, SAttributes.ToString(AttributeQuote, CheckEntity), '?>', StartNewLine)
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring; AttrArray: array of TZAttrArray;
  StartNewLine: Boolean; CheckEntity: Boolean = true);
var
  t: TZAttributes;
begin
  t := TZAttributes.Create();
  try
    t.Add(AttrArray, true);
    WriteInstruction(InstructionName, t, StartNewLine, CheckEntity);
  finally
    FreeAndNil(t);
  end;
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring; StartNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  WriteInstruction(InstructionName, Attributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring; SAttributes: TZAttributes);
begin
  WriteInstruction(InstructionName, SAttributes, true, true);
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring; AttrArray: array of TZAttrArray);
begin
  WriteInstruction(InstructionName, AttrArray, true, true);
end;

procedure TZsspXMLWriter.WriteInstruction(const InstructionName: ansistring);
begin
  WriteInstruction(InstructionName, Attributes, true, true);
end;

/// /::::::::::::: TZsspXMLReader :::::::::::::::::////

constructor TZsspXMLReader.Create();
begin
  inherited Create();
  FAttributes := TZAttributes.Create();
  FMaxBufferLength := 4096;
  FBuffer := '';
  ClearAll();
  FIgnoreCase := false;
  FQuotesEqual := false;
  FMaxTagCount := 20;
  FAttributesMatch := true;
  setlength(FTags, FMaxTagCount);
end;

destructor TZsspXMLReader.Destroy();
begin
  EndRead();
  FreeAndNil(FAttributes);
  setlength(FTags, 0);
  FTags := nil;
  inherited Destroy();
end;

// Изменяет размер массива с тэгами
// INPUT
// NewSize: integer - новый размер
procedure TZsspXMLReader.ResizeTagArray(NewSize: integer);
var
  delta: integer;
begin
  delta := 0;
  if (NewSize >= FMaxTagCount) then
  begin
    delta := NewSize;
    if (NewSize < 50) then
      delta := delta * 4
    else if (NewSize < 100) then
      delta := delta * 2
    else
      delta := delta + 20;
  end
  else if (NewSize > 50) then
  begin
    if (FMaxTagCount - NewSize > 200) then
      delta := NewSize + 100;
  end;
  if (delta > 0) then
    setlength(FTags, delta);
end; // ResizeTagArray

procedure TZsspXMLReader.SetAttributesMatch(Value: Boolean);
begin
  if (not FInProcess) then
    FAttributesMatch := Value;
end;

function TZsspXMLReader.Eof(): Boolean;
begin
  result := false;
  if InProcess then
  begin
    if FPFirst > FPLast then
    begin
      ReadBuffer();
      if FPFirst > FPLast then
        result := true;
    end;
  end
  else
    result := true;
end;

procedure TZsspXMLReader.SetQuotesEqual(Value: Boolean);
begin
  if (not InProcess) then
    FQuotesEqual := Value;
end;

procedure TZsspXMLReader.SetAttributes(Value: TZAttributes);
begin
  if Value <> nil then
    FAttributes.Assign(Value);
end;

function TZsspXMLReader.GetTag(num: integer): ansistring;
begin
  if (num >= 0) and (num < TagCount) then
    result := FTags[num]
  else
    result := '';
end;

// Устанавливает максимальную длинну буфера
procedure TZsspXMLReader.SetMaxBufferLength(Value: integer);
begin
  if Value > 513 then
    if not InProcess then
    begin
      FMaxBufferLength := Value;
      SetString(FBuffer, nil, FMaxBufferLength);
      // SetLength(FBuffer, FMaxBufferLength);
      // FPFirst := 0;
      // FPlast := 0;
    end;
end;

function TZsspXMLReader.BeginReadStream(Stream: TStream): integer;
var
  S: ansistring;
begin
  if InProcess then
    result := 1
  else
  begin
    result := 0;
    try
      FStream := Stream;
    except
      result := 2;
    end;
    if FStream = nil then
      result := 3;
    if result = 0 then
    begin
      FInProcess := true;
      if FSourceType = 111 then
        FSourceType := 1
      else
        FSourceType := 2;
      ClearAll();
      ReadBuffer();
      // выбор энкодера для чтения
      S := copy(FBuffer, 1, FPLast);
      RecognizeEncoding(S);
    end;
  end;
end;

// закончить чтение
procedure TZsspXMLReader.EndRead();
begin
  if InProcess then
  begin
    FInProcess := false;
    // if FSourceType = 1 then
    // FStream := nil;
    if FSourceType = 1 then
      FreeAndNil(FStream);
  end;
end;

// Очистка (кроме кол-ва тэгов и стека)
procedure TZsspXMLReader.Clear();
begin
  FTextBeforeTag := '';
  FRawTextTag := '';
  FTagName := '';
  FTagType := TXmlTagType.xttUnknown;
  FValue := '';
  FErrorCode := 0;
  FAttributes.Clear();
end;

// Очистка всего
procedure TZsspXMLReader.ClearAll();
var
  t: integer;
begin
  Clear();
  if FSourceType <> 3 then
    SetString(FBuffer, nil, MaxBufferLength)
  else
  begin
    t := Length(FBuffer);
    if t > 0 then
    begin
      MaxBufferLength := t;
      FPFirst := 1;
      FPLast := t;
    end;
  end;
  FTagCount := 0;
  ResizeTagArray(0);
  FStreamEnd := false;
end;

// Читает 1 символ из буфера
procedure TZsspXMLReader.GetOneChar(var OneChar: ansistring; var err: Boolean);
begin
  err := false;
  OneChar := '';
  if FPFirst <= FPLast then
  begin
    OneChar := FBuffer[FPFirst];
    inc(FPFirst);
  end
  else
  begin
    ReadBuffer();
    if FPFirst <= FPLast then
    begin
      OneChar := FBuffer[FPFirst];
      inc(FPFirst);
    end
    else
      err := true;
  end;
end;

// Читает буфер из потока
procedure TZsspXMLReader.ReadBuffer();
var
  t: integer;
begin
  if InProcess then
    if (FSourceType = 1) or (FSourceType = 2) then // только для файла/потока
    begin
      t := 0;
    // SetString(FBuffer, nil, MaxBufferLength);
      if Assigned(FStream) and (not FStreamEnd) then
      begin
        t := FStream.Read(Pointer(FBuffer)^, MaxBufferLength);
        if t < MaxBufferLength then
          FStreamEnd := true;
      end;
      if t >= 1 then
      begin
        FPFirst := 1;
        FPLast := t;
      end
      else
      begin
        FPFirst := 1;
        FPLast := 0;
      end;
    end;
end;

// Попытка распознания кодировки
// Берём текст из буфера и пытаемся распознать кодировку
// INPUT:
// var txt: ansistring - анализируемый текст
procedure TZsspXMLReader.RecognizeEncoding(var txt: ansistring);
var
  BOM: integer;
  cpname: ansistring;
  recognized: Boolean; // флаг точного определения кодировки
  codepagenum: integer; // кодировка
  ftype: integer;

begin
  // кодировка текста
  // codepagenum:
  // 0 - неопределена
  // 1 - UTF-8
  // 2 - UTF-16BE
  // 3 - UTF-16LE
  // 4 - UTF-32BE
  // 5 - UTF-32LE
  // 6 - Windows-1251
  // 7 - CP866
  recognized := RecognizeEncodingXML(txt, BOM, codepagenum, cpname, ftype);
  if recognized then
  begin
    FCharReader := nil;
    case codepagenum of
      1:
        begin
          FCharReader := @ReadCharUTF8;
          FCharConverter := @conv_UTF8ToLocal;
          if BOM = 1 then
            self.FPFirst := 4;
        end;
      2:
        begin
          FCharReader := @ReadCharUTF16BE;
          FCharConverter := @conv_UTF16BEToLocal;
          if BOM = 2 then
            self.FPFirst := 3;
        end;
      3:
        begin
          FCharReader := @ReadCharUTF16LE;
          FCharConverter := @conv_UTF16LEToLocal;
          if BOM = 3 then
            self.FPFirst := 3;
        end;
      4:
        begin
          FCharReader := @ReadCharUTF32;
          FCharConverter := @conv_UTF32BEToLocal;
        end;
      5:
        begin
          FCharReader := @ReadCharUTF32;
          FCharConverter := @conv_UTF32LEToLocal;
        end;
      6:
        FCharConverter := @conv_WIN1251ToLocal;
      7:
        FCharConverter := @conv_CP866ToLocal;
    end;
  end
  else
  begin
    { tut }// вручную перепроверка і всё такое
    FCharReader := nil;
    case codepagenum of
      1:
        begin
          FCharReader := @ReadCharUTF8;
          FCharConverter := @conv_UTF8ToLocal;
          if BOM = 1 then
            self.FPFirst := 4;
        end;
      2:
        begin
          FCharReader := @ReadCharUTF16BE;
          FCharConverter := @conv_UTF16BEToLocal;
          if BOM = 2 then
            self.FPFirst := 3;
        end;
      3:
        begin
          FCharReader := @ReadCharUTF16LE;
          FCharConverter := @conv_UTF16LEToLocal;
          if BOM = 3 then
            self.FPFirst := 3;
        end;
      4:
        begin
          FCharReader := @ReadCharUTF32;
          FCharConverter := @conv_UTF32BEToLocal;
        end;
      5:
        begin
          FCharReader := @ReadCharUTF32;
          FCharConverter := @conv_UTF32LEToLocal;
        end;
      6:
        FCharConverter := @conv_WIN1251ToLocal;
      7:
        FCharConverter := @conv_CP866ToLocal;
    end;
  end;
end;

function TZsspXMLReader.ReadTag(): Boolean;
var
  ch: ansistring;
  ss: ansistring;
  err: Boolean;
  RawEncodingBeforeTag: ansistring;
  RawTextTagNonDecoded: ansistring;
  end_tag: Boolean; // найден конец тэга
  _isClosedTag: Boolean;
  _isInstruction: Boolean;

  procedure _get_char();
  begin
    if Assigned(FCharReader) then
      FCharReader(GetOneChar, ch, err)
    else
      GetOneChar(ch, err);
    if err then
      exit;
    if Assigned(FCharConverter) then
      ss := FCharConverter(ch)
    else
      ss := ch;
  end;

  procedure RawTag();
  begin
    RawTextTagNonDecoded := RawTextTagNonDecoded + ch;
    if ss[1] <> #0 then // todo: что делать, если в тексте действительно будет #0?
      FRawTextTag := FRawTextTag + ss;
  end;

  // срабатывает после обнаружения "="
  function GetParamValue(): ansistring;
  var
    count_quote: byte;
    _openquote: ansistring;

  begin
    count_quote := 0;
    result := '';
    while true do
    begin
      _get_char();
      if err then
        break;
      if Length(ss) >= 1 then
      begin
        RawTag();
        case ss[1] of
          ' ', #13, #10, #9:
            begin
              if count_quote = 1 then
                result := result + ss[1]
              else if Length(result) > 0 then
              begin
                FErrorCode := FErrorCode or 1;
                break;
              end
            end;
          '>':
            begin
              if count_quote = 1 then
                result := result + ss[1] // tut
              else
              begin
                if Length(result) > 0 then
                  FErrorCode := FErrorCode or 1
                else
                  FErrorCode := FErrorCode or 2;
                end_tag := true;
                break;
              end;
            end;
          '=':
            begin
              if count_quote = 1 then
                result := result + ss[1]
              else
              begin
                if Length(result) > 0 then
                  FErrorCode := FErrorCode or 4
                else
                  FErrorCode := FErrorCode or 2;
                break;
              end;
            end;
          '?':
            begin
              if count_quote = 1 then
                result := result + ss[1]
              else
              begin
                if Length(result) > 0 then
                  _isInstruction := true
                else
                  FErrorCode := FErrorCode or 1024;
                break;
              end;
            end;
          '/':
            begin
              if count_quote = 1 then
                result := result + ss[1]
              else
              begin
                if Length(result) > 0 then
                  _isClosedTag := true
                else
                  FErrorCode := FErrorCode or 8;
                break;
              end;
            end;
          '<':
            begin
              if count_quote = 1 then
                result := result + ss
              else
                // будем игнорировать
                FErrorCode := FErrorCode or 32;
            end;
          '"', '''':
            begin
              if count_quote = 0 then
              begin
                if Length(result) > 0 then // <tag ... param = value"...>
                begin
                  FErrorCode := FErrorCode or 16;
                  break;
                end
                else
                begin
                  _openquote := ss[1];
                  inc(count_quote);
                end;
              end
              else
              begin
                if (QuotesEqual or (ss[1] = _openquote)) then
                  break
                else
                  result := result + ss;
              end;
            end;
        else
          result := result + ss;
        end;
      end; // if
    end; // while
  end;

  procedure Delete2end(var svalue: ansistring);
  var
    t: integer;
  begin
    t := Length(svalue);
    if t >= 2 then
      delete(svalue, t - 1, 2);
  end;

  // комментарий / CDATA
  function GetCommentCDATA(): integer;
  var
    _type_comment: integer; // сразу в result
    _tmp: integer;
    _last2: ansistring;
    S, sl: ansistring;

  begin
    setlength(_last2, 2);
    _last2[1] := #0;
    _last2[2] := #0;
    _type_comment := 0;
    S := '';
    sl := '';
    while true do
    begin
      if end_tag then
        break;
      _get_char();
      if err then
        break;
      if Length(ss) >= 1 then
      begin
        RawTag();
        case ss[1] of
          '>':
            begin
              case _type_comment of
                - 1: // <!unknown_tag ....> - сразу заканчиваем
                  begin
                    end_tag := true;
                    break;
                  end;
                0: // O_o <!ta>...
                  begin
                    end_tag := true;
                    FErrorCode := FErrorCode or 128;
                    break;
                  end;
                1: // <!--...-->
                  begin
                    if _last2 = '--' then
                    begin
                      Delete2end(FValue);
                      end_tag := true;
                      break;
                    end
                    else
                      FValue := FValue + ss;
                  end;
                2: // <![CDATA[...]]>
                  begin
                    if _last2 = ']]' then
                    begin
                      Delete2end(FValue);
                      end_tag := true;
                      break;
                    end
                    else
                      FValue := FValue + ss;
                  end;
              end;
            end;
        else
          begin
            if _type_comment <> 0 then
              FValue := FValue + ss
            else
            begin
              _tmp := Length(S);
              if _tmp = 2 then
              begin
                if S = '--' then
                begin
                  _type_comment := 1;
                  FValue := FValue + ss;
                end
                else
                  S := S + UTF8Encode(UpperCase(UTF8ToString(ansistring(ss[1]))));
              end
              else if _tmp = 7 then
              begin
                if S = '[CDATA[' then
                begin
                  _type_comment := 2;
                  FValue := FValue + ss;
                end
                else
                begin
                  FErrorCode := FErrorCode or 128;
                  _type_comment := -1;
                  FValue := sl + ss;
                end;
                  // FValue := FValue + ss;
              end
              else
              begin
                S := S + UTF8Encode(UpperCase(UTF8ToString(ansistring(ss[1]))));
                sl := sl + ss[1];
              end;
            end;
          end; // else
        end; // case

        // '-->' ']]>'
        _last2[2] := _last2[1];
        _last2[1] := ss[1];
      end;
    end;
    result := _type_comment;
  end;

  procedure CheckClose(var some_par: Boolean; err_code: integer);
  begin
    if some_par then
    begin
      if not(ss[1] in [' ', #0, #9, #10, #13, '>']) then
      begin
        some_par := false;
        FErrorCode := FErrorCode or err_code;
      end;
    end;
  end;

  procedure ProcessTag();
  var
    _isTagName: byte;
    _isParam: Boolean;
    S: ansistring;
    _tmp: integer;

  begin
    _isTagName := 0;
    _isInstruction := false;
    _isParam := false;
    S := '';
    while true do
    begin
      if end_tag then
        break;
      _get_char();
      if err then
        break;
      if Length(ss) >= 1 then
      begin
        RawTag();
        CheckClose(_isClosedTag, 64);
        CheckClose(_isInstruction, 1024);

        case ss[1] of
          '>':
            begin
              if _isTagName = 0 then
              begin
                if Length(S) > 0 then
                begin
                  FTagName := S;
                  S := '';
                end
                else
                  FErrorCode := FErrorCode or 256;
              end;
              //
              if _isClosedTag then
              begin
                if (FTagType <> TXmlTagType.xttDeclare) and (FTagType <> TXmlTagType.xttEnd) then
                  FTagType := TXmlTagType.xttClosed
                else
                  FErrorCode := FErrorCode or 8192;
              end
              else if _isInstruction then
              begin
                if FTagType <> TXmlTagType.xttDeclare then
                  FErrorCode := FErrorCode or 16384;
              end;
              end_tag := true;
            end;
          '<':
            FErrorCode := FErrorCode or 32768; // warning/error/ignore? {tut}
          '=':
            begin
              _isParam := false;
              if _isTagName > 0 then
              begin
                if Length(S) > 0 then
                begin
                  Attributes.Add(S, GetParamValue(), FAttributesMatch);
                  S := '';
                end;
              end
              else
              begin
                // < =...>
                if Length(S) > 0 then
                begin
                  FTagName := S;
                  S := '';
                  GetParamValue(); // может, забыли имя параметра ввести? {tut}
                end;
                FErrorCode := FErrorCode or 256;
                _isTagName := 13;
              end;
            end;
          '!': // комментарий / CDATA
            begin
              if (_isTagName = 0) and (Length(S) = 0) then
              begin
                _tmp := GetCommentCDATA();
                if end_tag then
                  case _tmp of
                    1:
                      FTagType := TXmlTagType.xttComment;
                    2:
                      FTagType := TXmlTagType.xttCData;
                  end;
                break;
              end
              else
              begin
                // <TAG!>
                FErrorCode := FErrorCode or 512;
              end;
            end;
          '?':
            begin
              if (_isTagName = 0) and (Length(S) = 0) then
                FTagType := TXmlTagType.xttDeclare
              else
                _isInstruction := true;
            end;
          '/':
            begin
              // сделать проверку, если текст тэга не пустой
              if (_isTagName = 0) and (Length(S) = 0) then
                FTagType := TXmlTagType.xttEnd
              else
              begin
                if _isTagName = 0 then
                  FTagName := S;
                _isClosedTag := true;
              end;
            end;
          '"', '''':
            FErrorCode := FErrorCode or 65536;
          ' ', #13, #10, #9:
            begin
              if (_isTagName = 0) then
              begin
                if Length(S) > 0 then
                begin
                  FTagName := S;
                  _isTagName := 1;
                  S := '';
                end
                else
                  FErrorCode := FErrorCode or 2048;
              end
              else
              begin
                if Length(S) > 0 then
                  _isParam := true;
              end;
            end;
          #0:
            ; // и что с ним делать? {tut}
        else
          begin
            if _isParam then
            begin
              if Length(S) > 0 then
              begin
                _isParam := false;
                S := '';
                FErrorCode := FErrorCode or 4096;
              end;
            end;
            S := S + ss;
          end;
        end; // case
      end; // if
    end; // while
  end;

begin
  result := true;
  end_tag := false;
  _isClosedTag := false;
  RawEncodingBeforeTag := '';
  err := false;
  Clear();
  while true do
  begin
    _get_char();
    if err then
      break;

    if Length(ss) >= 1 then
    begin
      case ss[1] of
        '<':
          begin
            FRawTextTag := ss;
            RawTextTagNonDecoded := ch;
            ProcessTag();

            if err then
              FErrorCode := FErrorCode or 131072
            else if end_tag then
              if (FTagType = TXmlTagType.xttUnknown) and (Length(FTagName) > 0) then
                FTagType := TXmlTagType.xttStart;

            break;
          end;
      else
        begin
          FTextBeforeTag := FTextBeforeTag + ss;
          RawEncodingBeforeTag := RawEncodingBeforeTag + ch;
        end;
      end; // case
    end; // if
  end; // while
  if FTagType = TXmlTagType.xttStart then
    AddTag(FTagName)
  else if FTagType = TXmlTagType.xttEnd then
    DeleteClosedTag();
  if Eof() then
    if TagCount > 0 then
      FErrorCode := FErrorCode or 524288;
end;

// Добавить тэг в стэк
procedure TZsspXMLReader.AddTag(const Value: ansistring);
begin
  inc(FTagCount);
  ResizeTagArray(FTagCount);
  FTags[FTagCount - 1] := Value;
end;

procedure TZsspXMLReader.SetIgnoreCase(Value: Boolean);
begin
  if not InProcess then
    FIgnoreCase := Value
end;

procedure TZsspXMLReader.DeleteTag();
begin
  if TagCount > 0 then
  begin
    dec(FTagCount);
    ResizeTagArray(FTagCount);
  end;
end;

procedure TZsspXMLReader.DeleteClosedTag();
var
  b: Boolean;
begin
  b := false;
  if TagCount = 0 then
  begin
    FErrorCode := FErrorCode or 262144;
    exit;
  end;
  if IgnoreCase then
  begin
    if (AnsiUpperCase(UTF8ToString(TagName)) = AnsiUpperCase(UTF8ToString(Tags[TagCount - 1]))) then
      b := true;
  end
  else
  begin
    if TagName = Tags[TagCount - 1] then
      b := true;
  end;
  if b then
    DeleteTag()
  else
    FErrorCode := FErrorCode or 262144;
end;

constructor TZAttributesH.Create();
begin
  inherited;
  FAttributes := TZAttributes.Create();
end;

destructor TZAttributesH.Destroy();
begin
  FreeAndNil(FAttributes);
  inherited;
end;

procedure TZAttributesH.Assign(Source: TPersistent);
begin
  if (Source is TZAttributesH) then
    self.FAttributes.Assign((Source as TZAttributesH).FAttributes)
  else if (Source is TZAttributes) then
    self.FAttributes.Assign((Source as TZAttributes));
end;

function TZAttributesH.GetAttrCount(): integer;
begin
  result := FAttributes.Count;
end;

// Получает значение атрибута по его названию
function TZAttributesH.GetAttrS(Att: string): string;
begin
  result := UTF8ToString(FAttributes.ItemsByName[UTF8Encode(Att)]);
end;

// Устанавливает значение атрибута по названию
procedure TZAttributesH.SetAttrS(Att: string; const Value: string);
begin
  FAttributes.ItemsByName[UTF8Encode(Att)] := UTF8Encode(Value);
end;

// Получает знаечние атрибута по номеру
function TZAttributesH.GetAttrI(num: integer): string;
begin
  result := UTF8ToString(FAttributes.ItemsByNum[num]);
end;

// Устанавливает значение атрибута по номеру
procedure TZAttributesH.SetAttrI(num: integer; const Value: string);
begin
  FAttributes.ItemsByNum[num] := UTF8Encode(Value);
end;

// Получает значение атрибут-значение по номеру
function TZAttributesH.GetAttr(num: integer): TZAttrArrayH;
var
  t: TZAttrArrayH;
  a: TZAttrArray;
  i: integer;
begin
  a := FAttributes.Items[num];
  for i := 0 to 1 do
    t[i] := UTF8ToString(a[i]);
  result := t;
end;

// Устанавливает атрибут-значение по номеру
procedure TZAttributesH.SetAttr(num: integer; const Value: TZAttrArrayH);
var
  t: TZAttrArrayH;
  a: TZAttrArray;
  i: integer;
begin
  for i := 0 to 1 do
    a[i] := UTF8Encode(t[i]);
  FAttributes.Items[num] := a;
end;

procedure TZAttributesH.Add(const AttrName: string; const Value: string; TestMatch: Boolean = true);
begin
  FAttributes.Add(UTF8Encode(AttrName), UTF8Encode(Value), TestMatch);
end;

procedure TZAttributesH.Add(const Attr: TZAttrArrayH; TestMatch: Boolean = true);
begin
  Add(Attr[0], Attr[1], TestMatch);
end;

procedure TZAttributesH.Add(Att: array of TZAttrArrayH; TestMatch: Boolean = true);
var
  i: integer;
begin
  for i := Low(Att) to High(Att) do
    Add(Att[i][0], Att[i][1], TestMatch);
end;

procedure TZAttributesH.Clear();
begin
  FAttributes.Clear();
end;

procedure TZAttributesH.DeleteItem(Index: integer);
begin
  FAttributes.DeleteItem(Index);
end;

procedure TZAttributesH.Insert(Index: integer; const AttrName: string; const Value: string; TestMatch: Boolean = true);
begin
  FAttributes.Insert(Index, UTF8Encode(AttrName), UTF8Encode(Value), TestMatch);
end;

procedure TZAttributesH.Insert(Index: integer; const Attr: TZAttrArrayH; TestMatch: Boolean = true);
begin
  Insert(Index, Attr[0], Attr[1], TestMatch);
end;

function TZAttributesH.ToString(quote: char; CheckEntity: Boolean; addempty: Boolean): string;
begin
  result := UTF8ToString(FAttributes.ToString(ansichar(quote), CheckEntity, addempty));
end;

function TZAttributesH.ToString(quote: char; CheckEntity: Boolean): string;
begin
  result := UTF8ToString(FAttributes.ToString(ansichar(quote), CheckEntity));
end;

function TZAttributesH.ToString(quote: char): string;
begin
  result := UTF8ToString(FAttributes.ToString(ansichar(quote)));
end;

function TZAttributesH.ToString(CheckEntity: Boolean): string;
begin
  result := UTF8ToString(FAttributes.ToString(CheckEntity));
end;

function TZAttributesH.ToString(): string;
begin
  result := UTF8ToString(FAttributes.ToString());
end;

function TZAttributesH.IsContainsAttribute(const AttrName: string; CaseSensitivity: Boolean = true): Boolean;
begin
  result := FAttributes.IsContainsAttribute(AttrName, CaseSensitivity);
end;

/// /////////////////////////////////////////////////////////////////////////////
/// / TZsspXMLWriterH
/// /////////////////////////////////////////////////////////////////////////////

constructor TZsspXMLWriterH.Create(Stream: TStream);
begin
  inherited Create();
  FXMLWriter := TZsspXMLWriter.Create(Stream);
  FAttributes := TZAttributesH.Create();
end;

destructor TZsspXMLWriterH.Destroy();
begin
  FreeAndNil(FXMLWriter);
  FreeAndNil(FAttributes);
  inherited;
end;

function TZsspXMLWriterH.GetXMLBuffer(): string;
begin
  result := UTF8ToString(FXMLWriter.Buffer);
end;

function TZsspXMLWriterH.GetAttributeQuote(): char;
begin
  result := char(FXMLWriter.AttributeQuote);
end;

function TZsspXMLWriterH.GetInProcess(): Boolean;
begin
  result := FXMLWriter.InProcess
end;

function TZsspXMLWriterH.GetMaxBufferLength(): integer;
begin
  result := FXMLWriter.MaxBufferLength;
end;

function TZsspXMLWriterH.GetNewLine(): Boolean;
begin
  result := FXMLWriter.NewLine;
end;

function TZsspXMLWriterH.GetTabLength(): integer;
begin
  result := FXMLWriter.TabLength;
end;

function TZsspXMLWriterH.GetTagCount(): integer;
begin
  result := FXMLWriter.TagCount;
end;

function TZsspXMLWriterH.GetTextConverter(): TCPToAnsiConverter;
begin
  result := TextConverter;
end;

function TZsspXMLWriterH.GetUnixNLSeparator(): Boolean;
begin
  result := FXMLWriter.UnixNLSeparator;
end;

function TZsspXMLWriterH.GetTag(num: integer): string;
begin
  result := UTF8ToString(FXMLWriter.Tags[num]);
end;

function TZsspXMLWriterH.GetTabSymbol(): char;
begin
  result := char(FXMLWriter.TabSymbol);
end;

procedure TZsspXMLWriterH.SetAttributeQuote(Value: char);
begin
  FXMLWriter.AttributeQuote := ansichar(Value);
end;

procedure TZsspXMLWriterH.SetMaxBufferLength(Value: integer);
begin
  FXMLWriter.MaxBufferLength := Value;
end;

procedure TZsspXMLWriterH.SetNewLine(Value: Boolean);
begin
  FXMLWriter.NewLine := Value;
end;

procedure TZsspXMLWriterH.SetTabLength(Value: integer);
begin
  FXMLWriter.TabLength := Value;
end;

procedure TZsspXMLWriterH.SetTabSymbol(Value: char);
begin
  FXMLWriter.TabSymbol := ansichar(Value);
end;

procedure TZsspXMLWriterH.SetTextConverter(Value: TAnsiToCPConverter);
begin
  FXMLWriter.TextConverter := Value;
end;

procedure TZsspXMLWriterH.SetUnixNLSeparator(Value: Boolean);
begin
  FXMLWriter.UnixNLSeparator := Value;
end;

procedure TZsspXMLWriterH.SetAttributes(Value: TZAttributesH);
begin
  if (Value <> nil) then
    FAttributes.Assign(Value);
end;

procedure TZsspXMLWriterH.EndSaveTo();
begin
  FXMLWriter.EndSaveTo();
end;

procedure TZsspXMLWriterH.FlushBuffer();
begin
  FXMLWriter.FlushBuffer();
end;

procedure TZsspXMLWriterH.WriteCDATA(CDATA: string; CorrectCDATA: Boolean; StartNewLine: Boolean = true);
begin
  FXMLWriter.WriteCDATA(UTF8Encode(CDATA), CorrectCDATA, StartNewLine);
end;

procedure TZsspXMLWriterH.WriteCDATA(CDATA: string);
begin
  FXMLWriter.WriteCDATA(UTF8Encode(CDATA));
end;

procedure TZsspXMLWriterH.WriteComment(Comment: string; StartNewLine: Boolean = true);
begin
  FXMLWriter.WriteComment(UTF8Encode(Comment), StartNewLine);
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteEmptyTag(UTF8Encode(TagName), SAttributes.FAttributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
  CheckEntity: Boolean = true);
var
  a: array of TZAttrArray;
  kol, _start: integer;
  i, num: integer;
begin
  kol := High(AttrArray);
  _start := Low(AttrArray);
  try
    setlength(a, kol - _start + 1);
    num := 0;
    for i := _start to kol do
    begin
      a[num][0] := UTF8Encode(AttrArray[i][0]);
      a[num][1] := UTF8Encode(AttrArray[i][1]);
    end;
    FXMLWriter.WriteEmptyTag(UTF8Encode(TagName), a, StartNewLine, CheckEntity);
  finally
    setlength(a, 0);
    a := nil;
  end;
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string; StartNewLine: Boolean; CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteEmptyTag(UTF8Encode(TagName), FAttributes.FAttributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string; SAttributes: TZAttributesH);
begin
  FXMLWriter.WriteEmptyTag(UTF8Encode(TagName), SAttributes.FAttributes);
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string; AttrArray: array of TZAttrArrayH);
begin
  WriteEmptyTag(TagName, AttrArray, true, true);
end;

procedure TZsspXMLWriterH.WriteEmptyTag(TagName: string);
begin
  WriteEmptyTag(TagName, FAttributes, true, true);
end;

procedure TZsspXMLWriterH.WriteEndTagNode();
begin
  FXMLWriter.WriteEndTagNode();
end;

procedure TZsspXMLWriterH.WriteEndTagNode(isForce: Boolean; CloseTagNewLine: Boolean);
begin
  FXMLWriter.WriteEndTagNode(isForce, CloseTagNewLine);
end;

procedure TZsspXMLWriterH.WriteHeader(CodePageName: string; BOM: ansistring);
begin
  WriteRaw(BOM, false, false);
  Attributes.Clear();

  Attributes.Add('version', '1.0');
  if (CodePageName > '') then
    Attributes.Add('encoding', CodePageName);
  WriteInstruction('xml', false);

  Attributes.Clear();
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteInstruction(UTF8Encode(InstructionName), SAttributes.FAttributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string; AttrArray: array of TZAttrArrayH;
  StartNewLine: Boolean; CheckEntity: Boolean = true);
var
  a: array of TZAttrArray;
  kol, _start: integer;
  i, num: integer;
begin
  kol := High(AttrArray);
  _start := Low(AttrArray);
  try
    setlength(a, kol - _start + 1);
    num := 0;
    for i := _start to kol do
    begin
      a[num][0] := UTF8Encode(AttrArray[i][0]);
      a[num][1] := UTF8Encode(AttrArray[i][1]);
    end;
    FXMLWriter.WriteInstruction(UTF8Encode(InstructionName), a, StartNewLine, CheckEntity);
  finally
    setlength(a, 0);
    a := nil;
  end;
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string; StartNewLine: Boolean; CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteInstruction(UTF8Encode(InstructionName), FAttributes.FAttributes, StartNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string; SAttributes: TZAttributesH);
begin
  FXMLWriter.WriteInstruction(UTF8Encode(InstructionName), SAttributes.FAttributes);
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string; AttrArray: array of TZAttrArrayH);
begin
  WriteInstruction(InstructionName, AttrArray, true, true);
end;

procedure TZsspXMLWriterH.WriteInstruction(InstructionName: string);
begin
  FXMLWriter.WriteInstruction(UTF8Encode(InstructionName), FAttributes.FAttributes);
end;

procedure TZsspXMLWriterH.WriteRaw(text: string; UseConverter: Boolean; StartNewLine: Boolean = true);
begin
  FXMLWriter.WriteRaw(UTF8Encode(text), UseConverter, StartNewLine);
end;

procedure TZsspXMLWriterH.WriteRaw(text: ansistring; UseConverter: Boolean; StartNewLine: Boolean = true);
begin
  FXMLWriter.WriteRaw(text, UseConverter, StartNewLine);
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteTag(UTF8Encode(TagName), UTF8Encode(text), SAttributes.FAttributes, StartNewLine, CloseTagNewLine,
    CheckEntity);
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string; AttrArray: array of TZAttrArrayH;
  StartNewLine: Boolean; CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  a: array of TZAttrArray;
  kol, _start: integer;
  i, num: integer;
begin
  kol := High(AttrArray);
  _start := Low(AttrArray);
  try
    setlength(a, kol - _start + 1);
    num := 0;
    for i := _start to kol do
    begin
      a[num][0] := UTF8Encode(AttrArray[i][0]);
      a[num][1] := UTF8Encode(AttrArray[i][1]);
    end;
    FXMLWriter.WriteTag(UTF8Encode(TagName), UTF8Encode(text), a, StartNewLine, CloseTagNewLine, CheckEntity);
  finally
    setlength(a, 0);
    a := nil;
  end;
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string; StartNewLine: Boolean; CloseTagNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  WriteTag(TagName, text, FAttributes, StartNewLine, CloseTagNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string; SAttributes: TZAttributesH);
begin
  FXMLWriter.WriteTag(UTF8Encode(TagName), UTF8Encode(text), FAttributes.FAttributes);
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string; AttrArray: array of TZAttrArrayH);
begin
  WriteTag(TagName, text, AttrArray, true, false, true);
end;

procedure TZsspXMLWriterH.WriteTag(TagName: string; text: string);
begin
  WriteTag(TagName, text, FAttributes, true, false, true);
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string; SAttributes: TZAttributesH; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
begin
  FXMLWriter.WriteTagNode(UTF8Encode(TagName), SAttributes.FAttributes, StartNewLine, CloseTagNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string; AttrArray: array of TZAttrArrayH; StartNewLine: Boolean;
  CloseTagNewLine: Boolean; CheckEntity: Boolean = true);
var
  a: array of TZAttrArray;
  kol, _start: integer;
  i, num: integer;
begin
  kol := High(AttrArray);
  _start := Low(AttrArray);
  try
    setlength(a, kol - _start + 1);
    num := 0;
    for i := _start to kol do
    begin
      a[num][0] := UTF8Encode(AttrArray[i][0]);
      a[num][1] := UTF8Encode(AttrArray[i][1]);
    end;
    FXMLWriter.WriteTagNode(UTF8Encode(TagName), a, StartNewLine, CloseTagNewLine, CheckEntity);
  finally
    setlength(a, 0);
    a := nil;
  end;
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string; StartNewLine: Boolean; CloseTagNewLine: Boolean;
  CheckEntity: Boolean = true);
begin
  WriteTagNode(TagName, FAttributes, StartNewLine, CloseTagNewLine, CheckEntity);
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string; SAttributes: TZAttributesH);
begin
  WriteTagNode(TagName, SAttributes, true, false, true);
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string; AttrArray: array of TZAttrArrayH);
begin
  WriteTagNode(TagName, AttrArray, true, false, true);
end;

procedure TZsspXMLWriterH.WriteTagNode(TagName: string);
begin
  WriteTagNode(TagName, FAttributes, true, false, true);
end;

/// /////////////////////////////////////////////////////////////////////////////
/// / TZsspXMLReaderH
/// /////////////////////////////////////////////////////////////////////////////

constructor TZsspXMLReaderH.Create();
begin
  inherited;
  FAttributes := TZAttributesH.Create();
  FXMLReader := TZsspXMLReader.Create();
end;

destructor TZsspXMLReaderH.Destroy();
begin
  FreeAndNil(FXMLReader);
  FreeAndNil(FAttributes);
  inherited;
end;

procedure TZsspXMLReaderH.SetAttributesMatch(Value: Boolean);
begin
  FXMLReader.AttributesMatch := Value;
end;

function TZsspXMLReaderH.GetAttributesMatch: Boolean;
begin
  result := FXMLReader.AttributesMatch;
end;

procedure TZsspXMLReaderH.SetQuotesEqual(Value: Boolean);
begin
  FXMLReader.QuotesEqual := Value;
end;

function TZsspXMLReaderH.GetQuotesEqual(): Boolean;
begin
  result := FXMLReader.QuotesEqual;
end;

function TZsspXMLReaderH.GetAttributes(): TZAttributesH;
begin
  result := FAttributes;
end;

function TZsspXMLReaderH.GetInProcess(): Boolean;
begin
  result := FXMLReader.InProcess;
end;

function TZsspXMLReaderH.GetIsTagClosed(): Boolean;
begin
  result := (TagType = TXmlTagType.xttClosed);
end;

function TZsspXMLReaderH.GetIsTagEnd(): Boolean;
begin
  result := (TagType = TXmlTagType.xttEnd);
end;

function TZsspXMLReaderH.GetIsTagOfData: Boolean;
begin
  result := (TagType = TXmlTagType.xttStart) or (TagType = TXmlTagType.xttClosed) or (TagType = TXmlTagType.xttEnd);
end;

function TZsspXMLReaderH.GetIsTagStart(): Boolean;
begin
  result := (TagType = TXmlTagType.xttStart);
end;

function TZsspXMLReaderH.GetIsTagStartOrClosed: Boolean;
begin
  result := (TagType = TXmlTagType.xttStart) or (TagType = TXmlTagType.xttClosed);
end;

function TZsspXMLReaderH.GetRawTextTag(): string;
begin
  result := UTF8ToString(FXMLReader.RawTextTag);
end;

function TZsspXMLReaderH.GetErrorCode(): integer;
begin
  result := FXMLReader.ErrorCode;
end;

function TZsspXMLReaderH.GetIgnoreCase(): Boolean;
begin
  result := FXMLReader.IgnoreCase;
end;

function TZsspXMLReaderH.GetValue(): string;
begin
  result := UTF8ToString(FXMLReader.TagValue);
end;

function TZsspXMLReaderH.IsTagClosedByName(TagName: string): Boolean;
begin
  result := (string(FXMLReader.TagName) = TagName) and (FXMLReader.TagType = TXmlTagType.xttClosed);
end;

function TZsspXMLReaderH.IsTagEndByName(TagName: string): Boolean;
begin
  result := (string(FXMLReader.TagName) = TagName) and (FXMLReader.TagType = TXmlTagType.xttEnd);
end;

function TZsspXMLReaderH.IsTagStartByName(TagName: string): Boolean;
begin
  result := (string(FXMLReader.TagName) = TagName) and (FXMLReader.TagType = TXmlTagType.xttStart);
end;

function TZsspXMLReaderH.IsTagStartOrClosedByName(TagName: string): Boolean;
begin
  result := (string(FXMLReader.TagName) = TagName) and
    ((FXMLReader.TagType = TXmlTagType.xttStart) or (FXMLReader.TagType = TXmlTagType.xttClosed));
end;

function TZsspXMLReaderH.GetTagType(): TXmlTagType;
begin
  result := FXMLReader.TagType;
end;

function TZsspXMLReaderH.GetTagCount(): integer;
begin
  result := FXMLReader.TagCount;
end;

function TZsspXMLReaderH.GetTextBeforeTag(): string;
begin
  result := UTF8ToString(FXMLReader.TextBeforeTag);
end;

function TZsspXMLReaderH.GetTagName(): string;
begin
  result := UTF8ToString(FXMLReader.TagName);
end;

function TZsspXMLReaderH.GetTag(num: integer): string;
begin
  result := UTF8ToString(FXMLReader.Tags[num]);
end;

procedure TZsspXMLReaderH.SetMaxBufferLength(Value: integer);
begin
  FXMLReader.MaxBufferLength := Value;
end;

function TZsspXMLReaderH.GetMaxBufferLength(): integer;
begin
  result := FXMLReader.MaxBufferLength;
end;

procedure TZsspXMLReaderH.SetAttributes(Value: TZAttributesH);
begin
  if (Value <> nil) then
    FAttributes.Assign(Value);
end;

procedure TZsspXMLReaderH.SetIgnoreCase(Value: Boolean);
begin
  FXMLReader.IgnoreCase := Value;
end;

function TZsspXMLReaderH.BeginReadStream(Stream: TStream): integer;
begin
  result := FXMLReader.BeginReadStream(Stream);
end;

function TZsspXMLReaderH.ReadTag(): Boolean;
begin
  if FXMLReader.Eof then
    exit(false);
  result := FXMLReader.ReadTag();
  FAttributes.Assign(FXMLReader.Attributes);
end;

function TZsspXMLReaderH.ReadToEndTagByName(TagName: string): Boolean;
begin
  if self.Eof() then
    exit(false);

  result := not IsTagEndByName(TagName);
  if result then
    self.ReadTag();
end;

procedure TZsspXMLReaderH.EndRead();
begin
  FXMLReader.EndRead();
end;

function TZsspXMLReaderH.Eof(): Boolean;
begin
  result := FXMLReader.Eof();
end;

end.
