unit Excel4Delphi.Common;

interface

uses
  System.SysUtils, System.Types, System.Classes, Excel4Delphi, Excel4Delphi.Xml;

const
  ZE_MMinInch: real = 25.4;

type
  TTempFileStream = class(THandleStream)
  private
    FFileName: string;
  public
    constructor Create();
    destructor Destroy; override;
    property FileName: string read FFileName;
  end;

// Попытка преобразовать строку в число
function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double; overload;
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double; overload;

// Попытка преобразовать строку в boolean
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;

// заменяет все запятые на точки
function ZEFloatSeparator(st: string): string;

// Проверяет заголовки страниц, при необходимости корректирует
function ZECheckTablesTitle(var XMLSS: TZWorkBook; const SheetsNumbers: array of integer;
  const SheetsNames: array of string; out _pages: TIntegerDynArray; out _names: TStringDynArray;
  out retCount: integer): boolean;

// Очищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray; var _names: TStringDynArray);

// Переводит строку в boolean
function ZEStrToBoolean(const val: string): boolean;

// Заменяет в строке последовательности на спецсимволы
function ZEReplaceEntity(const st: string): string;

// despite formal angle datatype declaration in default "range check off" mode
// it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;

/// <summary>
/// Despite formal angle datatype declaration in default "range check off" mode it can be anywhere -32K to +32K
/// This fn brings it back into 0 .. +179 range
/// </summary>
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;

implementation

uses
  {$IFDEF MSWINDOWS}
  Winapi.windows,
  {$ENDIF}
  System.DateUtils, System.IOUtils;

function FileCreateTemp(var tempName: string): THandle;
{$IFNDEF MSWINDOWS}
var
  FS: TFileStream;
{$ENDIF}
begin
  Result := INVALID_HANDLE_VALUE;
  tempName := TPath.GetTempFileName();
  if tempName <> '' then
  begin
    {$IFDEF MSWINDOWS}
    Result := CreateFile(PChar(tempName), GENERIC_READ or GENERIC_WRITE, 0, nil, OPEN_EXISTING,
      FILE_ATTRIBUTE_TEMPORARY or FILE_FLAG_DELETE_ON_CLOSE, 0);
    {$ELSE}
    Result := FileCreate(tempName, fmCreate);
    {$ENDIF}
    if Result = INVALID_HANDLE_VALUE then
      TFile.Delete(tempName);
  end;
end;

constructor TTempFileStream.Create();
var
  FileHandle: THandle;
begin
  FileHandle := FileCreateTemp(FFileName);
  if FileHandle = INVALID_HANDLE_VALUE then
    raise Exception.Create('Не удалось создать временный файл');
  inherited Create(FileHandle);
end;

destructor TTempFileStream.Destroy;
begin
  if THandle(Handle) <> INVALID_HANDLE_VALUE then
    FileClose(Handle);
  inherited Destroy;
end;

// despite formal angle datatype declaration in default "range check off" mode
// it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range for Excel XML
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;
var
  Neg: boolean;
  A: integer;
begin
  if (value >= -90) and (value <= +90) then
    Result := value
  else
  begin (* Special values: 270; 450; -450; 180; -180; 135 *)
    Neg := value < 0; (* F, F, T, F, T, F *)
    A := Abs(value) mod 360; // 0..359       (*  270, 90, 90, 180, 180, 135  *)
    if A > 180 then
      A := A - 360; // -179..+180   (*  -90, 90, 90, 180, 180, 135 *)
    if A < 0 then
    begin
      Neg := not Neg; (* T,  -"- F, T, F, T, F *)
      A := -A; // 0..180       (*  90, -"- 90, 90, 180, 180, 135 *)
    end;
    if A > 90 then
      A := A - 180; // 91..180 -> -89..0 (* 90, 90, 90, 0, 0, -45 *)
    Result := A;
    If Neg then
      Result := -Result; (* -90, +90, -90, 0, 0, -45 *)
  end;
end;

// despite formal angle datatype declaration in default "range check off" mode
// it can be anywhere -32K to +32K
// This fn brings it back into 0 .. +180 range
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;
begin
  Result := ZENormalizeAngle90(value);
  If Result < 0 then
    Result := 90 - Result;
end;

// Заменяет в строке последовательности на спецсимволы
// INPUT
// const st: string - входящая строка
// RETURN
// string - обработанная строка
function ZEReplaceEntity(const st: string): string;
var
  s, s1: string;
  i: integer;
  isAmp: boolean;
  ch: char;

  procedure CheckS();
  begin
    s1 := UpperCase(s);
    if (s1 = '&GT;') then
      s := '>'
    else if (s1 = '&LT;') then
      s := '<'
    else if (s1 = '&AMP;') then
      s := '&'
    else if (s1 = '&APOS;') then
      s := ''''
    else if (s1 = '&QUOT;') then
      s := '"';
  end; // _checkS

begin
  s := '';
  Result := '';
  isAmp := false;
  for i := 1 to length(st) do
  begin
    ch := st[i];
    case ch of
      '&':
        begin
          if (isAmp) then
          begin
            Result := Result + s;
            s := ch;
          end
          else
          begin
            isAmp := true;
            s := ch;
          end;
        end;
      ';':
        begin
          if (isAmp) then
          begin
            s := s + ch;
            CheckS();
            Result := Result + s;
            s := '';
            isAmp := false;
          end
          else
          begin
            Result := Result + s + ch;
            s := '';
          end;
        end;
    else
      if (isAmp) then
        s := s + ch
      else
        Result := Result + ch;
    end; // case
  end; // for
  if (s > '') then
  begin
    CheckS();
    Result := Result + s;
  end;
end; // ZEReplaceEntity

// Переводит строку в boolean
// INPUT
// const val: string - переводимая строка
function ZEStrToBoolean(const val: string): boolean;
begin
  if (val = '1') or (UpperCase(val) = 'TRUE') then
    Result := true
  else
    Result := false;
end;

// Попытка преобразовать строку в boolean
// const st: string        - строка для распознавания
// valueIfError: boolean - значение, которое подставляется при ошибке преобразования
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;
begin
  Result := valueIfError;
  if (st > '') then
  begin
    if (CharInSet(st[1], ['T', 't', '1', '-'])) then
      Result := true
    else if (CharInSet(st[1], ['F', 'f', '0'])) then
      Result := false
    else
      Result := valueIfError;
  end;
end; // ZETryStrToBoolean

function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
begin
  retValue := ZETryStrToFloat(st, Result);
end;

// Попытка преобразовать строку в число
// INPUT
// const st: string        - строка
// out isOk: boolean       - если true - ошибки небыло
// valueIfError: double  - значение, которое подставляется при ошибке преобразования
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double;
var
  s: string;
  i: integer;
begin
  Result := 0;
  isOk := true;
  if (length(trim(st)) <> 0) then
  begin
    s := '';
    for i := 1 to length(st) do
      if (CharInSet(st[i], ['.', ','])) then
        s := s + FormatSettings.DecimalSeparator
      else if (st[i] <> ' ') then
        s := s + st[i];

    isOk := TryStrToFloat(s, Result);
    if (not isOk) then
      Result := valueIfError;
  end;
end; // ZETryStrToFloat

// Попытка преобразовать строку в число
// INPUT
// const st: string        - строка
// valueIfError: double  - значение, которое подставляется при ошибке преобразования
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double;
var
  s: string;
  i: integer;
begin
  Result := 0;
  if (trim(st) <> '') then
  begin
    s := '';
    for i := 1 to length(st) do
      if (CharInSet(st[i], ['.', ','])) then
        s := s + FormatSettings.DecimalSeparator
      else if (st[i] <> ' ') then
        s := s + st[i];
    try
      Result := StrToFloat(s);
    except
      Result := valueIfError;
    end;
  end;
end; // ZETryStrToFloat

// заменяет все запятые на точки
function ZEFloatSeparator(st: string): string;
var
  k: integer;
begin
  Result := '';
  for k := 1 to length(st) do
    if (st[k] = ',') then
      Result := Result + '.'
    else
      Result := Result + st[k];
end;

// Очищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray; var _names: TStringDynArray);
begin
  SetLength(_pages, 0);
  SetLength(_names, 0);
  _names := nil;
  _pages := nil;
end;

resourcestring
  DefaultSheetName = 'Sheet';

// делает уникальную строку, добавляя к строке '(num)'
// топорно, но работает
// INPUT
// var st: string - строка
// n: integer - номер
procedure ZECorrectStrForSave(var st: string; n: integer);
var
  l, i, m, num: integer;
  s: string;
begin
  if trim(st) = '' then
    st := DefaultSheetName; // behave uniformly with ZECheckTablesTitle

  l := length(st);
  if st[l] <> ')' then
    st := st + '(' + inttostr(n) + ')'
  else
  begin
    m := l;
    for i := l downto 1 do
      if st[i] = '(' then
      begin
        m := i;
        break;
      end;
    if m <> l then
    begin
      s := copy(st, m + 1, l - m - 1);
      try
        num := StrToInt(s) + 1;
      except
        num := n;
      end;
      Delete(st, m, l - m + 1);
      st := st + '(' + inttostr(num) + ')';
    end
    else
      st := st + '(' + inttostr(n) + ')';
  end;
end; // ZECorrectStrForSave

// делаем уникальные значения массивов
// INPUT
// var mas: array of string - массив со значениями
procedure ZECorrectTitles(var mas: array of string);
var
  i, num, k, _kol: integer;
  s: string;
begin
  num := 0;
  _kol := High(mas);
  while (num < _kol) do
  begin
    s := UpperCase(mas[num]);
    k := 0;
    for i := num + 1 to _kol do
    begin
      if (s = UpperCase(mas[i])) then
      begin
        inc(k);
        ZECorrectStrForSave(mas[i], k);
      end;
    end;
    inc(num);
    if k > 0 then
      num := 0;
  end;
end; // CorrectTitles

// Проверяет заголовки страниц, при необходимости корректирует
// INPUT
// var XMLSS: TZWorkBook
// const SheetsNumbers:array of integer
// const SheetsNames: array of string
// var _pages: TIntegerDynArray
// var _names: TStringDynArray
// var retCount: integer
// RETURN
// boolean - true - всё нормально, можно продолжать дальше
// false - что-то не то подсунули, дальше продолжать нельзя
function ZECheckTablesTitle(var XMLSS: TZWorkBook; const SheetsNumbers: array of integer;
  const SheetsNames: array of string; out _pages: TIntegerDynArray; out _names: TStringDynArray;
  out retCount: integer): boolean;
var
  t1, t2, i: integer;
  // '!' is allowed; ':' is not; whatever else ?
  procedure SanitizeTitle(var s: string);
  var
    i: integer;
  begin
    s := trim(s);
    for i := 1 to length(s) do
      if s[i] = ':' then
        s[i] := ';';
  end;
  function CoalesceTitle(const i: integer; const checkArray: boolean): string;
  begin
    if checkArray then
    begin
      Result := SheetsNames[i];
      SanitizeTitle(Result);
    end
    else
      Result := '';

    if Result = '' then
    begin
      Result := XMLSS.Sheets[_pages[i]].Title;
      SanitizeTitle(Result);
    end;

    if Result = '' then
      Result := DefaultSheetName + ' ' + inttostr(_pages[i] + 1);
  end;

begin
  Result := false;
  t1 := Low(SheetsNumbers);
  t2 := High(SheetsNumbers);
  retCount := 0;
  // если пришёл пустой массив SheetsNumbers - берём все страницы из Sheets
  if t1 = t2 + 1 then
  begin
    retCount := XMLSS.Sheets.Count;
    SetLength(_pages, retCount);
    for i := 0 to retCount - 1 do
      _pages[i] := i;
  end
  else
  begin
    // иначе берём страницы из массива SheetsNumbers
    for i := t1 to t2 do
    begin
      if (SheetsNumbers[i] >= 0) and (SheetsNumbers[i] < XMLSS.Sheets.Count) then
      begin
        inc(retCount);
        SetLength(_pages, retCount);
        _pages[retCount - 1] := SheetsNumbers[i];
      end;
    end;
  end;

  if (retCount <= 0) then
    exit;

  // названия страниц
// t1 :=  Low(SheetsNames); // we anyway assume later that Low(_names) == t1 - then let us just skip this.
  t2 := High(SheetsNames);
  SetLength(_names, retCount);
// if t1 = t2 + 1 then
// begin
// for i := 0 to retCount - 1 do
// begin
// _names[i] := XMLSS.Sheets[_pages[i]].Title;
// if trim(_names[i]) = '' then _names[i] := 'list';
// end;
// end else
// begin
// if (t2 > retCount) then
// t2 := retCount - 1;
// for i := t1 to t2 do
// _names[i] := SheetsNames[i];
// if (t2 < retCount) then
// for i := t2 + 1 to retCount - 1 do
// begin
// _names[i] := XMLSS.Sheets[_pages[i]].Title;
// if trim(_names[i]) = '' then _names[i] := 'list';
// end;
// end;
  for i := Low(_names) to High(_names) do
  begin
    _names[i] := CoalesceTitle(i, i <= t2);
  end;

  ZECorrectTitles(_names);
  Result := true;
end; // ZECheckTablesTitle

end.
