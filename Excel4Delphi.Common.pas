unit Excel4Delphi.Common;

interface

uses
  SysUtils,
  Types,
  Classes,
  Excel4Delphi,
  Excel4Delphi.Xml;

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

//Попытка преобразовать строку в число
function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double; overload;
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double; overload;
/// <summary>
/// Try convert string (any formats) to datetime
/// </summary>
function ZETryParseDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;
/// <summary>
/// Try convert string (YYYY-MM-DDTHH:MM:SS[.mmm]) to datetime
/// </summary>
function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;

//Попытка преобразовать строку в boolean
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;

//заменяет все запятые на точки
function ZEFloatSeparator(st: string): string;

//Проверяет заголовки страниц, при необходимости корректирует
function ZECheckTablesTitle(var XMLSS: TZWorkBook; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; out _pages: TIntegerDynArray;
                            out _names: TStringDynArray; out retCount: integer): boolean;

//Очищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);

//Переводит строку в boolean
function ZEStrToBoolean(const val: string): boolean;

/// <summary>
/// Decodes the input HTML data and returns the decoded HTML data.
/// </summary>
function ZEReplaceEntity(const text: string): string;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;

/// <summary>
/// Despite formal angle datatype declaration in default "range check off" mode it can be anywhere -32K to +32K
/// This fn brings it back into 0 .. +179 range
/// </summary>
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;

implementation

uses
  DateUtils, IOUtils, Winapi.Windows, Variants, VarUtils, NetEncoding;

function FileCreateTemp(var tempName: string): THandle;
begin
  Result := INVALID_HANDLE_VALUE;
  TempName := TPath.GetTempFileName();
  if TempName <> '' then begin
    Result := CreateFile(PChar(TempName), GENERIC_READ or GENERIC_WRITE, 0, nil,
      OPEN_EXISTING, FILE_ATTRIBUTE_TEMPORARY or FILE_FLAG_DELETE_ON_CLOSE, 0);
    if Result = INVALID_HANDLE_VALUE then
      TFile.Delete(TempName);
  end;
end;

constructor TTempFileStream.Create();
var FileHandle: THandle;
begin
  FileHandle := FileCreateTemp(FFileName);
  if FileHandle = INVALID_HANDLE_VALUE then
    raise  Exception.Create('The file cannot be created.');
  inherited Create(FileHandle);
end;

destructor TTempFileStream.Destroy;
begin
  if THandle(Handle) <> INVALID_HANDLE_VALUE then
    FileClose(Handle);
  inherited Destroy;
end;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into -90 .. +90 range for Excel XML
function ZENormalizeAngle90(const value: TZCellTextRotate): integer;
var Neg: boolean; A: integer;
begin
   if (value >= -90) and (value <= +90)
      then Result := value
   else begin                             (* Special values: 270; 450; -450; 180; -180; 135 *)
      Neg := Value < 0;                             (*  F, F, T, F, T, F *)
      A := Abs(value) mod 360;      // 0..359       (*  270, 90, 90, 180, 180, 135  *)
      if A > 180 then A := A - 360; // -179..+180   (*  -90, 90, 90, 180, 180, 135 *)
      if A < 0 then begin
         Neg := not Neg;                            (*  T,  -"- F, T, F, T, F  *)
         A := - A;                  // 0..180       (*  90, -"- 90, 90, 180, 180, 135 *)
      end;
      if A > 90 then A := A - 180; // 91..180 -> -89..0 (* 90, 90, 90, 0, 0, -45 *)
      Result := A;
      If Neg then Result := - Result;               (* -90, +90, -90, 0, 0, -45 *)
   end;
end;

// despite formal angle datatype declaration in default "range check off" mode
//    it can be anywhere -32K to +32K
// This fn brings it back into 0 .. +180 range
function ZENormalizeAngle180(const value: TZCellTextRotate): integer;
begin
  Result := ZENormalizeAngle90(value);
  If Result < 0 then Result := 90 - Result;
end;


function ZEReplaceEntity(const text: string): string;
begin
  result := TNetEncoding.HTML.Decode(text);
end;

//Переводит строку в boolean
//INPUT
//  const val: string - переводимая строка
function ZEStrToBoolean(const val: string): boolean;
begin
  if (val = '1') or (UpperCase(val) = 'TRUE')  then
    result := true
  else
    result := false;
end;

//Попытка преобразовать строку в boolean
//  const st: string        - строка для распознавания
//    valueIfError: boolean - значение, которое подставляется при ошибке преобразования
function ZETryStrToBoolean(const st: string; valueIfError: boolean = false): boolean;
begin
  result := valueIfError;
  if (st > '') then begin
    if CharInSet(st[1], ['T', 't', '1', '-']) then
      result := true
    else
    if CharInSet(st[1], ['F', 'f', '0']) then
      result := false
    else
      result := valueIfError;
  end;
end; //ZETryStrToBoolean

function ZEIsTryStrToFloat(const st: string; out retValue: double): boolean;
begin
  retValue := ZETryStrToFloat(st, Result);
end;

//Попытка преобразовать строку в число
//INPUT
//  const st: string        - строка
//  out isOk: boolean       - если true - ошибки небыло
//    valueIfError: double  - значение, которое подставляется при ошибке преобразования
function ZETryStrToFloat(const st: string; out isOk: boolean; valueIfError: double = 0): double;
var
  s: string;
  i: integer;
  n: integer;
  c: integer;
  done: boolean;
  hasSep: boolean;
begin
  Result := valueIfError;
  SetLength(s, Length(st));
  n := 0;
  done := false;
  hasSep := false;
  isOk := false;
  for i := 1 to Length(st) do
  begin
    c := Ord(st[i]);
    if c = 32 then
    begin
      if n > 0 then
        done := true;
      continue;
    end
    else if (c >= 48) and (c <= 57) then
    begin
      if done then
        exit; // Если уже был пробел и опять пошли цифры, то это ошибка, поэтому выходим.
      Inc(n);
      s[n] := Char(c);
    end
    else if c in [44, 46] then
    begin
      if done or hasSep then
        exit; // Если уже был пробел или разделитель и опять попался разделитель, то это ошибка, поэтому выходим.
      Inc(n);
      s[n] := FormatSettings.DecimalSeparator;
      hasSep := true;
    end;
  end;
  if n > 0 then
  begin
    SetLength(s, n);
    isOk := TryStrToFloat(s, Result);
    if (not isOk) then
      Result := valueIfError;
  end;
end; //ZETryStrToFloat

//Попытка преобразовать строку в число
//INPUT
//  const st: string        - строка
//    valueIfError: double  - значение, которое подставляется при ошибке преобразования
function ZETryStrToFloat(const st: string; valueIfError: double = 0): double;
var isOk: boolean;
begin
  Result := ZETryStrToFloat(st, isOk, valueIfError);
end; //ZETryStrToFloat

//заменяет все запятые на точки
function ZEFloatSeparator(st: string): string;
var k: integer;
begin
  result := '';
  for k := 1 to length(st) do
    if (st[k] = ',') then
      result := result + '.'
    else
      result := result + st[k];
end;

function ZETryParseDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;
var err: Integer;
  doubleVal: Double;
  LResult: HResult;
begin
  result := false;

  if TryZEStrToDateTime(AStrDateTime, retDateTime) then
    exit(true);

  retDateTime := StrToDateDef(AStrDateTime, 0);
  if retDateTime > 0 then
    exit(true);

  retDateTime := StrToDateTimeDef(AStrDateTime, 0);
  if retDateTime > 0 then
    exit(true);

  Val(AStrDateTime, retDateTime, err);
  if err = 0 then
    exit(true);

  LResult := VarDateFromStr(PWideChar(AStrDateTime), VAR_LOCALE_USER_DEFAULT, 0, retDateTime);
  if LResult = VAR_OK then
    exit(true);

  if LResult = VAR_TYPEMISMATCH then begin
    if not TryStrToDate(PWideChar(AStrDateTime), retDateTime) then begin
      if TryStrToFloat(AStrDateTime, doubleVal) then begin
        retDateTime := doubleVal;
        exit(true);
      end;
    end;
  end;

  Val(AStrDateTime, doubleVal, err);
  if err = 0 then begin
    retDateTime := doubleVal;
    exit(true);
  end;

  retDateTime := 0;
end;

function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;
var a: array [0..10] of word;
  i, l: integer;
  s, ss: string;
  count: integer;
  ch: char;
  datedelimeters: integer;
  istimesign: boolean;
  timedelimeters: integer;
  istimezone: boolean;
  lastdateindex: integer;
  tmp: integer;
  msindex: integer;
  tzindex: integer;
  timezonemul: integer;
  _ms: word;

  function TryAddToArray(const ST: string): boolean;
  begin
    if (count > 10) then begin
      Result := false;
      exit;
    end;
    Result := TryStrToInt(ST, tmp);
    if (Result) then begin
      a[Count] := word(tmp);
      inc(Count);
    end
  end;

  procedure _CheckDigits();
  var _l: integer;
  begin
    _l := length(s);
    if (_l > 0) then begin
      if (_l > 4) then begin//it is not good
        if (istimesign) then begin
          // HHMMSS?
          if (_l = 6) then begin
            ss := copy(s, 1, 2);
            if (TryAddToArray(ss)) then begin
              ss := copy(s, 3, 2);
              if (TryAddToArray(ss)) then begin
                ss := copy(s, 5, 2);
                if (not TryAddToArray(ss)) then
                  Result := false;
              end else
                Result := false;
            end else
              Result := false
          end else
            Result := false;
        end else begin
          // YYYYMMDD?
          if (_l = 8) then begin
            ss := copy(s, 1, 4);
            if (not TryAddToArray(ss)) then
              Result := false
            else begin
              ss := copy(s, 5, 2);
              if (not TryAddToArray(ss)) then
                Result := false
              else begin
                ss := copy(s, 7, 2);
                if (not TryAddToArray(ss)) then
                  Result := false;
              end;
            end;
          end else
            Result := false;
        end;
      end else
        if (not TryAddToArray(s)) then
          Result := false;
    end; //if
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
    if (count > 0) then
      lastdateindex := count;

    _CheckDigits();
  end;

  procedure _processTimeDelimiter();
  begin
    _CheckDigits();
    inc(timedelimeters)
  end;

  procedure _processDateDelimiter();
  begin
    _CheckDigits();
    if (istimesign) then begin
      tzindex := count;
      istimezone := true;
      timezonemul := -1;
    end else
      inc(datedelimeters);
  end;

  procedure _processMSDelimiter();
  begin
    _CheckDigits();
    msindex := count;
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

  function _TryGetDateTime(): boolean;
  var _time, _date: TDateTime;
  begin
    //Result := true;
    if (msindex >= 0) then
      _ms := a[msindex];
    if (lastdateindex >= 0) then begin
      Result := TryEncodeDate(a[0], a[1], a[2], _date);
      if (Result) then begin
        Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, _time);
        if (Result) then
          retDateTime := _date + _time;
      end;
    end else
      Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, retDateTime);
  end;

  function _TryGetDate(): boolean;
  begin
    if (datedelimeters = 0) and (timedelimeters >= 2) then begin
      if (msindex >= 0) then
        _ms := a[msindex];
      result := TryEncodeTime(a[0], a[1], a[2], _ms, retDateTime);
    end else if (count >= 3) then
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

  l := length(AStrDateTime);
  s := '';
  count := 0;
  for i := 1 to l do begin
    ch := AStrDateTime[i];
    case (ch) of
      '0'..'9': _processDigit();
      't', 'T': _processTimeSign();
      '-':      _processDateDelimiter();
      ':':      _processTimeDelimiter();
      '.', ',': _processMSDelimiter();
      'z', 'Z': _processTimeZoneSign();
      '+':      _processTimeZonePlus();
    end;
    if (not Result) then
      break
  end;

  if (Result and (s <> '')) then
    _CheckDigits();

  if (Result) then begin
    if (istimesign) then
      Result := _TryGetDateTime()
    else
      Result := _TryGetDate();
  end;
end; //TryZEStrToDateTime

//Очищает массивы
procedure ZESClearArrays(var _pages: TIntegerDynArray;  var _names: TStringDynArray);
begin
  SetLength(_pages, 0);
  SetLength(_names, 0);
  _names := nil;
  _pages := nil;
end;

resourcestring DefaultSheetName = 'Sheet';

//делает уникальную строку, добавляя к строке '(num)'
//топорно, но работает
//INPUT
//  var st: string - строка
//      n: integer - номер
procedure ZECorrectStrForSave(var st: string; n: integer);
var l, i, m, num: integer; s: string;
begin
  if Trim(st) = '' then
     st := DefaultSheetName;  // behave uniformly with ZECheckTablesTitle

  l := length(st);
  if st[l] <> ')' then
    st := st + '(' + inttostr(n) + ')'
  else
  begin
    m := l;
    for i := l downto 1 do
    if st[i] = '(' then begin
      m := i;
      break;
    end;
    if m <> l then begin
      s := copy(st, m+1, l-m - 1);
      try
        num := StrToInt(s) + 1;
      except
        num := n;
      end;
      delete(st, m, l-m + 1);
      st := st + '(' + inttostr(num) + ')';
    end else
      st := st + '(' + inttostr(n) + ')';
  end;
end; //ZECorrectStrForSave

//делаем уникальные значения массивов
//INPUT
//  var mas: array of string - массив со значениями
procedure ZECorrectTitles(var mas: array of string);
var i, num, k, _kol: integer; s: string;
begin
  num := 0;
  _kol := High(mas);
  while (num < _kol) do begin
    s := UpperCase(mas[num]);
    k := 0;
    for i := num + 1 to _kol do begin
      if (s = UpperCase(mas[i])) then begin
        inc(k);
        ZECorrectStrForSave(mas[i], k);
      end;
    end;
    inc(num);
    if k > 0 then num := 0;
  end;
end; //CorrectTitles

//Проверяет заголовки страниц, при необходимости корректирует
//INPUT
//  var XMLSS: TZWorkBook
//  const SheetsNumbers:array of integer
//  const SheetsNames: array of string
//  var _pages: TIntegerDynArray
//  var _names: TStringDynArray
//  var retCount: integer
//RETURN
//      boolean - true - всё нормально, можно продолжать дальше
//                false - что-то не то подсунули, дальше продолжать нельзя
function ZECheckTablesTitle(var XMLSS: TZWorkBook; const SheetsNumbers:array of integer;
                            const SheetsNames: array of string; out _pages: TIntegerDynArray;
                            out _names: TStringDynArray; out retCount: integer): boolean;
var t1, t2, i: integer;
  // '!' is allowed; ':' is not; whatever else ?
  procedure SanitizeTitle(var s: string);   var i: integer;
  begin
    s := Trim(s);
    for i := 1 to length(s) do
       if s[i] = ':' then s[i] := ';';
  end;
  function CoalesceTitle(const i: integer; const checkArray: boolean): string;
  begin
    if checkArray then begin
       Result := SheetsNames[i];
       SanitizeTitle(Result);
    end else
       Result := '';

    if Result = '' then begin
       Result := XMLSS.Sheets[_pages[i]].Title;
       SanitizeTitle(Result);
    end;

    if Result = '' then
       Result := DefaultSheetName + ' ' + IntToStr(_pages[i] + 1);
  end;

begin
  result := false;
  t1 :=  Low(SheetsNumbers);
  t2 := High(SheetsNumbers);
  retCount := 0;
  //если пришёл пустой массив SheetsNumbers - берём все страницы из Sheets
  if t1 = t2 + 1 then
  begin
    retCount := XMLSS.Sheets.Count;
    setlength(_pages, retCount);
    for i := 0 to retCount - 1 do
      _pages[i] := i;
  end else
  begin
    //иначе берём страницы из массива SheetsNumbers
    for i := t1 to t2 do
    begin
      if (SheetsNumbers[i] >= 0) and (SheetsNumbers[i] < XMLSS.Sheets.Count) then
      begin
        inc(retCount);
        setlength(_pages, retCount);
        _pages[retCount-1] := SheetsNumbers[i];
      end;
    end;
  end;

  if (retCount <= 0) then
    exit;

  //названия страниц
//  t1 :=  Low(SheetsNames); // we anyway assume later that Low(_names) == t1 - then let us just skip this.
  t2 := High(SheetsNames);
  setlength(_names, retCount);
//  if t1 = t2 + 1 then
//  begin
//    for i := 0 to retCount - 1 do
//    begin
//      _names[i] := XMLSS.Sheets[_pages[i]].Title;
//      if trim(_names[i]) = '' then _names[i] := 'list';
//    end;
//  end else
//  begin
//    if (t2 > retCount) then
//      t2 := retCount - 1;
//    for i := t1 to t2 do
//      _names[i] := SheetsNames[i];
//    if (t2 < retCount) then
//    for i := t2 + 1 to retCount - 1 do
//    begin
//      _names[i] := XMLSS.Sheets[_pages[i]].Title;
//      if trim(_names[i]) = '' then _names[i] := 'list';
//    end;
//  end;
  for i := Low(_names) to High(_names) do begin
      _names[i] := CoalesceTitle(i, i <= t2);
  end;


  ZECorrectTitles(_names);
  result := true;
end; //ZECheckTablesTitle

end.
