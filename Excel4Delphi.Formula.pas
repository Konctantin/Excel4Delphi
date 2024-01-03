unit Excel4Delphi.Formula;

interface

type
  TZEFormula = class
    /// <summary>
    /// Return absolute AA styled cell address.
    /// </summary>
    class function GetColAddres(const ColumnIndex: integer; FromZero: boolean = true): string; static;
    /// <summary>
    /// Extracts column index from it's AA address.
    /// </summary>
    class function GetColIndex(ColumnAdress: string; FromZero: boolean = true): integer; static;
    /// <summary>
    /// Extracts cell column and row from A1 styled address.
    /// </summary>
    class function GetCellCoords(const cell: string; out column, row: integer): boolean; static;
    /// <summary>
    /// Extracts range top, left right and bottom from 'А1:B2' based string.
    /// If string is 'A1' format then right=left, bottom=top.
    /// </summary>
    class function GetCellRange(const range: string; out left, top, right, bottom: integer): boolean; static;
  end;

implementation

uses
  System.SysUtils,
  System.Math;

const
  CHARS: array [0..25] of char = (
  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
  'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
  );

{ TZEFormula }

class function TZEFormula.GetColAddres(const columnIndex: integer; FromZero: boolean): string;
var t, n: integer; s: string;
begin
  t := columnIndex;
  if (not FromZero) then
    dec(t);
  result := '';
  s := '';
  while t >= 0 do begin
    n := t mod 26;
    t := (t div 26) - 1;
    s := s + CHARS[n];
  end;
  for t := length(s) downto 1 do
    result := result + s[t];
end;

class function TZEFormula.GetColIndex(ColumnAdress: string; FromZero: boolean): integer;
var i: integer; num, t, s: integer;
begin
  result := -1;
  num := 0;
  ColumnAdress := UpperCase(ColumnAdress);
  s := 1;
  for i := length(ColumnAdress) downto 1 do begin
    if not CharInSet(ColumnAdress[I], ['A'..'Z']) then
      continue;
    t := ord(ColumnAdress[i]) - ord('A');
    num := num + (t + 1) * s;
    s := s * 26;
    if (s < 0) or (num < 0) then
      exit;
  end;
  result := num;
  if FromZero then
    result := result - 1;
end;

class function TZEFormula.GetCellCoords(const cell: string; out column, row: integer): boolean;
var right, bottom: integer;
begin
  result := GetCellRange(cell, column, row, right, bottom);
end;

class function TZEFormula.GetCellRange(const range: string; out left, top, right, bottom: integer): boolean;
var i, p: integer;
  cols, rows: TArray<string>;
begin
  left := -1; top := -1; right := -1; bottom := -1;
  cols := ['',''];
  rows := ['',''];
  result := true;
  p := 0;
  for i := 1 to length(range) do
    case range[i] of
      'A'..'Z', 'a'..'z':
        cols[p] := cols[p] + range[i];
      '0'..'9':
        begin
          if cols[p] = '' then
            exit(false);
          rows[p] := rows[p] + range[i];
        end;
      ':':
        if p = 0 then
          inc(p)
        else
          exit(false);
    else
      exit(false);
    end;

  if not TryStrToInt(rows[0], top) then
    exit(false);

  left := GetColIndex(cols[0]);
  if left < 0 then
    exit(false);

  bottom := Max(StrToIntDef(rows[1], -1), top);
  right  := Max(GetColIndex(cols[1]), left);
  dec(top);
  dec(bottom);
end;

end.
