program d1;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  Excel4Delphi,
  Excel4Delphi.Stream,
  System.SysUtils;

procedure CreateNewBook;
// Creating new workbook
var
  workBook: TZWorkBook;
begin
  workBook := TZWorkBook.Create(nil);
  try
    workBook.Sheets.Add('My sheet');
    workBook.Sheets[0].ColCount := 10;
    workBook.Sheets[0].RowCount := 10;
    workBook.Sheets[0].CellRef['A', 0].AsString := 'Hello';
    workBook.Sheets[0].RangeRef['A', 0, 'B', 2].Merge();
    workBook.SaveToFile('file.xlsx');
  finally
    workBook.Free();
  end;
end;

begin
  try
    { TODO -oUser -cConsole Main : Insert code here }
    CreateNewBook;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;

end.
