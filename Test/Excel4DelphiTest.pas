unit Excel4DelphiTest;

interface

uses DUnitX.TestFramework
  , Excel4Delphi
  , Excel4Delphi.Xml
  , Excel4Delphi.Stream
  ;

const FILE_1 = '../../Examples/Test1.xlsx';

type
  [TestFixture]
  TTestObject = class
  public
    [Test]
    [TestCase('TestDropColumnAndRows 1', FILE_1)]
    procedure TestDropColumnAndRows(fileName: string);
  end;

implementation

{ TTestObject }

procedure TTestObject.TestDropColumnAndRows(fileName: string);
var book: TZWorkBook;
begin
  book := TZWorkBook.Create(nil);
  try
    book.LoadFromFile(fileName);

    Assert.AreEqual(book.Sheets[0].MergeCells.Count, 6, 'Merged count');

    book.Sheets[0].DeleteColumns(9, 1);
    Assert.AreEqual(book.Sheets[0].MergeCells.Count, 5, 'Merged count');

    book.Sheets[0].DeleteColumns(6, 1);
    Assert.AreEqual(book.Sheets[0].MergeCells.Count, 5, 'Merged count');

    book.Sheets[0].DeleteRows(4, 1);
    Assert.AreEqual(book.Sheets[0].MergeCells.Count, 4, 'Merged count');

    book.SaveToFile('test_del_rows.xlsx');
  finally
    book.Free();
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TTestObject);

end.
