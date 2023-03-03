unit Excel4DelphiTest.Formula;

interface

uses DUnitX.TestFramework;

type
  [TestFixture]
  TTestFormulaObject = class
  public
    [Test]
    [TestCase('GetColAddres 1', '222')]
    procedure TestGetColAddres(value: integer);

    [Test]
    [TestCase('GetColIndex 1', 'HO')]
    procedure TestGetColIndex(value: string);

    [Test]
    [TestCase('GetCellCoords 1', 'B1')]
    procedure TestGetCellCoords(value: string);

    [Test]
    [TestCase('CellRange 1', 'B1:C2')]
    procedure TestCellRange(value: string);
  end;

implementation

uses System.Variants, System.SysUtils, Soap.XSBuiltIns
  , Excel4Delphi
  , Excel4Delphi.Formula
  ;

{ TTestFormulaObject }

procedure TTestFormulaObject.TestCellRange(value: string);
var left, top, right, bottom: integer;
begin
  TZEFormula.GetCellRange(value, left, top, right, bottom);
  Assert.AreEqual(left, 1);
  Assert.AreEqual(top, 0);
  Assert.AreEqual(right, 2);
  Assert.AreEqual(bottom, 1);
end;

procedure TTestFormulaObject.TestGetCellCoords(value: string);
var left, top: integer;
begin
  TZEFormula.GetCellCoords(value, left, top);
  Assert.AreEqual(left, 1);
  Assert.AreEqual(top, 0);
end;

procedure TTestFormulaObject.TestGetColAddres(value: integer);
var res: string;
begin
  res := TZEFormula.GetColAddres(value);
  Assert.AreEqual(res, 'HO');
end;

procedure TTestFormulaObject.TestGetColIndex(value: string);
var res: Integer;
begin
  res := TZEFormula.GetColIndex(value);
  Assert.AreEqual(res, 222);
end;

initialization
  TDUnitX.RegisterTestFixture(TTestFormulaObject);

end.
