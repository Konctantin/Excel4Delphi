unit Excel4DelphiTest.Converters;

interface

uses DUnitX.TestFramework;

type
  [TestFixture]
  TTestConvertrsObject = class
  public
    [Test]
    [TestCase('String to TDate 1', '12.10.2005')]
    [TestCase('String to TDate 2', '12-10-2005')]
    [TestCase('String to TDate 3', '12/10/2005')]
    [TestCase('String to TDate 4', '38637')]
    [TestCase('String to TDate 5', '20051012')]
    procedure TestConvertStringToDate(value: string);

    [Test]
    [TestCase('String to TDateTime 1', '12.10.2005 22:55:16')]
    [TestCase('String to TDateTime 2', '12-10-2005 22:55:16')]
    [TestCase('String to TDateTime 3', '12/10/2005 22:55:16')]
    //[TestCase('String to TDate 4', '38637,9550462963')]
    [TestCase('String to TDateTime 5', '38637.9550462963')]
    [TestCase('String to TDateTime 6', '2005-10-12T22:55:16')]
    procedure TestConvertStringToDateTime(value: string);
  end;

implementation

uses System.Variants, System.SysUtils, Soap.XSBuiltIns
  , Excel4Delphi
  , Excel4Delphi.Common
  ;

{ TTestConvertrsObject }

procedure TTestConvertrsObject.TestConvertStringToDate(value: string);
var dt, dt2: TDateTime;
begin
  dt := 0;
  dt2 := EncodeDate(2005, 10, 12);
  Assert.AreEqual(ZETryParseDateTime(value, dt), true, 'String to TDate: bool');
  Assert.AreEqual(dt, dt2, 'String to TDate');
end;

procedure TTestConvertrsObject.TestConvertStringToDateTime(value: string);
var dt, dt2: TDateTime;
begin
  dt := 0;
  dt2 := EncodeDate(2005, 10, 12) + EncodeTime(22, 55, 16, 0);
  Assert.AreEqual(ZETryParseDateTime(value, dt), true, 'String to TDateTime: bool');
  Assert.AreEqual(dt, dt2, 'String to TDateTime');
end;

initialization
  TDUnitX.RegisterTestFixture(TTestConvertrsObject);

end.
