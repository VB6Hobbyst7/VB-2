unit m.test;

interface

uses
  System.Classes, System.SysUtils
  ;

type
  TTest = class
  private
  public
    Date: TDateTime;
    TestNum,
    BatchNum: Integer;
    Operator: String;
    constructor Create(const ADate: TDateTime; const ATestNum, ABatchNum: Integer; const AOperator: String);
  end;

implementation

{ TTest }

constructor TTest.Create(const ADate: TDateTime; const ATestNum, ABatchNum: Integer; const AOperator: String);
begin
  Date := ADate;
  TestNum := ATestNum;
  BatchNum := ABatchNum;
  Operator := AOperator;
end;

end.
