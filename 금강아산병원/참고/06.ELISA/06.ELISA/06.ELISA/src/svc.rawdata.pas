unit svc.rawdata;

interface

uses
  m.rawdata,
  System.Classes, System.SysUtils, System.UITypes, Spring
  ;

type
  TsvcRawdata = class(TRawdataContainer);
//  private
//    function GetCellColors(c, r: Integer): TColor;
//    function GetCellMaterial(c, r: Integer): String;
//    function GetCellStd(c, r: Integer): Boolean;
//    function GetCellFontStyles(c, r: Integer): TFontStyles;
//  public
//    procedure AssignRawdata(const ASrc: TStringList);
//    procedure AssignMaterial(const ACriteriaMaterial: TCriteriaMaterial; const ASrc: TStringList);
//
//    property CellColors[c, r: Integer]: TColor read GetCellColors;
//    property CellMaterials[c, r: Integer]: String read GetCellMaterial;
//    property CellStd[c, r: Integer]: Boolean read GetCellStd;
//    property CellFontStyles[c, r: Integer]: TFontStyles read GetCellFontStyles;
//  end;

implementation

uses
  svc, m.elisa,

  System.Math, CodeSiteLogging, mCodeSiteHelper, System.StrUtils
  ;

{ TSvcQuantiFileFmt }

//procedure TsvcRawdata.AssignMaterial(const ACriteriaMaterial: TCriteriaMaterial; const ASrc: TStringList);
//var
//  LCell: TSectionCell;
//  i, c, r: Integer;
//  LRows: TStringList;
//  LCols: TStringList;
//  LStdRange: Nullable<Boolean>;
//  LPoint: TCellPoint;
//begin
//  i := -1;
//  FSectionCells.Clear;
//  LRows := ASrc;
//  LCols := TStringList.Create;
//  try
//    LCols.StrictDelimiter := True;
//    for r := 0 to LRows.Count -1 do
//    begin
//      LCols.CommaText := LRows[r];
//      for c := 0 to LCols.Count -1 do
//      begin
//        case ACriteriaMaterial of
//          cmNil_Antigen:
//          begin
//            LStdRange := TDefaultMaterial.IsM2StdRange(c, r);
//            if LStdRange.Value then
//              i := TDefaultMaterial.NM2Len + IfThen(c > 5, 1)
//            else
//              i := (c * 4) + (r div 2) - IfThen((c >= 5), 2 + IfThen(c > 5, 2));
//            if not FSectionCells.ContainsKey(i) then
//              FSectionCells.Add(i, TSectionCell.Create(i,
//                IfThen(LStdRange, cmStandard, cmNil_Antigen),
//                IfThen(LStdRange, btStandard, btCriteria)));
//            LCell := FSectionCells[i];
//            LPoint := LCell.AddPoint(c, r, LCols[c]);
//            LCell.ID := IfThen(LStdRange, Format('Std %d', [r mod 4 +1]), Format('ID %d', [i +1]));
//            FSectionCells[i] := LCell;
//          end;
//
//          cmNil_Antigen_Mitogen:
//          begin
//            LStdRange := TDefaultMaterial.IsM3StdRange(c, r);
//            if LStdRange.Value then
//              i := TDefaultMaterial.NM3Len + IfThen(c >= 3, 1 + IfThen(c >= 4, 1 + IfThen(c >= 5, 1)))
//            else
//              i := IfThen(c < 3, (c div 3), 4 + 8 * ((c div 3) -1)) + r;
//            if not FSectionCells.ContainsKey(i) then
//              FSectionCells.Add(i, TSectionCell.Create(i,
//                IfThen(LStdRange, cmStandard, cmNil_Antigen_Mitogen),
//                IfThen(LStdRange, btStandard, btCriteria)));
//            LCell := FSectionCells[i];
//            LPoint := LCell.AddPoint(c, r, LCols[c]);
//            LCell.ID := IfThen(LStdRange, Format('Std %d', [r mod 4 +1]), Format('ID %d', [i +1]));
//            FSectionCells[i] := LCell;
//          end;
//        end;
//        if (i = -1) or not LStdRange.HasValue then
//          raise Exception.Create('Exception occured when process the ' + ACriteriaMaterial.ToString);
//
//        FCellMat[c, r] := i;
//        FCellColorIdx[c, r] := IfThen(not LStdRange.Value, i, -1);
//        FStdColorIdx[c, r] := IfThen(LStdRange, r mod 4, -1);
//        FCellPoints[c, r] := LPoint;
//        FCriteriaMaterial[c, r] := LCell.CriteriaMaterial;
//        FBlockTypes[c, r] := LCell.BlockType;
//        FIDs[c, r] := LCell.ID;
//        FValues[c, r] := LPoint.Value;
////        FIds[c, r] := Format('[%d, %d] %d', [c, r, i]);
////        FValues[c, r] := Format('%d, %d', [FCellColorIdx[c, r], FStdColorIdx[c, r]]);
//      end;
//    end;
//  finally
//    FreeAndNil(LCols);
//  end;
//end;
//
//procedure TsvcRawdata.AssignRawdata(const ASrc: TStringList);
//var
//  LCell: TSectionCell;
//  i, c, r: Integer;
//  LRows: TStringList;
//  LCols: TStringList;
//  LPoint: TCellPoint;
//begin
//  FSectionCells.Clear;
//  LRows := ASrc;
//  LCols := TStringList.Create;
//  try
//    i := 0;
//    LCols.StrictDelimiter := True;
//    for r := 0 to LRows.Count -1 do
//    begin
//      LCols.CommaText := LRows[r];
//      for c := 0 to LCols.Count -1 do
//      begin
//        if not FSectionCells.ContainsKey(i) then
//          FSectionCells.Add(i, TSectionCell.Empty(i));
//        LCell := FSectionCells[i];
//        LPoint := LCell.AddPoint(c, r, LCols[c]);
//        FCellMat[c, r] := i;
//        FCellColorIdx[c, r] := -1;
//        FStdColorIdx[c, r] := -1;
//        FCellPoints[c, r] := LPoint;
//        FCriteriaMaterial[c, r] := LCell.CriteriaMaterial;
//        FBlockTypes[c, r] := LCell.BlockType;
//        FIDs[c, r] := LCell.ID;
//        FValues[c, r] := LPoint.Value;
//        Inc(i);
//      end;
//    end;
//  finally
//    FreeAndNil(LCols);
//  end;
//end;
//
//function TsvcRawdata.GetCellColors(c, r: Integer): TColor;
//var
//  LCell: TSectionCell;
//begin
//  Result := $00000000;
//  LCell := CellMat[c, r];
//  case LCell.BlockType of
//    btCriteria:
//      case LCell.CriteriaMaterial of
//        cmNil_Antigen        : Result := TDefaultMaterial.clM2[FCellColorIdx[c,r]];
//        cmNil_Antigen_Mitogen: Result := TDefaultMaterial.clM3[FCellColorIdx[c,r]];
//      end;
//
//    btStandard:
//      Result := TDefaultMaterial.clStdMaterial[FStdColorIdx[c, r]];
//
//    btEmpty:
//      Result := $00FFFFFF;
//  end;
//end;
//
//function TsvcRawdata.GetCellFontStyles(c, r: Integer): TFontStyles;
//begin
//  Result := [];
//  case CellMat[c, r].BlockType of
//    btStandard: Result := [TFontStyle.fsBold, TFontStyle.fsUnderline];
//  end;
//end;
//
//function TsvcRawdata.GetCellMaterial(c, r: Integer): String;
//begin
//  case CellMat[c, r].BlockType of
//    btCriteria,
//    btStandard: Result := CellMat[c, r].MaterialText[FCellPoints[c, r].Idx];
//    btEmpty: Result := ' ';
//  end;
//end;
//
//function TsvcRawdata.GetCellStd(c, r: Integer): Boolean;
//begin
//  Result := CellMat[c, r].CriteriaMaterial = cmStandard;
//end;

end.
