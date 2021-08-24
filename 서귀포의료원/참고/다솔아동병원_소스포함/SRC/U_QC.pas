unit U_QC;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ADODB, StdCtrls, DB, DBGrids, Grids, BaseGrid, AdvGrid,
  AdvStringGrid_Ex, TeEngine, Series, TeeProcs, Chart, ComCtrls, ExtCtrls,
  Buttons, Menus, StringLib, Spin, U_IFClass;

type
  TF_QC = class(TForm)
    Panel7: TPanel;
    dtpF: TDateTimePicker;
    dtpT: TDateTimePicker;
    Panel4: TPanel;
    cmbxENM: TComboBox;
    Panel5: TPanel;
    Panel6: TPanel;
    cmbxLot: TComboBox;
    Panel10: TPanel;
    Chart1: TChart;
    Series1: TLineSeries;
    Series2: TLineSeries;
    Series3: TLineSeries;
    Series4: TLineSeries;
    Series5: TLineSeries;
    Series6: TLineSeries;
    Series7: TLineSeries;
    Series8: TLineSeries;
    Panel11: TPanel;
    Panel1: TPanel;
    Panel13: TPanel;
    gdList: TAdvStringGrid_Ex;
    gbData: TGroupBox;
    Panel14: TPanel;
    Panel3: TPanel;
    Panel2: TPanel;
    Panel12: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    edENM: TEdit;
    edLOT: TEdit;
    edRES: TEdit;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    dtpEDT: TDateTimePicker;
    Panel18: TPanel;
    dtpETM: TDateTimePicker;
    Panel19: TPanel;
    edDate: TEdit;
    edSEQ: TEdit;
    MainMenu1: TMainMenu;
    FILE1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    mnData: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    Panel20: TPanel;
    Panel21: TPanel;
    cmbxType: TComboBox;
    Panel22: TPanel;
    Memo1: TMemo;
    seIdx: TSpinEdit;
    btnPrt: TSpeedButton;
    ckbxVal: TCheckBox;
    btnList: TSpeedButton;
    btnSave: TSpeedButton;
    btnDel: TSpeedButton;
    pnCV: TPanel;
    pnMean: TPanel;
    pnSD: TPanel;
    pn2SD: TPanel;
    pn3SD: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure btnListClick(Sender: TObject);
    procedure cmbxENMChange(Sender: TObject);
    procedure gdListCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure gdListEditCellDone(Sender: TObject; ACol, ARow: Integer);
    procedure N5Click(Sender: TObject);
    procedure mnDataClick(Sender: TObject);
    procedure gdListClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure gdListClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure edENMChange(Sender: TObject);
    procedure edLOTChange(Sender: TObject);
    procedure dtpEDTCloseUp(Sender: TObject);
    procedure dtpEDTChange(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure Chart1ClickSeries(Sender: TCustomChart; Series: TChartSeries;
      ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X,
      Y: Integer);
    procedure N6Click(Sender: TObject);
    procedure btnPrtClick(Sender: TObject);
  private
    { Private declarations }
    procedure ViewList(RowIdx:integer=0);
    procedure ViewResOne(TQC:TIfMaster);
    procedure DeleteOne(EDT, SEQ:string);
    procedure InitCombo;
    procedure PanelClear;
    procedure InitLabelCombo(ENM:string);
    procedure InitChart;
    procedure ViewListAll;
    procedure ViewSetting(Lot:string);
    procedure SaveRes(EDT, SEQ, ECD, RES:string);
  public
    { Public declarations }
  end;

var
  F_QC: TF_QC;

implementation

uses Math, U_ENV, GlobalVar, U_DM, U_QC_PRT, QRCtrls, U_CodeInfo;

{$R *.dfm}

const
  C_CKB = 0;
  C_ENM = C_CKB+1;
  C_TYP = C_ENM+1;
  C_ETM = C_TYP+1;
  C_LOT = C_ETM+1;
  C_RES = C_LOT+1;
  C_EDT = C_RES+1;
  C_SEQ = C_EDT+1;
//  C_LVL = C_SEQ+1;

procedure TF_QC.InitCombo;
var
  ENM:string;
begin
  cmbxENM.Clear;

  with DM.qryC2 do begin
      Close;
      SQL.Text:= ' Select distinct upcode From TB_QC Order By upcode';
      Open;

      cmbxENM.Items.BeginUpdate;
      try
          if RecordCount > 0 then begin
              while Not Eof do begin
                  cmbxENM.Items.Add(Fields[0].AsString);
                  Next;
              end;
          end;
      finally
          cmbxENM.Items.EndUpdate;
      end;
  end;

  cmbxENM.ItemIndex:=0;

  cmbxENM.OnChange(nil);
end;

procedure TF_QC.FormCreate(Sender: TObject);
begin
  InitCombo;
  dtpF.Date:= now-30;
  dtpT.Date:= now;


  ViewListAll;

  gdList.HideColumns(C_EDT, C_SEQ);
end;

procedure TF_QC.btnListClick(Sender: TObject);
begin
  seIdx.Value:= 0;

  ViewList(seIdx.Value);
end;

procedure TF_QC.InitChart;
var
  i:integer;
  sNm:string;
begin
  for i:=0 to Chart1.SeriesCount-1 do begin
      Chart1.Series[i].Clear;

      sNm:= UpperCase(Chart1.Series[i].Title);
      if sNm = 'VAL' then
          Chart1.Series[i].SeriesColor:= TGlobal.FGrpColor.Res
      else if sNm = 'MEAN' then
          Chart1.Series[i].SeriesColor:= TGlobal.FGrpColor.Mean
      else if (sNm = '- SD1') or (sNm = '+ SD1') then
          Chart1.Series[i].SeriesColor:= TGlobal.FGrpColor.Sd1
      else if (sNm = '- SD2') or (sNm = '+ SD2') then
          Chart1.Series[i].SeriesColor:= TGlobal.FGrpColor.Sd2
      else if (sNm = '- SD3') or (sNm = '+ SD3') then
          Chart1.Series[i].SeriesColor:= TGlobal.FGrpColor.Sd3
      else
          ShowMessage(sNm);
  end;
end;

procedure TF_QC.cmbxENMChange(Sender: TObject);
begin
  InitLabelCombo(cmbxENM.Text);
end;

procedure TF_QC.InitLabelCombo(ENM:string);
begin
  cmbxLot.Clear;

  cmbxLot.Items.Add('전체');
  
  with DM.qryC2 do begin
      Close;
      SQL.Text:= ' Select distinct Lot From TB_QC '+
                 ' Where UpCode = '''+ENM+''' '+
                 '   And Mid(BarCode,17,2) = ''01''   ';
      Open;

      cmbxLot.Items.BeginUpdate;
      try
          if RecordCount > 0 then begin
              while Not Eof do begin
                  cmbxLot.Items.Add(Fields[0].AsString);
                  Next;
              end;
          end;
      finally
          cmbxLot.Items.EndUpdate;
      end;
  end;

  cmbxLot.ItemIndex:=0;
end;

procedure TF_QC.ViewListAll;
var
  i:integer;
begin
{
  gdAll.ClearNormalCells;
  gdAll.Row:=1;
  gdAll.RowCount:=2;

  with DM.qryC2 do begin
      Close;
      SQL.Text:= ' Select examtime, upcode, barcode, result, flag from tb_qc '+
                 ' order by examtime ';
      Open;

      if RecordCount = 0 then exit;

      gdAll.RowCount:= RecordCount + 1;
      i:=0;
      while Not Eof do begin
          inc(i);

          gdAll.Cells[0, i]:= fieldByName('examtime').AsString;
          gdAll.Cells[1, i]:= fieldByName('upcode').AsString;
          gdAll.Cells[2, i]:= fieldByName('barcode').AsString;
          gdAll.Cells[3, i]:= fieldByName('result').AsString;
          gdAll.Cells[4, i]:= fieldByName('flag').AsString;
          Next;
      end;
  end;     }
end;

procedure TF_QC.ViewSetting(Lot:string);
var
  i:integer;
begin
{
  gdSet.ClearNormalCells;
  gdSet.Row:=1;
  gdSet.RowCount:=2;

  with DM.qryC2 do begin
      Close;
      SQL.Text:= ' Select iname, ENm, fdt, tdt, lot, lev, r_low, r_high, sd1, sd2 from tb_lot '+
                 ' where lot =  '''+Lot+''' '+
                 '  order by fdt, tdt ';
      Open;

      if RecordCount = 0 then exit;

      gdSet.RowCount:= RecordCount + 1;
      i:=0;
      while Not Eof do begin
          inc(i);

          gdSet.Cells[0, i]:= fieldByName('ENm').AsString;
          gdSet.Cells[1, i]:= fieldByName('lot').AsString;
          gdSet.Cells[2, i]:= fieldByName('fdt').AsString+'~'+fieldByName('tdt').AsString;
          gdSet.Cells[3, i]:= fieldByName('lev').AsString;
          gdSet.Cells[4, i]:= fieldByName('r_low').AsString;
          gdSet.Cells[5, i]:= fieldByName('r_high').AsString;
          gdSet.Cells[6, i]:= fieldByName('sd1').AsString;
          gdSet.Cells[7, i]:= fieldByName('sd2').AsString;
          Next;
      end;
  end;
        }
end;

procedure TF_QC.gdListCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  if ARow > 0 then begin
      if ACol = 3 then
          CanEdit:= True
      else
          CanEdit:= False;
  end;
end;

procedure TF_QC.gdListEditCellDone(Sender: TObject; ACol, ARow: Integer);
var
  EDT, SEQ, ECD, RES:string;
begin
{
  if ARow > 0 then begin
      if ACol = 3 then begin
          EDT:= gdList.Cells[4, ARow];
          SEQ:= gdList.Cells[5, ARow];
          ECD:= gdList.Cells[0, ARow];
          RES:= gdList.Cells[3, ARow];
          if (EDT = '') or (SEQ='') or (ECD='') or (RES='') then exit
          else
              SaveRes(EDT, SEQ, ECD, RES);
      end;
  end;  }
end;

procedure TF_QC.SaveRes(EDT, SEQ, ECD, RES: string);
begin
  with DM.qryC2 do begin
     Close;
     SQL.Text:= ' Update tb_QC Set result = '''+RES+''' '+
                ' where ExamDate =  '''+EDT+''' '+
                '   And ExamSeq = '''+SEQ+'''   '+
                '   And UpCode  = '''+ECD+'''   ';
                //'   And Lev     = ''
     ExecSql;
  end;
end;

procedure TF_QC.N5Click(Sender: TObject);
begin
  Close;
end;

procedure TF_QC.mnDataClick(Sender: TObject);
begin
  gbData.Visible:= mnData.Checked;
end;

procedure TF_QC.gdListClickCell(Sender: TObject; ARow, ACol: Integer);
var
  TQC:TIfMaster;
begin
  if ARow > 0 then begin
      TQC:= TIfMaster.Create;
      try
          TQC.FRcvTime:= gdList.Cells[C_ETM, ARow];
          TQC.FExamDate:= gdList.Cells[C_EDT, ARow];
          TQC.FExamSeq:= gdList.Cells[C_SEQ, ARow];
          TQC.FLotNo:= gdList.Cells[C_LOT, ARow];
          TQC.FResult:= gdList.Cells[C_RES, ARow];
          TQC.FIfCode:= gdList.Cells[C_ENM, ARow];
          //TQC.FLVL:= gdList.Cells[C_LVL, ARow];

          ViewResOne(TQC);

      finally
          TQC.Free;
      end;
  end;

end;

procedure TF_QC.gdListClick(Sender: TObject);
begin
  seIdx.Value:= gdList.Row;
  gdList.OnClickCell(Sender, gdList.Row, gdList.Col);
end;

procedure TF_QC.ViewResOne(TQC: TIfMaster);
var
  Dt, tm: TDateTime;
begin
  if TQC.FLotNo = '' then exit;

  with TQC do begin
      edENM.Text:= FIfCode;
      edDate.Text:= FExamDate;
      edSEQ.Text:= FExamSeq;
      edRES.Text:= FResult;
      edLOT.Text:= FLotNo;

      dt:= StrToDateTime(FRcvTime);
      tm:= StrToDateTime(FRcvTime);

      dtpEDT.DateTime:= dt;
      dtpETM.DateTime:= tm;
  end;
end;

procedure TF_QC.btnSaveClick(Sender: TObject);
var
  TQC:TIfMaster;
begin

  TQC:= TIfMaster.Create;
  try
      //TQC.FETM:= FormatDateTime('yyyy-mm-dd', dtpEDT.Date)+' '+FormatDateTime('hh:nn:ss', dtpETM.DateTime)
      TQC.FRcvTime:= FormatDateTime('yyyy-mm-dd hh:nn:ss', dtpEDT.DateTime);
      TQC.FExamDate:= Trim(edDate.Text);
      TQC.FExamSeq:= Trim(edSEQ.Text);
      TQC.FLotNo:= Trim(edLOT.Text);
      TQC.FTYP:= UpperCase(Copy(cmbxType.Text,1,1));

      if StrToFloatDef(edRES.Text, -100) < -99 then
      begin
          ShowMessage('결과를 숫자로 입력하세요!');
          exit;
      end;
      TQC.FResult:= edRes.Text;
      TQC.FIfCode:= Trim(edENM.Text);
      //TQC.FLVL:= gdList.Cells[C_LVL, ARow];

      if TQC.FExamDate = '' then TQC.FExamDate:= FormatDateTime('yyyymmdd', StrToDateTime(TQC.FRcvTime));
      if TQC.FExamSeq = '' then
          TQC.FExamSeq:= DM.GetExamSeq(TQC.FExamDate);

      DM.SaveQC(TQC);

  finally
      TQC.Free;
  end;

  ViewList(seIdx.Value);

end;

procedure TF_QC.edENMChange(Sender: TObject);
begin
  if (Trim(edENM.Text) <> '') and (Trim(edLOT.Text) <> '') then
  begin
      btnSave.Enabled:= True;
      btnDel.Enabled:= True;
  end
  else begin
      btnSave.Enabled:= False;
      btnDel.Enabled:= False;
  end;
end;

procedure TF_QC.edLOTChange(Sender: TObject);
begin
  if (Trim(edENM.Text) <> '') and (Trim(edLOT.Text) <> '') then
  begin
      btnSave.Enabled:= True;
      btnDel.Enabled:= True;
  end
  else begin
      btnSave.Enabled:= False;
      btnDel.Enabled:= False;
  end;
end;

procedure TF_QC.PanelClear;
begin
  edENM.Text := '';
  edDate.Text:= '';
  edSEQ.Text := '';
  edRES.Text := '';
  edLOT.Text := '';
  dtpEDT.DateTime:= dtpF.DateTime;
  dtpETM.DateTime:= dtpF.DateTime;
end;

procedure TF_QC.dtpEDTCloseUp(Sender: TObject);
begin
  dtpETM.Date:= dtpEDT.Date;
end;

procedure TF_QC.dtpEDTChange(Sender: TObject);
begin
  dtpEDT.DateTime:= dtpETM.DateTime;
end;

procedure TF_QC.btnDelClick(Sender: TObject);
var
  EDT, SEQ:string;
begin
  EDT:= edDate.Text;
  SEQ:= edSEQ.Text;

  if (EDT='') or (SEQ='') then exit;

  if MessageDlg('삭제한후 복구 할수 없습니다. 계속하시겠습니까?', mtWarning, mbOKCancel, 1) = mrOk then
  begin
      DeleteOne(EDT, SEQ);
      btnList.Click;
  end;

end;

procedure TF_QC.DeleteOne(EDT, SEQ: string);
begin
  with DM.qryC2 do begin
     Close;
     SQL.Text:= ' Delete From tb_QC '+
                ' Where ExamDate = '''+EDT+''' '+
                '   And ExamSeq = '''+SEQ+'''  ';
     ExecSql;
  end;
end;

procedure TF_QC.Chart1ClickSeries(Sender: TCustomChart;
  Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Series = Chart1.Series[0] then begin
      gdList.Row:= ValueIndex+1;
  end;

end;

procedure TF_QC.ViewList(RowIdx:integer);
var
  WF, WT, ENM, Res, Min, Sd, Cv, Lot:string;
  dAll, dVal, dMean, dSd1, dSd2, dSd3, dCv:double;
  i, R, dCnt, nStr, nEnd:integer;
  aVal1, aVal2, aLSd1, aHSd1, aLSd2, aHSd2, aLSd3, aHSd3, aMean:array of double;
  aLot, aDate1, aDate2:array of string;
  dLSd1, dHSd1, dLSd2, dHSd2, dMin, dMax:double;
  OldLot, NewLot:string;
begin
  gdList.ClearNormalCells;
  gdList.Row:=1;
  gdList.RowCount:= 2;

  dLSd1:=0;
  dHSd1:=0;
  dLSd2:=0;
  dHSd2:=0;

  pnMean.Caption:= '';
  pnSD.Caption:= '';
  pnCV.Caption:= '';
  pn2SD.Caption:= '';
  pn3SD.Caption:= '';

  InitChart;

  ViewSetting(cmbxLot.Text);

  WF:= FormatDateTime('yyyymmdd', dtpF.Date);
  WT:= FormatDateTime('yyyymmdd', dtpT.Date);
  ENM:= cmbxENM.Text;
  if cmbxLot.ItemIndex = 0 then
      Lot:= ''
  else
      Lot:= cmbxLot.Text;

  dVal:= 0;
  dMin:= 0;
  dMax:= 0;

  Chart1.Series[0].Marks.Visible:= ckbxVal.Checked;

  with DM.qryC2 do begin
      Close;
      SQL.Text:= ' Select Lot, ExamDate, ExamSeq, TYP, UpCode, LOT, ExamTime, Result, Flag '+
                 ' From TB_QC  '+
                 ' Where Examdate Between '''+WF+''' '+
                 '                    And '''+WT+''' '+
                 '   And UpCode = '''+ENM+''' '+
                 '   And TYP = ''Q''  ';
      if Lot <> '' then
          SQL.Text:= SQL.Text + '   And LOT = '''+Lot+''' ';

      SQL.Text:= SQL.Text + ' Order by ExamTime ';
      Open;

      if RecordCount = 0 then exit;

      dAll:=0;
      dVal:=0;
      dCnt:=0;

      SetLength(aVal1, Recordcount+1);
      SetLength(aDate1, RecordCount+1);
      SetLength(aLot, RecordCount+1);

      for i:=0 to High(aVal1) do begin
          aVal1[i]:= -1;
          aDate1[i]:= '';
      end;

      gdList.RowCount:= Recordcount +1;

      R:=0;

      while Not Eof do begin
          Inc(R);
          //gdList.AddCheckBox(C_CKB, R, False, False);
          gdList.Cells[C_CKB, R]:= IntToStr(R);
          gdList.Cells[C_ENM, R]:= FieldByName('UpCode').AsString;
          gdList.Cells[C_ETM, R]:= FieldByName('ExamTime').AsString;
          gdList.Cells[C_LOT, R]:= FieldByName('LOT').AsString;
          gdList.Cells[C_EDT, R]:= FieldByName('ExamDate').AsString;
          gdList.Cells[C_SEQ, R]:= FieldByName('ExamSeq').AsString;
          if FIeldByName('TYP').AsString = 'C' then
              gdList.Cells[C_TYP, R]:= 'CAL'
          else
          if FieldByname('TYP').AsString = 'Q' then
              gdList.Cells[C_TYP, R]:= 'QC'
          else
          if FieldByName('TYP').AsString = 'S' then
              gdList.Cells[C_TYP, R]:= 'SAMPLE';

          Res:= FieldByName('Result').AsString;
          if Pos('(', Res) > 0 then begin
              nStr:= Pos('(',Res)+1;    //10
              nEnd:= Pos(')',Res);      //15
              Res:= Copy(Res, nStr, nEnd-nStr);
              //Res:= Copy(Res, Pos('(',Res)+1, (Length(Res)- Pos(')',Res) - (Pos('(',Res)+2)  ));
          end;

          Res:= StringReplace(Res, '<', '', [rfReplaceAll]);
          Res:= StringReplace(Res, '>', '', [rfReplaceAll]);

          dVal:= StrToFloatDef(Res, -1);
          if dVal < 0 then begin
              Next;
              Continue;
          end;

          Inc(dCnt);
          aVal1[dCnt-1]:= dVal;
          aDate1[dCnt-1]:= StringReplace(Copy(FieldByName('ExamTime').AsString,1,10),'-','',[rfReplaceAll]);
          aLot[dCnt-1]:= FieldByName('LOT').AsString;
          
          //데이터가 모두 같으면 Min과 Max가 다른상황이 나오므로.. 초기화..
          if dCnt = 1 then begin
              dMin:= dVal;
              dMax:= dVal;
          end;

          if dVal < dMin then
              dMin:= dVal;
          if dVal > dMax then
              dMax:= dVal;

          gdList.Cells[C_RES, R]:= Trim(Format('%5.1f', [dVal]));;

          Next;
      end;


      SetLength(aVal2, dCnt);
      SetLength(aLSD1, dCnt);
      SetLength(aHSD1, dCnt);
      SetLength(aLSD2, dCnt);
      SetLength(aHSD3, dCnt);
      SetLength(aLSD3, dCnt);
      SetLength(aHSD2, dCnt);
      SetLength(aMean, dCnt);
      SetLength(aDate2, dCnt);

      for i:=0 to dCnt -1 do begin
          aVal2[i]:= aVal1[i];
          aDate2[i]:= aDate1[i];
      end;

      dMean:= mean(aVal2);
      pnMean.Caption:= Trim(Format('%5.1f', [dMean]));

      //표준편차는 2건 이상인경우만 계산된다.
      if dMin <> dMax then begin
          Chart1.LeftAxis.Automatic:= false;
          chart1.LeftAxis.Maximum:= dMean*1.8;
          chart1.LeftAxis.Minimum:= dMean*-0.2;

          dSd1:= stddev(aVal2);
          pnSD.Caption:= Trim(Format('%5.1f', [dSd1]));

          dCv:= dSd1 / dMean * 100;
          pnCV.Caption:= Trim(Format('%5.1f', [dCv]));

          for i:=0 to High(aMean) do begin
              aMean[i]:= dMean;
              aLSD1[i]:= dMean - stddev(aVal2);
              aHSD1[i]:= dMean + stddev(aVal2);

              aLSD2[i]:= ( dMean - ( stddev(aVal2)+ stddev(aVal2)));
              aHSD2[i]:= ( dMean + ( stddev(aVal2)+ stddev(aVal2)));

              aLSD3[i]:= ( dMean - ( stddev(aVal2)+ stddev(aVal2)+ stddev(aVal2)));
              aHSD3[i]:= ( dMean + ( stddev(aVal2)+ stddev(aVal2)+ stddev(aVal2)));
          end;
      end
      else begin
          //편차가 0인경우 처리 할방법..
          pnSD.Caption:= '0';
          Chart1.LeftAxis.Automatic:= false;
          chart1.LeftAxis.Maximum:= dMax*2;
          chart1.LeftAxis.Minimum:= dMin*-1;
      end;

      pn2SD.Caption:= Format('%3.1f', [aLSD2[0]]) +' ~ '+ Format('%3.1f', [aHSD2[0]]);
      pn3SD.Caption:= Format('%3.1f', [aLSD3[0]]) +' ~ '+ Format('%3.1f', [aHSD3[0]]);

      OldLot:= aLot[0];
      for i:=0 to High(aMean) do begin
          //NewLot:= aLot[i];
          //if OldLot <> NewLot then begin
          //  Chart1.Series[8].AddXY(i, chart1.LeftAxis.Maximum);
          //end;

          TLineSeries( Chart1.Series[0] ).AddXY(i, aVal2[i]);
          TLineSeries( Chart1.Series[1] ).AddXY(i, aMean[i], aDate2[i]);
      end;

      TLineSeries( Chart1.Series[2] ).AddArray(aLSD1);
      TLineSeries( Chart1.Series[3] ).AddArray(aHSD1);
      TLineSeries( Chart1.Series[4] ).AddArray(aLSD2);
      TLineSeries( Chart1.Series[5] ).AddArray(aHSD2);
      TLineSeries( Chart1.Series[6] ).AddArray(aLSD3);
      TLineSeries( Chart1.Series[7] ).AddArray(aHSD3);
  end;

  if RowIdx > 0 then
      gdList.Row:= RowIdx;
      
  

end;

procedure TF_QC.N6Click(Sender: TObject);
begin
  if F_Env = nil then
      F_ENV:= TF_ENV.Create(Self);

  F_ENV.ShowModal;
end;

procedure TF_QC.btnPrtClick(Sender: TObject);
var
  i:integer;
  aETM, aLOT, aRES:array[1..90] of string;
begin
  TChart( Chart1 ).SaveToMetafileEnh(TGlobal.AppPath+'chart.WMF');

  for i:=1 to 90 do begin
      aETM[i]:= '';
      aLOT[i]:= '';
      aRES[i]:= '';
  end;

  with F_QC_PRT do begin
      qrlTitle.Caption:= 'CHORUS-TRIO '+cmbxENM.Text+' QC REPORT';
      qrlTestGG.Caption:= '기관명: '+ TGlobal.FHospNm;
      qrlTestDt.Caption:= 'TEST 기간: '+FormatDateTime('yyyy-mm-dd', dtpF.Date)+' ~ '+ FormatDateTime('yyyy-mm-dd', dtpT.Date);
      if FileExists(TGlobal.AppPath+'\chart.WMF') then
          qrlGrp.Picture.LoadFromFile(TGlobal.AppPath+'\chart.WMF')
      else
          qrlGrp.Picture:= nil;

      qrlENM.Caption := '항목 : ' + cmbxENM.Text;
      qrlUnit.Caption:= '단위 : ' + TCode.GetUnit_IF(cmbxENM.Text);
      qrlCor.Caption := '회사 : ' + TGlobal.FCor;
      qrlLot.Caption := 'LOT  : ' + cmbxLot.Text;
      qrlMean.Caption:= 'Mean : ' + pnMean.Caption;
      qrlSD.Caption  := 'SD   : ' + pnSD.Caption;
      qrlCV.Caption  := 'CV(%): ' + pnCV.Caption;
      qrl2SD.Caption := '2SD: ' + pn2SD.Caption;
      qrl3SD.Caption := '3SD: ' + pn3SD.Caption;

      for i:=1 to gdList.RowCount -1 do begin
          if i <= 90 then begin
              aETM[i]:= gdList.Cells[C_ETM, i];
              aLOT[i]:= gdList.Cells[C_LOT, i];
              aRES[i]:= gdList.Cells[C_RES, i];
          end;
      end;

      for i:=1 to 90 do begin
          TQRLabel( FindComponent('T'+IntToStr(i)) ).Caption:= aETM[i];
          TQRLabel( FindComponent('L'+IntToStr(i)) ).Caption:= aLOT[i];
          TQRLabel( FindComponent('R'+IntToStr(i)) ).Caption:= aRES[i];
      end;

      QuickRep1.Preview;
  end;


end;

end.
