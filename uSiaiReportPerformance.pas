unit uSiaiReportPerformance;

interface

uses RLReport, Data.DB;

procedure MeasureReport(AReport: TRLReport; const AName: string);
procedure MeasureReportDataSet(AReport: TRLReport; ADataSet: TDataSet;
  const AName: string);
procedure MeasureReportBand(AReport: TRLReport; ABand: TRLCustomBand;
  const AName: string);

implementation

uses System.Classes, System.SysUtils, uSiaiPerformance;

type
  TReportDataSetTiming = class(TComponent)
  private
    FName: string;
    FDataSet: TDataSet;
    FBeforeOpen: TDataSetNotifyEvent;
    FAfterOpen: TDataSetNotifyEvent;
    FBeforeScroll: TDataSetNotifyEvent;
    FAfterScroll: TDataSetNotifyEvent;
    FOpenStarted: UInt64;
    FOpening: Boolean;
    FOpenCount: Integer;
    FOpenTime: UInt64;
    FScrollStarted: UInt64;
    FScrolling: Boolean;
    FScrollCount: Integer;
    FScrollTime: UInt64;
    procedure BeforeOpen(DataSet: TDataSet);
    procedure AfterOpen(DataSet: TDataSet);
    procedure BeforeScroll(DataSet: TDataSet);
    procedure AfterScroll(DataSet: TDataSet);
  public
    destructor Destroy; override;
    procedure Reset;
    procedure Flush;
  end;

  TReportBandTiming = class(TComponent)
  private
    FName: string;
    FBand: TRLCustomBand;
    FBefore: TRLBeforePrintEvent;
    FAfter: TRLAfterPrintEvent;
    FBeforeCount: Integer;
    FAfterCount: Integer;
    FBeforeTime: UInt64;
    FAfterTime: UInt64;
    procedure BeforeBand(Sender: TObject; var PrintIt: Boolean);
    procedure AfterBand(Sender: TObject);
  public
    destructor Destroy; override;
    procedure Flush;
  end;

  TReportTiming = class(TComponent)
  private
    FName: string;
    FStarted: UInt64;
    FBefore: TRLBeforePrintEvent;
    FAfter: TRLAfterPrintEvent;
    procedure BeforeReport(Sender: TObject; var PrintIt: Boolean);
    procedure AfterReport(Sender: TObject);
  end;

procedure TReportDataSetTiming.BeforeOpen(DataSet: TDataSet);
begin
  FOpenStarted := PerformanceStart;
  FOpening := True;
  if Assigned(FBeforeOpen) then
    FBeforeOpen(DataSet);
end;

procedure TReportDataSetTiming.AfterOpen(DataSet: TDataSet);
var
  LFinished: UInt64;
begin
  try
    if Assigned(FAfterOpen) then
      FAfterOpen(DataSet);
  finally
    if FOpening then
    begin
      LFinished := PerformanceStart;
      Inc(FOpenCount);
      Inc(FOpenTime, LFinished - FOpenStarted);
      FOpening := False;
    end;
  end;
end;

procedure TReportDataSetTiming.BeforeScroll(DataSet: TDataSet);
begin
  FScrollStarted := PerformanceStart;
  FScrolling := True;
  if Assigned(FBeforeScroll) then
    FBeforeScroll(DataSet);
end;

procedure TReportDataSetTiming.AfterScroll(DataSet: TDataSet);
var
  LFinished: UInt64;
begin
  try
    if Assigned(FAfterScroll) then
      FAfterScroll(DataSet);
  finally
    if FScrolling then
    begin
      LFinished := PerformanceStart;
      Inc(FScrollCount);
      Inc(FScrollTime, LFinished - FScrollStarted);
      FScrolling := False;
    end;
  end;
end;

destructor TReportDataSetTiming.Destroy;
var
  LCurrent, LOwn: TDataSetNotifyEvent;
begin
  if FDataSet <> nil then
  begin
    LCurrent := FDataSet.BeforeOpen;
    LOwn := BeforeOpen;
    if (TMethod(LCurrent).Code = TMethod(LOwn).Code) and
       (TMethod(LCurrent).Data = Self) then
      FDataSet.BeforeOpen := FBeforeOpen;
    LCurrent := FDataSet.AfterOpen;
    LOwn := AfterOpen;
    if (TMethod(LCurrent).Code = TMethod(LOwn).Code) and
       (TMethod(LCurrent).Data = Self) then
      FDataSet.AfterOpen := FAfterOpen;
    LCurrent := FDataSet.BeforeScroll;
    LOwn := BeforeScroll;
    if (TMethod(LCurrent).Code = TMethod(LOwn).Code) and
       (TMethod(LCurrent).Data = Self) then
      FDataSet.BeforeScroll := FBeforeScroll;
    LCurrent := FDataSet.AfterScroll;
    LOwn := AfterScroll;
    if (TMethod(LCurrent).Code = TMethod(LOwn).Code) and
       (TMethod(LCurrent).Data = Self) then
      FDataSet.AfterScroll := FAfterScroll;
  end;
  inherited Destroy;
end;

procedure TReportDataSetTiming.Reset;
begin
  FOpenStarted := 0;
  FOpening := False;
  FOpenCount := 0;
  FOpenTime := 0;
  FScrollStarted := 0;
  FScrolling := False;
  FScrollCount := 0;
  FScrollTime := 0;
end;

procedure TReportDataSetTiming.Flush;
var
  LNow: UInt64;
begin
  if (FOpenCount = 0) and (FScrollCount = 0) then
    Exit;
  LNow := PerformanceStart;
  if FOpenCount > 0 then
    PerformanceElapsed(Format('%s: aberturas=%d', [FName, FOpenCount]),
      LNow - FOpenTime);
  if FScrollCount > 0 then
    PerformanceElapsed(Format('%s: navegacoes=%d', [FName, FScrollCount]),
      LNow - FScrollTime);
end;

procedure TReportBandTiming.BeforeBand(Sender: TObject; var PrintIt: Boolean);
var
  LStarted: UInt64;
begin
  LStarted := PerformanceStart;
  if Assigned(FBefore) then
    FBefore(Sender, PrintIt);
  Inc(FBeforeCount);
  Inc(FBeforeTime, PerformanceStart - LStarted);
end;

procedure TReportBandTiming.AfterBand(Sender: TObject);
var
  LStarted: UInt64;
begin
  LStarted := PerformanceStart;
  if Assigned(FAfter) then
    FAfter(Sender);
  Inc(FAfterCount);
  Inc(FAfterTime, PerformanceStart - LStarted);
end;

destructor TReportBandTiming.Destroy;
var
  LCurrent, LOwn: TRLBeforePrintEvent;
  LCurrentAfter, LOwnAfter: TRLAfterPrintEvent;
begin
  if FBand <> nil then
  begin
    LCurrent := FBand.BeforePrint;
    LOwn := BeforeBand;
    if (TMethod(LCurrent).Code = TMethod(LOwn).Code) and
       (TMethod(LCurrent).Data = Self) then
      FBand.BeforePrint := FBefore;
    LCurrentAfter := FBand.AfterPrint;
    LOwnAfter := AfterBand;
    if (TMethod(LCurrentAfter).Code = TMethod(LOwnAfter).Code) and
       (TMethod(LCurrentAfter).Data = Self) then
      FBand.AfterPrint := FAfter;
  end;
  inherited Destroy;
end;

procedure TReportBandTiming.Flush;
var
  LNow: UInt64;
begin
  LNow := PerformanceStart;
  if FBeforeCount > 0 then
    PerformanceElapsed(Format('%s: eventos BeforePrint=%d',
      [FName, FBeforeCount]), LNow - FBeforeTime);
  if FAfterCount > 0 then
    PerformanceElapsed(Format('%s: eventos AfterPrint=%d',
      [FName, FAfterCount]), LNow - FAfterTime);
end;

procedure TReportTiming.BeforeReport(Sender: TObject; var PrintIt: Boolean);
begin
  PerformanceCheckpoint(FName + ': inicio da geracao de paginas');
  FlushPerformanceLog;
  FStarted := PerformanceStart;
  if Assigned(FBefore) then FBefore(Sender, PrintIt);
end;

procedure TReportTiming.AfterReport(Sender: TObject);
var
  I: Integer;
begin
  try
    if Assigned(FAfter) then FAfter(Sender);
  finally
    for I := 0 to Owner.ComponentCount - 1 do
      if Owner.Components[I] is TReportDataSetTiming then
        TReportDataSetTiming(Owner.Components[I]).Flush
      else if Owner.Components[I] is TReportBandTiming then
        TReportBandTiming(Owner.Components[I]).Flush;
    PerformanceElapsed(FName + ': geracao de paginas (sem espera na visualizacao)', FStarted);
    FlushPerformanceLog;
  end;
end;

procedure MeasureReport(AReport: TRLReport; const AName: string);
var
  I: Integer;
  Timing: TReportTiming;
begin
  if not PerformanceEnabled or (AReport = nil) then Exit;
  for I := 0 to AReport.ComponentCount - 1 do
    if AReport.Components[I] is TReportTiming then Exit;
  Timing := TReportTiming.Create(AReport);
  Timing.FName := AName;
  Timing.FBefore := AReport.BeforePrint;
  Timing.FAfter := AReport.AfterPrint;
  AReport.BeforePrint := Timing.BeforeReport;
  AReport.AfterPrint := Timing.AfterReport;
end;

procedure MeasureReportDataSet(AReport: TRLReport; ADataSet: TDataSet;
  const AName: string);
var
  I: Integer;
  Timing: TReportDataSetTiming;
begin
  if not PerformanceEnabled or (AReport = nil) or (ADataSet = nil) then
    Exit;
  for I := 0 to AReport.ComponentCount - 1 do
    if (AReport.Components[I] is TReportDataSetTiming) and
       (TReportDataSetTiming(AReport.Components[I]).FDataSet = ADataSet) then
      Exit;
  Timing := TReportDataSetTiming.Create(AReport);
  Timing.FName := AName;
  Timing.FDataSet := ADataSet;
  Timing.FBeforeOpen := ADataSet.BeforeOpen;
  Timing.FAfterOpen := ADataSet.AfterOpen;
  Timing.FBeforeScroll := ADataSet.BeforeScroll;
  Timing.FAfterScroll := ADataSet.AfterScroll;
  ADataSet.BeforeOpen := Timing.BeforeOpen;
  ADataSet.AfterOpen := Timing.AfterOpen;
  ADataSet.BeforeScroll := Timing.BeforeScroll;
  ADataSet.AfterScroll := Timing.AfterScroll;
end;

procedure MeasureReportBand(AReport: TRLReport; ABand: TRLCustomBand;
  const AName: string);
var
  I: Integer;
  Timing: TReportBandTiming;
begin
  if not PerformanceEnabled or (AReport = nil) or (ABand = nil) then
    Exit;
  for I := 0 to AReport.ComponentCount - 1 do
    if (AReport.Components[I] is TReportBandTiming) and
       (TReportBandTiming(AReport.Components[I]).FBand = ABand) then
      Exit;
  Timing := TReportBandTiming.Create(AReport);
  Timing.FName := AName;
  Timing.FBand := ABand;
  Timing.FBefore := ABand.BeforePrint;
  Timing.FAfter := ABand.AfterPrint;
  ABand.BeforePrint := Timing.BeforeBand;
  ABand.AfterPrint := Timing.AfterBand;
end;

end.
