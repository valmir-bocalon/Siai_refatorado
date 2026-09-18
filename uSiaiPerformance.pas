unit uSiaiPerformance;

interface

uses Winapi.Windows;

type
  TSiaiMetric = (smLookupNome, smLookupCPF, smLookupLoteamento,
    smCalculoFinanceiro, smRecalculoExplicito, smPrimeiroTitulo,
    smProximoTitulo, smRestaurarTitulo, smReativarControles, smLookupParticipante);

function PerformanceEnabled: Boolean;
function PerformanceStart: UInt64;
procedure PerformanceSample(AMetric: TSiaiMetric; AStart: UInt64);
procedure PerformanceElapsed(const AStage: string; AStart: UInt64);
procedure PerformanceCheckpoint(const AStage: string);
procedure FlushPerformanceLog;

implementation

uses System.SysUtils, System.Classes, System.Diagnostics;

var
  GEnabled: Boolean;
  GStarted, GLastCheckpoint: UInt64;
  GLines: TStringList;
  GFileName: string;
  GClock: TStopwatch;
  GMetricCount, GMetricTime: array[TSiaiMetric] of UInt64;

const
  MetricNames: array[TSiaiMetric] of string = (
    'lookup participante/nome', 'lookup participante/CPF', 'lookup loteamento',
    'formulas financeiras (sem lookups)', 'recalculo explicito (inclui calculos e lookups)',
    'First (inclui eventos)', 'Next (inclui eventos)', 'Locate (inclui eventos)',
    'EnableControls (inclui eventos)', 'lookup participante/nome e CPF juntos');

function PerformanceEnabled: Boolean;
begin
  Result := GEnabled;
end;

procedure PerformanceSample(AMetric: TSiaiMetric; AStart: UInt64);
begin
  if not GEnabled then Exit;
  Inc(GMetricCount[AMetric]);
  Inc(GMetricTime[AMetric], GClock.ElapsedMilliseconds - AStart);
end;

function PerformanceTicks: UInt64;
begin
  Result := GClock.ElapsedMilliseconds;
end;

function PerformanceStart: UInt64;
begin
  Result := 0;
  if GEnabled then
    Result := PerformanceTicks;
end;

procedure PerformanceElapsed(const AStage: string; AStart: UInt64);
begin
  if not GEnabled then Exit;
  try
    GLines.Add(Format('%d ms | %s | duracao=%d ms',
      [PerformanceTicks - GStarted, AStage, PerformanceTicks - AStart]));
  except
    { O diagnostico nunca deve impedir o funcionamento do sistema. }
  end;
end;

procedure PerformanceCheckpoint(const AStage: string);
begin
  if not GEnabled then Exit;
  PerformanceElapsed(AStage, GLastCheckpoint);
  GLastCheckpoint := PerformanceTicks;
end;

procedure FlushPerformanceLog;
var
  Metric: TSiaiMetric;
begin
  if not GEnabled then Exit;
  try
    for Metric := Low(TSiaiMetric) to High(TSiaiMetric) do
      if GMetricCount[Metric] > 0 then
      begin
        GLines.Add(Format('%d ms | Cobranca detalhe: %s | chamadas=%d | total=%d ms',
          [PerformanceTicks - GStarted, MetricNames[Metric],
           GMetricCount[Metric], GMetricTime[Metric]]));
        GMetricCount[Metric] := 0;
        GMetricTime[Metric] := 0;
      end;
    GLines.SaveToFile(GFileName, TEncoding.UTF8);
  except
    { Sem permissao no TEMP, o diagnostico permanece apenas em memoria. }
  end;
end;

procedure InitializePerformanceLog;
var
  LPath: array[0..MAX_PATH] of Char;
  LLength: DWORD;
begin
  GEnabled := FindCmdLineSwitch('diagnostico', ['/', '-'], True);
  if not GEnabled then Exit;
  try
    GClock := TStopwatch.StartNew;
    GStarted := PerformanceTicks;
    GLastCheckpoint := GStarted;
    GLines := TStringList.Create;
    LLength := GetTempPath(Length(LPath), LPath);
    if (LLength = 0) or (LLength >= DWORD(Length(LPath))) then
    begin
      GEnabled := False;
      Exit;
    end;
    GFileName := IncludeTrailingPathDelimiter(string(LPath)) +
      Format('SIAI-startup-%d.log', [GetCurrentProcessId]);
    GLines.Add('SIAI: tempos em ms; dialogos incluem o tempo do usuario.');
    GLines.Add('Inicio: ' + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
  except
    GEnabled := False;
  end;
end;

initialization
  InitializePerformanceLog;

finalization
  FlushPerformanceLog;
  GLines.Free;

end.
