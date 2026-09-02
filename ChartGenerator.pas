unit ChartGenerator;

interface

uses
  Windows, Classes, Graphics, Controls, ZConnection, ZDataset, SysUtils, Math,
  Dialogs, GraphUtil, DB, Printers, ExtCtrls;

type
  TChartType = (ctBar, ctLine, ctPie);

  TChartData = record
    XValue: string;
    YValue: Double;
    Color: TColor;
  end;

  TChartDataArray = array of TChartData;

  TChartSeriesData = record
    Title: string;
    Data: TChartDataArray;
    Color: TColor;
  end;

  TChartSeriesArray = array of TChartSeriesData;

procedure LoadChartDataFromSQL(Connection: TZConnection; SQLQuery: string;
  XField, YField: string; var Data: TChartDataArray);
procedure LoadChartDataFromQuery(Query: TZQuery; XField, YField: string;
  var Data: TChartDataArray);
procedure LoadChartSeriesFromDataSet(DataSet: TDataSet; const XField: string;
  const SeriesTitles: array of string; const YFields: array of string;
  const SeriesColors: array of TColor; var Series: TChartSeriesArray);
procedure GenerateChart(Canvas: TCanvas; Data: TChartDataArray;
  ChartType: TChartType; Width, Height: Integer; Title: string = '';
  ShowPercentage: Boolean = True; ShowValuesOnBars: Boolean = True;
  ShowYValuesInLegend: Boolean = False; BackgroundColor: TColor = clWhite;
  ShowXAxisLabels: Boolean = True; ShowYAxisLabels: Boolean = True);
procedure GenerateMultiLineChart(Canvas: TCanvas; Series: TChartSeriesArray;
  Width, Height: Integer; Title: string = ''; BackgroundColor: TColor = clWhite;
  ShowXAxisLabels: Boolean = True; ShowYAxisLabels: Boolean = True);
procedure GenerateMultiSeriesChart(Canvas: TCanvas; Series: TChartSeriesArray;
  ChartType: TChartType; Width, Height: Integer; Title: string = '';
  BackgroundColor: TColor = clWhite; ShowXAxisLabels: Boolean = True;
  ShowYAxisLabels: Boolean = True; LegendBySeriesForBars: Boolean = True;
  DashboardLegend: Boolean = False);
procedure PrintPaintBoxChart(PaintBox: TPaintBox; const Title: string = '';
  Landscape: Boolean = True);

implementation

uses
  UPrinterUtils;

const
  DEFAULT_MARGIN = 40;
  COLOR_PALETTE: array[0..11] of TColor = (
    $0F39C9C, $0BAFEDD, $07474E7, $0C5C58B,
    $0C58BC5, $08BC5C5, $0C090C5, $0F9CEA4,
    $0677C52, $05C758B, $0340044, $0171717
  );

function FormatYValue(Value: Double): string;
var
  FS: TFormatSettings;
begin
  FS := TFormatSettings.Create;
  FS.DecimalSeparator := ',';
  FS.ThousandSeparator := '.';
  Result := FormatFloat('###,###,##0.00', Value, FS);
end;

procedure DrawTitle(Canvas: TCanvas; const Title: string; Width: Integer);
begin
  if Title <> '' then
  begin
    Canvas.Brush.Style := bsClear;
    Canvas.Font.Size := 14;
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clNavy;
    Canvas.TextOut((Width - Canvas.TextWidth(Title)) div 2, 15, Title);
    Canvas.Font.Style := [];
  end;
end;

procedure DrawValueBox(Canvas: TCanvas; const Text: string; X, Y: Integer);
var
  R: TRect;
begin
  Canvas.Font.Size := 8;
  R := Rect(X, Y, X + Canvas.TextWidth(Text) + 8, Y + Canvas.TextHeight(Text) + 4);
  Canvas.Brush.Style := bsSolid;
  Canvas.Brush.Color := clWhite;
  Canvas.Pen.Color := $00D8D8D8;
  Canvas.Rectangle(R);
  Canvas.Brush.Style := bsClear;
  Canvas.Font.Color := clNavy;
  Canvas.TextOut(R.Left + 4, R.Top + 2, Text);
end;

procedure LoadChartDataFromSQL(Connection: TZConnection; SQLQuery: string;
  XField, YField: string; var Data: TChartDataArray);
var
  ZQuery: TZQuery;
  I: Integer;
begin
  ZQuery := TZQuery.Create(nil);
  try
    ZQuery.Connection := Connection;
    ZQuery.SQL.Text := SQLQuery;
    ZQuery.Open;
    SetLength(Data, 0);
    while not ZQuery.Eof do
    begin
      SetLength(Data, Length(Data) + 1);
      I := High(Data);
      Data[I].XValue := ZQuery.FieldByName(XField).AsString;
      Data[I].YValue := ZQuery.FieldByName(YField).AsFloat;
      Data[I].Color := COLOR_PALETTE[I mod Length(COLOR_PALETTE)];
      ZQuery.Next;
    end;
  finally
    ZQuery.Free;
  end;
end;

procedure LoadChartDataFromQuery(Query: TZQuery; XField, YField: string;
  var Data: TChartDataArray);
var
  I: Integer;
begin
  SetLength(Data, 0);
  if Query = nil then
    Exit;
  if not Query.Active then
    Query.Open;
  Query.First;
  while not Query.Eof do
  begin
    SetLength(Data, Length(Data) + 1);
    I := High(Data);
    Data[I].XValue := Query.FieldByName(XField).AsString;
    Data[I].YValue := Query.FieldByName(YField).AsFloat;
    Data[I].Color := COLOR_PALETTE[I mod Length(COLOR_PALETTE)];
    Query.Next;
  end;
end;

procedure LoadChartSeriesFromDataSet(DataSet: TDataSet; const XField: string;
  const SeriesTitles: array of string; const YFields: array of string;
  const SeriesColors: array of TColor; var Series: TChartSeriesArray);
var
  I, Row: Integer;
  Bmk: TBookmark;
begin
  SetLength(Series, Length(YFields));
  for I := 0 to High(YFields) do
  begin
    if I <= High(SeriesTitles) then
      Series[I].Title := SeriesTitles[I]
    else
      Series[I].Title := YFields[I];
    if I <= High(SeriesColors) then
      Series[I].Color := SeriesColors[I]
    else
      Series[I].Color := COLOR_PALETTE[I mod Length(COLOR_PALETTE)];
    SetLength(Series[I].Data, 0);
  end;

  if (DataSet = nil) or (not DataSet.Active) or DataSet.IsEmpty then
    Exit;

  Bmk := DataSet.GetBookmark;
  DataSet.DisableControls;
  try
    DataSet.First;
    Row := 0;
    while not DataSet.Eof do
    begin
      for I := 0 to High(YFields) do
      begin
        SetLength(Series[I].Data, Row + 1);
        Series[I].Data[Row].XValue := DataSet.FieldByName(XField).AsString;
        Series[I].Data[Row].YValue := DataSet.FieldByName(YFields[I]).AsFloat;
        Series[I].Data[Row].Color := COLOR_PALETTE[(I * 31 + Row) mod Length(COLOR_PALETTE)];
      end;
      Inc(Row);
      DataSet.Next;
    end;
  finally
    if DataSet.BookmarkValid(Bmk) then
      DataSet.GotoBookmark(Bmk);
    DataSet.FreeBookmark(Bmk);
    DataSet.EnableControls;
  end;
end;

procedure GenerateMultiLineChart(Canvas: TCanvas; Series: TChartSeriesArray;
  Width, Height: Integer; Title: string = ''; BackgroundColor: TColor = clWhite;
  ShowXAxisLabels: Boolean = True; ShowYAxisLabels: Boolean = True);
begin
  GenerateMultiSeriesChart(Canvas, Series, ctLine, Width, Height, Title,
    BackgroundColor, ShowXAxisLabels, ShowYAxisLabels);
end;

procedure GenerateMultiSeriesChart(Canvas: TCanvas; Series: TChartSeriesArray;
  ChartType: TChartType; Width, Height: Integer; Title: string = '';
  BackgroundColor: TColor = clWhite; ShowXAxisLabels: Boolean = True;
  ShowYAxisLabels: Boolean = True; LegendBySeriesForBars: Boolean = True;
  DashboardLegend: Boolean = False);
var
  ChartArea: TRect;
  I, J, K, XPos, YPos, PrevX, PrevY, GroupWidth, BarWidth, BarLeft: Integer;
  PieLeft, PieTop, PieSize: Integer;
  LegendRows, LegendMonthColumns, LegendItemHeight: Integer;
  LegendColumns, LegendColumnWidth, LegendSideWidth, LegendFontSize: Integer;
  LegendTopPos, LegendLeft: Integer;
  MinY, MaxY, RangeY, ScaleFactorY, TotalY, StartAngle, EndAngle, MidAngle: Double;
  ZeroY: Integer;
  S, YLabel: string;
  LabelValue: Double;
  PieData: TChartDataArray;

  function DataCount: Integer;
  begin
    Result := 0;
    if Length(Series) > 0 then
      Result := Length(Series[0].Data);
  end;

  function HasData: Boolean;
  begin
    Result := (Length(Series) > 0) and (DataCount > 0);
  end;

  function MonthColorIndex(const MonthLabel: string; Fallback: Integer): Integer;
  var
    I: Integer;
  begin
    Result := Fallback mod 12;
    for I := 1 to 12 do
      if SameText(Trim(MonthLabel), Trim(FormatSettings.ShortMonthNames[I])) then
      begin
        Result := I - 1;
        Exit;
      end;
  end;

  function ShadeChannel(BaseValue, Tint: Byte): Byte;
  begin
    Result := BaseValue + ((255 - BaseValue) * Tint div 100);
  end;

  function ShadeSeriesColor(BaseColor: TColor; MonthIndex: Integer): TColor;
  const
    MONTH_TINT: array[0..11] of Byte =
      (0, 20, 40, 60, 12, 32, 52, 72, 24, 44, 64, 8);
  var
    BaseRGB: COLORREF;
    Tint: Byte;
  begin
    BaseRGB := ColorToRGB(BaseColor);
    Tint := MONTH_TINT[MonthIndex mod 12];
    Result := RGB(
      ShadeChannel(GetRValue(BaseRGB), Tint),
      ShadeChannel(GetGValue(BaseRGB), Tint),
      ShadeChannel(GetBValue(BaseRGB), Tint));
  end;

  function ItemColor(SeriesIndex, DataIndex: Integer): TColor;
  begin
    if (SeriesIndex >= 0) and (SeriesIndex <= High(Series)) and
       (DataIndex >= 0) and (DataIndex <= High(Series[SeriesIndex].Data)) then
      Result := ShadeSeriesColor(Series[SeriesIndex].Color,
        MonthColorIndex(Series[SeriesIndex].Data[DataIndex].XValue, DataIndex))
    else
      Result := COLOR_PALETTE[(DataIndex + (SeriesIndex * 5)) mod Length(COLOR_PALETTE)];
  end;

  function SeriesDisplayTitle(SeriesIndex: Integer): string;
  begin
    Result := Series[SeriesIndex].Title;
    if SameText(Result, 'Pedido de Vendas') then
      Result := 'Ped. Venda';
  end;

  function DataLegendText(SeriesIndex, DataIndex: Integer): string;
  begin
    Result := SeriesDisplayTitle(SeriesIndex) + ' - ' +
      Series[SeriesIndex].Data[DataIndex].XValue + ' - R$ ' +
      FormatYValue(Series[SeriesIndex].Data[DataIndex].YValue);
  end;

  function DashboardLegendItemCount: Integer;
  var
    I, J: Integer;
  begin
    Result := 0;
    for I := 0 to High(Series) do
      for J := 0 to High(Series[I].Data) do
        if (ChartType <> ctPie) or (Series[I].Data[J].YValue > 0) then
          Inc(Result);
  end;

  function DashboardLegendTextWidth: Integer;
  var
    I, J, TextWidthValue: Integer;
  begin
    Result := 0;
    Canvas.Font.Size := LegendFontSize;
    for I := 0 to High(Series) do
      for J := 0 to High(Series[I].Data) do
      begin
        if (ChartType = ctPie) and (Series[I].Data[J].YValue <= 0) then
          Continue;
        TextWidthValue := Canvas.TextWidth(DataLegendText(I, J));
        if TextWidthValue > Result then
          Result := TextWidthValue;
      end;
  end;

  function FitText(const Text: string; MaxWidth: Integer): string;
  begin
    Result := Text;
    while (Result <> '') and (Canvas.TextWidth(Result) > MaxWidth) do
      Delete(Result, Length(Result), 1);
    if Result <> Text then
    begin
      while (Result <> '') and (Canvas.TextWidth(Result + '...') > MaxWidth) do
        Delete(Result, Length(Result), 1);
      Result := Result + '...';
    end;
  end;

  function XLabelStep: Integer;
  var
    Slots: Integer;
  begin
    { Os dados do dashboard sao mensais; ate 12 categorias cabem nas
      areas atuais e devem permanecer visiveis no eixo X. }
    if DataCount <= 12 then
    begin
      Result := 1;
      Exit;
    end;
    Slots := Max(1, ChartArea.Width div 65);
    Result := Max(1, Ceil(DataCount / Slots));
  end;

  procedure DrawXAxisLabel(const LabelText: string; CenterX, LabelY: Integer);
  var
    LabelX, LabelWidth, MinLabelX, MaxLabelX: Integer;
  begin
    Canvas.Brush.Style := bsClear;
    Canvas.Font.Color := clBlack;
    Canvas.Font.Size := 8;
    LabelWidth := Canvas.TextWidth(LabelText);
    LabelX := CenterX - (LabelWidth div 2);
    { Mantem o texto inteiro dentro da area do grafico, sem invadir a
      legenda lateral, inclusive no primeiro e no ultimo mes. }
    MinLabelX := ChartArea.Left;
    MaxLabelX := ChartArea.Right - LabelWidth;
    if MaxLabelX < MinLabelX then
      MaxLabelX := MinLabelX;
    if LabelX < MinLabelX then
      LabelX := MinLabelX;
    if LabelX > MaxLabelX then
      LabelX := MaxLabelX;
    Canvas.TextOut(LabelX, LabelY, LabelText);
  end;

  function XAxisLabelsNeedSecondRow: Boolean;
  var
    I, SlotWidth: Integer;
  begin
    Result := False;
    if not DashboardLegend or (DataCount <= 1) then
      Exit;
    Canvas.Font.Size := 8;
    if (ChartType = ctLine) and (DataCount > 1) then
      SlotWidth := Max(1, ChartArea.Width div (DataCount - 1))
    else
      SlotWidth := Max(1, ChartArea.Width div DataCount);
    for I := 0 to High(Series[0].Data) do
      if Canvas.TextWidth(Series[0].Data[I].XValue) + 4 > SlotWidth then
      begin
        Result := True;
        Exit;
      end;
  end;

  function XAxisLabelY(Index: Integer): Integer;
  begin
    Result := ChartArea.Bottom + 8;
    if XAxisLabelsNeedSecondRow then
      Result := ChartArea.Bottom + 3 + ((Index mod 2) * 10);
  end;

  function LastNonZeroIndex(SeriesIndex: Integer): Integer;
  begin
    Result := High(Series[SeriesIndex].Data);
    while (Result > 0) and (Series[SeriesIndex].Data[Result].YValue <= 0) do
      Dec(Result);
  end;

  function ValueToY(Value: Double): Integer;
  begin
    Result := ChartArea.Bottom - Round((Value - MinY) * ScaleFactorY);
  end;

  procedure CalculateMetrics;
  var
    I, J, LeftMargin, LegendCount, RowsAvailable, MaxTextWidth: Integer;
  begin
    MinY := 0;
    MaxY := 0;
    for I := 0 to High(Series) do
      for J := 0 to High(Series[I].Data) do
      begin
        if Series[I].Data[J].YValue > MaxY then
          MaxY := Series[I].Data[J].YValue;
        if Series[I].Data[J].YValue < MinY then
          MinY := Series[I].Data[J].YValue;
      end;

    if (MaxY = 0) and (MinY = 0) then
      MaxY := 1;
    RangeY := MaxY - MinY;
    if RangeY <= 0 then
      RangeY := 1;

    Canvas.Font.Size := 8;
    LeftMargin := Max(62, Max(Canvas.TextWidth(FormatYValue(MaxY)), Canvas.TextWidth(FormatYValue(MinY))) + 16);
    LegendMonthColumns := 2;
    LegendItemHeight := 16;
    LegendSideWidth := 0;
    LegendLeft := 0;
    LegendTopPos := DEFAULT_MARGIN + IfThen(Title <> '', 42, 20);
    if DashboardLegend then
    begin
      { A dashboard tem varios graficos e muitas entradas. A legenda fica
        na lateral, liberando a faixa inferior para os eixos e botoes. }
      LegendFontSize := 7;
      LegendItemHeight := 13;
      LegendCount := DashboardLegendItemCount;
      MaxTextWidth := DashboardLegendTextWidth;
      LegendColumnWidth := Max(140, MaxTextWidth + 18);
      RowsAvailable := Max(1, (Height - LegendTopPos - 8) div LegendItemHeight);
      LegendColumns := Max(1, Ceil(LegendCount / RowsAvailable));
      LegendRows := Max(1, Ceil(LegendCount / LegendColumns));
      LegendSideWidth := LegendColumns * LegendColumnWidth + 8;
      ChartArea := Rect(LeftMargin, LegendTopPos,
        Width - LegendSideWidth - 8,
        Height - 28);
      LegendLeft := ChartArea.Right + 8;
    end
    else
    begin
      LegendRows := Max(1, Ceil(DataCount / LegendMonthColumns));
      ChartArea := Rect(LeftMargin, LegendTopPos,
        Width - 24, Height - (LegendRows * LegendItemHeight) - 28);
    end;
    if (Width < 420) or (ChartArea.Right - ChartArea.Left < 120) then
      ChartArea.Right := Width - 25;
    if ChartArea.Right <= ChartArea.Left then
      ChartArea.Right := Width - 20;
    if ChartArea.Bottom <= ChartArea.Top then
      ChartArea.Bottom := Height - 20;
    ScaleFactorY := (ChartArea.Bottom - ChartArea.Top) / RangeY;
    ZeroY := ValueToY(0);
  end;

  procedure DrawAxes;
  var I:integer;
  begin
    Canvas.Pen.Width := 1;
    Canvas.Pen.Style := psSolid;
    Canvas.Pen.Color := clBlack;
    Canvas.MoveTo(ChartArea.Left, ChartArea.Top);
    Canvas.LineTo(ChartArea.Left, ChartArea.Bottom);
    Canvas.LineTo(ChartArea.Right, ChartArea.Bottom);
    if ShowYAxisLabels then
    begin
      Canvas.Font.Size := 8;
      Canvas.Pen.Color := clSilver;
      Canvas.Pen.Style := psDot;
      for I := 0 to 5 do
      begin
        LabelValue := MinY + (I * RangeY) / 5;
        YPos := ValueToY(LabelValue);
        Canvas.MoveTo(ChartArea.Left, YPos);
        Canvas.LineTo(ChartArea.Right, YPos);
        YLabel := FormatYValue(LabelValue);
        Canvas.Brush.Style := bsClear;
        Canvas.Font.Color := clNavy;
        Canvas.TextOut(ChartArea.Left - Canvas.TextWidth(YLabel) - 8, YPos - 7, YLabel);
      end;
    end;
    if (MinY < 0) and (MaxY > 0) then
    begin
      Canvas.Pen.Style := psSolid;
      Canvas.Pen.Color := clGray;
      Canvas.MoveTo(ChartArea.Left, ZeroY);
      Canvas.LineTo(ChartArea.Right, ZeroY);
    end;
  end;

  procedure DrawLegend;
  var
    I, J, X, Y, MonthColumn, Row, ColumnWidth, LegendTop, ItemOrdinal: Integer;
    Txt: string;
  begin
    if Width < 420 then
      Exit;

    if DashboardLegend then
    begin
      if DashboardLegendItemCount = 0 then
        Exit;
      Canvas.Font.Size := LegendFontSize;
      LegendTop := LegendTopPos;
      ItemOrdinal := 0;
      for J := 0 to DataCount - 1 do
        for I := 0 to High(Series) do
        begin
          if J > High(Series[I].Data) then
            Continue;
          if (ChartType = ctPie) and (Series[I].Data[J].YValue <= 0) then
            Continue;
          MonthColumn := ItemOrdinal div LegendRows;
          Row := ItemOrdinal mod LegendRows;
          X := LegendLeft + MonthColumn * LegendColumnWidth + 4;
          Y := LegendTop + Row * LegendItemHeight;
          Canvas.Brush.Style := bsSolid;
          Canvas.Brush.Color := ItemColor(I, J);
          Canvas.Pen.Color := clGray;
          Canvas.Rectangle(X, Y, X + 10, Y + 10);
          Canvas.Brush.Style := bsClear;
          Canvas.Font.Color := clBlack;
          Txt := DataLegendText(I, J);
          { A largura foi calculada pelo TextWidth; nao truncar valores. }
          Canvas.TextOut(X + 14, Y - 1, Txt);
          Inc(ItemOrdinal);
        end;
      Exit;
    end;

    LegendTop := ChartArea.Bottom + 24;
    Canvas.Font.Size := 7;
    LegendItemHeight := Max(14, LegendItemHeight);
    LegendRows := Max(1, LegendRows);
    ColumnWidth := Max(60, (Width div LegendMonthColumns) div Max(1, Length(Series)));

    for J := 0 to DataCount - 1 do
      for I := 0 to High(Series) do
      begin
        if J > High(Series[I].Data) then
          Continue;
        if (ChartType = ctPie) and (Series[I].Data[J].YValue <= 0) then
          Continue;
        MonthColumn := J div LegendRows;
        Row := J mod LegendRows;
        X := MonthColumn * (Width div LegendMonthColumns) + I * ColumnWidth + 4;
        Y := LegendTop + Row * LegendItemHeight;
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := ItemColor(I, J);
        Canvas.Pen.Color := clGray;
        Canvas.Rectangle(X, Y, X + 10, Y + 10);
        Canvas.Brush.Style := bsClear;
        Canvas.Font.Color := clBlack;
        Txt := DataLegendText(I, J);
        Canvas.TextOut(X + 14, Y - 1, FitText(Txt, ColumnWidth - 18));
      end;
  end;

  procedure DrawBarChart;
  var
    DrawValues: Boolean;
    J,I:integer;
  begin
    DrawAxes;
    GroupWidth := Max(14, ChartArea.Width div Max(1, DataCount));
    BarWidth := Max(5, (GroupWidth - 10) div Max(1, Length(Series)));
    DrawValues := (BarWidth >= 18) and ((DataCount * Max(1, Length(Series))) <= 16);

    for J := 0 to DataCount - 1 do
    begin
      BarLeft := ChartArea.Left + J * GroupWidth + 5;
      for I := 0 to High(Series) do
      begin
        if J > High(Series[I].Data) then
          Continue;
        XPos := BarLeft + I * BarWidth;
        YPos := ValueToY(Series[I].Data[J].YValue);
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := ItemColor(I, J);
        Canvas.Pen.Color := clGray;
        Canvas.Rectangle(XPos, Min(YPos, ZeroY), XPos + BarWidth - 2, Max(YPos, ZeroY));
        if DrawValues and (Series[I].Data[J].YValue > 0) then
        begin
          S := FormatYValue(Series[I].Data[J].YValue);
          DrawValueBox(Canvas, S, XPos + Max(0, (BarWidth div 2) - (Canvas.TextWidth(S) div 2)), Max(ChartArea.Top, YPos - 22));
        end;
      end;
      if ShowXAxisLabels and (Length(Series[0].Data) > J) and ((J mod XLabelStep) = 0) then
        DrawXAxisLabel(Series[0].Data[J].XValue,
          BarLeft + (GroupWidth div 2), XAxisLabelY(J));
    end;
    DrawLegend;
  end;

  procedure DrawLineChart;
  var
    DrawAllValues, DrawThisValue: Boolean;
    I,J,LastIdx: Integer;
  begin
    DrawAxes;
    DrawAllValues := DataCount <= 8;
    for I := 0 to High(Series) do
    begin
      Canvas.Pen.Width := 3;
      Canvas.Pen.Color := ItemColor(I, 0);
      PrevX := 0;
      PrevY := 0;
      LastIdx := LastNonZeroIndex(I);
      for J := 0 to High(Series[I].Data) do
      begin
        if High(Series[I].Data) > 0 then
          XPos := ChartArea.Left + Round((J / High(Series[I].Data)) * ChartArea.Width)
        else
          XPos := ChartArea.Left + (ChartArea.Width div 2);
        YPos := ValueToY(Series[I].Data[J].YValue);
        if J > 0 then
        begin
          Canvas.Pen.Color := ItemColor(I, J);
          Canvas.MoveTo(PrevX, PrevY);
          Canvas.LineTo(XPos, YPos);
        end;
        Canvas.Brush.Style := bsSolid;
        Canvas.Brush.Color := ItemColor(I, J);
        Canvas.Pen.Color := clGray;
        Canvas.Ellipse(XPos - 4, YPos - 4, XPos + 4, YPos + 4);
        DrawThisValue := (Series[I].Data[J].YValue > 0) and (DrawAllValues or (J = LastIdx));
        if DrawThisValue then
          DrawValueBox(Canvas, FormatYValue(Series[I].Data[J].YValue), XPos + 7, Max(ChartArea.Top, YPos - 12));
        PrevX := XPos;
        PrevY := YPos;
      end;
    end;
    if ShowXAxisLabels and (DataCount > 0) then
      for J := 0 to High(Series[0].Data) do
      begin
        if (J mod XLabelStep) <> 0 then
          Continue;
        if High(Series[0].Data) > 0 then
          XPos := ChartArea.Left + Round((J / High(Series[0].Data)) * ChartArea.Width)
        else
          XPos := ChartArea.Left + (ChartArea.Width div 2);
        DrawXAxisLabel(Series[0].Data[J].XValue, XPos, XAxisLabelY(J));
      end;
    DrawLegend;
  end;

  procedure DrawPieChart;
  var
    DrawLabels: Boolean;
    J,I:integer;
  begin
    SetLength(PieData, 0);
    TotalY := 0;
    for I := 0 to High(Series) do
      for J := 0 to High(Series[I].Data) do
      begin
        if Series[I].Data[J].YValue <= 0 then
          Continue;
        K := Length(PieData);
        SetLength(PieData, K + 1);
        PieData[K].XValue := DataLegendText(I, J);
        PieData[K].YValue := Series[I].Data[J].YValue;
        PieData[K].Color := ItemColor(I, J);
        TotalY := TotalY + PieData[K].YValue;
      end;
    if TotalY <= 0 then
      Exit;

    PieSize := Min(ChartArea.Width, ChartArea.Height) - 16;
    if PieSize < 40 then
      PieSize := Min(ChartArea.Width, ChartArea.Height);
    PieLeft := ChartArea.Left + (ChartArea.Width - PieSize) div 2;
    PieTop := ChartArea.Top + (ChartArea.Height - PieSize) div 2;
    DrawLabels := (Length(PieData) <= 8) and (PieSize >= 130);
    StartAngle := 0;
    Canvas.Pen.Color := clWhite;
    for I := 0 to High(PieData) do
    begin
      EndAngle := StartAngle + (PieData[I].YValue / TotalY) * 360 * 16;
      Canvas.Brush.Style := bsSolid;
      Canvas.Brush.Color := PieData[I].Color;
      Pie(Canvas.Handle, PieLeft, PieTop, PieLeft + PieSize, PieTop + PieSize,
        PieLeft + PieSize div 2 + Round((PieSize div 2) * Cos(StartAngle * Pi / (180 * 16))),
        PieTop + PieSize div 2 - Round((PieSize div 2) * Sin(StartAngle * Pi / (180 * 16))),
        PieLeft + PieSize div 2 + Round((PieSize div 2) * Cos(EndAngle * Pi / (180 * 16))),
        PieTop + PieSize div 2 - Round((PieSize div 2) * Sin(EndAngle * Pi / (180 * 16))));
      if DrawLabels then
      begin
        MidAngle := (StartAngle + EndAngle) / 2;
        S := FormatFloat('0.0', (PieData[I].YValue / TotalY) * 100) + '%';
        DrawValueBox(Canvas, S,
          PieLeft + PieSize div 2 + Round((PieSize div 2 + 8) * Cos(MidAngle * Pi / (180 * 16))),
          PieTop + PieSize div 2 - Round((PieSize div 2 + 8) * Sin(MidAngle * Pi / (180 * 16))));
      end;
      StartAngle := EndAngle;
    end;
    DrawLegend;
  end;

begin
  Canvas.Brush.Style := bsSolid;
  Canvas.Brush.Color := BackgroundColor;
  Canvas.FillRect(Rect(0, 0, Width, Height));
  if not HasData then
    Exit;
  CalculateMetrics;
  DrawTitle(Canvas, Title, Width);
  case ChartType of
    ctBar: DrawBarChart;
    ctLine: DrawLineChart;
    ctPie: DrawPieChart;
  end;
end;

procedure GenerateChart(Canvas: TCanvas; Data: TChartDataArray;
  ChartType: TChartType; Width, Height: Integer; Title: string = '';
  ShowPercentage: Boolean = True; ShowValuesOnBars: Boolean = True;
  ShowYValuesInLegend: Boolean = False; BackgroundColor: TColor = clWhite;
  ShowXAxisLabels: Boolean = True; ShowYAxisLabels: Boolean = True);
var
  Series: TChartSeriesArray;
  I: Integer;
begin
  SetLength(Series, 1);
  Series[0].Title := Title;
  Series[0].Color := clNavy;
  SetLength(Series[0].Data, Length(Data));
  for I := 0 to High(Data) do
  begin
    Series[0].Data[I] := Data[I];
    if Series[0].Data[I].Color = 0 then
      Series[0].Data[I].Color := COLOR_PALETTE[I mod Length(COLOR_PALETTE)];
  end;
  GenerateMultiSeriesChart(Canvas, Series, ChartType, Width, Height, Title,
    BackgroundColor, ShowXAxisLabels, ShowYAxisLabels);
end;

procedure PrintPaintBoxChart(PaintBox: TPaintBox; const Title: string = '';
  Landscape: Boolean = True);
type
  THiddenControlInfo = record
    Control: TControl;
    WasVisible: Boolean;
  end;
var
  Bmp: TBitmap;
  R: TRect;
  OldOrientation: TPrinterOrientation;
  NomeImpressora: string;
  Margem, PageW, PageH, DrawW, DrawH: Integer;
  HiddenControls: array of THiddenControlInfo;
  ParentControl: TWinControl;
  SourceDC: HDC;
  PaintRect, ControlRect, IntersectionRect: TRect;
  ControlItem: TControl;
  procedure HideOverlappingButtons;
  var
    J: Integer;
  begin
    SetLength(HiddenControls, 0);
    ParentControl := PaintBox.Parent;
    if ParentControl = nil then
      Exit;
    PaintRect := PaintBox.BoundsRect;
    for J := 0 to ParentControl.ControlCount - 1 do
    begin
      ControlItem := ParentControl.Controls[J];
      if (ControlItem = PaintBox) or (not ControlItem.Visible) then
        Continue;
      { Os botoes de impressao sao controles irmaos que ficam sobre o
        PaintBox no dashboard. Somente botoes sobrepostos sao ocultados. }
      if (Pos('BUTTON', UpperCase(ControlItem.ClassName)) = 0) then
        Continue;
      ControlRect := ControlItem.BoundsRect;
      if IntersectRect(IntersectionRect, PaintRect, ControlRect) then
      begin
        SetLength(HiddenControls, Length(HiddenControls) + 1);
        HiddenControls[High(HiddenControls)].Control := ControlItem;
        HiddenControls[High(HiddenControls)].WasVisible := ControlItem.Visible;
        ControlItem.Visible := False;
      end;
    end;
  end;

  procedure RestoreOverlappingButtons;
  var
    J: Integer;
  begin
    for J := High(HiddenControls) downto 0 do
      if HiddenControls[J].Control <> nil then
        HiddenControls[J].Control.Visible := HiddenControls[J].WasVisible;
    SetLength(HiddenControls, 0);
  end;
begin
  if PaintBox = nil then
    Exit;
  if not SelecionarImpressoraPadraoWindows(NomeImpressora) then
  begin
    MessageDlg('Não foi possível imprimir o grafico: nenhuma impressora valida esta disponivel.', mtWarning, [mbOK], 0);
    Exit;
  end;
  Bmp := TBitmap.Create;
  try
    Bmp.SetSize(PaintBox.Width, PaintBox.Height);
    Bmp.Canvas.Brush.Color := clWhite;
    Bmp.Canvas.FillRect(Rect(0, 0, Bmp.Width, Bmp.Height));
    { O TPaintBox e um controle grafico e desenha no DC do pai. Oculta
      temporariamente somente botoes que estejam sobrepostos para que
      eles nao entrem na captura; o estado original e restaurado abaixo. }
    try
      HideOverlappingButtons;
      PaintBox.Invalidate;
      PaintBox.Update;
      ParentControl := PaintBox.Parent;
      if ParentControl <> nil then
      begin
        SourceDC := GetDC(ParentControl.Handle);
        if SourceDC <> 0 then
        try
          BitBlt(Bmp.Canvas.Handle, 0, 0, PaintBox.Width, PaintBox.Height,
            SourceDC, PaintBox.Left, PaintBox.Top, SRCCOPY);
        finally
          ReleaseDC(ParentControl.Handle, SourceDC);
        end
        else
          BitBlt(Bmp.Canvas.Handle, 0, 0, PaintBox.Width, PaintBox.Height,
            PaintBox.Canvas.Handle, 0, 0, SRCCOPY);
      end
      else
        BitBlt(Bmp.Canvas.Handle, 0, 0, PaintBox.Width, PaintBox.Height,
          PaintBox.Canvas.Handle, 0, 0, SRCCOPY);
    finally
      RestoreOverlappingButtons;
    end;
    if Title <> '' then
      Printer.Title := Title;
    OldOrientation := Printer.Orientation;
    try
      if Landscape then
        Printer.Orientation := poLandscape;
      try
        Printer.BeginDoc;
        try
          { Ajusta o grafico para caber na pagina sem deformar sua
            proporcao. A orientacao (inclusive paisagem) continua sendo
            definida pelo chamador, como antes. }
          Margem := 100;
          PageW := Printer.PageWidth - (Margem * 2);
          PageH := Printer.PageHeight - (Margem * 2);
          if (PageW > 0) and (PageH > 0) and
             (Bmp.Width > 0) and (Bmp.Height > 0) then
          begin
            if (Int64(PageW) * Bmp.Height) <=
               (Int64(PageH) * Bmp.Width) then
            begin
              DrawW := PageW;
              DrawH := MulDiv(PageW, Bmp.Height, Bmp.Width);
            end
            else
            begin
              DrawH := PageH;
              DrawW := MulDiv(PageH, Bmp.Width, Bmp.Height);
            end;
            R := Rect(Margem + ((PageW - DrawW) div 2),
              Margem + ((PageH - DrawH) div 2),
              Margem + ((PageW - DrawW) div 2) + DrawW,
              Margem + ((PageH - DrawH) div 2) + DrawH);
          end
          else
            R := Rect(Margem, Margem, Printer.PageWidth - Margem,
              Printer.PageHeight - Margem);
          Printer.Canvas.StretchDraw(R, Bmp);
        finally
          Printer.EndDoc;
        end;
      except
        on E: Exception do
          if E is EPrinter then
            MessageDlg('Não foi possível imprimir o grafico porque a impressora padrao nao esta disponivel.', mtWarning, [mbOK], 0)
          else
            raise;
      end;
    finally
      Printer.Orientation := OldOrientation;
    end;
  finally
    Bmp.Free;
  end;
end;

end.











