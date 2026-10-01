unit uFinancialChartPrint;

interface

uses Windows, SysUtils, Classes, Graphics, Math, Printers, Dialogs, ChartGenerator;

type
  TFinancialPrintRange = record
    StartIndex, Count: Integer;
  end;
  TFinancialPrintRanges = array of TFinancialPrintRange;

  TFinancialPrintPlan = record
    Width, Height, CategoryCount, CategoriesPerPage, PageCount: Integer;
    AxisMinimum, AxisMaximum: Double;
    CategoryPixels: Integer;
    Pages: TFinancialPrintRanges;
  end;

function PlanFinancialChartPrint(Canvas: TCanvas; const Series: TChartSeriesArray;
  ChartType: TChartType; const Options: TChartRenderOptions;
  PageWidth, PageHeight: Integer): TFinancialPrintPlan;
function RenderFinancialChartPrintPage(const Series: TChartSeriesArray;
  ChartType: TChartType; const Title: string; const Options: TChartRenderOptions;
  const Plan: TFinancialPrintPlan; PageIndex: Integer): TBitmap;
procedure PrintFinancialSeriesChart(const Series: TChartSeriesArray;
  ChartType: TChartType; const Title: string; const Options: TChartRenderOptions);

implementation

function PlanFinancialChartPrint(Canvas: TCanvas; const Series: TChartSeriesArray;
  ChartType: TChartType; const Options: TChartRenderOptions;
  PageWidth, PageHeight: Integer): TFinancialPrintPlan;
var
  I, J, LeftMargin, RightMargin, PlotWidth, PlotHeight, LegendHeight: Integer;
  CategoryWidth, W, LowWidth, HighWidth, MidWidth, VisibleCount: Integer;
  HasMarks: Boolean;
  Lines: TStringList;
  CategoryWidths: array of Integer;
  StartIndex, Count, PageIndex, PageCategoryWidth: Integer;
begin
  Result.Width := Max(1, PageWidth);
  Result.Height := Max(1, PageHeight);
  Result.CategoryCount := 0;
  if Length(Series) > 0 then Result.CategoryCount := Length(Series[0].Data);
  SetLength(Result.Pages, 0);
  SetLength(CategoryWidths, Result.CategoryCount);
  for J := 0 to High(CategoryWidths) do CategoryWidths[J] := 70;
  Result.AxisMinimum := 0;
  Result.AxisMaximum := 0;
  Result.CategoryPixels := 0;
  HasMarks := False;
  Canvas.Font.PixelsPerInch := 96;
  Canvas.Font.Name := Options.AxisFontName;
  Canvas.Font.Size := Options.AxisFontSize;
  Canvas.Font.Style := Options.AxisFontStyle;
  RightMargin := 24;
  CategoryWidth := 70;
  for I := 0 to High(Series) do
  begin
    HasMarks := HasMarks or (Series[I].MarkStyle <> cmsNone);
    for J := 0 to High(Series[I].Data) do
    begin
      Result.AxisMinimum := Min(Result.AxisMinimum, Series[I].Data[J].YValue);
      Result.AxisMaximum := Max(Result.AxisMaximum, Series[I].Data[J].YValue);
      W := Canvas.TextWidth(FormatFloat('#,##0.00', Series[I].Data[J].YValue));
      if Series[I].MarkStyle <> cmsNone then
      begin
        RightMargin := Max(RightMargin, W + 18);
        CategoryWidth := Max(CategoryWidth, (W + 16) * Max(1, Length(Series)));
        if J < Length(CategoryWidths) then
          CategoryWidths[J] := Max(CategoryWidths[J], (W + 16) * Max(1, Length(Series)));
      end;
      if ChartType in [ctBar, ctLine] then
      begin
        LowWidth := 1;
        HighWidth := Max(1, Canvas.TextWidth(Series[I].Data[J].XValue));
        while LowWidth < HighWidth do
        begin
          MidWidth := (LowWidth + HighWidth) div 2;
          Lines := WrapChartText(Canvas, Series[I].Data[J].XValue, MidWidth);
          try
            if Lines.Count <= 4 then HighWidth := MidWidth
            else LowWidth := MidWidth + 1;
          finally Lines.Free end;
        end;
        CategoryWidth := Max(CategoryWidth, LowWidth + 12);
        if J < Length(CategoryWidths) then
          CategoryWidths[J] := Max(CategoryWidths[J], LowWidth + 12);
      end;
    end;
  end;
  if (Result.AxisMinimum = 0) and (Result.AxisMaximum = 0) then
    Result.AxisMaximum := 1;
  W := Max(Canvas.TextWidth(FormatFloat('#,##0.00', Result.AxisMaximum)),
    Canvas.TextWidth(FormatFloat('#,##0.00', Result.AxisMinimum)));
  RightMargin := Max(RightMargin, (W div 2) + 8);
  LeftMargin := Max(62, W + 16);
  LegendHeight := 0;
  if Options.ShowLegend then LegendHeight := ((Length(Series) + 1) div 2) * 16;
  if ChartType = ctHorizontalBar then
  begin
    VisibleCount := HorizontalSeriesCount(Series, Options.CompactZeroSeries);
    if VisibleCount > 1 then Result.CategoryPixels := 6 + (VisibleCount * 18)
    else Result.CategoryPixels := Max(20, Canvas.TextHeight('Ag') + 6);
    PlotHeight := Max(1, Result.Height - 110 - LegendHeight);
    Result.CategoriesPerPage := Max(1, PlotHeight div Result.CategoryPixels);
  end
  else if (ChartType = ctLine) and not HasMarks then
    Result.CategoriesPerPage := Max(1, Result.CategoryCount)
  else
  begin
    PlotWidth := Max(1, Result.Width - LeftMargin - RightMargin);
    Result.CategoriesPerPage := Max(1, PlotWidth div CategoryWidth);
  end;
  if Result.CategoryCount > 0 then
    Result.CategoriesPerPage := Min(Result.CategoriesPerPage, Result.CategoryCount);
  StartIndex := 0;
  repeat
    if (ChartType = ctHorizontalBar) or ((ChartType = ctLine) and not HasMarks) then
      Count := Min(Result.CategoriesPerPage, Result.CategoryCount - StartIndex)
    else
    begin
      Count := 0;
      PageCategoryWidth := 70;
      // Long labels on one page must not reduce the capacity of every other page.
      // The longest fitting prefix gives the fewest consecutive pages: removing
      // categories from any later page cannot increase its required width.
      while StartIndex + Count < Result.CategoryCount do
      begin
        W := Max(PageCategoryWidth, CategoryWidths[StartIndex + Count]);
        if ((Count + 1) * W > PlotWidth) and (Count > 0) then Break;
        PageCategoryWidth := W;
        Inc(Count);
      end;
    end;
    PageIndex := Length(Result.Pages);
    SetLength(Result.Pages, PageIndex + 1);
    Result.Pages[PageIndex].StartIndex := StartIndex;
    Result.Pages[PageIndex].Count := Count;
    Inc(StartIndex, Count);
  until StartIndex >= Result.CategoryCount;
  Result.PageCount := Length(Result.Pages);
end;

function RenderFinancialChartPrintPage(const Series: TChartSeriesArray;
  ChartType: TChartType; const Title: string; const Options: TChartRenderOptions;
  const Plan: TFinancialPrintPlan; PageIndex: Integer): TBitmap;
var
  PageSeries: TChartSeriesArray;
  PageOptions: TChartRenderOptions;
  I, StartIndex: Integer;
  PageTitle: string;
begin
  if (PageIndex < 0) or (PageIndex >= Plan.PageCount) then
    raise ERangeError.Create('Pagina do grafico fora do intervalo.');
  PageSeries := Copy(Series, 0, Length(Series));
  StartIndex := Plan.Pages[PageIndex].StartIndex;
  for I := 0 to High(PageSeries) do
  begin
    PageSeries[I].Data := Copy(Series[I].Data, StartIndex, Plan.Pages[PageIndex].Count);
    if Options.ShowZeroSeriesValues and not ChartSeriesHasValue(Series[I]) then
      PageSeries[I].Title := PageSeries[I].Title + ' (0,00)';
  end;
  PageOptions := Options;
  PageOptions.ShowZeroSeriesValues := False;
  PageOptions.OverrideAxisScale := True;
  PageOptions.AxisMinimum := Plan.AxisMinimum;
  PageOptions.AxisMaximum := Plan.AxisMaximum;
  PageOptions.HorizontalCategoryPixels := Plan.CategoryPixels;
  PageOptions.PrintLabelLines := 4;
  PageTitle := Title;
  if Plan.PageCount > 1 then
    PageTitle := PageTitle + ' - Pagina ' + IntToStr(PageIndex + 1) + '/' + IntToStr(Plan.PageCount);
  Result := TBitmap.Create;
  try
    Result.PixelFormat := pf24bit;
    Result.SetSize(Plan.Width, Plan.Height);
    Result.Canvas.Font.PixelsPerInch := 96;
    GenerateStyledMultiSeriesChart(Result.Canvas, PageSeries, ChartType,
      Plan.Width, Plan.Height, PageTitle, PageOptions, True, True);
  except Result.Free; raise end;
end;

procedure PrintFinancialSeriesChart(const Series: TChartSeriesArray;
  ChartType: TChartType; const Title: string; const Options: TChartRenderOptions);
var
  OldOrientation: TPrinterOrientation;
  Bitmap: TBitmap;
  Measure: TBitmap;
  Plan: TFinancialPrintPlan;
  DpiX, DpiY, MarginX, MarginY, PageW, PageH, I: Integer;
begin
  Printer.Refresh;
  if Printer.Printers.Count = 0 then
  begin
    MessageDlg('Nenhuma impressora valida esta disponivel para imprimir o grafico.', mtWarning, [mbOK], 0);
    Exit;
  end;
  if (Printer.PrinterIndex < 0) or (Printer.PrinterIndex >= Printer.Printers.Count) then
    Printer.PrinterIndex := 0;
  OldOrientation := Printer.Orientation;
  Measure := TBitmap.Create;
  try
    Printer.Orientation := poLandscape;
    Printer.Title := Title;
    try
      Printer.BeginDoc;
      try
        DpiX := Max(1, GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSX));
        DpiY := Max(1, GetDeviceCaps(Printer.Canvas.Handle, LOGPIXELSY));
        MarginX := DpiX div 4;
        MarginY := DpiY div 4;
        PageW := Printer.PageWidth - (MarginX * 2);
        PageH := Printer.PageHeight - (MarginY * 2);
        if (PageW <= 0) or (PageH <= 0) then raise EPrinter.Create('Area de impressao insuficiente.');
        Plan := PlanFinancialChartPrint(Measure.Canvas, Series, ChartType, Options,
          MulDiv(PageW, 96, DpiX), MulDiv(PageH, 96, DpiY));
        for I := 0 to Plan.PageCount - 1 do
        begin
          if I > 0 then Printer.NewPage;
          Bitmap := RenderFinancialChartPrintPage(Series, ChartType, Title, Options, Plan, I);
          try
            Printer.Canvas.StretchDraw(Rect(MarginX, MarginY,
              MarginX + PageW, MarginY + PageH), Bitmap);
          finally Bitmap.Free end;
        end;
        Printer.EndDoc;
      except Printer.Abort; raise end;
    except
      on E: EPrinter do
        MessageDlg('Nao foi possivel imprimir o grafico na impressora padrao.', mtWarning, [mbOK], 0);
    end;
  finally
    Measure.Free;
    Printer.Orientation := OldOrientation;
  end;
end;

end.
