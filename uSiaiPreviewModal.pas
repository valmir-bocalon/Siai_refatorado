unit uSiaiPreviewModal;

interface

uses Vcl.Forms, RLReport;

function ExibirRelatorioModal(AOrigem: TCustomForm;
  ARelatorio: TRLCustomReport): Boolean;

implementation

uses System.Classes, RLPreviewForm;

type
  PPreviewContext = ^TPreviewContext;
  TPreviewContext = record
    Origem: TCustomForm;
    Anterior: PPreviewContext;
    AoCriarAnterior: TOnCreatePreviewFormProc;
  end;

var
  ContextoAtual: PPreviewContext;

procedure ConfigurarPreview(APreview: TRLPreviewForm);
var
  Contexto: PPreviewContext;
begin
  Contexto := ContextoAtual;
  { Uma chamada aninhada preserva o callback externo e o executa uma vez. }
  while (Contexto <> nil) and
    (@Contexto.AoCriarAnterior = @ConfigurarPreview) do
    Contexto := Contexto.Anterior;
  if (Contexto <> nil) and Assigned(Contexto.AoCriarAnterior) then
    Contexto.AoCriarAnterior(APreview);

  Contexto := ContextoAtual;
  if (Contexto <> nil) and (Contexto.Origem <> nil) and
    not (csDestroying in Contexto.Origem.ComponentState) then
  begin
    APreview.PopupMode := pmExplicit;
    APreview.PopupParent := Contexto.Origem;
  end;
end;

function ExibirRelatorioModal(AOrigem: TCustomForm;
  ARelatorio: TRLCustomReport): Boolean;
var
  Contexto: TPreviewContext;
begin
  Result := False;
  if ARelatorio = nil then
    Exit;
  if AOrigem = nil then
  begin
    Result := ARelatorio.PreviewModal;
    Exit;
  end;

  Contexto.Origem := AOrigem;
  Contexto.Anterior := ContextoAtual;
  Contexto.AoCriarAnterior := OnCreatePreviewForm;
  ContextoAtual := @Contexto;
  OnCreatePreviewForm := ConfigurarPreview;
  try
    Result := ARelatorio.PreviewModal;
  finally
    OnCreatePreviewForm := Contexto.AoCriarAnterior;
    ContextoAtual := Contexto.Anterior;
  end;
end;

end.
