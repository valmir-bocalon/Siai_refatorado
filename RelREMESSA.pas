unit RelREMESSA;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, TFlatCheckBoxUnit, CheckLst,
  DB, DBClient, Grids, DBGrids, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, ExtCtrls, RLReport;

type
  TFrm_RelREMESSA = class(TForm)
    XBanner10: TXBanner;
    Label2: TLabel;
    BTFechar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    GroupBox2: TGroupBox;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    DXBMarcEmpree: TdxButtonArround;
    DS_Remess: TDataSource;
    CDSRemess: TClientDataSet;
    DBGEmpre: TDBGrid;
    CDSRemessmarca: TIntegerField;
    CDSRemessidbanco_remessa: TLargeintField;
    CDSRemessremessa: TIntegerField;
    CDSRemessidbanco: TIntegerField;
    CDSRemessgerado: TDateField;
    CDSRemessnomearq: TStringField;
    CDSRemessarquivo: TStringField;
    dxButton1: TdxButtonArround;
    Timer1: TTimer;
    procedure BTFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DXBMarcEmpreeClick(Sender: TObject);
    procedure DBGEmpreDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGEmpreKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGEmpreMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure mudamarcaEmpreend;
    procedure BtRelatorioClick(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FIdArquivo: Int64;
    FIdConta: Int64;
    FCabecalho: TZReadOnlyQuery;
    FDetalhes: TZReadOnlyQuery;
    FLoteamento: TZReadOnlyQuery;
    FDSDetalhes: TDataSource;
    FDSLoteamento: TDataSource;
    FImprimindo: Boolean;
    FAntesImprimir: TRLBeforePrintEvent;
    FAposImprimir: TRLAfterPrintEvent;
    FInicioPreparo: UInt64;
    procedure AtualizarBotoes;
    procedure CarregarDetalhes;
    procedure VisualizarArquivo(ARetorno: Boolean);
    procedure AntesImprimir(Sender: TObject; var PrintIt: Boolean);
    procedure AposImprimir(Sender: TObject);

    procedure AfterConstruction; override;
  public
    { Public declarations }
    procedure SelecionarArquivo(AIdArquivo, AIdConta: Int64);
  end;

var
  Frm_RelREMESSA: TFrm_RelREMESSA;

implementation

uses tabelas, RelImovel01, RelImovelREMESSA, UFrmContasBancarias,
  RelImovelREMESSA2, uRuntimeFields, uSiaiPreviewModal, uSiaiPerformance, funcoes, FINANmsg;

{$R *.dfm}

procedure TFrm_RelREMESSA.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_RelREMESSA.FormShow(Sender: TObject);
var Inicio: UInt64;
begin
  Timer1.Enabled := False;
  CDSRemess.Close;
  CDSRemess.CreateDataSet;
  FCabecalho.Close;
  FCabecalho.ParamByName('idarquivo').AsLargeInt := FIdArquivo;
  FCabecalho.ParamByName('idconta').AsLargeInt := FIdConta;
  Inicio := PerformanceStart;
  try
    FCabecalho.Open;
    if not FCabecalho.IsEmpty then
    begin
      CDSRemess.Append;
      CDSRemessmarca.AsInteger := 0;
      CDSRemessremessa.AsInteger := FCabecalho.FieldByName('remessa').AsInteger;
      CDSRemessgerado.Value := FCabecalho.FieldByName('gerado').Value;
      CDSRemessidbanco_remessa.AsLargeInt := FCabecalho.FieldByName('idbanco_remessa').AsLargeInt;
      CDSRemessidbanco.AsInteger := FCabecalho.FieldByName('idbanco').AsInteger;
      CDSRemessnomearq.AsString := FCabecalho.FieldByName('nomearq').AsString;
      CDSRemess.Post;
    end;
  finally
    FCabecalho.Close;
    PerformanceElapsed('Rel remessas: buscar arquivo selecionado', Inicio);
  end;
  AtualizarBotoes;
  Timer1.Enabled := not CDSRemess.IsEmpty;
end;

procedure TFrm_RelREMESSA.DXBMarcEmpreeClick(Sender: TObject);
begin
  if DXBMarcEmpree.Caption = 'Desmarcar' Then Begin
    CDSRemess.First;
    while not CDSRemess.Eof do begin
      CDSRemess.Edit;
      CDSRemessmarca.Value := 1;
      CDSRemess.Post;
      CDSRemess.Next;
    end;
    DXBMarcEmpree.Caption := 'Marcar';
  end
  Else
  Begin
    CDSRemess.First;
    while not CDSRemess.Eof do begin
      CDSRemess.Edit;
      CDSRemessmarca.Value := 0;
      CDSRemess.Post;
      CDSRemess.Next;
    end;
    DXBMarcEmpree.Caption := 'Desmarcar';
  End;
end;

procedure TFrm_RelREMESSA.DBGEmpreDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGEmpre.DefaultDrawDataCell(Rect, DBGEmpre.columns[datacol].field, State);
  if Column.Field = CDSRemessmarca then begin
    DBGEmpre.Canvas.FillRect(Rect);
    if CDSRemessmarca.Value = 0 then
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 0)
    else
      DM_Tabelas.ImageList1.Draw(DBGEmpre.Canvas, Rect.Left + 15,Rect.Top + 1, 1);
  end;
end;

procedure TFrm_RelREMESSA.DBGEmpreKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 Then
    mudamarcaEmpreend;
end;

procedure TFrm_RelREMESSA.DBGEmpreMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  mudamarcaEmpreend;
end;

procedure TFrm_RelREMESSA.mudamarcaEmpreend;
Begin
  if DBGEmpre.SelectedField.FieldName = 'marca' Then Begin
    CDSRemess.Edit;
    if CDSRemessmarca.Value = 0 Then
      CDSRemessmarca.Value := 1
    else
      CDSRemessmarca.Value := 0;
    CDSRemess.Post;
  end;
End;

procedure TFrm_RelREMESSA.BtRelatorioClick(Sender: TObject);
begin
  VisualizarArquivo(False);
end;

procedure TFrm_RelREMESSA.dxButton1Click(Sender: TObject);
begin
  VisualizarArquivo(True);
end;

procedure TFrm_RelREMESSA.Timer1Timer(Sender: TObject);
var Posicao: Integer; Extensao: string;
begin
  //simula pressionamento dos botões
  Timer1.Enabled := False;
  if FImprimindo or not CDSRemess.Active or CDSRemess.IsEmpty then Exit;
  Posicao := Pos('.', CDSRemessnomearq.AsString);
  Extensao := UpperCase(Copy(CDSRemessnomearq.AsString, Posicao+1, 3));
  VisualizarArquivo(not ((Extensao='REM') or (Extensao='TST')));
end;

procedure TFrm_RelREMESSA.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FImprimindo then
  begin
    Action := caNone;
    Exit;
  end;
  Timer1.Enabled := False;
  Frm_RelREMESSA := nil;
  Action := caFree;
end;

procedure TFrm_RelREMESSA.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
  FCabecalho := TZReadOnlyQuery.Create(Self);
  FCabecalho.Connection := DM_Tabelas.zconeccao;
  FCabecalho.SQL.Text := 'select idbanco_remessa,remessa,idbanco,gerado,nomearq '+
    'from banco_remessa where idbanco_remessa=:idarquivo and idbanco=:idconta';
  FDetalhes := TZReadOnlyQuery.Create(Self);
  FDetalhes.Connection := DM_Tabelas.zconeccao;
  FDetalhes.SQL.Text :=
    'select m.idremessa_receb,m.remessa,m.idrec,r.idrecebimento,r.documento,'+
    'r.ordem,r.Dt_Entrada,r.Dt_Vencimento,r.Valor,r.adversa,r.idloteamento,'+
    'r.quadralote,m.Nossonumero,m.ocorrencia,m.desc_motivo,m.credito,m.dt_ocorrencia,'+
    'coalesce(p.nome_parte,r.nomeadversa) as nomeadversa '+
    'from remessa_receb m join recebimento r on m.idrec=r.idrecebimento '+
    'left join participante p on p.idpaticipante=r.adversa '+
    'where m.remessa=:idarquivo '+
    'order by r.idrecebimento,m.idremessa_receb';
  FLoteamento := TZReadOnlyQuery.Create(Self);
  FLoteamento.Connection := DM_Tabelas.zconeccao;
  FLoteamento.SQL.Text := 'select idloteamento,apelido,logotipo '+
    'from loteamento where idloteamento=:idloteamento';
  FDSDetalhes := TDataSource.Create(Self);
  FDSDetalhes.DataSet := FDetalhes;
  FDSLoteamento := TDataSource.Create(Self);
  FDSLoteamento.DataSet := FLoteamento;
end;

procedure TFrm_RelREMESSA.SelecionarArquivo(AIdArquivo, AIdConta: Int64);
begin
  FIdArquivo := AIdArquivo;
  FIdConta := AIdConta;
end;

procedure TFrm_RelREMESSA.AtualizarBotoes;
var Extensao: string; Posicao: Integer; Disponivel: Boolean;
begin
  Disponivel := CDSRemess.Active and not CDSRemess.IsEmpty and not FImprimindo;
  Extensao := '';
  if Disponivel then
  begin
    Posicao := Pos('.', CDSRemessnomearq.AsString);
    Extensao := UpperCase(Copy(CDSRemessnomearq.AsString, Posicao+1, 3));
  end;
  BtRelatorio.Enabled := Disponivel and ((Extensao='REM') or (Extensao='TST'));
  dxButton1.Enabled := Disponivel and not ((Extensao='REM') or (Extensao='TST'));
  BTFechar.Enabled := not FImprimindo;
end;

procedure TFrm_RelREMESSA.CarregarDetalhes;
var Inicio: UInt64; Campo: TField;
begin
  FDetalhes.Close;
  FDetalhes.ParamByName('idarquivo').AsLargeInt := CDSRemessidbanco_remessa.AsLargeInt;
  Inicio := PerformanceStart;
  FDetalhes.Open;
  FDetalhes.FetchAll;
  Campo := FDetalhes.FieldByName('Valor');
  if Campo is TNumericField then
    TNumericField(Campo).DisplayFormat := '###,###,##0.00';
  for Campo in FDetalhes.Fields do
    if Campo is TDateTimeField then
      TDateTimeField(Campo).DisplayFormat := 'dd/mm/yyyy';
  FDetalhes.First;
  PerformanceElapsed('Rel remessas: detalhes do arquivo selecionado', Inicio);
  FLoteamento.Close;
  if FDetalhes.IsEmpty then
    FLoteamento.ParamByName('idloteamento').AsInteger := 0
  else
    FLoteamento.ParamByName('idloteamento').Value := FDetalhes.FieldByName('idloteamento').Value;
  Inicio := PerformanceStart;
  FLoteamento.Open;
  PerformanceElapsed('Rel remessas: logotipo do empreendimento', Inicio);
end;

procedure TFrm_RelREMESSA.AntesImprimir(Sender: TObject; var PrintIt: Boolean);
begin
  FInicioPreparo := PerformanceStart;
  if Assigned(FAntesImprimir) then FAntesImprimir(Sender, PrintIt);
end;

procedure TFrm_RelREMESSA.AposImprimir(Sender: TObject);
begin
  if Assigned(FAposImprimir) then FAposImprimir(Sender);
  PerformanceElapsed('Rel remessas: preparar paginas do relatorio', FInicioPreparo);
end;

procedure TFrm_RelREMESSA.VisualizarArquivo(ARetorno: Boolean);
var ModeloRemessa: TRLReportREMESSA; ModeloRetorno: TRLReportREMESSA2;
    Relatorio: TRLReport; Modelo: TForm; Aviso: TFormMensagem;
begin
  Timer1.Enabled := False;
  if FImprimindo or not CDSRemess.Active or CDSRemess.IsEmpty then Exit;
  FImprimindo := True;
  Modelo := nil;
  try
    AtualizarBotoes;
    CDSRemess.Filtered := False;
    CDSRemess.Filter := 'marca=0';
    CDSRemess.Filtered := True;
    if CDSRemess.IsEmpty then Exit;
    CarregarDetalhes;
    if FDetalhes.IsEmpty then
    begin
      Aviso := TFormMensagem.Create(Self);
      try
        Aviso.Frase.Caption := 'Nenhum título disponível para este relatório.';
        AbrirModal(Self, Aviso);
      finally
        Aviso.Free;
      end;
      Exit;
    end;
    if ARetorno then
    begin
      ModeloRetorno := TRLReportREMESSA2.Create(Self);
      Modelo := ModeloRetorno;
      ModeloRetorno.ConfigurarFontes(FDSDetalhes, FDSLoteamento);
      ModeloRetorno.RLLabel2.Caption := CDSRemessnomearq.AsString;
      Relatorio := ModeloRetorno.RLReportREMESSA2;
    end
    else
    begin
      ModeloRemessa := TRLReportREMESSA.Create(Self);
      Modelo := ModeloRemessa;
      ModeloRemessa.ConfigurarFontes(FDSDetalhes, FDSLoteamento);
      ModeloRemessa.RLLabel2.Caption := CDSRemessnomearq.AsString;
      Relatorio := ModeloRemessa.RLReportREMESSA;
    end;
    FAntesImprimir := Relatorio.BeforePrint;
    FAposImprimir := Relatorio.AfterPrint;
    Relatorio.BeforePrint := AntesImprimir;
    Relatorio.AfterPrint := AposImprimir;
    ExibirRelatorioModal(Self, Relatorio);
  finally
    Modelo.Free;
    FAntesImprimir := nil;
    FAposImprimir := nil;
    FLoteamento.Close;
    FDetalhes.Close;
    CDSRemess.Filtered := False;
    FImprimindo := False;
    AtualizarBotoes;
    if PerformanceEnabled then FlushPerformanceLog;
  end;
end;

end.
