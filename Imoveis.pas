unit Imoveis;

interface

uses ButtonDxArround, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, XBanner, ExtCtrls, Grids, DBGrids,
  DBCtrls, Mask, XDBNum, ComCtrls, TabNotBk, DB, XNum,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, XLabel3D;

type
  TFrm_Imoveis = class(TForm)
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButtonArround;
    BtEditar: TdxButtonArround;
    BtIncluir: TdxButtonArround;
    BtCancelar: TdxButtonArround;
    BtGravar: TdxButtonArround;
    BTPesquisar: TdxButtonArround;
    BtRelatorio: TdxButtonArround;
    BtUltimo: TdxButtonArround;
    BtProximo: TdxButtonArround;
    BTAnterior: TdxButtonArround;
    BtPrimeiro: TdxButtonArround;
    BtProcessar: TdxButtonArround;
    Label2: TLabel;
    DBGImovel: TDBGrid;
    Pag_Imovel: TTabbedNotebook;
    XBanner10: TXBanner;
    Label1: TLabel;
    DBEcod: TDBEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    DBELote: TDBEdit;
    Label6: TLabel;
    Label7: TLabel;
    DBEFernte: TDBEdit;
    Label8: TLabel;
    DBEFundo: TDBEdit;
    Label9: TLabel;
    DBEDireita: TDBEdit;
    Label10: TLabel;
    DBEEsquerda: TDBEdit;
    Label11: TLabel;
    Label19: TLabel;
    Label13: TLabel;
    DBEConfFrente: TDBEdit;
    Label14: TLabel;
    DBEConfFundo: TDBEdit;
    Label15: TLabel;
    DBEConfEsqerda: TDBEdit;
    Label16: TLabel;
    DBEConfDireita: TDBEdit;
    Label20: TLabel;
    DBEConfEsquina: TDBEdit;
    XBanner1: TXBanner;
    XBanner2: TXBanner;
    Label21: TLabel;
    DBEObs: TDBMemo;
    DBEValor: TXDBNumEdit;
    DBImage1: TDBImage;
    DBELinha: TDBEdit;
    Label17: TLabel;
    Label18: TLabel;
    DBECurva: TDBEdit;
    Eloteamento: TEdit;
    BTFechar: TdxButtonArround;
    DBCBLado: TDBComboBox;
    DBCBDisp: TDBComboBox;
    Label12: TLabel;
    DBCBQuadra: TDBComboBox;
    DBEArea: TXDBNumEdit;
    Registros: TLabel;
    LReg: TLabel;
    Label22: TLabel;
    DBCBMatricula: TDBComboBox;
    DBEMATRICULA: TDBEdit;
    Label23: TLabel;
    ZQRecebimento: TZQuery;
    DS_Recebimento: TDataSource;
    XBanner3: TXBanner;
    XLabel3D1: TXLabel3D;
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    Label24: TLabel;
    XDBNumEdit1: TXDBNumEdit;
    procedure AtivaCampos;
    procedure DesativaCampos;
    procedure BTFecharClick(Sender: TObject);
    procedure Atualiza;
    procedure botoes;
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtProcessarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure BTPesquisarClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGImovelEnter(Sender: TObject);
    procedure DBGImovelMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGImovelKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EloteamentoExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBEObsExit(Sender: TObject);
    procedure DBGImovelDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBCBLadoExit(Sender: TObject);
    procedure DBCBDispExit(Sender: TObject);
    procedure DBCBMatriculaExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Imoveis: TFrm_Imoveis;

implementation

uses tabelas, funcoes, AchaImovel, RelImovel, Achaimoveis, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Imoveis.AtivaCampos;
Begin
  Eloteamento.ReadOnly    := False;
  DBCBQuadra.ReadOnly     := False;
  DBELote.ReadOnly        := False;
  DBEArea.ReadOnly        := False;
  DBEValor.ReadOnly       := False;
  DBEFernte.ReadOnly      := False;
  DBEFundo.ReadOnly       := False;
  DBEEsquerda.ReadOnly    := False;
  DBEDireita.ReadOnly     := False;
  DBCBLado.ReadOnly       := False;
  DBCBDisp.ReadOnly       := False;
  DBEObs.ReadOnly         := False;
  DBEConfFrente.ReadOnly  := False;
  DBEConfFundo.ReadOnly   := False;
  DBEConfEsqerda.ReadOnly := False;
  DBEConfDireita.ReadOnly := False;
  DBEConfEsquina.ReadOnly := False;
  DBELinha.ReadOnly       := False;
  DBECurva.ReadOnly       := False;
  BTPrimeiro.Enabled      := False;
  BTAnterior.Enabled      := False;
  BTProximo.Enabled       := False;
  BtUltimo.Enabled        := False;
  BtGravar.Enabled        := True;
  BtCancelar.Enabled      := True;
  BtProcessar.Enabled     := False;
  BtIncluir.Enabled       := False;
  BtEditar.Enabled        := False;
  BtExcluir.Enabled       := False;
  BtPesquisar.Enabled     := False;
  BtRelatorio.Enabled     := False;
  DBGImovel.Enabled       := False;
end;

procedure TFrm_Imoveis.DesativaCampos;
Begin
  Eloteamento.ReadOnly    := True;
  DBCBQuadra.ReadOnly     := True;
  DBELote.ReadOnly        := True;
  DBEArea.ReadOnly        := True;
  DBEValor.ReadOnly       := True;
  DBEFernte.ReadOnly      := True;
  DBEFundo.ReadOnly       := True;
  DBEEsquerda.ReadOnly    := True;
  DBEDireita.ReadOnly     := True;
  DBCBLado.ReadOnly       := False;
  DBCBDisp.ReadOnly       := False;
  DBEObs.ReadOnly         := True;
  DBEConfFrente.ReadOnly  := True;
  DBEConfFundo.ReadOnly   := True;
  DBEConfEsqerda.ReadOnly := True;
  DBEConfDireita.ReadOnly := True;
  DBEConfEsquina.ReadOnly := True;
  DBELinha.ReadOnly       := True;
  DBECurva.ReadOnly       := True;
  BTPrimeiro.Enabled      := true;
  BTAnterior.Enabled      := true;
  BTProximo.Enabled       := true;
  BtUltimo.Enabled        := true;
  BtGravar.Enabled        := False;
  BtCancelar.Enabled      := False;
  BtProcessar.Enabled     := true;
  BtIncluir.Enabled       := true;
  BtEditar.Enabled        := true;
  BtExcluir.Enabled       := true;
  BtPesquisar.Enabled     := true;
  BtRelatorio.Enabled     := true;
  DBGImovel.Enabled       := True;
end;

procedure TFrm_Imoveis.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Imoveis.Atualiza;
Begin
  Eloteamento.Text := DM_Tabelas.ZQImovel.FieldByName('nomeloteamento').AsString;
  LReg.Caption := inttostr(DM_Tabelas.ZQImovel.RecordCount);
End;

procedure TFrm_Imoveis.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZQImovel.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZQImovel.RecNo=DM_Tabelas.ZQImovel.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  Atualiza;
End;

procedure TFrm_Imoveis.BtPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQImovel.First;
  botoes;
end;

procedure TFrm_Imoveis.BTAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQImovel.Prior;
  botoes;
end;

procedure TFrm_Imoveis.BtProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQImovel.Next;
  botoes;
end;

procedure TFrm_Imoveis.BtUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQImovel.Last;
  botoes;
end;

procedure TFrm_Imoveis.BtGravarClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Gravar inclusão ou edição','Imóvel: '+DM_Tabelas.ZQimovel.FieldByName('idimovel').Text+' - Qd.'+DM_Tabelas.ZQimovel.FieldByName('quadra').AsString+' - Lt.'+DM_Tabelas.ZQimovel.FieldByName('lote').AsString) Then Exit;
  DM_tabelas.ZQImovel.Post;
  DesativaCampos;
  botoes;
end;

procedure TFrm_Imoveis.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Cancelar inclusão ou edição','Imóvel: '+DM_Tabelas.ZQimovel.FieldByName('idimovel').Text+' - Qd.'+DM_Tabelas.ZQimovel.FieldByName('quadra').AsString+' - Lt.'+DM_Tabelas.ZQimovel.FieldByName('lote').AsString) Then Exit;
  DM_tabelas.ZQImovel.Cancel;
  DesativaCampos;
  botoes;
end;

procedure TFrm_Imoveis.BtProcessarClick(Sender: TObject);
begin
//  processo
end;

procedure TFrm_Imoveis.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Incluir','') Then Exit;
  DM_tabelas.ZQImovel.Insert;
  AtivaCampos;
  Pag_Imovel.PageIndex := 0;
  Eloteamento.Clear;
  Eloteamento.SetFocus;
End;

procedure TFrm_Imoveis.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Editar','Imóvel: '+DM_Tabelas.ZQimovel.FieldByName('idimovel').Text+' - Qd.'+DM_Tabelas.ZQimovel.FieldByName('quadra').AsString+' - Lt.'+DM_Tabelas.ZQimovel.FieldByName('lote').AsString) Then Exit;
  DM_tabelas.ZQImovel.Edit;
  AtivaCampos;
  if Pag_Imovel.PageIndex = 0 Then
    Eloteamento.SetFocus
  else if Pag_Imovel.PageIndex = 1 Then
    DBEConfFrente.SetFocus;

end;

procedure TFrm_Imoveis.BTExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Excluir','Imóvel: '+DM_Tabelas.ZQimovel.FieldByName('idimovel').Text+' - Qd.'+DM_Tabelas.ZQimovel.FieldByName('quadra').AsString+' - Lt.'+DM_Tabelas.ZQimovel.FieldByName('lote').AsString) Then Exit;
  if simnao('Confirma a exclusão do imóvel ?','SIM') then
  begin
    if DM_Tabelas.ZQImovel.RecordCount>0 Then
      DM_tabelas.ZQImovel.Delete;
    botoes;
  end;
end;

procedure TFrm_Imoveis.BTPesquisarClick(Sender: TObject);
begin
{  Frm_AchaImovel.Top := Frm_Imoveis.Top+105;
  Frm_AchaImovel.Left := Frm_Imoveis.Left+(Frm_Imoveis.Width-Frm_AchaImovel.Width)  ;
  Frm_AchaImovel.showmodal;
  DBGImovel.SetFocus;}

  Frm_AchaImoveis.Top := Frm_Imoveis.Top+105;
  Frm_AchaImoveis.Left := Frm_Imoveis.Left+(Frm_Imoveis.Width-Frm_AchaImoveis.Width)  ;
  Frm_AchaImoveis.showmodal;
  DBGImovel.SetFocus;


end;

procedure TFrm_Imoveis.BtRelatorioClick(Sender: TObject);
begin
  if not Verif_senha('Imóvel','Relatório','') Then Exit;
//  Frm_RelImovel.Top := Frm_Imoveis.Top+105;
//  Frm_RelImovel.Left := Frm_Imoveis.Left-5+(Frm_Imoveis.Width-Frm_RelImovel.Width);
  if Frm_RelImovel=nil then
     Frm_RelImovel:=TFrm_RelImovel.Create(Application);
  Frm_RelImovel.showmodal;
  Frm_RelImovel:=nil;
end;

procedure TFrm_Imoveis.FormShow(Sender: TObject);
begin
  DesativaCampos;
  Pag_Imovel.PageIndex := 0;                           
  DM_Tabelas.ZQImovel.open;
  DM_Tabelas.ZQQuadras.open;
  DM_Tabelas.ZQLoteamento.open;
  DM_Tabelas.ZQRescisao.open;
  ZQRecebimento.Close;                                                                         
  ZQRecebimento.SQL.Clear;
  ZQRecebimento.SQL.Add('Select idrecebimento,Dt_Entrada,Dt_Vencimento,saldo,recpag,numordem,venda_idvenda,quadralote,cliente,usuario,RefBaixa,refvinda,empresa,origem,adversa,idloteamento,numboleto,');
  ZQRecebimento.SQL.Add('sum(valor) as VRRECEBIDO from Recebimento where quadralote='+quotedstr(DM_Tabelas.ZQRescisao.FieldByName('quadralote').AsString)+' and saldo=0 order by DT_Vencimento');
  ZQRecebimento.open;
  DBGImovel.SetFocus;
  if DM_Tabelas.ZQimovel.FieldByName('matri').AsString='SIM' then
  begin
    Label23.Visible:=True;
    DBEMATRICULA.Visible:=True;
  end
  else
  begin
    Label23.Visible:=False;
    DBEMATRICULA.Visible:=False;
  end;
end;

procedure TFrm_Imoveis.DBGImovelEnter(Sender: TObject);
begin
  botoes;
end;

procedure TFrm_Imoveis.DBGImovelMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  botoes;
end;

procedure TFrm_Imoveis.DBGImovelKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes;
end;

procedure TFrm_Imoveis.EloteamentoExit(Sender: TObject);
var
  varquadra : string;
begin
  if (DM_Tabelas.ZQImovel.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_loteamento(Eloteamento.Text)  then
      Eloteamento.SetFocus
    else Begin
      Eloteamento.Text := DM_tabelas.ZQAchaLotea.FieldByName('apelido').AsString;
      DM_Tabelas.ZQimovel.FieldByName('loteamento_idloteamento').AsLargeInt := DM_Tabelas.ZQAchaLotea.FieldByName('idloteamento').AsLargeInt;
      DM_Tabelas.ZQLoteamento.Locate('idloteamento',DM_Tabelas.ZQAchaLotea.FieldByName('idloteamento').AsLargeInt,[]);
      varquadra := DBCBQuadra.Text;
      DBCBQuadra.Clear;
      DM_Tabelas.ZQQuadras.First;
      while not DM_Tabelas.ZQQuadras.Eof do begin
        DBCBQuadra.Items.Add(DM_Tabelas.ZQquadras.FieldByName('descricao').AsString);
        DM_Tabelas.ZQQuadras.Next;
      end;
      DBCBQuadra.Text := varquadra;

    end;
  end;
end;

procedure TFrm_Imoveis.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Imoveis.DBEObsExit(Sender: TObject);
begin
  if DM_tabelas.ZQImovel.State in [DsInsert, DsEdit] Then Begin
    Pag_Imovel.PageIndex := 1;
    DBEConfFrente.SetFocus;
  end;
end;

procedure TFrm_Imoveis.DBGImovelDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = DM_Tabelas.ZQimovel.FieldByName('idimovel').Text Then Begin
    DBGImovel.Canvas.Brush.Color :=$006CFFFF;
    DBGImovel.Canvas.Font.Color := $00A80000;
    DBGImovel.Canvas.Font.Style := [FsBold];
  end;

  if DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString = 'SIM' then
  begin
    DBGImovel.Canvas.Font.Color := clGreen;
  end;
  if DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString <> 'SIM' then
  begin
    DBGImovel.Canvas.Font.Color := clRed;
  end;

  DBGImovel.DefaultDrawDataCell(Rect, DBGImovel.columns[datacol].field, State);

end;

procedure TFrm_Imoveis.DBCBLadoExit(Sender: TObject);
begin
  IF (DM_Tabelas.ZQImovel.State in [DsInsert, DsEdit]) and (empty(DM_Tabelas.ZQimovel.FieldByName('lado').AsString)) Then
    DBCBLado.SetFocus;
end;

procedure TFrm_Imoveis.DBCBDispExit(Sender: TObject);
begin
  IF (DM_Tabelas.ZQImovel.State in [DsInsert, DsEdit]) and (empty(DM_Tabelas.ZQimovel.FieldByName('disponivel').AsString)) Then
    DBCBLado.SetFocus;

end;

procedure TFrm_Imoveis.DBCBMatriculaExit(Sender: TObject);
begin
  if DBCBMatricula.text='SIM' then
  begin
    Label23.Visible:=True;
    DBEMATRICULA.Visible:=True;
    DBEMATRICULA.SetFocus;
  end
  else
  if DBCBMatricula.text<>'SIM' then
  begin
    Label23.Visible:=False;
    DBEMATRICULA.Visible:=False;
  end;
end;

procedure TFrm_Imoveis.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQImovel.close;
  DM_Tabelas.ZQQuadras.close;
  DM_Tabelas.ZQLoteamento.close;
  DM_Tabelas.ZQRescisao.Close;
  ZQRecebimento.close;
  Frm_Imoveis:=nil;
  Action:=CaFree;
end;

procedure TFrm_Imoveis.DBGrid2DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
   R   : TRect;
begin
  R := Rect;
  Dec(R.Bottom,2);
  if Column.Field = DM_Tabelas.ZQRescisao.FieldByName('historico')  then
  begin
    if not (gdSelected in State)
    then DBGrid2.Canvas.FillRect(Rect)
    else DBGrid2.Canvas.FillRect(Rect);
    {Desenha o texto do campo}
    DrawText(DBGrid2.Canvas.Handle,
    pchar(DM_Tabelas.ZQRescisao.FieldByName('historico').AsString),Length(DM_Tabelas.ZQRescisao.FieldByName('historico').AsString),R,DT_WORDBREAK);
  end;
end;


procedure TFrm_Imoveis.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Imoveis, 'ZQRecebimento', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentonomecli', 'nomecli', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoadversanome', 'adversanome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentonome_loteamento', 'nome_loteamento', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Imoveis, 'ZQRecebimento', 'ZQRecebimentoVRRECEBIDO', 'VRRECEBIDO', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);

end.
