unit UFrmContasBancarias;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,   dxButton, XBanner, StdCtrls, ExtCtrls, Mask, DBCtrls,
  JvExControls, JvButton, JvNavigationPane, Grids, DBGrids, DB, XDBNum,
  XDBDate, wwdbdatetimepicker, TFlatPanelUnit, ComCtrls, TabNotBk,
  TFlatGroupBoxUnit, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  dxCore2, XNum, XDate,acbrutil;

type
  TFrmCad_ContasBancarias = class(TForm)
    Panel1: TPanel;
    DXBExcluir: TdxButton;
    DXBEditar: TdxButton;
    DXBCancelar: TdxButton;
    DXBGravar: TdxButton;
    DXBPesquisar: TdxButton;
    DXBRelatorios: TdxButton;
    DXBFechar: TdxButton;
    DXBUltimo: TdxButton;
    DXBProximo: TdxButton;
    DXBAnterior: TdxButton;
    DXBPrimeiro: TdxButton;
    DXBIncluir: TdxButton;
    Label1: TLabel;
    DBGConta_Bancaria: TDBGrid;
    PagContaBanco: TTabbedNotebook;
    FlatPanel1: TFlatPanel;
    FlatPanel2: TFlatPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    JNBMostrasennha: TJvNavPanelButton;
    Label11: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    DBECodConta: TDBEdit;
    DBEBanco: TDBEdit;
    DBENomeBanco: TDBEdit;
    DBEAgencia: TDBEdit;
    DBEConta: TDBEdit;
    DBENomeTit: TDBEdit;
    DBESenhainterna: TDBEdit;
    DBESenhaBancaria: TDBEdit;
    GroupBox1: TGroupBox;
    DBGlimiteEspecial: TDBGrid;
    DXBIncluirEspecial: TdxButton;
    DXBExcluirEspecial: TdxButton;
    DXBEditarEspecial: TdxButton;
    XNEvalorespecial: TXDBNumEdit;
    DXBGravarespecial: TdxButton;
    DXBCancelarEspecial: TdxButton;
    DBEDataValida: TwwDBDateTimePicker;
    dbcbativa: TDBComboBox;
    DBEAbertura: TwwDBDateTimePicker;
    DBECodnoBanco: TDBEdit;
    DBENAgeV: TDBEdit;
    DBENContaV: TDBEdit;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    XDBEJuros: TXDBNumEdit;
    FGBMora: TFlatGroupBox;
    RBFixo: TRadioButton;
    RBVariavel: TRadioButton;
    Label19: TLabel;
    XDBNumEdit2: TXDBNumEdit;
    Label18: TLabel;
    DBECarteira: TDBEdit;
    DBGRemessa: TDBGrid;
    DBGrid2: TDBGrid;
    Label20: TLabel;
    DBCBBoleto: TDBComboBox;
    Label21: TLabel;
    XBanner2: TXBanner;
    XBanner4: TXBanner;
    XBanner5: TXBanner;
    XBanner6: TXBanner;
    XBanner1: TXBanner;
    XBanner3: TXBanner;
    Label12: TLabel;
    Eparticipante: TEdit;
    DBESeq: TDBEdit;
    dxButtonREMESSA: TdxButton;
    numero: TLabel;
    Label22: TLabel;
    DBEdit1: TDBEdit;
    psnum: TMaskEdit;
    Label23: TLabel;
    ZQRemes_Receb: TZQuery;
    DS_Remes_Receb: TDataSource;










    ZQBancRemes: TZQuery;







    DS_BancRemes: TDataSource;
    ZQContaBancaria: TZQuery;






















    DS_ContaBancaria: TDataSource;
    DBEdit2: TDBEdit;
    Label24: TLabel;
    Label25: TLabel;
    DBEdit3: TDBEdit;
    Label26: TLabel;
    DBEdit4: TDBEdit;




    Label27: TLabel;
    DBEdit5: TDBEdit;
    Label28: TLabel;
    DBEdit6: TDBEdit;
    Label29: TLabel;
    DBEdit7: TDBEdit;



    Label30: TLabel;
    XDBNumEdit1: TXDBNumEdit;
    Label31: TLabel;
    DBEdit8: TDBEdit;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    DBEdit9: TDBEdit;


    Label35: TLabel;
    Label36: TLabel;
    XDBNumEdit3: TXDBNumEdit;
    XDBDateEdit1: TXDBDateEdit;


    Label37: TLabel;
    DBEdit10: TDBEdit;
    Label38: TLabel;
    DBEdit11: TDBEdit;


    procedure botoes_setas;
    procedure Atualiza_pagina;
    procedure Filtra_limite;
    procedure desativando_botoes_em_edicao_limite;
    procedure Ativar_campos;
    procedure Desativar_campos;
    procedure DXBPrimeiroClick(Sender: TObject);
    procedure DXBAnteriorClick(Sender: TObject);
    procedure DXBProximoClick(Sender: TObject);
    procedure DXBUltimoClick(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DXBFecharClick(Sender: TObject);
    procedure DXBIncluirClick(Sender: TObject);
    procedure DXBEditarClick(Sender: TObject);
    procedure DXBExcluirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure JNBMostrasennhaClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBESenhaBancariaExit(Sender: TObject);
    procedure DXBGravarespecialClick(Sender: TObject);
    procedure DXBCancelarEspecialClick(Sender: TObject);
    procedure DXBIncluirEspecialClick(Sender: TObject);
    procedure DXBEditarEspecialClick(Sender: TObject);
    procedure DXBExcluirEspecialClick(Sender: TObject);
    procedure DBGConta_BancariaCellClick(Column: TColumn);
    procedure DBGConta_BancariaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGConta_BancariaDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure RBFixoClick(Sender: TObject);
    procedure DBGRemessaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGRemessaKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EparticipanteExit(Sender: TObject);
    procedure DBGRemessaDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure dxButtonREMESSAClick(Sender: TObject);
    procedure psnumExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBECarteiraExit(Sender: TObject);
    procedure DXBRelatoriosClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBEdit8Enter(Sender: TObject);

  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  FrmCad_ContasBancarias: TFrmCad_ContasBancarias;

implementation

uses tabelas, funcoes, RelREMESSA, RelContaBancaria, uRuntimeFields;

{$R *.dfm}

procedure TFrmCad_ContasBancarias.botoes_setas;
Begin
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  if DM_Tabelas.ZQContaBancaria.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  if DM_Tabelas.ZQContaBancaria.RecNo = DM_Tabelas.ZQContaBancaria.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
End;

procedure TFrmCad_ContasBancarias.Atualiza_pagina;
Begin
  Eparticipante.Text := DM_Tabelas.ZQContaBancaria.FieldByName('nomeparticip').AsString;
  if DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString = 'F' Then Begin
    RBFixo.Checked := True;
    Label19.Caption := 'Valor. Fixo';
  end
  else Begin
    RBVariavel.Checked := True;
    Label19.Caption :='% p/ Calculo do Vlr. Mora Di�ria';
  end;
//  Label20.Caption := INTTOSTR(DM_Tabelas.ZQRemes_Receb.RecordCount);
  botoes_setas;
End;

procedure TFrmCad_ContasBancarias.Filtra_limite;
Begin
  if DM_tabelas.ZQContaBancaria.RecordCount>0 Then Begin
    DM_tabelas.ZQEspecial.SQL.Clear;
    DM_tabelas.ZQEspecial.SQL.Add('select idespecial,data_valida,Conta_Bancaria_cod_banco,valor from limite_especial where conta_bancaria_cod_banco='+DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text);
    DM_tabelas.ZQEspecial.Open;
  end;
End;

procedure TFrmCad_ContasBancarias.desativando_botoes_em_edicao_limite;
Begin
  DBGlimiteEspecial.Enabled := False;
  DBGConta_Bancaria.Enabled := False;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DXBGravar.Enabled := false;
  DXBCancelar.Enabled := false;
  DXBGravarespecial.Enabled := true;
  DXBCancelarEspecial.Enabled := true;
  DXBIncluirEspecial.Enabled := false;
  DXBEditarEspecial.Enabled := false;
  DXBExcluirEspecial.Enabled := false;
  DBEDataValida.readOnly := false;
  XNEvalorespecial.readOnly := false;
end;

procedure TFrmCad_ContasBancarias.Ativar_campos;
Begin
  Eparticipante.ReadOnly := False;
  DBGConta_Bancaria.Enabled := False;
  DBEBanco.ReadOnly := False;
  DBENomeBanco.ReadOnly := False;
  XDBEJuros.ReadOnly := False;
  DBECarteira.ReadOnly := False;
  DBEAgencia.ReadOnly := False;
  DBEConta.ReadOnly := False;
  DBENAgeV.ReadOnly := False;
  DBENContaV.ReadOnly := False;
  DBECodnoBanco.ReadOnly := False;
  DBEAbertura.ReadOnly := False;
  DBENomeTit.ReadOnly := False;
  FGBMora.Enabled := True;
  DBCBBoleto.ReadOnly := False;
  DBESenhainterna.ReadOnly := False;
  DBESenhaBancaria.ReadOnly := False;
  dbcbativa.ReadOnly := False;
  DXBPrimeiro.Enabled := False;
  DXBAnterior.Enabled := False;
  DXBProximo.Enabled := False;
  DXBUltimo.Enabled := False;
  DXBIncluir.Enabled := False;
  DXBEditar.Enabled := False;
  DXBExcluir.Enabled := False;
  DXBGravar.Enabled := True;
  DXBCancelar.Enabled := True;
  DXBGravarespecial.Enabled := false;
  DXBCancelarEspecial.Enabled := false;
  DXBIncluirEspecial.Enabled := false;
  DXBEditarEspecial.Enabled := false;
  DXBExcluirEspecial.Enabled := false;
End;

procedure TFrmCad_ContasBancarias.Desativar_campos;
Begin
  Eparticipante.ReadOnly := True;
  DBGlimiteEspecial.Enabled := True;
  DBGConta_Bancaria.Enabled := True;
  FGBMora.Enabled := False;
  dbcbativa.ReadOnly := True;
  DBECarteira.ReadOnly := True;
  DBEBanco.ReadOnly := True;
  XDBEJuros.ReadOnly := True;
  DBCBBoleto.ReadOnly := True;
  DBENomeBanco.ReadOnly := True;
  DBEAgencia.ReadOnly := True;
  DBENAgeV.ReadOnly := True;
  DBENContaV.ReadOnly := True;
  DBECodnoBanco.ReadOnly := True;
  DBEConta.ReadOnly := True;
  DBEAbertura.ReadOnly := True;
  DBENomeTit.ReadOnly := True;
  DBESenhainterna.ReadOnly := True;
  DBESenhaBancaria.ReadOnly := True;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := False;
  DXBPrimeiro.Enabled := True;
  DXBAnterior.Enabled := True;
  DXBProximo.Enabled := True;
  DXBUltimo.Enabled := True;
  DXBIncluir.Enabled := True;
  DXBEditar.Enabled := True;
  DXBExcluir.Enabled := True;
  DXBGravarespecial.Enabled := false;
  DXBCancelarEspecial.Enabled := false;
  DXBIncluirEspecial.Enabled := true;
  DXBEditarEspecial.Enabled := true;
  DXBExcluirEspecial.Enabled := true;
  DBEDataValida.readOnly := true;
  XNEvalorespecial.readOnly := true;
End;

procedure TFrmCad_ContasBancarias.DXBPrimeiroClick(Sender: TObject);
begin
  DM_tabelas.ZQContaBancaria.First;
  DXBPrimeiro.Enabled := false;
  DXBAnterior.Enabled := false;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  DXBGravar.Enabled := False;
  DXBCancelar.Enabled := false;
  Filtra_limite;
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DXBAnteriorClick(Sender: TObject);
begin
  DM_tabelas.ZQContaBancaria.Prior;
  if DM_tabelas.ZQContaBancaria.RecNo = 1 then begin
    DXBPrimeiro.Enabled := false;
    DXBAnterior.Enabled := false;
  end;
  DXBProximo.Enabled := true;
  DXBUltimo.Enabled := true;
  Filtra_limite;
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DXBProximoClick(Sender: TObject);
begin
  DM_tabelas.ZQContaBancaria.Next;
  if DM_tabelas.ZQContaBancaria.RecNo = DM_tabelas.ZQContaBancaria.RecordCount then begin
    DXBProximo.Enabled := false;
    DXBUltimo.Enabled := false;
  end;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  Filtra_limite;
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DXBUltimoClick(Sender: TObject);
begin
  DM_tabelas.ZQContaBancaria.Last;
  DXBProximo.Enabled := false;
  DXBUltimo.Enabled := false;
  DXBPrimeiro.Enabled := true;
  DXBAnterior.Enabled := true;
  Filtra_limite;
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DXBGravarClick(Sender: TObject);
begin
  if not Verif_senha('Conta Banc�ria','Confirma��o de Inclus�o ou Edi��o  (Bot�o GRAVAR)','') then exit;
  if RBFixo.Checked = True Then Begin
    DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString := 'F';
    Label19.Caption := 'Valor. Fixo';
  end
  else Begin
    DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString := 'V';
    Label19.Caption :='% p/ Calculo do Vlr. Mora Di�ria';
  end;
  DM_tabelas.ZQContaBancaria.Post;

  DM_tabelas.ZQsequencia_remessa.close;
  DM_tabelas.ZQsequencia_remessa.SQL.Clear;
  DM_tabelas.ZQsequencia_remessa.SQL.Add('select idremessa_sequencia,idconta_bancaria,sequencia from remessa_sequencia where idconta_bancaria='+quotedstr(DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text));
  DM_tabelas.ZQsequencia_remessa.Open;
  if DM_tabelas.ZQsequencia_remessa.RecordCount=0 then
  begin
    DM_tabelas.ZQsequencia_remessa.Insert;
    DM_tabelas.ZQsequencia_remessa.FieldByName('idconta_bancaria').AsLargeInt:=DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
    DM_tabelas.ZQsequencia_remessa.FieldByName('sequencia').AsLargeInt:=0;
    DM_tabelas.ZQsequencia_remessa.Post;
  end;
  DM_tabelas.ZQContaBancaria.Refresh;
  Desativar_campos;
end;

procedure TFrmCad_ContasBancarias.DXBCancelarClick(Sender: TObject);
begin
   if not Verif_senha('Conta Banc�ria','Cancelamento de Inclus�o ou Edi��o (Bot�o CANCELAR) ','') then exit;
   DM_tabelas.ZQContaBancaria.Cancel;
   Desativar_campos;
end;

procedure TFrmCad_ContasBancarias.DXBFecharClick(Sender: TObject);
begin
   Close;
end;

procedure TFrmCad_ContasBancarias.DXBIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Conta Banc�ria','Incluir Conta Banc�ria','') then exit;
  Ativar_Campos;
  DM_tabelas.ZQContaBancaria.Insert;
  PagContaBanco.PageIndex := 0;
  DBEBanco.SetFocus;
end;

procedure TFrmCad_ContasBancarias.DXBEditarClick(Sender: TObject);
begin
   if not Verif_senha('Conta Banc�ria','Editar Conta Banc�ria','C�d. interno: '+DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text) then exit;
   Ativar_campos;
   DM_tabelas.ZQContaBancaria.Edit;
   if PagContaBanco.PageIndex = 0 Then
     DBEBanco.SetFocus
   else
     DBECarteira.SetFocus;
end;

procedure TFrmCad_ContasBancarias.DXBExcluirClick(Sender: TObject);
var
  varcod : string;
begin
  if PagContaBanco.PageIndex = 0 Then Begin
    if not Verif_senha('Conta Banc�ria','Excluir Conta Banc�ria','C�d.: '+ DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text) then exit;
    if DM_tabelas.ZQContaBancaria.RecordCount>0 then begin
      varcod := DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text;
      DM_tabelas.ZQContaBancaria.SQL.Clear;
      DM_tabelas.ZQContaBancaria.SQL.Add('delete from mov_banco where conta_bancaria_cod_banco ='+quotedstr(varcod));
      DM_tabelas.ZQContaBancaria.ExecSQL;
      DM_tabelas.ZQContaBancaria.SQL.Clear;
      DM_tabelas.ZQContaBancaria.SQL.Add('select idconta_bancaria,n_banco,n_agencia,n_agencia_v,n_conta,n_conta_v,nomebanco,n_no_banco,nomeresposavel,nometitular,aberta,senha_interna,senha_conta,ativa,');
      DM_tabelas.ZQContaBancaria.SQL.Add('       jurosemboleto,taxadiaria,valoroupercent,carteira,boleto,idparticipante,n_dif_empreed,convenio,doc_titular,n_sequencial,nosso_numero,Apelido,cod_transmissao,');
      DM_tabelas.ZQContaBancaria.SQL.Add('       complemento,moramensal,variacao,data_maxima_desconto,Perc_descontos,layoutversaoArquivo,layoutversaoLote ');
      DM_tabelas.ZQContaBancaria.SQL.Add('   from Conta_Bancaria');
      DM_tabelas.ZQContaBancaria.Open;

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('delete from remessa_sequencia where idconta_bancaria='+quotedstr(varcod));
      DM_tabelas.ZQsequencia_remessa.ExecSQL;

      DM_tabelas.ZQsequencia_remessa.close;
      DM_tabelas.ZQsequencia_remessa.SQL.Clear;
      DM_tabelas.ZQsequencia_remessa.SQL.Add('select idremessa_sequencia,idconta_bancaria,sequencia from remessa_sequencia');
      DM_tabelas.ZQsequencia_remessa.Open;

      DM_Tabelas.ZQContaBancaria.Locate('idConta_Bancaria',varcod,[]);
      DM_tabelas.ZQContaBancaria.Delete;
    end;
  end;
  if PagContaBanco.PageIndex = 1 Then Begin
    if not Verif_senha('Conta Banc�ria','Excluir Remessa',' n� da remessa '+DM_tabelas.ZQBancRemes.FieldByName('remessa').Text+' da conta '+DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text) then exit;
    DM_Tabelas.ZQRemesRec.SQL.Clear;
    DM_Tabelas.ZQRemesRec.SQL.Add('delete from remessa_receb where remessa='+quotedstr(DM_Tabelas.ZQBancRemes.FieldByName('remessa').Text));
    DM_Tabelas.ZQRemesRec.ExecSQL;
    DM_Tabelas.ZQRemesRec.SQL.Clear;
    DM_Tabelas.ZQRemesRec.SQL.Add('Select idremessa_receb,remessa,idrec,retorno_motivo,acao,Nossonumero,ocorrencia,desc_motivo,credito,dt_ocorrencia ');
    DM_Tabelas.ZQRemesRec.SQL.Add('  from remessa_receb');
    DM_Tabelas.ZQRemesRec.Open;
    DM_Tabelas.ZQBancRemes.Delete;
  end;
end;

procedure TFrmCad_ContasBancarias.FormShow(Sender: TObject);
begin
  EnsureContaBancariaLookupFields(DM_Tabelas);
  DM_Tabelas.ZQContaBancaria.close;
  DM_Tabelas.ZQContaBancaria.Open;
  DM_Tabelas.ZQBancRemes.close;
  DM_Tabelas.ZQBancRemes.Open;
  DM_Tabelas.ZQRemesRec.Close;
  DM_Tabelas.ZQRemes_Receb.close;
  DM_Tabelas.ZQRemes_Receb.Open;
  DM_Tabelas.ZQEspecial.Open;
  DM_tabelas.ZQsequencia_remessa.Open;
  DXBGravarespecial.Enabled := false;
  DXBCancelarEspecial.Enabled := false;
  DXBRelatorios.Enabled := True;
  DBESenhainterna.Visible := False;
  DBESenhaBancaria.Visible := False;
  DBECodConta.Enabled := false;
  DXBGravarespecial.Enabled := false;
  DXBCancelarEspecial.Enabled := false;
//  DM_Tabelas.ZQRemes_Receb.Refresh;
  Desativar_campos;
  PagContaBanco.PageIndex := 0;
  Atualiza_pagina;

  Label26.Visible:=false;
  DBEdit4.Visible:=false;

end;

procedure TFrmCad_ContasBancarias.JNBMostrasennhaClick(Sender: TObject);
begin
  if DBESenhainterna.Visible then Begin
    DBESenhainterna.Visible := False;
    DBESenhaBancaria.Visible := False;
    JNBMostrasennha.Caption := 'Visualiza Senha';
  end
  else
  begin
    if not Verif_senha('Conta Banc�ria','Mostrar Senhas','C�d. Conta: '+ DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text) then exit;
    DBESenhainterna.Visible := True;
    DBESenhaBancaria.Visible := True;
    IF DM_tabelas.ZQContaBancaria.State in [dsEdit, dsInsert] then Begin
      DBESenhainterna.ReadOnly := False;
      DBESenhaBancaria.ReadOnly := False;
    end;
    JNBMostrasennha.Caption := 'Esconde Senha';
  end;
end;

procedure TFrmCad_ContasBancarias.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmCad_ContasBancarias.DBESenhaBancariaExit(Sender: TObject);
begin
  DXBGravar.SetFocus;
end;

procedure TFrmCad_ContasBancarias.DXBGravarespecialClick(Sender: TObject);
begin
  if not Verif_senha('BANCO - Hist. Lim. Especial','Confirma��o de Inclus�o ou Edi��o','Conta: '+ DM_tabelas.ZQEspecial.FieldByName('Conta_Bancaria_cod_banco').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('data_valida').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('valor').Text) then exit;
  DM_tabelas.ZQEspecial.FieldByName('Conta_Bancaria_cod_banco').AsInteger := DM_tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
  DM_tabelas.ZQEspecial.Post;
  DM_tabelas.ZQEspecial.Refresh;
  DM_tabelas.ZQEspecial.Last;
  desativar_campos;
end;

procedure TFrmCad_ContasBancarias.DXBCancelarEspecialClick(Sender: TObject);
begin
  if not Verif_senha('BANCO - Hist. Lim. Especial','Cancelamento de Inclus�o ou Edi��o','Conta: '+ DM_tabelas.ZQEspecial.FieldByName('Conta_Bancaria_cod_banco').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('data_valida').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('valor').Text) then exit;
  DM_tabelas.ZQEspecial.Cancel;
  desativar_campos;
end;

procedure TFrmCad_ContasBancarias.DXBIncluirEspecialClick(Sender: TObject);
begin
  if not Verif_senha('BANCO - Hist. Lim. Especial','Incluir Limite Especial','') then exit;
  DM_tabelas.ZQEspecial.Insert;
  desativando_botoes_em_edicao_limite;
  DBEDataValida.SetFocus;

end;

procedure TFrmCad_ContasBancarias.DXBEditarEspecialClick(Sender: TObject);
begin
  if not Verif_senha('BANCO - Hist. Lim. Especial','Editar Limite Especial','Conta: '+ DM_tabelas.ZQEspecial.FieldByName('Conta_Bancaria_cod_banco').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('data_valida').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('valor').Text) then exit;
   DM_tabelas.ZQEspecial.Edit;
   desativando_botoes_em_edicao_limite;
   DBEDataValida.SetFocus;
end;

procedure TFrmCad_ContasBancarias.DXBExcluirEspecialClick(Sender: TObject);
begin
  if not Verif_senha('BANCO - Hist. Lim. Especial','Excluir Limite Especial','Conta: '+ DM_tabelas.ZQEspecial.FieldByName('Conta_Bancaria_cod_banco').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('data_valida').Text+'  '+DM_tabelas.ZQEspecial.FieldByName('valor').Text) then exit;
  if DM_tabelas.ZQEspecial.RecordCount>0 then begin
    DM_tabelas.ZQEspecial.Delete;
  end;
end;

procedure TFrmCad_ContasBancarias.DBGConta_BancariaCellClick(
  Column: TColumn);
begin
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DBGConta_BancariaKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  Atualiza_pagina;
end;

procedure TFrmCad_ContasBancarias.DBGConta_BancariaDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECodConta.Text = DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Text Then Begin
    DBGConta_Bancaria.Canvas.Brush.Color :=$006CFFFF;
    DBGConta_Bancaria.Canvas.Font.Color := $00A80000;
    DBGConta_Bancaria.Canvas.Font.Style := [FsBold];
  end;
  DBGConta_Bancaria.DefaultDrawDataCell(Rect, DBGConta_Bancaria.columns[datacol].field, State);

end;

procedure TFrmCad_ContasBancarias.RBFixoClick(Sender: TObject);
begin
  if DM_Tabelas.ZQContaBancaria.State in [DsEdit,DsInsert] Then Begin
    if RBFixo.Checked Then Begin
      Label19.Caption := 'Valor. Fixo';
      DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString := 'F';
    end
    else Begin
      Label19.Caption :='% p/ Calculo do Vlr. Mora Di�ria';
      DM_Tabelas.ZQContaBancaria.FieldByName('valoroupercent').AsString := 'V';
    end;
  end;
end;

procedure TFrmCad_ContasBancarias.DBGRemessaMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
//  Label20.Caption := INTTOSTR(DM_Tabelas.ZQRemes_Receb.RecordCount);
end;

procedure TFrmCad_ContasBancarias.DBGRemessaKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
//  Label20.Caption := INTTOSTR(DM_Tabelas.ZQRemes_Receb.RecordCount);
end;

procedure TFrmCad_ContasBancarias.EparticipanteExit(Sender: TObject);
begin
  if DM_Tabelas.ZQContaBancaria.State in [DsInsert, DsEdit] Then
    if empty(Eparticipante.Text) Then
      DM_tabelas.ZQContaBancaria.FieldByName('idparticipante').AsInteger := 0
    else
      if AchaCLiente(500,15,Eparticipante.Text) Then Begin
        Eparticipante.Text := DM_TAbelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        DM_tabelas.ZQContaBancaria.FieldByName('idparticipante').AsInteger := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
      end;
end;

procedure TFrmCad_ContasBancarias.DBGRemessaDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBESeq.Text = DM_Tabelas.ZQBancRemes.FieldByName('idbanco_remessa').Text Then Begin
    DBGRemessa.Canvas.Brush.Color :=$006CFFFF;
    DBGRemessa.Canvas.Font.Color := $00A80000;
    DBGRemessa.Canvas.Font.Style := [FsBold];
  end;
  DBGRemessa.DefaultDrawDataCell(Rect, DBGRemessa.columns[datacol].field, State);

end;

procedure TFrmCad_ContasBancarias.dxButtonREMESSAClick(Sender: TObject);
var
rg:TBookMark;
begin
  if not Verif_senha('Conta Banc�ria','Remessas','') Then Exit;
//  Frm_RelREMESSA.Top := Frm_RelREMESSA.Top+105;
  //Frm_RelREMESSA.Left := Frm_RelREMESSA.Left-5+(Frm_RelREMESSA.Width-Frm_RelREMESSA.Width);
  numero.Caption:=DM_Tabelas.ZQBancRemes.FieldByName('gerado').Text+DM_Tabelas.ZQBancRemes.FieldByName('nomearq').AsString;
  rg := DM_Tabelas.ZQBancRemes.GetBookMark;
  if Frm_RelREMESSA=nil then
     Frm_RelREMESSA:=TFrm_RelREMESSA.Create(Self);
  Frm_RelREMESSA.showmodal;
  Frm_RelREMESSA:=nil;
  Frm_RelREMESSA.Free;
  DM_Tabelas.ZQBancRemes.GotoBookMark(rg);
  DM_Tabelas.ZQBancRemes.FreeBookMark(rg);
  DM_Tabelas.ZQRemes_Receb2.close;
end;

procedure TFrmCad_ContasBancarias.psnumExit(Sender: TObject);
begin
  if not empty(psnum.Text) then
  begin
    ZQBancRemes.Open;
    ZQContaBancaria.Open;
    ZQRemes_Receb.Close;
    ZQRemes_Receb.SQL.Clear;
    ZQRemes_Receb.SQL.Add('SELECT * FROM remessa_receb where nossonumero like '+quotedstr(psnum.text+'%'));
    ZQRemes_Receb.open;
    ZQRemes_Receb.First;
//    ZQRemes_Receb.Locate('nossonumero',psnum.text,[loCaseInsensitive]);
    DM_Tabelas.ZQContaBancaria.Locate('idconta_bancaria',ZQContaBancaria.FieldByName('idconta_bancaria').Text,[loCaseInsensitive]);
    DM_Tabelas.ZQBancRemes.Locate('idbanco_remessa',ZQBancRemes.FieldByName('idbanco_remessa').Text,[loCaseInsensitive]);
    DM_Tabelas.ZQRemes_Receb.Locate('idremessa_receb',ZQRemes_Receb.FieldByName('idremessa_receb').Text,[loCaseInsensitive]);
    ZQBancRemes.Close;
    ZQContaBancaria.Close;
    ZQRemes_Receb.Close;
    psnum.Clear;
  end;
end;

procedure TFrmCad_ContasBancarias.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQContaBancaria.Close;
  DM_Tabelas.ZQBancRemes.close;
  DM_Tabelas.ZQRemes_Receb.close;
  DM_Tabelas.ZQEspecial.Close;
  DM_tabelas.ZQsequencia_remessa.close;
  FrmCad_ContasBancarias:=nil;
  Action:=CaFree;
end;

procedure TFrmCad_ContasBancarias.DBECarteiraExit(Sender: TObject);
var
t:integer;
xcrt:string;
begin
  if not empty(trim(DBECarteira.Text)) then
  begin
    xcrt:=OnlyNumber(trim(DBECarteira.Text));
    if xcrt<>emptystr then
    begin
      t:=strtoint(DBECarteira.Text);
      DBECarteira.Text:=strzero(t,3);
    end;
  end;
end;

procedure TFrmCad_ContasBancarias.DXBRelatoriosClick(Sender: TObject);
begin
  DM_Tabelas.ZQContaBancaria.Filtered:=false;
  DM_Tabelas.ZQContaBancaria.Filter:='n_conta='+quotedstr(DBEConta.Text)+' and n_conta_v='+quotedstr(DBENContaV.Text);
  DM_Tabelas.ZQContaBancaria.Filtered:=true;
  if not DM_Tabelas.ZQEmpresa.Active then
     DM_Tabelas.ZQEmpresa.open;
  if Frm_RelBancaria=nil then
     Frm_RelBancaria:=TFrm_RelBancaria.Create(Self);
  Frm_RelBancaria.RLReport1.PreviewModal;
  Frm_RelBancaria:=nil;
end;

procedure TFrmCad_ContasBancarias.FormKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if key = VK_F12 then
  begin
    if Label26.Visible=false then
       Label26.Visible:=true;
    if DBEdit4.Visible=false then
       DBEdit4.Visible:=true;
  end;
end;

procedure TFrmCad_ContasBancarias.DBEdit8Enter(Sender: TObject);
begin
  if not Verif_senha('Conta Banc�ria','Nosso N�mero','') then exit;
end;


procedure TFrmCad_ContasBancarias.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmCad_ContasBancarias, 'ZQRemes_Receb', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebidremessa_receb', 'idremessa_receb', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebidrec', 'idrec', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebretorno_motivo', 'retorno_motivo', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebacao', 'acao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_RecebNossonumero', 'Nossonumero', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebocorrencia', 'ocorrencia', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebdesc_motivo', 'desc_motivo', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebcredito', 'credito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQRemes_Receb', 'ZQRemes_Recebdt_ocorrencia', 'dt_ocorrencia', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_ContasBancarias, 'ZQBancRemes', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesidbanco_remessa', 'idbanco_remessa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesremessa', 'remessa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesidbanco', 'idbanco', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesgerado', 'gerado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesnomearq', 'nomearq', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesarquivo', 'arquivo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQBancRemes', 'ZQBancRemesdigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmCad_ContasBancarias, 'ZQContaBancaria', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaidconta_bancaria', 'idconta_bancaria', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_banco', 'n_banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_agencia', 'n_agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_agencia_v', 'n_agencia_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_conta', 'n_conta', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_conta_v', 'n_conta_v', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarianomebanco', 'nomebanco', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_no_banco', 'n_no_banco', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarianomeresposavel', 'nomeresposavel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarianometitular', 'nometitular', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaaberta', 'aberta', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariasenha_interna', 'senha_interna', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariasenha_conta', 'senha_conta', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaativa', 'ativa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariajurosemboleto', 'jurosemboleto', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariataxadiaria', 'taxadiaria', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariavaloroupercent', 'valoroupercent', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariacarteira', 'carteira', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaboleto', 'boleto', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaidparticipante', 'idparticipante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarianomeparticip', 'nomeparticip', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idparticipante', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_dif_empreed', 'n_dif_empreed', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaconvenio', 'convenio', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariadoc_titular', 'doc_titular', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarian_sequencial', 'n_sequencial', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarianosso_numero', 'nosso_numero', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaApelido', 'Apelido', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariacod_transmissao', 'cod_transmissao', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariacomplemento', 'complemento', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariamoramensal', 'moramensal', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariavariacao', 'variacao', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariadata_maxima_desconto', 'data_maxima_desconto', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancariaPerc_descontos', 'Perc_descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarialayoutversaoArquivo', 'layoutversaoArquivo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmCad_ContasBancarias, 'ZQContaBancaria', 'ZQContaBancarialayoutversaoLote', 'layoutversaoLote', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.
