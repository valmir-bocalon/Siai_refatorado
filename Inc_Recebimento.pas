unit Inc_Recebimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBClient, StdCtrls, Grids, DBGrids, XDate, XNum, Mask,
  XEdit, XBanner, dxButton, DBCtrls, XDBEdit, XDBNum,
  wwdbdatetimepicker, ComCtrls, JvExComCtrls, JvDateTimePicker,
  JvDBDateTimePicker, ExtCtrls, jpeg, TFlatRadioButtonUnit,
  ZAbstractRODataset, ZDataset, dxCore2;

type
  TFrm_Inc_Recebimento = class(TForm)
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    XNEParcelas: TXNumEdit;
    XNEValor: TXNumEdit;
    XNEIntervado: TXNumEdit;
    DS_Parcelas: TDataSource;
    CDSParcelas: TClientDataSet;
    CDSParcelasTipDoc: TStringField;
    CDSParcelasDocum: TStringField;
    CDSParcelasOrdem: TStringField;
    CDSParcelasVenci: TDateField;
    CDSParcelasVrParc: TFloatField;
    CDSParcelascliente: TIntegerField;
    CDSParcelasusuario: TIntegerField;
    CDSParcelasentrada: TDateField;
    DXBGravar: TdxButton;
    DXBCancelar: TdxButton;
    Label1: TLabel;
    Label2: TLabel;
    WDBDVenci: TwwDBDateTimePicker;
    XDBNVrParc: TXDBNumEdit;
    XDBEDocumen: TXDBEdit;
    DBCBTIPDOC: TDBComboBox;
    DBText1: TDBText;
    Label3: TLabel;
    MObser: TMemo;
    Label4: TLabel;
    CDSParcelasbanco: TStringField;
    CDSParcelasagencia: TStringField;
    CDSParcelasdono: TStringField;
    CDSParcelascpf_cnpj: TStringField;
    GBCheque: TGroupBox;
    DBEAgencia: TDBEdit;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    DBEBanco: TDBEdit;
    DBEDono: TDBEdit;
    DBECpf_Cnpj: TDBEdit;
    EContabil: TEdit;
    Label19: TLabel;
    CDSParcelascontabil: TIntegerField;
    XDEEntrada: TDateTimePicker;
    Label20: TLabel;
    Label21: TLabel;
    Edit1: TEdit;
    Label22: TLabel;
    Label23: TLabel;
    Panel2: TPanel;
    LabelEmpresa: TLabel;
    CBEscolheEmpresa: TComboBox;
    Eadversa: TEdit;
    Label24: TLabel;
    DBEConta: TDBEdit;
    CDSParcelasconta: TStringField;
    Enomes: TEdit;
    Ecodnome: TEdit;
    Label16: TLabel;
    GBConta: TGroupBox;
    CDSParcelasdesccontabil: TStringField;
    DBGParcelas: TDBGrid;
    Image2: TImage;
    Image1: TImage;
    DS_dividirpor: TDataSource;
    DBEOrdem: TDBEdit;
    Panel_particpantes: TPanel;
    Label25: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label26: TLabel;
    Erecebedor: TEdit;
    XDBNumEdit2: TXDBNumEdit;
    EContabil2: TEdit;
    DBCBTipdocdiv: TDBComboBox;
    DBGDiv: TDBGrid;
    XPartes: TXNumEdit;
    DBECod: TDBEdit;
    Label30: TLabel;
    XBanner2: TXBanner;
    XBanner1: TXBanner;
    XBanner3: TXBanner;
    XBanner4: TXBanner;
    XBanner5: TXBanner;
    RBRec: TFlatRadioButton;
    RBPag: TFlatRadioButton;
    ZQdividirpor: TClientDataSet;
    ZQdividirporcodcli: TIntegerField;
    ZQdividirpornomecli: TStringField;
    ZQdividirporcodconta: TIntegerField;
    ZQdividirpordescriconta: TStringField;
    ZQdividirportipdoc: TStringField;
    ZQdividirporpercent: TFloatField;
    ZQdividirporid: TIntegerField;
    EQuadralote: TEdit;
    Label32: TLabel;
    ZROQReceb: TZReadOnlyQuery;
    DS_ROReceb1: TDataSource;

    CBTipRec: TComboBox;
    CDSParcelasTip: TStringField;









































    procedure XNEIntervadoExit(Sender: TObject);
    procedure CBTipDocExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DXBCancelarClick(Sender: TObject);
    procedure DBCBTIPDOCExit(Sender: TObject);
    procedure DXBGravarEnter(Sender: TObject);
    procedure DXBGravarClick(Sender: TObject);
    procedure XDBEDocumenExit(Sender: TObject);
    procedure XDBNVrParcExit(Sender: TObject);
    procedure XNEParcelasExit(Sender: TObject);
    procedure DBEBancoExit(Sender: TObject);
    procedure DBEAgenciaExit(Sender: TObject);
    procedure EContabilExit(Sender: TObject);
    procedure EContabilEnter(Sender: TObject);
    procedure EadversaExit(Sender: TObject);
    procedure DBECpf_CnpjExit(Sender: TObject);
    procedure EnomesExit(Sender: TObject);
    procedure DBCBTIPDOCEnter(Sender: TObject);
    procedure EContabil2Exit(Sender: TObject);
    procedure XPartesExit(Sender: TObject);
    procedure DBGParcelasDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGDivDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure ErecebedorExit(Sender: TObject);
    procedure XDBNumEdit2Exit(Sender: TObject);
    procedure XDBNumEdit2Enter(Sender: TObject);
    procedure ErecebedorEnter(Sender: TObject);
    procedure EContabil2Enter(Sender: TObject);
    procedure XDBNVrParcEnter(Sender: TObject);
    procedure RBRecClick(Sender: TObject);
    procedure RBPagClick(Sender: TObject);
    procedure DBCBTipdocdivExit(Sender: TObject);
    procedure DBCBTipdocdivEnter(Sender: TObject);
    procedure EQuadraloteExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
  VarEmpresa : integer;
    { Public declarations }
  end;

var
  Frm_Inc_Recebimento: TFrm_Inc_Recebimento;
  VarXCliente, varXadversa, VarXusuario : integer;
  VArDoc, vartipodoc : String;
implementation

uses Tabelas, Funcoes, uRuntimeFields;

{$R *.dfm}

procedure TFrm_Inc_Recebimento.XNEIntervadoExit(Sender: TObject);
Var
  VarSoma : Double;
  VarXy, VarParc, VarParcTot: integer;
  VarVenc : TDateTime;
  VarOrd  : string;
begin
  CDSParcelas.close;
  CDSParcelas.CreateDataSet;
  CDSParcelas.IndexFieldNames := 'ordem';
  VarVenc:=XDEEntrada.DateTime;
  VarParcTot := strtoint(floattostr(XNEParcelas.Value));
  VarSoma := 0;
  for VarParc:=1 to VarParcTot do Begin
    CDSParcelas.Insert;
    CDSParcelasTipDoc.Value := DBCBTIPDOC.Text;
    Varord := inttostr(VarParc);
    for VarXy := length(inttostr(VarParc)) to length(XNEParcelas.Text)-1 do
       Varord := '0'+Varord;
    CDSParcelasOrdem.Value := VarOrd+'/'+floattostr(XNEParcelas.Value);
    if VarParc=1 Then Begin
      CDSParcelasVenci.Value := XDEEntrada.DateTime;
      if RBRec.Checked Then
        CDSParcelascontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecinc').AsInteger
      else
        CDSParcelascontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpaginc').AsInteger;
    end
    else
    Begin
      if (XNEIntervado.Value=30) and (DM_tabelas.ZQConfiguracoes.FieldByName('mesmodiames').AsString = 'S') Then
        CDSParcelasVenci.Value := IncMonth(XDEEntrada.DateTime,VarParc-1)
      else
        CDSParcelasVenci.Value := VarVenc+XNEIntervado.Value;
      VarVenc := CDSParcelasVenci.Value;
      CDSParcelasVrParc.Value := Truncar(XNEValor.Value/XNEParcelas.Value,2);
      VarSoma := VarSoma + CDSParcelasVrParc.Value;
    end;
  End;
  if XNEValor.Value<>VarSoma Then Begin
    CDSParcelas.First;
    CDSParcelas.Edit;
    CDSParcelasVrParc.Value := CDSParcelasVrParc.Value+(XNEValor.Value-VarSoma);
  end;
  CDSParcelas.First;
  DXBGravar.Enabled := True;
  EContabil.SetFocus;
  DXBGravar.SetFocus;
end;

procedure TFrm_Inc_Recebimento.CBTipDocExit(Sender: TObject);
begin
  Label15.Visible := False;
  DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
  IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then
    Label15.Visible := True;
end;

procedure TFrm_Inc_Recebimento.FormShow(Sender: TObject);
Var
  VarX : Integer;
begin
  DM_Tabelas.ZQCheque.Open;
  DM_Tabelas.ZQNumOrdem.Open;
  ZQdividirpor.Close;
  ZQdividirpor.CreateDataSet;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
//    if DM_Tabelas.ZQTipodoclancinclus.Value = 'S' Then
      DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
      DBCBTipdocdiv.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  CDSParcelas.close;
  CDSParcelas.CreateDataSet;
  GBCheque.Visible := False;
  DXBGravar.Enabled := False;
  DM_tabelas.ZQEmpresa.First;
  VarEmpresa := DM_tabelas.ZQEmpresa.FieldByName('codigo').AsLargeInt;
  CBEscolheEmpresa.Clear;
  DM_tabelas.ZQEmpresa.DisableControls;
  while not DM_tabelas.ZQEmpresa.Eof do Begin
    CBEscolheEmpresa.Items.Add(DM_tabelas.ZQEmpresa.FieldByName('codigo').Text+' - '+DM_tabelas.ZQEmpresa.FieldByName('razao').AsString);
    DM_tabelas.ZQEmpresa.Next;
  end;
  DM_tabelas.ZQEmpresa.EnableControls;
  CBEscolheEmpresa.ItemIndex := 0;
  LabelEmpresa.Caption := CBEscolheEmpresa.Text;
  Enomes.Text := '';
  XNEValor.Value := 0;
  MObser.Text := '';
  XNEParcelas.Value := 1;
  XDEEntrada.DateTime := date;
  XNEIntervado.Value := 0;
//  XPartes.Value := DM_Tabelas.ZQConfiguracoesparticipantes.Value;
  RBRec.setfocus;
end;

procedure TFrm_Inc_Recebimento.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Frm_Inc_Recebimento:=nil;
  Action:=CaFree;
end;

procedure TFrm_Inc_Recebimento.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Inc_Recebimento.DXBCancelarClick(Sender: TObject);
begin
  Label20.Caption := 'cancelado';
  Close;
end;

procedure TFrm_Inc_Recebimento.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBCancelar.Focused Then Begin
    if not DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTIPDOC.Text,[])Then Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').AsString='S' Then Begin
      if CDSParcelasVenci.Value<>date Then BEgin
        showmessage('Este tipo de documento só aceita pagamento avista!!!');
        EContabil.SetFocus;
        Exit;
      end;
    end;
    Label15.Visible := False;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin
      Label15.Visible := True;
      GBCheque.Visible := True;
      DBEBanco.SetFocus;
    end;
    vartipodoc := DBCBTIPDOC.Text;
  end;
end;

procedure TFrm_Inc_Recebimento.DXBGravarEnter(Sender: TObject);
begin
  CDSParcelas.First;
  While not CDSParcelas.Eof do Begin
    if (empty(CDSParcelasContabil.Text))  Then Begin  // and (CDSParcelasTipDoc.Value<>'DI')
      EContabil.SetFocus;
      Break;
    end;
    if empty(CDSParcelasTipDoc.Value) Then Begin
      DBCBTIPDOC.SetFocus;
       Break;
    end;
    CDSParcelas.Next;
  End;
end;

procedure TFrm_Inc_Recebimento.DXBGravarClick(Sender: TObject);
Var
  Varsoma : Double;
  tparc,varX : integer;
begin
 CDSParcelas.cancel;
  CDSParcelas.First;
  varsoma := 0;
  if empty(Enomes.Text) then begin
    showmessage('O nome do participante deve estar preenchido...');
    Enomes.SetFocus;
    exit;
  end;
  if empty(Eadversa.Text) then begin
    showmessage('O nome da parte adversa deve estar preenchida...');
    Eadversa.SetFocus;
    exit;
  end;
  if XNEParcelas.Value= 0 Then Begin
    showmessage('Deve ser informado a quantidade de parcelas...');
    XNEParcelas.SetFocus;
    exit;
  end;
  if XNEValor.Value= 0 Then Begin
    showmessage('Deve ser informado o valor da negociação...');
    XNEValor.SetFocus;
    exit;
  end;
  while not CDSParcelas.Eof do begin
    if empty(CDSParcelascontabil.Text) Then Begin
      showmessage('Todo os planos de contas devem estar preenchidos...');
      EContabil.SetFocus;
      exit;
    end;
    if empty(CDSParcelasTipDoc.Value) Then Begin
      showmessage('Todo os tipos de documentos devem estar preenchidos...');
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    varsoma := Varsoma + CDSParcelasVrParc.Value;
    CDSParcelas.Next;
  end;
  if floattostr(varsoma) <> floattostr(XNEValor.Value) Then Begin
    Showmessage('Há uma diferença de '+floattostr(varsoma-XNEValor.Value)+' no valor das parcelas...  Refaça as parcelas...');
    XNEParcelas.Setfocus;
    exit;
  end;
  CDSParcelas.First;
//  if not Verif_senha('Financeiro','Gravar Inclusão ','participante: '+DM_Tabelas.ZqParticipanteidpaticipante.Text+'  Nº parc.:'+XNEParcelas.Text+'  1º Doc.:'+CDSParcelasDocum.Value+'  Entrada:'+Datetostr(XDEEntrada.DateTime)) then exit;
  DM_Tabelas.ZQNumOrdem.Insert;
  DM_Tabelas.ZQNumOrdem.Post;
  if Xpartes.Value >0 Then
  Begin
    ZQdividirpor.First;
    while not ZQdividirpor.Eof do
    Begin
      DM_TAbelas.ZQDividido.Insert;
      DM_Tabelas.ZQDividido.FieldByName('participante').AsInteger := ZQdividirporcodcli.Value;
      DM_Tabelas.ZQDividido.FieldByName('percentual').AsFloat := ZQdividirporpercent.Value;
      DM_Tabelas.ZQDividido.FieldByName('planodeconta').AsInteger := ZQdividirporcodconta.Value;
      DM_Tabelas.ZQDividido.FieldByName('tipodocumento').AsString := ZQdividirportipdoc.Value;
      DM_Tabelas.ZQDividido.FieldByName('ordem').AsInteger := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
      DM_TAbelas.ZQDividido.Post;
      DM_Tabelas.ZQParticipante.Locate('idpaticipante',ZQdividirporcodcli.Value,[]);
      if empty(DM_Tabelas.ZqParticipante.FieldByName('codpaginc').Text) Then Begin
        DM_Tabelas.ZQParticipante.Edit;
        DM_Tabelas.ZqParticipante.FieldByName('codpaginc').AsInteger := ZQdividirporcodconta.Value;
        DM_Tabelas.ZQParticipante.Post;
      end;
      DM_Tabelas.ZQPlanodeContas.Locate('codigo', ZQdividirporcodconta.Value,[]);
      if empty(DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString) Then Begin
        DM_Tabelas.ZQPlanodeContas.Edit;
        DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString := ZQdividirportipdoc.Value;
        DM_Tabelas.ZQPlanodeContas.Post;
      End;
      ZQdividirpor.Next;
    end;
  end;
  DM_Tabelas.ZQParticipante.Locate('idpaticipante',strtoint(Ecodnome.Text),[]);
  if (RBPag.Checked) and (empty(DM_Tabelas.ZqParticipante.FieldByName('codpaginc').Text)) Then Begin
    DM_Tabelas.ZqParticipante.Edit;
    DM_Tabelas.ZqParticipante.FieldByName('codpaginc').AsInteger := CDSParcelasContabil.Value;
    DM_Tabelas.ZQParticipante.Post;
  end;
  if (RBRec.Checked) and (empty(DM_Tabelas.ZqParticipante.FieldByName('codrecinc').Text)) Then Begin
    DM_Tabelas.ZQParticipante.Edit;
    DM_Tabelas.ZqParticipante.FieldByName('codrecinc').AsInteger := CDSParcelasContabil.Value;
    DM_Tabelas.ZQParticipante.Post;
  end;
  tparc:=CDSParcelas.recordcount;
  CDSParcelas.first;
  CDSParcelas.DisableControls;
  while not CDSParcelas.Eof do
  begin
    DM_Tabelas.ZQPlanodeContas.Locate('codigo', CDSParcelasContabil.Value,[]);
    if empty(DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString) Then
    Begin
      DM_Tabelas.ZQPlanodeContas.Edit;
      DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString := CDSParcelasTipDoc.Value;
      DM_Tabelas.ZQPlanodeContas.Post;
    End;
    DM_tabelas.ZQRecebimento.Insert;
    DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString := 'R';
    if RBPag.Checked then
      DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString := 'P';
    DM_tabelas.ZQRecebimento.FieldByName('empresa').AsInteger := VarEmpresa;
    DM_Tabelas.ZQRecebimento.FieldByName('adversa').AsInteger := VarXAdversa;
    DM_tabelas.ZQRecebimento.FieldByName('documento').AsString :=Edit1.text;//CDSParcelasDocum.Value;
    DM_tabelas.ZQRecebimento.FieldByName('cliente').AsInteger := strtoint(Ecodnome.Text);
    DM_tabelas.ZQRecebimento.FieldByName('usuario').AsInteger := DM_tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt;
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime := XDEEntrada.Date; //Date;
    DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime := CDSParcelasVenci.Value;
    DM_tabelas.ZQRecebimento.FieldByName('Valor').AsFloat := CDSParcelasVrparc.Value;
    DM_tabelas.ZQRecebimento.FieldByName('VrDoc').AsFloat := XNEValor.Value;
    DM_tabelas.ZQRecebimento.FieldByName('numordem').AsInteger := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').AsLargeInt;
    DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text+'-'+CDSParcelasOrdem.Value+'/'+strzero(tparc,3);//+CDSParcelasOrdem.Value;
    DM_tabelas.ZQRecebimento.FieldByName('TipDoc').AsString := CDSParcelasTipDoc.Value;
    DM_tabelas.ZQRecebimento.FieldByName('Observ').AsString := MObser.Text;
    DM_tabelas.ZQRecebimento.FieldByName('contabil').AsInteger := CDSParcelasContabil.Value;
    DM_tabelas.ZQRecebimento.FieldByName('origem').AsString := label23.Caption;
    if ZROQReceb.Active=true then
    begin
      DM_tabelas.ZQRecebimento.FieldByName('idloteamento').AsInteger:=ZROQReceb.FieldByName('idloteamento').AsInteger;
      DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').AsInteger:=ZROQReceb.FieldByName('venda_idvenda').AsInteger;
      DM_tabelas.ZQRecebimento.FieldByName('quadralote').AsString:=ZROQReceb.FieldByName('quadralote').AsString;
      DM_tabelas.ZQRecebimento.FieldByName('origem').AsString := ZROQReceb.FieldByName('origem').AsString;
      label23.Caption:=ZROQReceb.FieldByName('origem').AsString;
      DM_tabelas.ZQRecebimento.FieldByName('nomeadversa').AsString:=ZROQReceb.FieldByName('nomeadversa').AsString;
      DM_tabelas.ZQRecebimento.FieldByName('Parcelas_fixas').AsString:=ZROQReceb.FieldByName('Parcelas_fixas').AsString;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('vend_receb').AsString = 'N' Then Begin
      if DM_tabelas.ZQCobaRe.active = false then
         DM_tabelas.ZQCobaRe.open;
      DM_tabelas.ZQCobaRe.Insert;
      DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
      DM_tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime := XDEEntrada.Date;
      DM_tabelas.ZQCobaRe.Post;
      DM_tabelas.ZQRecebimento.FieldByName('RefBaixa').AsInteger := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
      DM_tabelas.ZQReceb_Baixa.Insert;
      DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').AsDateTime := Date;
      DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').AsFloat := CDSParcelasVrparc.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').AsString := CDSParcelasTipDoc.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('Docum').AsString := CDSParcelasDocum.Value;
      DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').AsInteger := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
      DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').AsInteger := CDSParcelasContabil.Value;
      DM_tabelas.ZQReceb_Baixa.Post;
    end
    else
      DM_tabelas.ZQRecebimento.FieldByName('saldo').AsFloat := CDSParcelasVrparc.Value;

    DM_tabelas.ZQRecebimento.FieldByName('tip').AsString:=CDSParcelasTip.Value;

    DM_tabelas.ZQRecebimento.Post;
    if DM_tabelas.ZQTipoDoc.FieldByName('vend_caixa').AsString='S' Then Begin
       if DM_tabelas.ZQCaixa.active=false then
          DM_tabelas.ZQCaixa.open;
      DM_tabelas.ZQCaixa.Insert;
      DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').AsInteger := CDSParcelasContabil.Value;
      DM_tabelas.ZQCaixa.FieldByName('data_lan').AsDateTime := date;
      DM_tabelas.ZQCaixa.FieldByName('vr_lan').AsFloat := CDSParcelasVrparc.Value;
      DM_tabelas.ZQCaixa.FieldByName('documento').AsString := CDSParcelasDocum.Value;
      DM_tabelas.ZQCaixa.FieldByName('empresa').AsInteger := 1;
      DM_tabelas.ZQCaixa.FieldByName('credeb').AsString := 'C';
      DM_tabelas.ZQCaixa.FieldByName('contracodigo').AsInteger := strtoint(Ecodnome.Text);
      DM_tabelas.ZQCaixa.FieldByName('vinculo').AsString := label23.Caption;
      DM_tabelas.ZQCaixa.FieldByName('obs').AsString := 'Recebimento - '+label23.Caption+' Ordem -'+CDSParcelasOrdem.Value+'/'+strzero(tparc,3);//+CDSParcelasOrdem.Value;
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
        DM_tabelas.ZQCaixa.FieldByName('obs').AsString :=DM_tabelas.ZQCaixa.FieldByName('obs').AsString + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
      end;
      DM_tabelas.ZQCaixa.Post;
    end;
    IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' Then Begin
      if DM_tabelas.ZQCheque.active=false then
         DM_tabelas.ZQCheque.open;
      DM_tabelas.ZQCheque.Insert;
      DM_tabelas.ZQCheque.FieldByName('idvinculo').AsInteger :=  DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
      DM_tabelas.ZQCheque.FieldByName('Banco').AsString := CdSParcelasBanco.Value;
      DM_tabelas.ZQCheque.FieldByName('agencia').AsString := CDSParcelasagencia.Value;
      DM_tabelas.ZQCheque.FieldByName('Dono').AsString := CDSParcelasdono.Value;
      DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').AsString := CDSParcelasCPF_CNPJ.Value;
      DM_tabelas.ZQCheque.FieldByName('valor').AsFloat := CDSParcelasVrparc.Value;
      DM_tabelas.ZQCheque.FieldByName('emissao').AsDateTime := XDEEntrada.DateTime;
      DM_Tabelas.ZQCheque.FieldByName('conta').AsString := CDSparcelasconta.Value;
      DM_tabelas.ZQCheque.FieldByName('deposito').AsDateTime := CDSParcelasVenci.Value;
      DM_Tabelas.ZQCheque.FieldByName('idrecebimento').AsInteger := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
      DM_tabelas.ZQCheque.Post;
    end;
    CDSParcelas.Next;
  end;
  CDSParcelas.EnableControls;
  DM_tabelas.ZQCobaRe.close;
  DM_tabelas.ZQCaixa.close;
  //DM_tabelas.ZQCheque.close;
  DM_Tabelas.ZQRecebimento.Cancel;
  DM_Tabelas.ZQRecebimento.Refresh;
  ZROQReceb.close;
  Close;
end;

procedure TFrm_Inc_Recebimento.XDBEDocumenExit(Sender: TObject);
begin
  if (not empty(XDBEDocumen.Text)) and (DM_tabelas.ZQRecebimento.Locate('documento',XDBEDocumen.Text,[])) Then Begin
    showmessage('Este documento já foi lançado anteriormente... corrija a numeração.....');
    XDBEDocumen.SetFocus;
    exit;
  End;
  if DM_tabelas.ZQCheque.Locate('Numero, Banco, Agencia, conta',VarArrayOf([XDBEDocumen.Text,CDSParcelasbanco.Value,CDSParcelasagencia.Value,CDSParcelasconta.Value]),[]) Then Begin
    showmessage('Este cheque já foi lançado anteriormente... corrija a numeração.....');
    XDBEDocumen.SetFocus;
    exit;
  end;
  DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
  if (DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S') and (pos(quotedstr(CDSParcelasDocum.Value),VarDoc)>0) Then Begin
    showmessage('Este cheque já foi lançado nesta inclusão... corrija a numeração.....');
    XDBEDocumen.SetFocus;
    exit;
  end;
  if (DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S') and (empty(CDSParcelasDocum.Value)) Then Begin
    showmessage('o campo Nº do cheque deve ser preenchido...');
    XDBEDocumen.SetFocus;
    exit;
  end;
  GBCheque.Visible := False;
end;

procedure TFrm_Inc_Recebimento.XDBNVrParcExit(Sender: TObject);
Var
  varregis: Integer;
  varsoma, vardif : Double;
begin
  if XNEValor.Value= 0 Then Begin
    showmessage('Deve ser informado o valor da negociação...');
    XNEValor.SetFocus;
    exit;
  end;
  varsoma := 0;
  if XNEParcelas.Value>1 then
  begin
    varregis := CDSParcelas.RecNo;
    CDSParcelas.First;
    VarDoc := '';
    while not (CDSParcelas.Eof) and (CDSParcelas.RecNo <= varregis) do Begin
      VarDoc := VarDoc + quotedstr(CDSParcelasDocum.Value)+',';
      varsoma := varsoma +CDSParcelasVrParc.Value;
      CDSParcelas.next;
    end;
    VarDoc := copy(VarDoc,1,length(VarDoc)-1);
    vardif := XNEValor.Value - varsoma;
    CDSParcelas.RecNo := varregis;
    if CDSParcelas.RecNO< XNEParcelas.Value Then
    Begin
      CDSParcelas.RecNO := varregis+1;
      while not CDSParcelas.Eof do Begin
        CDSParcelas.Edit;
        CDSParcelasVrParc.Value := truncar(vardif / (XNEParcelas.Value-varregis),2);
        CDSParcelas.Next;
      end;
    end;
    CDSParcelas.RecNo := varregis ;
    if not CDSParcelas.Eof Then Begin
      CDSParcelas.RecNo := varregis+1;
      EContabil.SetFocus;
    end
    else Begin
      DXBGravar.Enabled := True;
      DXBGravar.SetFocus;
    end;
  end
  else
  begin
    CDSParcelasOrdem.value:=strzero(strtoint(floattostr(XNEParcelas.value)),3);
    CDSParcelasTip.Value:=copy(CBTipRec.Text,1,1);
    CDSParcelas.post;
    CDSParcelas.insert;
    DXBGravar.Enabled := True;
    XNEParcelas.SetFocus;
  end;
end;

procedure TFrm_Inc_Recebimento.XNEParcelasExit(Sender: TObject);
begin
  if XNEParcelas.Value<1 Then
    XNEParcelas.SetFocus;
end;

procedure TFrm_Inc_Recebimento.DBEBancoExit(Sender: TObject);
begin
  if Empty(DBEBanco.Text) Then
    DBEBanco.SetFocus;
end;

procedure TFrm_Inc_Recebimento.DBEAgenciaExit(Sender: TObject);
begin
if Empty(DBEAgencia.Text) Then
  DBEAgencia.SetFocus;
end;

procedure TFrm_Inc_Recebimento.EContabilExit(Sender: TObject);
Var
  VarCD : string;
begin
  if not DXBCancelar.Focused Then Begin
    varCD := 'C';
    if RBPag.Checked Then
      varCD := 'D';
    if not AchaPlanoDeConta(300, 20, 'S', EContabil.Text, varCD) then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := DM_TAbelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' - '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    CDSParcelasContabil.Value := DM_TAbelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
    CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
  end;
end;

procedure TFrm_Inc_Recebimento.EContabilEnter(Sender: TObject);
Var
  VarTipDoc, VarDocum, varcontach, VarBanco, VarAgencia, VarCPF_CNPJ, VarDono: string;
  VarConta : Integer;
begin
  if CDSParcelas.Recno>1  Then
  begin
    CDSParcelas.Prior;
    VarTipDoc := CDSParcelasTipdoc.Value;
    VarDocum := CDSParcelasDocum.Value;
    VarBanco := CDSParcelasBanco.Value;
    varcontach := CDSParcelasconta.Value;
    VarAgencia := CDSParcelasagencia.Value;
    VarCPF_CNPJ := CDSParcelascpf_cnpj.Value;
    VarDono := CDSParcelasDono.Value;
    VarConta := CDSParcelasContabil.Value;
    CDSParcelas.Next;
    CDSParcelas.Edit;
    if Empty(CDSParcelasTipDoc.Value) Then
      CDSParcelasTipdoc.Value := VarTipDoc;
    if Empty(CDSParcelasDocum.Value) Then
      CDSParcelasDocum.Value :=  VarDocum;
    if Empty(CDSParcelasBanco.Value) Then
      CDSParcelasBanco.Value := VarBanco ;
    if Empty(CDSParcelasagencia.Value) Then
      CDSParcelasagencia.Value := VarAgencia;
    if Empty(CDSParcelascpf_cnpj.Value) then
      CDSParcelascpf_cnpj.Value := VarCPF_CNPJ;
    if Empty(CDSParcelasDono.Value) then
      CDSParcelasDono.Value := VarDono;
    if Empty(CDSParcelasConta.Value) then
      CDSParcelasConta.Value := VarContach;
    if Empty(CDSParcelasContabil.Text) then Begin
      DM_tabelas.ZQPlanoDeContas.Locate('codigo',varconta,[]);
      CDSParcelasContabil.Value := VarConta;
      EContabil.Text := DM_tabelas.ZQPlanoDeContas.FieldByName('descricao').Text;
    end;
  end
  else
  Begin
    CDSParcelas.Edit;
    CDSParcelasDono.Value := Enomes.Text;
//    CDSParcelascpf_cnpj.Value := DM_tabelas.ZQParticipante.Value;
  End;
end;

procedure TFrm_Inc_Recebimento.EadversaExit(Sender: TObject);
begin
  If not DXBCancelar.Focused Then Begin
    AchaCLiente(250, 140, Eadversa.Text);
    Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
    VarXAdversa := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
  end;
end;

procedure TFrm_Inc_Recebimento.DBECpf_CnpjExit(Sender: TObject);
begin
  If not DXBCancelar.Focused Then
    if not empty(CDSParcelascpf_cnpj.Value) then
      if not Verif_doc(CDSParcelascpf_cnpj.Value,True) then
        DBECpf_Cnpj.SetFocus
end;

procedure TFrm_Inc_Recebimento.EnomesExit(Sender: TObject);
begin
  If not DXBCancelar.Focused Then Begin
    if not AchaCLiente(180, 140,Enomes.Text) Then Begin
      Enomes.SetFocus;
      Exit;
    End;
    Ecodnome.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    Enomes.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
    if RBRec.Checked Then 
      EContabil.Text := DM_Tabelas.ZqParticipante.FieldByName('codrecinc').Text
    else
      EContabil.Text := DM_Tabelas.ZqParticipante.FieldByName('codpaginc').Text;
  End;
end;

procedure TFrm_Inc_Recebimento.DBCBTIPDOCEnter(Sender: TObject);
begin
  if (CDSParcelas.RecNo>1) and (empty(DBCBTIPDOC.Text)) Then
    CDSParcelasTipDoc.Value  := vartipodoc;
end;

procedure TFrm_Inc_Recebimento.EContabil2Exit(Sender: TObject);
Var
  VarCD : string;
begin
  if not DXBCancelar.Focused Then Begin
    if not AchaPlanoDeConta(300,20,'S',EContabil2.Text,'D') then Begin
      EContabil2.SetFocus;
      exit;
    end;
    EContabil2.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Text+' - '+DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    if not (ZQdividirpor.State in [DsInsert, DsEdit]) Then
      ZQdividirpor.Edit;
    ZQdividirporcodconta.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
    ZQdividirportipdoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
  end;
end;

procedure TFrm_Inc_Recebimento.XPartesExit(Sender: TObject);
begin
  if empty(Enomes.Text) then begin
    showmessage('O nome do participante deve estar preenchido...');
    Enomes.SetFocus;
    exit;
  end;
  ZQdividirpor.Close;
  if Xpartes.Value > 0 Then Begin
    ZQdividirpor.CreateDataSet;
    ZQdividirpor.DisableControls;
    while ZQdividirpor.RecordCount< XPartes.Value do Begin
      ZQdividirpor.Insert;
      ZQdividirporid.Value := ZQdividirpor.RecordCount;
      ZQdividirporpercent.Value := 100/Xpartes.Value;
      if ZQdividirporid.Value = 0 then
        ZQdividirporcodcli.Value := strtoint(Ecodnome.Text);
      ZQdividirpor.Post;
    end;
    ZQdividirpor.EnableControls;
    ZQdividirpor.IndexFieldNames := 'id';
    ZQdividirpor.first;
    Erecebedor.Enabled := True;
    EContabil2.Enabled := True;
    DBCBTipdocdiv.Enabled := True;
    XDBNumEdit2.Enabled := True;
    DBGDiv.Enabled := True;
    Erecebedor.SetFocus;
  end
  else Begin
    Erecebedor.Enabled := False;
    EContabil2.Enabled := False;
    DBCBTipdocdiv.Enabled := False;
    XDBNumEdit2.Enabled := False;
    DBGDiv.Enabled := False;
    DXBGravar.SetFocus;
  end;
end;

procedure TFrm_Inc_Recebimento.DBGParcelasDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBEOrdem.Text = CDSParcelasordem.Value Then Begin
    DBGParcelas.Canvas.Brush.Color := $006CFFFF;
    DBGParcelas.Canvas.Font.Color := $00A80000;
    DBGParcelas.Canvas.Font.Style := [FsBold];
  end;
  DBGParcelas.DefaultDrawDataCell(Rect, DBGParcelas.columns[datacol].field, State);
end;

procedure TFrm_Inc_Recebimento.DBGDivDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if DBECod.Text = ZQdividirporid.Text Then Begin
    DBGDiv.Canvas.Brush.Color := $006CFFFF;
    DBGDiv.Canvas.Font.Color := $00A80000;
    DBGDiv.Canvas.Font.Style := [FsBold];
  end;
  DBGDiv.DefaultDrawDataCell(Rect, DBGDiv.columns[datacol].field, State);

end;

procedure TFrm_Inc_Recebimento.ErecebedorExit(Sender: TObject);
begin
  If not DXBCancelar.Focused Then Begin
    if not Achacliente(180, 140,Erecebedor.Text) Then Begin
      Erecebedor.SetFocus;
      Exit;
    End;
    ZQdividirpor.Edit;
    ZQdividirporcodcli.Value := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt;
    Erecebedor.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
    if not empty(DM_Tabelas.ZQAchaParticip.FieldByName('codpaginc').Text) then
      EContabil2.Text := DM_Tabelas.ZQAchaParticip.FieldByName('codpaginc').Text;
  End;
end;

procedure TFrm_Inc_Recebimento.XDBNumEdit2Exit(Sender: TObject);
Var
  varregis: Integer;
  varsoma, vardif : Double;
begin
  varsoma := 0;
  varregis := ZQdividirpor.RecNo;
  ZQdividirpor.First;
  ZQdividirpor.DisableControls;
  while not (ZQdividirpor.Eof) and (ZQdividirpor.RecNo <= varregis) do Begin
    varsoma := varsoma +ZQdividirporpercent.Value;
    ZQdividirpor.next;
  end;
  ZQdividirpor.EnableControls;
  vardif := 100-varsoma;
  ZQdividirpor.RecNo := varregis;
  if ZQdividirpor.RecNO< XPartes.Value Then Begin
    ZQdividirpor.RecNO := varregis+1;
    while not ZQdividirpor.Eof do Begin
      ZQdividirpor.Edit;
      ZQdividirporpercent.Value := vardif / (XPartes.Value-varregis);
      ZQdividirpor.Next;
    end;
  end;
  ZQdividirpor.RecNo := varregis ;
  if not ZQdividirpor.Eof Then Begin
    ZQdividirpor.RecNo := varregis+1;
    Erecebedor.SetFocus;
  end
  else
    DXBGravar.SetFocus;
end;

procedure TFrm_Inc_Recebimento.XDBNumEdit2Enter(Sender: TObject);
begin
  if ZQdividirpor.RecNo=XPartes.Value Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Inc_Recebimento.ErecebedorEnter(Sender: TObject);
begin

  Erecebedor.Text := ZQdividirpornomecli.Value;
end;

procedure TFrm_Inc_Recebimento.EContabil2Enter(Sender: TObject);
begin
  if not empty(ZQdividirpordescriconta.Value) Then
    EContabil2.Text := ZQdividirpordescriconta.Value;
end;

procedure TFrm_Inc_Recebimento.XDBNVrParcEnter(Sender: TObject);
begin
  if CDSParcelas.RecNo = XNEParcelas.Value Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrm_Inc_Recebimento.RBRecClick(Sender: TObject);
begin
  if RBRec.Checked then
    Panel_particpantes.Visible := False
  else
    Panel_particpantes.Visible := False;

end;

procedure TFrm_Inc_Recebimento.RBPagClick(Sender: TObject);
begin
  if RBRec.Checked then
    Panel_particpantes.Visible := False
  else
    Panel_particpantes.Visible := False;

end;

procedure TFrm_Inc_Recebimento.DBCBTipdocdivExit(Sender: TObject);
begin
  if not DXBCancelar.Focused Then Begin
    if not DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipdocdiv.Text,[])Then Begin
      DBCBTipdocdiv.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').AsString='S' Then Begin
      if CDSParcelasVenci.Value<>date Then BEgin
        showmessage('Este tipo de documento só aceita pagamento avista!!!');
        EContabil.SetFocus;
        Exit;
      end;
    end;
    Label15.Visible := False;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin
      Label15.Visible := True;
      GBCheque.Visible := True;
      DBEBanco.SetFocus;
    end;
    vartipodoc  := ZQdividirportipdoc.Value;
  end;
end;

procedure TFrm_Inc_Recebimento.DBCBTipdocdivEnter(Sender: TObject);
begin
  if (ZQdividirpor.RecNo>1) and (empty(ZQdividirportipdoc.Text)) Then
    ZQdividirportipdoc.Value  := vartipodoc;
end;

procedure TFrm_Inc_Recebimento.EQuadraloteExit(Sender: TObject);
begin
  ZROQReceb.Close;
  ZROQReceb.SQL.Clear;
  ZROQReceb.SQL.Add('select idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  ZROQReceb.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  ZROQReceb.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  ZROQReceb.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
  ZROQReceb.SQL.Add(' from recebimento where quadralote='+quotedstr(EQuadralote.Text));
  ZROQReceb.Open;
  ZROQReceb.Last;
  if ZROQReceb.RecordCount>0 then
  begin
    Enomes.Text:= ZROQReceb.FieldByName('nome_parte').AsString;
    Ecodnome.Text:=ZROQReceb.FieldByName('cliente').Text;
    Eadversa.Text:=ZROQReceb.FieldByName('adversa').Text;
//    XNEValor.Value:=ZROQRecebVrDoc.Value;
    Edit1.Text:=ZROQReceb.FieldByName('documento').AsString;
  end;
end;


procedure TFrm_Inc_Recebimento.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Inc_Recebimento, 'ZROQReceb', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebnome_parte', 'nome_parte', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebdocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebcliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebsaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebmarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebrefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebcontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebcustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQReceborigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebrecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebnumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebvenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebnumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebsq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebsomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebnomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebjuros', 'juros', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebdescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebobserv_estorno', 'observ_estorno', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebtip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebsld_antes_reajuste', 'sld_antes_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Inc_Recebimento, 'ZROQReceb', 'ZROQRecebPercentual_reajuste', 'Percentual_reajuste', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.

