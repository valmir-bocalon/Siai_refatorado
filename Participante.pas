unit Participante;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, TabNotBk, Grids, DBGrids, StdCtrls,  dxButton,
  XBanner, ExtCtrls, wwdbdatetimepicker, DBCtrls, Mask, DB, XDBEdit, XDBNum,
  JvExControls, JvXPCore, JvXPBar, ZAbstractRODataset, ZAbstractDataset,
  ZDataset, dxCore2;

type
  TFrm_Participante = class(TForm)
    Panel2: TPanel;
    XBanner8: TXBanner;
    BTExcluir: TdxButton;
    BtEditar: TdxButton;
    BtIncluir: TdxButton;
    BtCancelar: TdxButton;
    BtGravar: TdxButton;
    BTPesquisar: TdxButton;
    BtRelatorio: TdxButton;
    BtUltimo: TdxButton;
    BtProximo: TdxButton;
    BTAnterior: TdxButton;
    BtPrimeiro: TdxButton;
    BTFechar: TdxButton;
    BtProcessar: TdxButton;
    Label2: TLabel;
    DBGPart: TDBGrid;
    LReg: TLabel;
    Registros: TLabel;
    Pagina: TPageControl;
    TabPag1: TTabSheet;
    TabPag2: TTabSheet;
    TabPag3: TTabSheet;
    TabPag4: TTabSheet;
    XBanner1: TXBanner;
    Label1: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Label19: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label16: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label27: TLabel;
    Label4: TLabel;
    Label57: TLabel;
    DBECod: TDBEdit;
    DBERazao: TDBEdit;
    DBEDoc2: TDBEdit;
    DBEFone1: TDBEdit;
    DBEFone2: TDBEdit;
    DBEFone3: TDBEdit;
    DBENacionalidade: TDBEdit;
    DBEEmail: TDBEdit;
    DBCBPessoa: TDBComboBox;
    DBEAniversario: TwwDBDateTimePicker;
    ECidNaturalidade: TEdit;
    Eestadonatural: TEdit;
    DBETrab: TDBEdit;
    DBEProf: TDBEdit;
    DBEDoc1: TXDBEdit;
    DBEEstCivil: TDBComboBox;
    DBMemo1: TDBMemo;
    DBEREnda: TXDBNumEdit;
    DBCBRegime: TDBComboBox;
    DBEFone4: TDBEdit;
    GBPacto: TGroupBox;
    XBanner6: TXBanner;
    Label29: TLabel;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    DBEDataCasou: TwwDBDateTimePicker;
    DBECartorio: TDBEdit;
    DBELivro: TDBEdit;
    DBEFolha: TDBEdit;
    ECidcasou: TEdit;
    ECidComarca: TEdit;
    EEstCart: TEdit;
    EEstComarca: TEdit;
    wwDBDateTimePicker1: TwwDBDateTimePicker;
    GBCodPag: TGroupBox;
    Label52: TLabel;
    Label53: TLabel;
    XBanner11: TXBanner;
    EContabil1: TEdit;
    EContabil2: TEdit;
    GBCodRec: TGroupBox;
    Label54: TLabel;
    Label55: TLabel;
    XBanner9: TXBanner;
    EContabil3: TEdit;
    EContabil4: TEdit;
    XBanner2: TXBanner;
    GroupBox1: TGroupBox;
    XBanner3: TXBanner;
    Label9: TLabel;
    Label10: TLabel;
    Label8: TLabel;
    Label25: TLabel;
    Label7: TLabel;
    Label51: TLabel;
    DBECep: TDBEdit;
    DBEEnde: TDBEdit;
    DBEBairr: TDBEdit;
    ECidEnder: TEdit;
    EEstado: TEdit;
    DBECompl: TDBEdit;
    GroupBox2: TGroupBox;
    XBanner4: TXBanner;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label26: TLabel;
    Label56: TLabel;
    DBEEndeCob: TDBEdit;
    DBEBairrCob: TDBEdit;
    ECidaddeCob: TEdit;
    EestadoCob: TEdit;
    DBECepCob: TDBEdit;
    DBEComplCob: TDBEdit;
    Label18: TLabel;
    Label20: TLabel;
    Label28: TLabel;
    Label31: TLabel;
    Label30: TLabel;
    Label37: TLabel;
    Label50: TLabel;
    DBEDoc1Resp: TDBEdit;
    DBEDoc2Resp: TXDBEdit;
    DBEEmailResp: TDBEdit;
    DBENomeResp: TDBEdit;
    DBENascconj: TwwDBDateTimePicker;
    ECidNatConj: TEdit;
    GroupBox4: TGroupBox;
    XBanner7: TXBanner;
    Label36: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    DBEProfResp: TDBEdit;
    DBECargo: TDBEdit;
    DBEAdmiss: TwwDBDateTimePicker;
    DBERendaResponsa: TXDBNumEdit;
    Eestconj: TEdit;
    CBAssina: TCheckBox;
    XBanner5: TXBanner;
    XBanner12: TXBanner;
    TabPag5: TTabSheet;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    XBanner13: TXBanner;
    DBERepDoc1: TDBEdit;
    DBERepDoc2: TXDBEdit;
    DBEEmailRep: TDBEdit;
    DBEResp: TDBEdit;
    DBENascRep: TwwDBDateTimePicker;
    ERepNatur: TEdit;
    ERespEst: TEdit;
    DBGResp: TDBGrid;
    PagResp: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    XBanner14: TXBanner;
    Label65: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    Label68: TLabel;
    DBEReprEmpr: TDBEdit;
    DBERepCargo: TDBEdit;
    DBERepAdmiss: TwwDBDateTimePicker;
    DBERespRenda: TXDBNumEdit;
    CBReprAss: TCheckBox;
    XBanner15: TXBanner;
    DBERepConjNome: TDBEdit;
    Label69: TLabel;
    ERepConjNatCid: TEdit;
    Label70: TLabel;
    ERepConjNatUF: TEdit;
    Label71: TLabel;
    DBERepConjDoc2: TDBEdit;
    DBERepConjDoc1: TXDBEdit;
    DBERepConjNasc: TwwDBDateTimePicker;
    Label72: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    CBRepConjAss: TCheckBox;
    DBERepConjCasDT: TwwDBDateTimePicker;
    Label75: TLabel;
    Label76: TLabel;
    DBERepConjCartorio: TDBEdit;
    Label77: TLabel;
    DBERepConjLivro: TDBEdit;
    Label78: TLabel;
    DBERepConjFolha: TDBEdit;
    ERepConjCasUF: TEdit;
    Label79: TLabel;
    ERepConjCasCid: TEdit;
    Label80: TLabel;
    Label81: TLabel;
    DBEdit1: TDBEdit;
    Label82: TLabel;
    DBEdit2: TDBEdit;
    DS_Participantes: TDataSource;
    ZqParticipantes: TZQuery;
    TabSheet3: TTabSheet;
    XBanner16: TXBanner;
    DS_ZQContratos: TDataSource;
    ZQContratos: TZQuery;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    ZQuery1: TZQuery;
    DS_BxEntrada: TDataSource;
    ZQBxEntrada: TZQuery;
    DS_BsParcela: TDataSource;
    ZQBxParcela: TZQuery;
    btnimprimir: TdxButton;
    Label83: TLabel;
    DBComboBox1: TDBComboBox;
    Label84: TLabel;
    DBSEXO: TDBComboBox;
    TabSheet4: TTabSheet;
    Label85: TLabel;
    cc: TwwDBDateTimePicker;
    Label86: TLabel;
    DBComboBox2: TDBComboBox;
    Label87: TLabel;
    cex: TwwDBDateTimePicker;
    Label88: TLabel;
    jb: TwwDBDateTimePicker;
    XBanner17: TXBanner;
    Label89: TLabel;
    Label90: TLabel;
    dbquadra: TDBEdit;
    dblote: TDBEdit;
    dxButton1: TdxButton;
    DS_zqpart: TDataSource;
    ZQpart: TZQuery;
    TabSheet5: TTabSheet;
    XBanner18: TXBanner;
    DS_zqpart2: TDataSource;
    ZQPart2: TZQuery;
    DBText1: TDBText;
    Timer1: TTimer;
    Jrel: TJvXPBar;
    XBanner10: TXBanner;
    ProgressBar1: TProgressBar;
    TabSheet6: TTabSheet;
    XBanner19: TXBanner;
    DBMemo2: TDBMemo;
    ZqParticipantes_doc: TZQuery;
    ZQConjuge: TZQuery;
    DS_Conjuge: TDataSource;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBCBPessoaChange(Sender: TObject);
    procedure atualiza_tela;
    procedure botoes;
    procedure BtPrimeiroClick(Sender: TObject);
    procedure BTAnteriorClick(Sender: TObject);
    procedure BtProximoClick(Sender: TObject);
    procedure BtUltimoClick(Sender: TObject);
    procedure BtGravarClick(Sender: TObject);
    procedure BtCancelarClick(Sender: TObject);
    procedure BtIncluirClick(Sender: TObject);
    procedure BtEditarClick(Sender: TObject);
    procedure BTExcluirClick(Sender: TObject);
    procedure BTPesquisarClick(Sender: TObject);
    procedure BtRelatorioClick(Sender: TObject);
    procedure BTFecharClick(Sender: TObject);
    procedure ativa_campos;
    procedure desativa_campos;
    procedure DBCBPessoaExit(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DBEEndeExit(Sender: TObject);
    procedure DBEBairrExit(Sender: TObject);
    procedure ECidEnderExit(Sender: TObject);
    procedure DBGPartEnter(Sender: TObject);
    procedure DBGPartKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGPartMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBGPartDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure ECidaddeCobExit(Sender: TObject);
    procedure ECidNaturalidadeExit(Sender: TObject);
    procedure DBEDoc1Exit(Sender: TObject);
    procedure DBEDoc1Enter(Sender: TObject);
    procedure DBEDoc2RespEnter(Sender: TObject);
    procedure DBEDoc2RespExit(Sender: TObject);
    procedure ERepNaturExit(Sender: TObject);
    procedure CBAssinaExit(Sender: TObject);
    procedure DBCBRegimeExit(Sender: TObject);
    procedure DBEEstCivilExit(Sender: TObject);
    procedure DBECepExit(Sender: TObject);
    procedure ECidcasouExit(Sender: TObject);
    procedure ECidComarcaExit(Sender: TObject);
    procedure GBPactoExit(Sender: TObject);
    procedure DBMemo1Exit(Sender: TObject);
    procedure DBEComplExit(Sender: TObject);
    procedure DBERespExit(Sender: TObject);
    procedure CBReprAssExit(Sender: TObject);
    procedure DBERepDoc2Exit(Sender: TObject);
    procedure DBERepDoc2Enter(Sender: TObject);
    procedure ECidNatConjExit(Sender: TObject);
    procedure atualiza_tela_responsavel;
    procedure DBGRespKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGRespMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PagRespExit(Sender: TObject);
    procedure CBRepConjAssExit(Sender: TObject);
    procedure ERepConjNatCidExit(Sender: TObject);
    procedure ERepConjCasCidExit(Sender: TObject);
    procedure JrelItems0Click(Sender: TObject);
    procedure DBEEndeEnter(Sender: TObject);
    procedure btnimprimirClick(Sender: TObject);
    procedure DBComboBox2Exit(Sender: TObject);
    procedure ccEnter(Sender: TObject);
    procedure cexEnter(Sender: TObject);
    procedure jbEnter(Sender: TObject);
    procedure dbquadraEnter(Sender: TObject);
    procedure dxButton1Click(Sender: TObject);
    procedure JrelItems1Click(Sender: TObject);
    procedure JrelItems2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure PaginaChange(Sender: TObject);
    procedure DBMemo2Enter(Sender: TObject);
    procedure DBMemo2Exit(Sender: TObject);
    procedure JrelItems3Click(Sender: TObject);
    procedure JrelItems4Click(Sender: TObject);
    procedure JrelItems5Click(Sender: TObject);
    procedure DBGPartDblClick(Sender: TObject);
    procedure JrelItems6Click(Sender: TObject);
  private
    { Private declarations }

    procedure AfterConstruction; override;
  public
    { Public declarations }
  end;

var
  Frm_Participante: TFrm_Participante;

implementation

uses tabelas, funcoes, hormain, PesqEndereco, Relparticipante, Umapa,
  RelReceb02_mensal, hormain2, Relparticipante1, Relparticipante_empre,
  Relparticipante1_composta, Relparticipante_empre_completa,
  Relparticipante_ficha, uRuntimeFields;

{$R *.dfm}

function LookupConjugeCidade(const ALookupField, AKeyField,
  AResultField: string): string;
var
  LField: TField;
  LValue: Variant;
begin
  Result := '';
  if (DM_Tabelas = nil) or (DM_Tabelas.ZQConjuge = nil) or
     (DM_Tabelas.ZQCidade = nil) or (not DM_Tabelas.ZQConjuge.Active) or
     DM_Tabelas.ZQConjuge.IsEmpty then
    Exit;

  { Use o campo calculado quando o runtime já o criou. }
  LField := DM_Tabelas.ZQConjuge.FindField(ALookupField);
  if (LField <> nil) and not LField.IsNull then
  begin
    Result := LField.AsString;
    if Result <> '' then
      Exit;
  end;

  { Fallback para a chave física, evitando que a tela dependa da ordem de
    criação do lookup durante a primeira abertura do cadastro. }
  LField := DM_Tabelas.ZQConjuge.FindField(AKeyField);
  if (LField = nil) or LField.IsNull or (not DM_Tabelas.ZQCidade.Active) then
    Exit;
  LValue := DM_Tabelas.ZQCidade.Lookup('idcidade', LField.Value, AResultField);
  if not VarIsNull(LValue) and not VarIsEmpty(LValue) then
    Result := VarToStr(LValue);
end;

procedure TFrm_Participante.FormShow(Sender: TObject);
begin
  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  DM_Tabelas.ZqParticipante.open;
  DM_Tabelas.ZQresponsavel.open;
  DM_TAbelas.ZQConjuge.Open;
//  ZQContratos.open;
  pagina.ActivePageIndex := 0;
  desativa_campos;
  { A abertura nao passa pelos botoes de navegacao. Atualize o primeiro
    registro depois que Participante, Responsaveis e Conjuge estiverem ativos. }
  botoes;
  DBGPart.SetFocus;
//
end;

procedure TFrm_Participante.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  DM_Tabelas.ZQresponsavel.close;
  DM_TAbelas.ZQConjuge.Close;
  DM_Tabelas.ZqParticipante.close;
  ZQContratos.close;
  Frm_Participante:=nil;
  Action:=CaFree;
end;

procedure TFrm_Participante.DBCBPessoaChange(Sender: TObject);
begin
  atualiza_tela;
end;

procedure TFrm_Participante.atualiza_tela;
Begin
  LReg.Caption := inttostr(DM_Tabelas.ZqParticipante.RecordCount);
  Label5.Visible := True;
  if not (DM_tabelas.ZQresponsavel.State in [DsEdit, DsInsert]) Then Begin
    DM_Tabelas.ZQresponsavel.Filter := 'paticipante_idpaticipante='+quotedstr(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text);
    DM_Tabelas.ZQresponsavel.Filtered := True;
  end;
  Label36.Visible := False;
  DBEProfResp.Visible := False;
  if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'F' Then Begin
    Label36.Visible := True;
    DBEProfResp.Visible := True;
    TabPag2.Caption := 'Cônjuge';
    DBCBPessoa.Text := 'FÍSICA';
    Label3.Caption := 'Nome';
    Label5.Caption := 'R.G.';
    Label6.Caption := 'C.P.F.';
    Label22.Caption := 'Aniversário';
    if not (DM_TAbelas.ZQConjuge.State in [DsInsert, DsEdit]) Then
    Begin
      if not empty(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text) then
      begin
        DM_Tabelas.ZQConjuge.Filter := 'idparti='+DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text;
        DM_Tabelas.ZQConjuge.Filtered := True;
      end;
      if (DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString = '') or (DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString = 'AMASIADO(A)') or (DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString = 'CASADO(A)') or (DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString = 'REL.ESTAVEL') Then Begin
        Pagina.Pages[2].TabVisible := True;
        Pagina.Pages[3].TabVisible := False;
        if DM_Tabelas.ZQConjuge.FieldByName('regime').AsString = 'COMUNHÃO UNIVERSAL DE BENS' Then Begin
          GBPacto.Visible := True;
          ECidcasou.Text := LookupConjugeCidade('cart_cidnome', 'cidade_cart', 'nomecid');
          EEstCart.Text := LookupConjugeCidade('cart_cidest', 'cidade_cart', 'estado');
          ECidComarca.Text := LookupConjugeCidade('comar_cidnome', 'comarca', 'nomecid');
          EEstComarca.Text := LookupConjugeCidade('comar_cidest', 'comarca', 'estado');
        end
        else
          GBPacto.Visible := False;
      end
      else Begin
        Pagina.Pages[2].TabVisible := False;
        Pagina.Pages[3].TabVisible := True;
      end;
    end;
    ECidNatConj.Text := LookupConjugeCidade('nomecida', 'natural', 'nomecid');
    Eestconj.Text := LookupConjugeCidade('estcida', 'natural', 'estado');
    if DM_Tabelas.ZQConjuge.FieldByName('assina').AsString = 'S' Then
      CBAssina.Checked := True
    else
      CBAssina.Checked := False;
    DBEDoc2.Visible := True;
    DBEEstCivil.Visible := True;
    Label16.Visible := True;
    if pos(DM_Tabelas.ZqParticipante.FieldByName('ESTADOCIVIL').AsString,'REL.ESTAVEL AMASIADO(A) CASADO(A)')>0 Then Begin
      DBCBRegime.Visible := True;
      Label27.Visible := True;
      DBCBRegime.enabled:=true;
      DBCBRegime.readonly:=false;
    end
    else Begin
      DBCBRegime.Visible := FAlse;
      Label27.Visible := False;
    end;
    Label33.Visible := True;
    Label34.Visible := True;
    DBETrab.Visible := True;
    DBEProf.Visible := True;
  end
  else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'J' Then Begin
    DBCBPessoa.Text := 'JURÍDICA';
    Label3.Caption := 'Razão';
    Label5.Caption := 'Insc. Est.';
    Label6.Caption := 'C.N.P.J.';
    Label22.Caption := 'Inaugurada';
    DBEDoc2.Visible := True;
    DBEEstCivil.Visible := False;
    Label16.Visible := False;
    GBPacto.Visible := False;
    DBCBRegime.Visible := False;
    Label27.Visible := False;
    Label33.Visible := False;
    Label34.Visible := False;
    DBETrab.Visible := False;
    DBEProf.Visible := False;
    Pagina.Pages[2].TabVisible := False;
    Pagina.Pages[3].TabVisible := True;

  end
  else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'R' Then Begin
    DBCBPessoa.Text := 'RURAL';
    Label3.Caption := 'Razão';
    Label5.Caption := 'Insc. Rural';
    Label6.Caption := 'C.N.P.J.';
    Label22.Caption := 'Inaugurada';
    DBEDoc2.Visible := True;
    DBEEstCivil.Visible := False;
    Label16.Visible := False;
    GBPacto.Visible := False;
    DBCBRegime.Visible := False;
    Label27.Visible := False;
    Label33.Visible := False;
    Label34.Visible := False;
    DBETrab.Visible := False;
    DBEProf.Visible := False;
  end
  else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'P' Then Begin
    DBCBPessoa.Text := 'Pública';
    Label3.Caption := 'Razão';
    Label6.Caption := 'C.N.P.J.';
    Label22.Caption := 'Inaugurada';
    Label5.Visible := False;
    DBEDoc2.Visible := False;
    DBEEstCivil.Visible := False;
    Label16.Visible := False;
    GBPacto.Visible := False;
    DBCBRegime.Visible := False;
    Label27.Visible := False;
    Label33.Visible := False;
    Label34.Visible := False;
    DBETrab.Visible := False;
    DBEProf.Visible := False;
  end
  else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'I' Then Begin
    Label36.Visible := True;
    DBEProfResp.Visible := True;
//    TabPag2.Caption := 'Cônjuge';
    DBCBPessoa.Text := 'IMPUBERE';
    Label3.Caption := 'Nome';
    Label5.Caption := 'R.G.';
    Label6.Caption := 'C.P.F.';
    Label22.Caption := 'Aniversário';

//        Pagina.Pages[1].TabVisible := False;
//        Pagina.Pages[2].TabVisible := True;

  end;

 // DBEDoc1.EditMask := '';
  DBEDoc2Resp.EditMask                                     := '';
  DBERepDoc2.EditMask                                      := '';
  ECidEnder.Text                                           := DM_Tabelas.ZqParticipante.FieldByName('nomecidade').AsString;
  EEstado.Text                                             := DM_Tabelas.ZqParticipante.FieldByName('nomeestado').AsString;
  ECidaddeCob.Text                                         := DM_Tabelas.ZqParticipante.FieldByName('nomecidadecob').AsString;
  EestadoCob.Text                                          := DM_Tabelas.ZqParticipante.FieldByName('nomeestadocob').AsString;
  ECidNaturalidade.Text                                    := DM_Tabelas.ZqParticipante.FieldByName('nomecidadenatural').AsString;
  Eestadonatural.Text                                      := DM_Tabelas.ZqParticipante.FieldByName('nomeestadonatural').AsString;
  DM_TAbelas.CDSResponsavelTemp.Close;
  DM_TAbelas.CDSResponsavelTemp.CreateDataSet;
  DM_TAbelas.ZQresponsavel.First;
  DM_TAbelas.ZQresponsavel.DisableControls;
  ProgressBar1.Visible                                     :=true;
  ProgressBar1.Position                                    :=0;
  ProgressBar1.Max                                         :=DM_TAbelas.ZQresponsavel.RecordCount;
  while not DM_TAbelas.ZQresponsavel.Eof do Begin
    ProgressBar1.Position                                  :=DM_TAbelas.ZQresponsavel.RecNo;
    DM_TAbelas.CDSResponsavelTemp.Insert;
    DM_TAbelas.CDSResponsavelTempdocumen1.Value            := DM_TAbelas.ZQresponsavel.FieldByName('documen1').AsString;
    DM_TAbelas.CDSResponsavelTempdocumen2.Value            := DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString;
    DM_Tabelas.CDSResponsavelTempnasicmento.Value          := DM_TAbelas.ZQresponsavel.FieldByName('nasicmento').AsDateTime;
    DM_TAbelas.CDSResponsavelTempcid_natural.Value         := DM_TAbelas.ZQresponsavel.FieldByName('cid_natural').AsLargeInt;
    DM_TAbelas.CDSResponsavelTempnomerepres.Value          := DM_TAbelas.ZQresponsavel.FieldByName('nomerepres').AsString;
    DM_TAbelas.CDSResponsavelTempemailrepres.Value         := DM_TAbelas.ZQresponsavel.FieldByName('emailrepres').AsString;
    DM_TAbelas.CDSResponsavelTempprofiss.Value             := DM_TAbelas.ZQresponsavel.FieldByName('profiss').AsString;
    DM_Tabelas.CDSResponsavelTempassina.Value              := DM_Tabelas.ZQresponsavel.FieldByName('assina').AsString;
    DM_Tabelas.CDSResponsavelTemprenda.Value               := DM_Tabelas.ZQresponsavel.FieldByName('renda').AsFloat;
    DM_TAbelas.CDSResponsavelTemptrabalho.Value            := DM_TAbelas.ZQresponsavel.FieldByName('trabalho').AsString;
    DM_Tabelas.CDSResponsavelTempadmissao.Value            := DM_Tabelas.ZQresponsavel.FieldByName('admissao').AsDateTime;
    DM_Tabelas.CDSResponsavelTempconjugenome.Value         := DM_Tabelas.ZQresponsavel.FieldByName('nomeconjuge').AsString;
    DM_Tabelas.CDSResponsavelTempconjugedoc1.Value         := DM_Tabelas.ZQresponsavel.FieldByName('doc1conjuge').AsString;
    DM_Tabelas.CDSResponsavelTempconjugedoc2.Value         := DM_Tabelas.ZQresponsavel.FieldByName('doc2conjuge').AsString;
    DM_Tabelas.CDSResponsavelTempconjugenasc.Value         := DM_TAbelas.ZQresponsavel.FieldByName('nascconjuge').AsDateTime;
    DM_TAbelas.CDSResponsavelTempconjugeass.Value          := DM_Tabelas.ZQresponsavel.FieldByName('assconjuge').AsString;
    DM_Tabelas.CDSResponsavelTempconjugenatur.Value        := DM_Tabelas.ZQresponsavel.FieldByName('cidconjuge').AsInteger;
    DM_Tabelas.CDSResponsavelTempconjugeDT.Value           := DM_Tabelas.ZQresponsavel.FieldByName('data').AsDateTime;
    DM_Tabelas.CDSResponsavelTempconjugecartorio.Value     := DM_Tabelas.ZQresponsavel.FieldByName('cartorio').AsString;
    DM_Tabelas.CDSResponsavelTempconjugelivro.Value        := DM_Tabelas.ZQresponsavel.FieldByName('livro').AsString;
    DM_Tabelas.CDSResponsavelTempconjugefolha.Value        := DM_Tabelas.ZQresponsavel.FieldByName('folha').AsString;
    DM_Tabelas.CDSResponsavelTempconjuge_cidcartorio.Value := DM_Tabelas.ZQresponsavel.FieldByName('cidade_cart').AsInteger ;
    DM_TAbelas.CDSResponsavelTemp.Post;
    DM_TAbelas.ZQresponsavel.Next;
  end;
  DM_TAbelas.ZQresponsavel.EnableControls;
  ProgressBar1.Position:=0;
  ProgressBar1.Visible:=false;
  atualiza_tela_responsavel;
End;

procedure TFrm_Participante.botoes;
Begin
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  if DM_Tabelas.ZqParticipante.Recno<=1 Then Begin
    BTPrimeiro.Enabled := False;
    BTAnterior.Enabled := False;
  end;
  if DM_Tabelas.ZqParticipante.RecNo=DM_Tabelas.ZqParticipante.RecordCount Then Begin
    BTProximo.Enabled := False;
    BtUltimo.Enabled := False;
  end;
  atualiza_tela;

End;

procedure TFrm_Participante.BtPrimeiroClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.First;
  botoes;
end;

procedure TFrm_Participante.BTAnteriorClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Prior;
  botoes;
end;

procedure TFrm_Participante.BtProximoClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Next;
  botoes;
end;

procedure TFrm_Participante.BtUltimoClick(Sender: TObject);
begin
  DM_Tabelas.ZqParticipante.Last;
  botoes;
end;

procedure TFrm_Participante.BtGravarClick(Sender: TObject);
var
  varinsert : boolean;
begin
  if not Verif_senha('Participante','Gravar inclusão ou edição','Participante: '+DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text+' - Qd.'+DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString) Then Exit;
  varinsert := False;
  if DM_Tabelas.ZqParticipante.State in [DsInsert] Then
    varinsert := True;
  if DM_Tabelas.zqparticipante.State in [DsInsert,DsEdit] Then
  begin
    DM_Tabelas.ZqParticipante.Post;
  end;
  DM_Tabelas.ZQresponsavel.First;
  while not DM_Tabelas.ZQresponsavel.Eof do
    DM_Tabelas.ZQresponsavel.Delete;
  DM_Tabelas.CDSResponsavelTemp.First;
  while not DM_Tabelas.CDSResponsavelTemp.Eof do begin
    DM_Tabelas.ZQresponsavel.Insert;
    DM_Tabelas.ZQresponsavel.FieldByName('paticipante_idpaticipante').AsLargeInt := DM_TAbelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
    DM_TAbelas.ZQresponsavel.FieldByName('documen1').AsString := DM_TAbelas.CDSResponsavelTempdocumen1.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString := DM_TAbelas.CDSResponsavelTempdocumen2.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('nasicmento').AsDateTime := DM_Tabelas.CDSResponsavelTempnasicmento.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('cid_natural').AsLargeInt := DM_TAbelas.CDSResponsavelTempcid_natural.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('nomerepres').AsString := DM_TAbelas.CDSResponsavelTempnomerepres.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('emailrepres').AsString:= DM_TAbelas.CDSResponsavelTempemailrepres.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('profiss').AsString := DM_TAbelas.CDSResponsavelTempprofiss.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('assina').AsString := DM_Tabelas.CDSResponsavelTempassina.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('renda').AsFloat := DM_Tabelas.CDSResponsavelTemprenda.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('trabalho').AsString := DM_TAbelas.CDSResponsavelTemptrabalho.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('admissao').AsDateTime := DM_Tabelas.CDSResponsavelTempadmissao.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('nomeconjuge').AsString := DM_Tabelas.CDSResponsavelTempconjugenome.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('doc1conjuge').AsString := DM_Tabelas.CDSResponsavelTempconjugedoc1.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('doc2conjuge').AsString := DM_Tabelas.CDSResponsavelTempconjugedoc2.Value;
    DM_TAbelas.ZQresponsavel.FieldByName('nascconjuge').AsDateTime := DM_Tabelas.CDSResponsavelTempconjugenasc.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('assconjuge').AsString := DM_TAbelas.CDSResponsavelTempconjugeass.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('cidconjuge').AsInteger := DM_Tabelas.CDSResponsavelTempconjugenatur.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('data').AsDateTime := DM_Tabelas.CDSResponsavelTempconjugeDT.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('cartorio').AsString := DM_Tabelas.CDSResponsavelTempconjugecartorio.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('livro').AsString := DM_Tabelas.CDSResponsavelTempconjugelivro.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('folha').AsString := DM_Tabelas.CDSResponsavelTempconjugefolha.Value;
    DM_Tabelas.ZQresponsavel.FieldByName('cidade_cart').AsInteger := DM_Tabelas.CDSResponsavelTempconjuge_cidcartorio.Value;
    DM_Tabelas.ZQresponsavel.Post;
    DM_TAbelas.CDSResponsavelTemp.Next;
  end;
  if  Pagina.Pages[2].TabVisible = True then
  begin
    if not empty(DM_Tabelas.ZQConjuge.FieldByName('nome').AsString) Then Begin
      DM_Tabelas.ZQConjuge.FieldByName('idparti').AsInteger := DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsLargeInt;
      DM_Tabelas.ZQConjuge.Post;
    end
    else
      DM_Tabelas.ZQConjuge.Cancel;
  end
  else
    DM_Tabelas.ZQConjuge.Cancel;

//  if empty(DM_Tabelas.ZQresponsaveldocumen2.Value) Then
//    DM_Tabelas.ZQresponsavel.Delete;
  desativa_campos;
//  ZQContratos.Refresh;
  DBGPart.SetFocus;
end;

procedure TFrm_Participante.BtCancelarClick(Sender: TObject);
begin
  if not Verif_senha('Participante','Cancelar inclusão ou edição','Participante: '+DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text+' - Qd.'+DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString) Then Exit;
  DM_Tabelas.ZqParticipante.Cancel;
  DM_Tabelas.ZQresponsavel.Cancel;
  DM_Tabelas.ZQConjuge.Cancel;
  desativa_campos;
end;

procedure TFrm_Participante.BtIncluirClick(Sender: TObject);
begin
  if not Verif_senha('Participante','Incluir','') Then Exit;
  pagina.ActivePageIndex := 1;
  ECidNaturalidade.Text := '';
  DM_Tabelas.ZqParticipante.Insert;
//  DM_Tabelas.ZQConjuge.Insert;
  ativa_campos;
  DM_Tabelas.ZqParticipante.FieldByName('cadastrado').AsDateTime := date;
  DBCBPessoa.SetFocus;
end;

procedure TFrm_Participante.BtEditarClick(Sender: TObject);
begin
  if not Verif_senha('Participante','Editar','Participante: '+DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text+' - Qd.'+DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString) Then Exit;
  ativa_campos;
  if Pagina.ActivePageIndex = 1 Then
    DBCBPessoa.SetFocus
  else if Pagina.ActivePageIndex = 2 Then Begin
    DBENomeResp.SetFocus;
  end
  else if Pagina.ActivePageIndex = 3 Then Begin
    DBEResp.SetFocus;
  end
  else if Pagina.ActivePageIndex = 4 Then
    DBEEnde.SetFocus;
  DM_Tabelas.ZqParticipante.Edit;
  DM_Tabelas.CDSResponsavelTemp.Edit;
  DM_TAbelas.ZQConjuge.Edit;
end;

procedure TFrm_Participante.BTExcluirClick(Sender: TObject);
begin
  if not Verif_senha('Participante','Exluir','Participante: '+DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text+' - Qd.'+DM_Tabelas.ZQParticipante.FieldByName('nome_parte').AsString) Then Exit;
  if simnao('Confirma a exclusão do participante ? ','SIM') then
     DM_Tabelas.ZqParticipante.Delete;
end;

procedure TFrm_Participante.BTPesquisarClick(Sender: TObject);
begin
  Self.FormStyle := fsNormal;
  try
    AchaCLiente(Self.Top+105,Self.Left+200,'');
    DM_Tabelas.ZqParticipante.Locate('idpaticipante',DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').AsLargeInt,[]);
    DBGPart.SetFocus;
  finally
    Self.FormStyle := fsStayOnTop;
    Self.BringToFront;
  end;
end;

procedure TFrm_Participante.BtRelatorioClick(Sender: TObject);
begin
  if not Verif_senha('Participante','Relatório','') Then Exit;
  jrel.Visible:=true;
end;

procedure TFrm_Participante.BTFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrm_Participante.ativa_campos;
Begin
  DBCBPessoa.ReadOnly := False;
  DBERazao.ReadOnly := False;
  DBEAniversario.ReadOnly := False;
  DBEDoc1.ReadOnly := False;
  DBEDoc2.ReadOnly := False;
  DBENacionalidade.ReadOnly := False;
  DBEFone1.ReadOnly := False;
  DBEFone2.ReadOnly := False;
  DBEFone3.ReadOnly := False;
  DBEFone4.ReadOnly := False;
  ECidNaturalidade.ReadOnly := False;
  DBETrab.ReadOnly := False;
  DBEProf.ReadOnly := False;
  DBEEmail.ReadOnly := False;
  DBEEnde.ReadOnly := False;
  DBEEndeCob.ReadOnly := False;
  DBEEnde.ReadOnly := False;
  DBEBairr.ReadOnly := False;
  ECidEnder.ReadOnly := False;
  DBECep.ReadOnly := False;
  DBEEndeCob.ReadOnly := False;
  DBEBairrCob.ReadOnly := False;
  ECidaddeCob.ReadOnly := False;
  DBECepCob.ReadOnly := False;
  ECidNatconj.ReadOnly := False;
  DBENomeResp.ReadOnly := False;
  DBEDoc1Resp.ReadOnly := False;
  DBEEstCivil.ReadOnly := False;
  DBEProfResp.ReadOnly := False;
  DBENascconj.ReadOnly := False;
  DBEEmailResp.ReadOnly := False;
  DBCBRegime.ReadOnly := False;
  DBEEstCivil.ReadOnly := False;
  DBEDataCasou.ReadOnly := False;
  DBECartorio.ReadOnly := False;
  DBELivro.ReadOnly := False;
  DBEFolha.ReadOnly := False;
  DBECargo.ReadOnly := False;
  DBEAdmiss.ReadOnly := False;
  DBERendaResponsa.ReadOnly := False;
  DBERenda.ReadOnly := False;
  ECidcasou.ReadOnly := False;
  ECidComarca.ReadOnly := False;
  CBAssina.Enabled := True;
  CBAssina.Enabled := True;
  BTPrimeiro.Enabled := False;
  BTAnterior.Enabled := False;
  BTProximo.Enabled := False;
  BtUltimo.Enabled := False;
  BtGravar.Enabled := True;
  BtCancelar.Enabled := True;
  BtProcessar.Enabled := False;
  BtFechar.Enabled := False;
  BtIncluir.Enabled := False;
  BtEditar.Enabled := False;
  BtExcluir.Enabled := False;
  BtPesquisar.Enabled := False;
  BtRelatorio.Enabled := False;
  DBGPart.Enabled := False;
  DBEResp.ReadOnly := False;
  DBERepDoc1.ReadOnly := False;
  DBERepDoc1.ReadOnly := False;
  ERepNatur.ReadOnly := False;
  DBENascRep.ReadOnly := False;
  CBReprAss.Enabled := True;
  DBEEmailRep.ReadOnly := False;
  DBEReprEmpr.ReadOnly := False;
  DBERepCargo.ReadOnly := False;
  DBERepAdmiss.ReadOnly := False;
  DBERespRenda.ReadOnly := False;
  DBERepConjNome.ReadOnly := False;
  ERepConjNatCid.ReadOnly := False;
  DBERepConjDoc2.ReadOnly := False;
  DBERepConjDoc1.ReadOnly := False;
  DBERepConjNasc.ReadOnly := False;
  DBERepConjCasDT.ReadOnly := False;
  DBERepConjCartorio.ReadOnly := False;
  DBERepConjLivro.ReadOnly := False;
  DBERepConjFolha.ReadOnly := False;
  ERepConjCasCid.ReadOnly := False;
  CBRepConjAss.Enabled := True;

  DBGResp.Enabled := False;
end;

procedure TFrm_Participante.desativa_campos;
Begin
  DBCBPessoa.ReadOnly := True;
  DBERazao.ReadOnly := True;
  DBEAniversario.ReadOnly := True;
  DBEDoc1.ReadOnly := True;
  DBEDoc2.ReadOnly := True;
  DBENacionalidade.ReadOnly := True;
  DBEFone1.ReadOnly := True;
  DBEFone2.ReadOnly := True;
  DBEFone3.ReadOnly := True;
  DBEFone4.ReadOnly := True;
  ECidNaturalidade.ReadOnly := True;
  DBETrab.ReadOnly := True;
  DBEProf.ReadOnly := True;
  DBEEmail.ReadOnly := True;
  DBEEnde.ReadOnly := True;
  DBEBairr.ReadOnly := True;
  DBEEnde.ReadOnly := True;
  DBEEndeCob.ReadOnly := True;
  ECidEnder.ReadOnly := True;
  DBECep.ReadOnly := True;
  DBEEndeCob.ReadOnly := True;
  DBEBairrCob.ReadOnly := True;
  ECidaddeCob.ReadOnly := True;
  DBECepCob.ReadOnly := True;
  ECidNatconj.ReadOnly := True;
  DBENomeResp.ReadOnly := True;
  DBEDoc1Resp.ReadOnly := True;
  DBEEstCivil.ReadOnly := True;
  DBEProfResp.ReadOnly := True;
  DBENascconj.ReadOnly := True;
  DBEEmailResp.ReadOnly := True;
  CBAssina.Enabled := True;
  DBCBRegime.ReadOnly := True;
  DBEEstCivil.ReadOnly := True;
  DBEDataCasou.ReadOnly := True;
  DBECartorio.ReadOnly := True;
  DBELivro.ReadOnly := True;
  DBEFolha.ReadOnly := True;
  DBECargo.ReadOnly := True;
  DBEAdmiss.ReadOnly := True;
  DBERendaResponsa.ReadOnly := True;
  DBERenda.ReadOnly := True;
  ECidcasou.ReadOnly := True;
  ECidComarca.ReadOnly := True;
  CBAssina.Enabled := False;
  BTPrimeiro.Enabled := True;
  BTAnterior.Enabled := True;
  BTProximo.Enabled := True;
  BtUltimo.Enabled := True;
  BtGravar.Enabled := False;
  BtCancelar.Enabled := False;
  BtFechar.Enabled := True;
  BtProcessar.Enabled := True;
  BtIncluir.Enabled := True;
  BtEditar.Enabled := True;
  BtExcluir.Enabled := True;
  BtPesquisar.Enabled := True;
  BtRelatorio.Enabled := True;
  DBGPart.Enabled := True;
  DBEResp.ReadOnly := True;
  DBERepDoc1.ReadOnly := True;
  DBERepDoc1.ReadOnly := True;
  ERepNatur.ReadOnly := True;
  DBENascRep.ReadOnly := True;
  CBReprAss.Enabled := False;
  DBEEmailRep.ReadOnly := True;
  DBEReprEmpr.ReadOnly := True;
  DBERepCargo.ReadOnly := True;
  DBERepAdmiss.ReadOnly := True;
  DBERespRenda.ReadOnly := True;
  DBERepConjNome.ReadOnly :=true;
  ERepConjNatCid.ReadOnly := true;
  DBERepConjDoc2.ReadOnly := true;
  DBERepConjDoc1.ReadOnly := true;
  DBERepConjNasc.ReadOnly := true;
  DBERepConjCasDT.ReadOnly := true;
  DBERepConjCartorio.ReadOnly := true;
  DBERepConjLivro.ReadOnly := true;
  DBERepConjFolha.ReadOnly := true;
  ERepConjCasCid.ReadOnly := true;
  CBRepConjAss.Enabled := False;
  DBGResp.Enabled := True;
end;

procedure TFrm_Participante.DBCBPessoaExit(Sender: TObject);
begin
  if pos(DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString,'FJPRI')=0 Then Begin
    DBCBPessoa.SetFocus;
    exit
  end;
  atualiza_tela;
end;

procedure TFrm_Participante.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  if not (ActiveControl is TDBGrid) then begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrm_Participante.DBEEndeExit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] then
  begin
    if empty(DM_tabelas.ZqParticipante.FieldByName('endereco').AsString) then
    begin
      if simnao('O Endereço Está em Branco. Deseja Cadastrar ?','SIM') then
      begin
        FrmPesqEndereco.Top := Self.Top+100;
        FrmPesqEndereco.Left := Self.Left;
        FrmPesqEndereco.showmodal;
        DM_tabelas.ZqParticipante.FieldByName('endereco').AsString:=DM_tabelas.localizacao.FieldByName('log_tipo_logradouro').AsString+' '+DM_tabelas.localizacao.FieldByName('log_no').AsString;
        DM_tabelas.ZqParticipante.FieldByName('bairro').AsString:=DM_tabelas.localizacao.FieldByName('bai_no').AsString;
        ECidEnder.Text:=DM_tabelas.localizacao.FieldByName('loc_no').AsString;
        EEstado.Text:=DM_tabelas.localizacao.FieldByName('ufe_sg').AsString;
        DM_tabelas.ZqParticipante.FieldByName('cep').AsString:=copy(DM_tabelas.localizacao.FieldByName('cep').AsString,1,2)+'.'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,3,3)+'-'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,6,3);
        DBECompl.SetFocus;
        DM_tabelas.localizacao.Close;
      end;
    end
  end;
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (Empty(DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString)) Then
  Begin
    DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString;
  end
  else if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString<>DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString) Then
  begin
    if simnao('Atualiza o Logradouro de Cobrança Igual ao Logradouro ?','SIM') then
    begin
        DM_Tabelas.ZqParticipante.FieldByName('ende_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('endereco').AsString;
        ECidaddeCob.Text := ECidEnder.Text;
        EestadoCob.Text := DM_Tabelas.ZqParticipante.FieldByName('nomeestado').AsString;
        DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString:=DM_Tabelas.ZqParticipante.FieldByName('cep').AsString;
    end;    
  end;
end;

procedure TFrm_Participante.DBEBairrExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (Empty(DM_Tabelas.ZqParticipante.FieldByName('bairro_cob').AsString)) Then Begin
    DM_Tabelas.ZqParticipante.FieldByName('bairro_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('bairro').AsString;
  end;

end;

procedure TFrm_Participante.ECidEnderExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then
  Begin
    if not Acha_Cidade(ECidEnder.text) Then
      ECidEnder.SetFocus
    else
    Begin
      ECidEnder.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstado.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      if empty(DM_Tabelas.ZqParticipante.FieldByName('cep').AsString) Then
        DM_Tabelas.ZqParticipante.FieldByName('cep').AsString := DM_tabelas.ZQAchaCidade.FieldByName('cepgeral').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('cidade').AsInteger := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
  if DM_Tabelas.ZqParticipante.FieldByName('cidade_cob').AsLargeInt=0 Then
    ECidaddeCob.Text := ECidEnder.Text;
  if (Empty(EestadoCob.Text)) Then
     EestadoCob.Text := DM_Tabelas.ZqParticipante.FieldByName('nomeestado').AsString;
end;

procedure TFrm_Participante.DBGPartEnter(Sender: TObject);
begin
  botoes;
end;

procedure TFrm_Participante.DBGPartKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  botoes;
end;

procedure TFrm_Participante.DBGPartMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
 botoes;

  if Pagina.TabIndex=7 then
  begin
    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('DROP TABLE IF EXISTS part');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('CREATE TABLE part ( `id` int(11) NOT NULL AUTO_INCREMENT, `observ_estorno` varchar(60000), `idpaticipante` int(11) NOT NULL,');
    ZQPart.SQL.Add(' PRIMARY KEY (`id`) ) ENGINE=MEMORY');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('INSERT INTO part (observ_estorno, idpaticipante) SELECT r.observ_estorno,p.idpaticipante');
    ZQPart.SQL.Add('FROM participante as p join recebimento as r on p.idpaticipante=r.adversa where p.idpaticipante='+quotedstr(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text)+' and r.observ_estorno is not null');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('select * from part');
    ZQPart.open;

    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('DROP TABLE IF EXISTS part2');
    ZQPart2.ExecSQL;


    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('CREATE TABLE part2 ( `id` int(11) NOT NULL AUTO_INCREMENT, `observ_estorno` varchar(60000), `idpaticipante` int(11) NOT NULL,');
    ZQPart2.SQL.Add(' PRIMARY KEY (`id`) ) ENGINE=MEMORY');
    ZQPart2.ExecSQL;

    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('select * from part2');
    ZQPart2.open;

    ZQPart.First;
    if ZQPart.RecordCount>0 then
    begin
      ZQPart2.Insert;
      ZQPart2.FieldByName('idpaticipante').AsInteger:=ZQpart.FieldByName('idpaticipante').AsInteger;
      ZQPart2.post;
      while not ZQPart.Eof do
      begin
        ZQPart2.Edit;
        ZQPart2.FieldByName('observ_estorno').AsString:=ZQPart2.FieldByName('observ_estorno').AsString+#10+#13+ZQpart.FieldByName('observ_estorno').AsString;
        ZQPart2.post;
        ZQPart.Next;
      end;
    end;
 end;
  
end;

procedure TFrm_Participante.DBGPartDblClick(Sender: TObject);
var
  IdParticipante: string;
  ImprimirFicha: Boolean;
begin
  jrel.Visible := False;
  if not Verif_senha('Participante', 'Relatórios', 'Ficha Completa') then
    Exit;

  IdParticipante := DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').AsString;

  Self.FormStyle := fsNormal;
  try
    ImprimirFicha := simnao('Deseja imprimir em formato de Ficha?', 'SIM');
  finally
    Self.FormStyle := fsStayOnTop;
  end;

  ZqParticipantes.Close;
  ZqParticipantes.SQL.Clear;
  ZqParticipantes.SQL.Add('Select * from participante where idpaticipante=' +
    IdParticipante);
  ZqParticipantes.Open;

  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;

  try
    Self.FormStyle := fsNormal;
    if ImprimirFicha then
    begin
      if FrmRelparticipante_ficha = nil then
        FrmRelparticipante_ficha := TFrmRelparticipante_ficha.Create(Application);
      FrmRelparticipante_ficha.RLReport1.PreviewModal;
      FreeAndNil(FrmRelparticipante_ficha);
    end
    else
    begin
      if FrmRelparticipante_composta = nil then
        FrmRelparticipante_composta := TFrmRelparticipante_composta.Create(Application);
      FrmRelparticipante_composta.RLReport1.PreviewModal;
      FreeAndNil(FrmRelparticipante_composta);
    end;
  finally
    Self.FormStyle := fsStayOnTop;
    ZqParticipantes.Close;
    ZqParticipantes.SQL.Clear;
    ZqParticipantes.SQL.Add('Select * from participante');
    ZqParticipantes.Open;
  end;
end;

procedure TFrm_Participante.DBGPartDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
{  if DBECod.Text = DM_Tabelas.ZqParticipanteidpaticipante.Text Then Begin
    DBGPart.Canvas.Brush.Color :=$006CFFFF;
    DBGPart.Canvas.Font.Color := $00A80000;
    DBGPart.Canvas.Font.Style := [FsBold];
  end;
  DBGPart.DefaultDrawDataCell(Rect, DBGPart.columns[datacol].field, State);
 }
end;

procedure TFrm_Participante.ECidaddeCobExit(Sender: TObject);
begin
{  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidaddeCob.text) Then
      ECidaddeCob.SetFocus
    else Begin
      ECidaddeCob.Text := DM_Tabelas.ZQAchaCidadenomecid.AsString;
      EestadoCob.Text := DM_Tabelas.ZQAchaCidadeestado.AsString;
      if empty(DM_Tabelas.ZqParticipantecep_cob.Value) Then
        DM_Tabelas.ZqParticipantecep_cob.Value := DM_tabelas.ZQAchaCidadecepgeral.Value;
      DM_Tabelas.ZqParticipantecidade_cob.Value := DM_Tabelas.ZQAchaCidadeidcidade.Value;
    end;
  end;}

  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then
  Begin
    if not Acha_Cidade(ECidaddeCob.text) Then
      ECidaddeCob.SetFocus
    else
    Begin
      ECidaddeCob.text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EestadoCob.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      if empty(DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString) Then
         DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString := DM_tabelas.ZQAchaCidade.FieldByName('cepgeral').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('cidade_cob').AsLargeInt := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Participante.ECidNaturalidadeExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidNaturalidade.text) Then
      ECidNaturalidade.SetFocus
    else Begin
      ECidNaturalidade.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      Eestadonatural.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZqParticipante.FieldByName('naturalidade').AsInteger := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
//    ECidNaturalidade.SetFocus;
  end;

end;

procedure TFrm_Participante.DBEDoc1Exit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [dsInsert, dsEdit] then
  begin
    // 🔹 Validação do documento
    if not Verif_doc(DBEDoc1.Text, True) then
    begin
      DBEDoc1.SetFocus;
      Exit;
    end;

    // 🔹 Consulta no banco
    with ZqParticipantes_doc do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT idpaticipante');
      SQL.Add('FROM participante');
      SQL.Add('WHERE doc1 = :doc');

      ParamByName('doc').AsString := Trim(DBEDoc1.Text);

      Open;

      // 🔹 Se encontrou
      if not IsEmpty then
      begin
        // 🔹 Evita validar contra o próprio registro
        if FieldByName('idpaticipante').AsLargeInt <> DM_Tabelas.ZqParticipante.FieldByName('idpaticipante').AsLargeInt then
        begin
          if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'F' then
            ShowMessage('CPF já cadastrado!')
          else
            ShowMessage('CNPJ já cadastrado!');

          DBEDoc1.SetFocus;
          close;
          Exit;
        end;
      end;
      close;
    end;
  end;
end;


procedure TFrm_Participante.DBEDoc1Enter(Sender: TObject);
var
  vartex : string;
begin
  IF (dM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit])  Then Begin // and (empty(DBEDoc1.Text))
    vartex := SemMascara(DM_Tabelas.ZqParticipante.FieldByName('doc1').AsString);
    DBEDoc1.EditMask := '';
    DBEDoc1.Text := vartex ;
    If DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'F' Then
      DBEDoc1.EditMask := '999.999.999-99;1;_'
    else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'J' Then
      DBEDoc1.EditMask := '99.999.999/9999-99;1;_'
    else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'R' Then
      DBEDoc1.EditMask := '999.999.999-99;1;_'
    else if DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'P' Then
      DBEDoc1.EditMask := '99.999.999/9999-99;1;_'
    else If DM_Tabelas.ZqParticipante.FieldByName('tipopessoa').AsString = 'I' Then
      DBEDoc1.EditMask := '999.999.999-99;1;_';

  end;
end;

procedure TFrm_Participante.DBEDoc2RespEnter(Sender: TObject);
var
  vartex : String;
begin
  IF (dM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit])  Then Begin // and (empty(DBEDoc1.Text))
    vartex := SemMascara(DM_Tabelas.ZQresponsavel.FieldByName('documen2').AsString);
    DBEDoc2Resp.EditMask := '';
    DBEDoc2Resp.Text := vartex ;
//    If DM_Tabelas.ZQresponsaveldocumen2.Value = 'F' Then
      DBEDoc2Resp.EditMask := '999.999.999-99;1;_'
//    else if DM_Tabelas.ZqParticipantetipopessoa.Value = 'J' Then
//      DBEDoc2Resp.EditMask := '99.999.999/9999-99;1;_'
//    else if DM_Tabelas.ZqParticipantetipopessoa.Value = 'R' Then
//      DBEDoc2Resp.EditMask := '999.999.999-99;1;_'
//    else if DM_Tabelas.ZqParticipantetipopessoa.Value = 'P' Then
//      DBEDoc2Resp.EditMask := '99.999.999/9999-99;1;_';
  End;
end;

procedure TFrm_Participante.DBEDoc2RespExit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] then
    if not Verif_doc(DBEDoc2Resp.text,True) Then
      DBEDoc2Resp.SetFocus;
end;

procedure TFrm_Participante.ERepNaturExit(Sender: TObject);
begin
  if (DM_Tabelas.CDSResponsavelTemp.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ERepNatur.text) Then
      ERepNatur.SetFocus
    else Begin
      ERepNatur.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      ERespEst.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.CDSResponsavelTempcid_natural.Value := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Participante.CBAssinaExit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert,DSEdit] Then Begin
    if CBAssina.Checked Then
      DM_Tabelas.ZQConjuge.FieldByName('assina').AsString := 'S'
    else
      DM_Tabelas.ZQConjuge.FieldByName('assina').AsString := '';
    if DM_Tabelas.ZqParticipante.State in [DsInsert] Then Begin
      Pagina.ActivePageIndex := 4;
      DBEEnde.SetFocus;
    end;
  end;
end;

procedure TFrm_Participante.DBCBRegimeExit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] Then BEgin
    if DBCBRegime.ItemIndex <0 Then Begin
      DBCBRegime.SetFocus;
      exit;
    end;
    if DBCBRegime.ItemIndex = 2 Then Begin
      GBPacto.Visible := True;
      DBEDataCasou.SetFocus;
    end
    else Begin
      GBPacto.Visible := False;
      pagina.ActivePageIndex := 2;
      DBENomeResp.SetFocus;
    end;
  end;
end;

procedure TFrm_Participante.DBEEstCivilExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (DBEEstCivil.ItemIndex <0) Then Begin
    DBEEstCivil.SetFocus;
    exit;
  end;
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) Then
    if (DBEEstCivil.ItemIndex =1) Then
    begin
      DBCBRegime.Visible := True;
      Label27.visible:=true;
      DBCBRegime.enabled:=true;
      DBCBRegime.readonly:=false;

    end
    else
    begin
      DBCBRegime.Visible := False;
      Label27.visible:=false;
    end;  
  if (DM_Tabelas.ZqParticipante.State in [DsInsert]) Then Begin
    if DBEEstCivil.ItemIndex <= 1 Then
    Begin
       if DM_Tabelas.Zqconjuge.State in [DsBrowse] then
          DM_Tabelas.Zqconjuge.Insert;
      Pagina.Pages[2].TabVisible := True;
      Pagina.Pages[3].TabVisible :=False ;
      DBCBRegime.enabled:=true;
      if (DBEEstCivil.ItemIndex = 1) or (DBEEstCivil.ItemIndex = 2) Then
        DBCBRegime.SEtFocus
      else Begin
        pagina.ActivePageIndex := 2;
        DBENomeResp.SetFocus;
      end;
    end
    else Begin
      Pagina.Pages[2].TabVisible := False;
      Pagina.Pages[3].TabVisible := True;
      DBCBRegime.Enabled := False;
      pagina.ActivePageIndex := 3;
      DBEResp.SetFocus;
    End;
  end;
end;

procedure TFrm_Participante.DBECepExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (Empty(DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString)) Then Begin
    DM_Tabelas.ZqParticipante.FieldByName('cep_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('cep').AsString;
  end;

end;

procedure TFrm_Participante.ECidcasouExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidcasou.text) Then
      ECidcasou.SetFocus
    else Begin
      ECidcasou.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstCart.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQConjuge.FieldByName('cidade_cart').AsFloat := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;

end;

procedure TFrm_Participante.ECidComarcaExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidComarca.text) Then
      ECidComarca.SetFocus
    else Begin
      ECidComarca.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      EEstComarca.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQConjuge.FieldByName('comarca').AsFloat := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;

end;

procedure TFrm_Participante.GBPactoExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert,DsEdit]) Then Begin
    if DBEEstCivil.ItemIndex<=1 Then Begin
      pagina.ActivePageIndex := 2;
      DBENomeResp.SetFocus;
    end
    else Begin
      pagina.ActivePageIndex := 3;
      DBEResp.SetFocus;
    end;
  end;
end;

procedure TFrm_Participante.DBMemo1Exit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert]) and (DBCBPessoa.ItemIndex <> 0) Then Begin
    pagina.ActivePageIndex := 3;
    DBEResp.SetFocus;
  End;
end;

procedure TFrm_Participante.DBMemo2Enter(Sender: TObject);
begin
  Frm_Participante.KeyPreview:=false;
end;

procedure TFrm_Participante.DBMemo2Exit(Sender: TObject);
begin
  Frm_Participante.KeyPreview:=true;
end;

procedure TFrm_Participante.DBEComplExit(Sender: TObject);
begin
  if (DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit]) and (Empty(DM_Tabelas.ZqParticipante.FieldByName('complemento_cob').AsString)) Then Begin
    DM_Tabelas.ZqParticipante.FieldByName('complemento_cob').AsString := DM_Tabelas.ZqParticipante.FieldByName('complemento').AsString;
  end;
end;

procedure TFrm_Participante.DBERespExit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] Then Begin
    if empty(DBEResp.Text) Then Begin
      DM_Tabelas.CDSResponsavelTemp.Cancel;
      if DM_Tabelas.ZqParticipante.State in [DsInsert] Then Begin
        Pagina.ActivePageIndex := 4;
        DBEEnde.SetFocus;
      end
      else
        BtGravar.SetFocus;
    end;
  end;
end;

procedure TFrm_Participante.CBReprAssExit(Sender: TObject);
begin
  if DM_Tabelas.CDSResponsavelTemp.State in [DsInsert,DSEdit] Then Begin
    if CBReprAss.Checked Then
      DM_Tabelas.CDSResponsavelTempassina.Value := 'S'
    else
      DM_Tabelas.CDSResponsavelTempassina.Value := '';
  end;
end;

procedure TFrm_Participante.DBERepDoc2Exit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] then
    if not Verif_doc(DBERepDoc2.text,True) Then
      DBERepDoc2.SetFocus;
end;

procedure TFrm_Participante.DBERepDoc2Enter(Sender: TObject);
var
  vartex : String;
begin
  IF (dM_Tabelas.CDSResponsavelTemp.State in [DsInsert, DsEdit])  Then Begin // and (empty(DBEDoc1.Text))
    vartex := SemMascara(DM_Tabelas.CDSResponsavelTempdocumen2.Value);
    DBERepDoc2.EditMask := '';
    DBERepDoc2.Text := vartex ;
    DBERepDoc2.EditMask := '999.999.999-99;1;_'
  End;
end;

procedure TFrm_Participante.ECidNatConjExit(Sender: TObject);
begin
  if (DM_Tabelas.ZQConjuge.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ECidNatConj.text) Then
      ECidNatConj.SetFocus
    else Begin
      ECidNatConj.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      Eestconj.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.ZQConjuge.FieldByName('natural').AsInteger := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;

end;

procedure TFrm_Participante.atualiza_tela_responsavel;
Begin
  ERepNatur.Text       := DM_Tabelas.CDSResponsavelTempnomecidade.AsString;
  ERespEst.Text        := DM_Tabelas.CDSResponsavelTempestcidade.AsString;
  ERepConjNatCid.Text  := DM_Tabelas.CDSResponsavelTempconjugenaturcidnome.AsString;
  ERepConjNatUF.Text   := DM_Tabelas.CDSResponsavelTempconjugenaturcidUF.AsString;
  ERepConjCasCid.Text  := DM_Tabelas.CDSResponsavelTempconjuge_cid_cart_nome.AsString;
  ERepConjCasUF.Text   := DM_Tabelas.CDSResponsavelTempconjuge_cid_cart_UF.AsString;
  CBRepConjAss.Checked := False;
  CBReprAss.Checked    := False;
  if DM_Tabelas.CDSResponsavelTempassina.Value = 'S' Then
    CBReprAss.Checked  := True;
  if DM_Tabelas.CDSResponsavelTempconjugeass.Value = 'S' Then
  CBRepConjAss.Checked := True;
end;
procedure TFrm_Participante.DBGRespKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  atualiza_tela_responsavel;
end;

procedure TFrm_Participante.DBGRespMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
atualiza_tela_responsavel;
end;

procedure TFrm_Participante.PagRespExit(Sender: TObject);
begin
  DM_Tabelas.CDSResponsavelTemp.Insert;
  DBEResp.SetFocus;
end;

procedure TFrm_Participante.CBRepConjAssExit(Sender: TObject);
begin
  if DM_Tabelas.CDSResponsavelTemp.State in [DsInsert,DSEdit] Then Begin
    if CBRepConjAss.Checked Then
      DM_Tabelas.CDSResponsavelTempconjugeass.Value := 'S'
    else
      DM_Tabelas.CDSResponsavelTempconjugeass.Value := '';
  end;
end;

procedure TFrm_Participante.ERepConjNatCidExit(Sender: TObject);
begin
  if (DM_Tabelas.CDSResponsavelTemp.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ERepConjNatCid.text) Then
      ERepConjNatCid.SetFocus
    else Begin
      ERepConjNatCid.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      ERepConjNatUF.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.CDSResponsavelTempconjugenatur.Value := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Participante.ERepConjCasCidExit(Sender: TObject);
begin
  if (DM_Tabelas.CDSResponsavelTemp.State in [DsInsert, DsEdit]) and not(BtCancelar.Focused or BtFechar.Focused) Then Begin
    if not Acha_Cidade(ERepConjCasCid.text) Then
      ERepConjCasCid.SetFocus
    else Begin
      ERepConjCasCid.Text := DM_Tabelas.ZQAchaCidade.FieldByName('nomecid').AsString;
      ERepConjCasUF.Text := DM_Tabelas.ZQAchaCidade.FieldByName('estado').AsString;
      DM_Tabelas.CDSResponsavelTempconjuge_cidcartorio.Value := DM_Tabelas.ZQAchaCidade.FieldByName('idcidade').AsLargeInt;
    end;
  end;
end;

procedure TFrm_Participante.JrelItems0Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Etiquetas') Then Exit;
  Self.FormStyle:=fsNormal;
  if frmMain=nil then
     frmMain:=TfrmMain.create(Application);
  frmMain.showmodal;
  frmMain:=nil;
  Self.FormStyle:=fsStayOnTop;
end;

procedure TFrm_Participante.DBEEndeEnter(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.State in [DsInsert, DsEdit] then
  begin
    if empty(DM_tabelas.ZqParticipante.FieldByName('endereco').AsString) then
    begin
      FrmPesqEndereco.Top := Self.Top+100;
      FrmPesqEndereco.Left := Self.Left;
      FrmPesqEndereco.showmodal;
      DM_tabelas.ZqParticipante.FieldByName('endereco').AsString:=DM_tabelas.localizacao.FieldByName('log_tipo_logradouro').AsString+' '+DM_tabelas.localizacao.FieldByName('log_no').AsString;
      DM_tabelas.ZqParticipante.FieldByName('bairro').AsString:=DM_tabelas.localizacao.FieldByName('bai_no').AsString;
      ECidEnder.Text:=DM_tabelas.localizacao.FieldByName('loc_no').AsString;
      EEstado.Text:=DM_tabelas.localizacao.FieldByName('ufe_sg').AsString;
      DM_tabelas.ZqParticipante.FieldByName('cep').AsString:=copy(DM_tabelas.localizacao.FieldByName('cep').AsString,1,2)+'.'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,3,3)+'-'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,6,3);
      DBEEnde.SetFocus;
      DM_tabelas.localizacao.Close;
    end
    else
    begin
      if simnao('Deseja Alterar o Endereço?','SIM') then
      begin
        FrmPesqEndereco.Top := Self.Top+100;
        FrmPesqEndereco.Left := Self.Left;
        FrmPesqEndereco.showmodal;
        DM_tabelas.ZqParticipante.FieldByName('endereco').AsString:=DM_tabelas.localizacao.FieldByName('log_tipo_logradouro').AsString+' '+DM_tabelas.localizacao.FieldByName('log_no').AsString;
        DM_tabelas.ZqParticipante.FieldByName('bairro').AsString:=DM_tabelas.localizacao.FieldByName('bai_no').AsString;
        ECidEnder.Text:=DM_tabelas.localizacao.FieldByName('loc_no').AsString;
        EEstado.Text:=DM_tabelas.localizacao.FieldByName('ufe_sg').AsString;
        DM_tabelas.ZqParticipante.FieldByName('cep').AsString:=copy(DM_tabelas.localizacao.FieldByName('cep').AsString,1,2)+'.'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,3,3)+'-'+copy(DM_tabelas.localizacao.FieldByName('cep').AsString,6,3);
        DBEEnde.SetFocus;
        DM_tabelas.localizacao.Close;
      end;
    end;
  end;
end;

procedure TFrm_Participante.btnimprimirClick(Sender: TObject);
begin
  Self.FormStyle:=fsNormal;
  if FrmRelparticipante=nil then
     FrmRelparticipante:=TFrmRelparticipante.Create(Application);
  FrmRelparticipante.RLReport1.PreviewModal;
  FrmRelparticipante:=nil;
  Self.FormStyle:=fsStayOnTop;
end;

procedure TFrm_Participante.DBComboBox2Exit(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString='SIM' then
  begin
     cc.ReadOnly:=false;
     cex.ReadOnly:=false;
     jb.ReadOnly:=false;
     dbquadra.ReadOnly:=false;
     dblote.ReadOnly:=false;
  end
  else
  begin
     DM_Tabelas.ZqParticipante.FieldByName('cc').Clear;
     DM_Tabelas.ZqParticipante.FieldByName('cex').Clear;
     DM_Tabelas.ZqParticipante.FieldByName('jb').Clear;
     DM_Tabelas.ZqParticipante.FieldByName('xquadra').AsString:='';
     DM_Tabelas.ZqParticipante.FieldByName('xlote').AsString:='';
     cc.Clear;
     cc.ReadOnly:=true;
     cex.Clear;
     cex.ReadOnly:=true;
     jb.Clear;
     jb.ReadOnly:=true;
     dbquadra.Clear;
     dbquadra.ReadOnly:=true;
     dblote.Clear;
     dblote.ReadOnly:=true;
     BtGravar.SetFocus;
  end;
end;

procedure TFrm_Participante.ccEnter(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString='SIM' then
  begin
     cc.ReadOnly:=false;
  end
  else
  begin
     cc.ReadOnly:=true;
  end;

end;

procedure TFrm_Participante.cexEnter(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString='SIM' then
  begin
     cex.ReadOnly:=false;
  end
  else
  begin
     cex.ReadOnly:=true;
  end;

end;

procedure TFrm_Participante.jbEnter(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString='SIM' then
  begin
     jb.ReadOnly:=false;
  end
  else
  begin
     jb.ReadOnly:=true;
  end;

end;

procedure TFrm_Participante.dbquadraEnter(Sender: TObject);
begin
  if DM_Tabelas.ZqParticipante.FieldByName('judicial').AsString='SIM' then
  begin
     dbquadra.ReadOnly:=false;
     dblote.ReadOnly:=false;
  end
  else
  begin
     dbquadra.ReadOnly:=true;
     dblote.ReadOnly:=true;
  end;

end;

procedure TFrm_Participante.dxButton1Click(Sender: TObject);
begin
  Self.FormStyle:=fsNormal;
  if frmgoogle=nil then
     frmgoogle:=Tfrmgoogle.Create(Application);
  frmgoogle.Cidade.Text:=ECidEnder.Text;
  frmgoogle.Estado.Text:=EEstado.Text;
  frmgoogle.Rua.Text:=DBEEnde.Text;
  frmgoogle.Cep.Text:=charrem(DBECep.Text);
  frmgoogle.Button1click(Nil);
  frmgoogle.ShowModal;
  frmgoogle:=nil;
  frmgoogle.Free;
  Self.FormStyle:=fsStayOnTop;
end;

procedure TFrm_Participante.JrelItems1Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Etiquetas') Then Exit;
  Self.FormStyle:=fsNormal;
  if frmMain2=nil then
     frmMain2:=TfrmMain2.create(Application);
  frmMain2.showModal;
  frmMain2:=nil;
  Self.FormStyle:=fsStayOnTop;
end;

procedure TFrm_Participante.JrelItems2Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Ficha') Then Exit;
  ZqParticipantes.close;
  ZqParticipantes.SQL.clear;
//  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
//  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
//  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico,idcidade,nomecid,estado');
//  ZqParticipantes.SQL.Add(' from participante as p join cidade as c on c.idcidade=p.cidade order by nome_parte');
  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  ZqParticipantes.SQL.Add(' from participante order by nome_parte');
  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  ZqParticipantes.open;
  try
    Self.FormStyle:=fsNormal;
    if FrmRelparticipante1=nil then
       FrmRelparticipante1:=TFrmRelparticipante1.create(Application);
    FrmRelparticipante1.RLReport1.Previewmodal;
    FreeAndNil(FrmRelparticipante1);
  finally
  Self.FormStyle:=fsStayOnTop;
  ZqParticipantes.close;
  ZqParticipantes.SQL.clear;
//  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
//  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
//  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico,idcidade,nomecid,estado');
//  ZqParticipantes.SQL.Add(' from participante as p join cidade as c on c.idcidade=p.cidade  ');

  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  ZqParticipantes.SQL.Add(' from participante');
  ZqParticipantes.open;

  end;
end;

procedure TFrm_Participante.JrelItems3Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Ficha por Empreendimento') Then Exit;

  try
    Self.FormStyle:=fsNormal;
    if FrmRelparticipante_empre=nil then
       FrmRelparticipante_empre:=TFrmRelparticipante_empre.create(Application);
    FrmRelparticipante_empre.RLReport1.Previewmodal;
    FreeAndNil(FrmRelparticipante_empre);
  finally
    Self.FormStyle:=fsStayOnTop;


  end;
end;

procedure TFrm_Participante.JrelItems4Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Ficha Completa') Then Exit;
  ZqParticipantes.close;
  ZqParticipantes.SQL.clear;
//  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
//  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
//  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico,idcidade,nomecid,estado');
//  ZqParticipantes.SQL.Add(' from participante as p join cidade as c on c.idcidade=p.cidade order by nome_parte');

  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  ZqParticipantes.SQL.Add(' from participante order by nome_parte');
  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  ZqParticipantes.open;
  try
    Self.FormStyle:=fsNormal;
    if FrmRelparticipante_composta=nil then
       FrmRelparticipante_composta:=TFrmRelparticipante_composta.create(Application);
    FrmRelparticipante_composta.RLReport1.Previewmodal;
    FreeAndNil(FrmRelparticipante_composta);
  finally
  Self.FormStyle:=fsStayOnTop;
  ZqParticipantes.close;
  ZqParticipantes.SQL.clear;
//  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
//  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
//  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico,idcidade,nomecid,estado');
//  ZqParticipantes.SQL.Add(' from participante as p join cidade as c on c.idcidade=p.cidade  ');
  ZqParticipantes.SQL.Add('Select  idpaticipante,nome_parte,doc1,doc2,endereco,bairro,cidade,cep,ende_cob,bairro_cob,cidade_cob,cep_cob,Fone1,fone2,fone3,tipopessoa,aniversario,nacionalidade,');
  ZqParticipantes.SQL.Add('        email,naturalidade,localdetrab,profissao,ESTADOCIVIL,renda,observacao,codpaginc,codrecinc,codpagBx,codrecBx,fone4,complemento,complemento_cob,cadastrado,Nome_Firma,');
  ZqParticipantes.SQL.Add('        Inativo,sexo,cc,cex,jb,judicial,xquadra,xlote,plantao,historico');
  ZqParticipantes.SQL.Add(' from participante');
  ZqParticipantes.open;
  end;
end;

procedure TFrm_Participante.JrelItems5Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Ficha por Empreendimento') Then Exit;

  try
    if DM_Tabelas.ZQCidade.Active then
      DM_Tabelas.ZQCidade.Close;
    DM_Tabelas.ZQCidade.Open;
    Self.FormStyle:=fsNormal;
    if FrmRelparticipante_empre_completa=nil then
       FrmRelparticipante_empre_completa:=TFrmRelparticipante_empre_completa.create(Application);
    FrmRelparticipante_empre_completa.RLReport1.Previewmodal;
    FreeAndNil(FrmRelparticipante_empre_completa);
  finally
    Self.FormStyle:=fsStayOnTop;


  end;
end;

procedure TFrm_Participante.JrelItems6Click(Sender: TObject);
begin
  jrel.Visible:=false;
  if not Verif_senha('Participante','Relatórios','Fichário Completo') Then Exit;
  ZqParticipantes.Close;
  ZqParticipantes.SQL.Clear;
  ZqParticipantes.SQL.Add('Select * from participante order by nome_parte');

  if DM_Tabelas.ZQCidade.Active then
    DM_Tabelas.ZQCidade.Close;
  DM_Tabelas.ZQCidade.Open;
  ZqParticipantes.Open;

  try
    Self.FormStyle := fsNormal;
    try
      if FrmRelparticipante_ficha = nil then
        FrmRelparticipante_ficha := TFrmRelparticipante_ficha.Create(Application);
      FrmRelparticipante_ficha.RLReport1.PreviewModal;
    finally
      FreeAndNil(FrmRelparticipante_ficha);
    end;
  finally
  Self.FormStyle := fsStayOnTop;
    ZqParticipantes.Close;
    ZqParticipantes.SQL.Clear;
    ZqParticipantes.SQL.Add('Select * from participante');
    ZqParticipantes.Open;
  end;
end;

procedure TFrm_Participante.Timer1Timer(Sender: TObject);
begin
  if Pagina.TabIndex=7 then
  begin
    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('DROP TABLE IF EXISTS part');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('CREATE TABLE part ( `id` int(11) NOT NULL AUTO_INCREMENT, `observ_estorno` varchar(60000), `idpaticipante` int(11) NOT NULL,');
    ZQPart.SQL.Add(' PRIMARY KEY (`id`) ) ENGINE=MEMORY');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('INSERT INTO part (observ_estorno, idpaticipante) SELECT r.observ_estorno,p.idpaticipante');
    ZQPart.SQL.Add('FROM participante as p join recebimento as r on p.idpaticipante=r.adversa where p.idpaticipante='+quotedstr(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text)+' and r.observ_estorno is not null');
    ZQPart.ExecSQL;

    ZQPart.Close;
    ZQPart.SQL.Clear;
    ZQPart.SQL.Add('select * from part');
    ZQPart.open;

    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('DROP TABLE IF EXISTS part2');
    ZQPart2.ExecSQL;


    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('CREATE TABLE part2 ( `id` int(11) NOT NULL AUTO_INCREMENT, `observ_estorno` varchar(60000), `idpaticipante` int(11) NOT NULL,');
    ZQPart2.SQL.Add(' PRIMARY KEY (`id`) ) ENGINE=MEMORY');
    ZQPart2.ExecSQL;

    ZQPart2.Close;
    ZQPart2.SQL.Clear;
    ZQPart2.SQL.Add('select * from part2');
    ZQPart2.open;

    ZQPart.First;
    if ZQPart.RecordCount>0 then
    begin
      ZQPart2.Insert;
      ZQPart2.FieldByName('idpaticipante').AsInteger:=ZQpart.FieldByName('idpaticipante').AsInteger;
      ZQPart2.post;
      ZQPart.DisableControls;
      while not ZQPart.Eof do
      begin
        ZQPart2.Edit;
        ZQPart2.FieldByName('observ_estorno').AsString:=ZQPart2.FieldByName('observ_estorno').AsString+#10+#13+ZQpart.FieldByName('observ_estorno').AsString;
        ZQPart2.post;
        ZQPart.Next;
      end;
      ZQPart.EnableControls;
    end;
  end;

end;

procedure TFrm_Participante.PaginaChange(Sender: TObject);
begin
  if Pagina.TabIndex=7 then
     Timer1.Enabled:=true
  else
    Timer1.Enabled:=false;

  if Pagina.ActivePage=TabSheet3 then
  begin
    ZQContratos.close;
    ZQContratos.SQL.Clear;
    ZQContratos.SQL.Add('SELECT * FROM participante as p join comprador as c on p.idpaticipante=c.paticipante_idpaticipante join venda as v on c.venda_idvenda=v.idvenda join imovel as i on v.imovel=i.idimovel   ');
    ZQContratos.SQL.Add(' join loteamento as l on l.idloteamento=i.loteamento_idloteamento where p.idpaticipante='+quotedstr(DM_Tabelas.ZQParticipante.FieldByName('idpaticipante').Text)+' order by p.idpaticipante');
    ZQContratos.Open;
  end;
end;


procedure TFrm_Participante.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrm_Participante, 'ZqParticipantes', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomecidade', 'nomecidade', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomeestado', 'nomeestado', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomecidadecob', 'nomecidadecob', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomeestadocob', 'nomeestadocob', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cob', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '00\.000\-999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '\(99\)9999\-9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantestipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipanteslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesESTADOCIVIL', 'ESTADOCIVIL', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '\(00\)0000\-0000;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesInativo', 'Inativo', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantessexo', 'sexo', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescc', 'cc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantescex', 'cex', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesjb', 'jb', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesjudicial', 'judicial', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesxquadra', 'xquadra', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesxlote', 'xlote', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomeestadonatural', 'nomeestadonatural', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True, True);
  RegisterRuntimeField(TFrm_Participante, 'ZqParticipantes', 'ZqParticipantesnomecidadenatural', 'nomecidadenatural', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'naturalidade', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True, True);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQContratos', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosidpaticipante', 'idpaticipante', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosnome_parte', 'nome_parte', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdoc1', 'doc1', TWideStringField, fkData, 18, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdoc2', 'doc2', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosendereco', 'endereco', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosbairro', 'bairro', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscidade', 'cidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscep', 'cep', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosende_cob', 'ende_cob', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosbairro_cob', 'bairro_cob', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscidade_cob', 'cidade_cob', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscep_cob', 'cep_cob', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosFone1', 'Fone1', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosfone2', 'fone2', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosfone3', 'fone3', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostipopessoa', 'tipopessoa', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosaniversario', 'aniversario', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosnacionalidade', 'nacionalidade', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosemail', 'email', TWideStringField, fkData, 150, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosnaturalidade', 'naturalidade', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoslocaldetrab', 'localdetrab', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosprofissao', 'profissao', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosestadocivil', 'estadocivil', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosrenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosobservacao', 'observacao', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscodrecinc', 'codrecinc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscodrecBx', 'codrecBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscodpaginc', 'codpaginc', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscodpagBx', 'codpagBx', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosfone4', 'fone4', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscomplemento', 'complemento', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscomplemento_cob', 'complemento_cob', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscadastrado', 'cadastrado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosNome_Firma', 'Nome_Firma', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosidcomprador', 'idcomprador', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratospaticipante_idpaticipante', 'paticipante_idpaticipante', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosvenda_idvenda', 'venda_idvenda', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratospromissario', 'promissario', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratospercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosvalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostabela_Price', 'tabela_Price', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosEscriturado', 'Escriturado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosidimovel', 'idimovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosloteamento_idloteamento', 'loteamento_idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosquadra', 'quadra', TWideStringField, fkData, 25, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoslote', 'lote', TWideStringField, fkData, 4, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoslado', 'lado', TWideStringField, fkData, 5, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmedidafrente', 'medidafrente', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmedidafundo', 'medidafundo', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmedidaesquerda', 'medidaesquerda', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmedidadedireita', 'medidadedireita', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosarea', 'area', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosconfrontofrente', 'confrontofrente', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosconfrontofundo', 'confrontofundo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosconfrontodireita', 'confrontodireita', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosconfrontoesquerda', 'confrontoesquerda', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosesquinanome', 'esquinanome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosesquinalinha', 'esquinalinha', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosesquinacurva', 'esquinacurva', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosvalorvenal', 'valorvenal', TFloatField, fkData, 0, 0, False, '', '###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosObervacaoloteamento', 'Obervacaoloteamento', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscornomapa', 'cornomapa', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdisponivel', 'disponivel', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmatri', 'matri', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmatricula', 'matricula', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosidloteamento', 'idloteamento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscidade_idcidade', 'cidade_idcidade', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmatriculaloteamento', 'matriculaloteamento', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdatacadastro', 'datacadastro', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdatainalguracao', 'datainalguracao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdataconclusao', 'dataconclusao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoslogotipo', 'logotipo', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosmapa', 'mapa', TBlobField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostestemunha1', 'testemunha1', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostestemunha2', 'testemunha2', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostestemunha3', 'testemunha3', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostestemunha4', 'testemunha4', TWideStringField, fkData, 120, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosObservincorp', 'Observincorp', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosbairro_1', 'bairro_1', TWideStringField, fkData, 40, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratospasta_mapa', 'pasta_mapa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosnomeloteamento', 'nomeloteamento', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratostipo', 'tipo', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratoscodcontabancaria', 'codcontabancaria', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosapelido', 'apelido', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQContratos', 'ZQContratosdigito_dif', 'digito_dif', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQBxEntrada', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradajr', 'jr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradade', 'de', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradadescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxEntrada', 'ZQBxEntradadata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQBxParcela', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelavr', 'vr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelajr', 'jr', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelade', 'de', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelaidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelarefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelaidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelavalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParceladescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelajuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParcelapercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQBxParcela', 'ZQBxParceladata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQpart', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQpart', 'ZQpartid', 'id', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQpart', 'ZQpartobserv_estorno', 'observ_estorno', TWideStringField, fkData, 60000, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQpart', 'ZQpartidpaticipante', 'idpaticipante', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQPart2', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQPart2', 'ZQPart2id', 'id', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQPart2', 'ZQPart2observ_estorno', 'observ_estorno', TWideStringField, fkData, 60000, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQPart2', 'ZQPart2idpaticipante', 'idpaticipante', TIntegerField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrm_Participante, 'ZQConjuge', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeidconjuge', 'idconjuge', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugenome', 'nome', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugedoc1', 'doc1', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugedoc2', 'doc2', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugenasc', 'nasc', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugenatural', 'natural', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeemail', 'email', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeassina', 'assina', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeempresa', 'empresa', TWideStringField, fkData, 100, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecargo', 'cargo', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeadmissa', 'admissa', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugerenda', 'renda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugenomecida', 'nomecida', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'natural', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeestcida', 'estcida', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'natural', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeidparti', 'idparti', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugeregime', 'regime', TWideStringField, fkData, 30, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecomarca', 'comarca', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecidade_cart', 'cidade_cart', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugedata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugefolha', 'folha', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugelivro', 'livro', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecomar_cidnome', 'comar_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'comarca', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecomar_cidest', 'comar_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'comarca', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecart_cidnome', 'cart_cidnome', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'cidade_cart', 'DM_Tabelas.ZQCidade', 'idcidade', 'nomecid', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecart_cidest', 'cart_cidest', TWideStringField, fkLookup, 2, 0, False, '', '', '', '', 0, 'cidade_cart', 'DM_Tabelas.ZQCidade', 'idcidade', 'estado', True);
  RegisterRuntimeField(TFrm_Participante, 'ZQConjuge', 'ZQConjugecartorio', 'cartorio', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);

End.
