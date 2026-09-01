

unit RecebBaixa;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,  dxButton, wwdbdatetimepicker, StdCtrls, Mask, DBCtrls,
  XBanner, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Grids,
  DBGrids, ImgList, XNum, DBClient, XDBDate, XDBNum, ComCtrls,
  JvExComCtrls, JvDateTimePicker, XEdit, XDBEdit, ExtCtrls, XDate, dxCore2,
  System.ImageList,threading;
type
  TFrmRecebBaixa = class(TForm)
    Label10: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label16: TLabel;
    DBEBaixaDocum: TDBEdit;
    DBEBaixaVrRec: TDBEdit;
    DXBBaixaGravar: TdxButton;
    DS_RecBai: TDataSource;
    ZQRecBai: TZQuery;
    DBGBaixando: TDBGrid;
    DXBFechar: TdxButton;
    ImageList1: TImageList;
    Label2: TLabel;
    Label3: TLabel;
    XNEAberto: TXNumEdit;
    XNEVencido: TXNumEdit;
    XNEMarcado: TXNumEdit;
    Label4: TLabel;
    DS_Parcelas: TDataSource;
    CDSParcelas: TClientDataSet;
    CDSParcelasTipDoc: TWideStringField;
    CDSParcelasDocum: TWideStringField;
    CDSParcelasOrdem: TWideStringField;
    CDSParcelasVenci: TDateField;
    CDSParcelasVrParc: TFloatField;
    CDSParcelascliente: TIntegerField;
    CDSParcelasusuario: TIntegerField;
    CDSParcelasentrada: TDateField;
    CDSParcelasbanco: TWideStringField;
    CDSParcelasagencia: TWideStringField;
    CDSParcelasdono: TWideStringField;
    CDSParcelascpf_cnpj: TWideStringField;
    DBGrid2: TDBGrid;
    XNEParcelas: TXNumEdit;
    Label5: TLabel;
    DBCBTIPDOC: TDBComboBox;
    Label6: TLabel;
    Label7: TLabel;
    XNEDesconto: TXNumEdit;
    XNEJuros: TXNumEdit;
    XNERecebido: TXNumEdit;
    Label8: TLabel;
    Label15: TLabel;
    Label9: TLabel;
    Label17: TLabel;
    Label19: TLabel;
    XNEResto: TXNumEdit;
    EMObsbx: TMemo;
    Label20: TLabel;
    EMNovostit: TMemo;
    Label21: TLabel;
    CDSParcelascodcontabil: TIntegerField;
    Label22: TLabel;
    EContabil: TEdit;
    CDSParcelasdesccontabil: TWideStringField;
    Label11: TLabel;
    GBCheque: TGroupBox;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    DBEAgencia: TDBEdit;
    DBEBanco: TDBEdit;
    DBEDono: TDBEdit;
    DBECpf_Cnpj: TDBEdit;
    CDSParcelaspercent: TFloatField;
    XDBNumEdit1: TXDBNumEdit;
    DBText1: TDBText;
    CDSParcelasconta: TWideStringField;
    Label28: TLabel;
    DEBConta: TDBEdit;
    GBContra: TGroupBox;
    Enomes: TEdit;
    Ecodnome: TEdit;
    RBParticipante: TRadioButton;
    RBCheque: TRadioButton;
    Label1: TLabel;
    Label23: TLabel;
    CDSParcelascodbanco: TIntegerField;
    GBBanco: TGroupBox;
    DBGBanco: TDBGrid;
    Label29: TLabel;
    EPlanocontabanco: TEdit;
    CDSParcelascontabanco: TIntegerField;
    CDSParcelassomapag: TWideStringField;
    RBPonteiro: TRadioButton;
    DBEidreg: TDBEdit;
    LREg: TLabel;
    CDSParcelasjuros: TFloatField;
    CDSParcelasdesconto: TFloatField;
    CBSaldoCorrigido: TCheckBox;
    DS_RecebBxTemp: TDataSource;
    ZQRecebBxTemp: TZQuery;

    XNJuros: TXNumEdit;
    Label30: TLabel;
    Label31: TLabel;
    Label33: TLabel;
    DS_dividdo: TDataSource;
    CDS_dividido: TClientDataSet;
    DS_ratear: TDataSource;
    ZQRatear: TZQuery;
    CDS_divididoparticipante: TIntegerField;
    CDS_divididovr_base: TFloatField;
    CDS_divididovr_calc: TFloatField;
    CDS_divididocontabil: TIntegerField;
    CDS_divididotipodoc: TWideStringField;
    CDS_divididoobs: TWideMemoField;
    DS_ItensRecibo: TDataSource;
    ZQItensRecibo: TZQuery;
    CDS_divididonomeparti: TWideStringField;
    CDS_divididodescrcontabil: TWideStringField;
    CDS_divididodescrcontabil2: TWideStringField;
    Memo1: TMemo;
    DS_Div_usado: TDataSource;
    ZQDiv_usado: TZQuery;


























































    CBPeriodoEnt: TComboBox;
    XBanner9: TXBanner;
    XBanner1: TXBanner;
    XBanner2: TXBanner;
    XBanner3: TXBanner;












    GBAdversa: TGroupBox;
    XBanner5: TXBanner;
    EcodAdversa: TEdit;
    GroupBox1: TGroupBox;
    XBanner4: TXBanner;
    EcodAdversaNova: TEdit;
    EAdversaNova: TEdit;
    XEEmpresend: TXEdit;
    CBTipobaixa: TComboBox;





    CDSParcelasnomedoempreendimento: TWideStringField;
    CDSParcelasvenda_idvenda: TIntegerField;
    CDSParcelasquadralote: TWideStringField;
    Label18: TLabel;
    DXBPesquisar: TdxButton;


    CDSParcelasSubstituicao: TWideStringField;

    CDSParcelassq: TIntegerField;
    DS_Receb_baixa: TDataSource;
    ZQReceb_Baixa: TZQuery;















    Label32: TLabel;

    ZQrecebimento: TZQuery;
    DataRecebimento: TDataSource;





























    Label35: TLabel;
    Elote: TEdit;
    btndes: TButton;
    Panel1: TPanel;
    fim: TLabel;
    dpl: TLabel;
    Label39: TLabel;
    EDLoteamento: TComboBox;
    Eadversa: TEdit;
    Label34: TLabel;
    Edit2: TEdit;
    DTPFim: TXDateEdit;
    DTPIni: TXDateEdit;
    JDEntrada: TXDateEdit;
    JDBaixa: TXDateEdit;
    WDTBaixaEntra: TXDBDateEdit;
    xnrecebido: TXNumEdit;
    YNEJuros: TXNumEdit;
    YNEDesconto: TXNumEdit;
    ZQCheque: TZQuery;

















































    DS_zqcheque: TDataSource;
    GroupBox2: TGroupBox;
    DBGrid1: TDBGrid;
    GroupBox3: TGroupBox;
    DBGrid3: TDBGrid;







    bar1: TProgressBar;
    CDSParcelasdata_quitacao: TDateField;
    CDSParcelasidrecebimento: TIntegerField;

    DS_Venda: TDataSource;
    ZQVenda: TZQuery;





    ZQEntrada: TZQuery;































    DS_Entrada: TDataSource;
    ZQParcela: TZQuery;





























    DS_Parcela: TDataSource;
    CDSParcelascontagem: TIntegerField;
    Label36: TLabel;
    DBEdit1: TDBEdit;
    CDSParcelasncheque: TWideStringField;
    CDSParcelastip: TWideStringField;



    Label37: TLabel;
    pini: TMaskEdit;
    Label38: TLabel;
    pfim: TMaskEdit;








    PB: TProgressBar;
    Lpassou: TLabel;
    Panel2: TPanel;
    Label40: TLabel;
    bxmanual: TComboBox;
    procedure Atualiza_tela;
    procedure DXBFecharClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DBGBaixandoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGBaixandoColEnter(Sender: TObject);
    procedure DBGBaixandoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure XNERecebidoExit(Sender: TObject);
    procedure DXBBaixaGravarClick(Sender: TObject);
    procedure DBCBTIPDOCExit(Sender: TObject);
    procedure DXBBaixaGravarEnter(Sender: TObject);
    procedure DBCBTIPDOCEnter(Sender: TObject);
    procedure DBEBaixaVrRecExit(Sender: TObject);
    procedure EContabilExit(Sender: TObject);
    procedure XDBNumEdit1Exit(Sender: TObject);
    procedure DBEBaixaDocumExit(Sender: TObject);
    procedure EContabilEnter(Sender: TObject);
    procedure DBECpf_CnpjExit(Sender: TObject);
    procedure EnomesExit(Sender: TObject);
    procedure Label1Click(Sender: TObject);
    procedure RBChequeClick(Sender: TObject);
    procedure GBBancoExit(Sender: TObject);
    procedure GBChequeExit(Sender: TObject);
    procedure EPlanocontabancoExit(Sender: TObject);
    procedure filtrabaixa;
    procedure GBContraExit(Sender: TObject);
    procedure RBParticipanteClick(Sender: TObject);
    procedure RBPonteiroClick(Sender: TObject);
    procedure XDBNumEdit1Enter(Sender: TObject);
    procedure JDEntradaExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ZQRecBaiCalcFields(DataSet: TDataSet);
    procedure CDSParcelasCalcFields(DataSet: TDataSet);
    procedure ZQRatearCalcFields(DataSet: TDataSet);
    procedure XNEJurosEnter(Sender: TObject);
    procedure CDS_divididoCalcFields(DataSet: TDataSet);
    procedure CBPeriodoEntExit(Sender: TObject);
    procedure EadversaExit(Sender: TObject);
    procedure EAdversaNovaExit(Sender: TObject);
    procedure CBTipobaixaExit(Sender: TObject);
    procedure DXBPesquisarClick(Sender: TObject);
    procedure EloteExit(Sender: TObject);
    procedure btndesClick(Sender: TObject);
    procedure EDLoteamentoDropDown(Sender: TObject);
    procedure Edit2Exit(Sender: TObject);
    procedure JDEntradaEnter(Sender: TObject);

    // procederes para arrastar componentes. aqui groupbox2
    procedure ControlMouseDown(Sender: TObject;
                             Button: TMouseButton;
                             Shift: TShiftState;
                             X, Y: Integer);
    procedure ControlMouseMove(Sender: TObject;
                             Shift: TShiftState;
                             X, Y: Integer);
    procedure ControlMouseUp(Sender: TObject;
                           Button: TMouseButton;
                           Shift: TShiftState;
                           X, Y: Integer);
    procedure XNEDescontoEnter(Sender: TObject);
    procedure DBGBaixandoDblClick(Sender: TObject);
    procedure DBGBaixandoExit(Sender: TObject);
    procedure DBGrid1ColEnter(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1Exit(Sender: TObject);
    procedure DBGrid1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid3ColEnter(Sender: TObject);
    procedure DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid3Exit(Sender: TObject);
    procedure DBGrid3KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGBaixandoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Label9MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pfimExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure XNEDescontoExit(Sender: TObject);
    procedure bxmanualExit(Sender: TObject);

  private
    { Private declarations }
  // variaveis para mudar componentes de lugar aqui o groupbox2
  inReposition: boolean;
  oldPos: TPoint;

    procedure AfterConstruction; override;

  public
    { Public declarations }
  end;

var
  FrmRecebBaixa: TFrmRecebBaixa;
  Vazio,VArDoc, varrecpag : String;
  xrec: integer;
  xdesc:double;
  Task : ITask;


implementation

uses tabelas, Cad_Recebimento, funcoes, ImpRecibo, PesqRecebimento_bx, uRuntimeFields;

{$R *.dfm}


// procedure de mover o formulario sem bordas
procedure MovimentaObject(Sender:TObject;Button:TMouseButton;Shift:TShiftState;X,Y:Integer;Formulario:TForm); 
 var ObjectPos,MousePosMov:TPoint;Pt:TPoint;fHandle:HWND; 
 begin 
   GetCursorPos(Pt); 
   ObjectPos.X:=Formulario.Left; 
   ObjectPos.Y:=Formulario.Top; 
   if(Sender is TForm)then 
     fHandle:=TWinControl(Sender).Handle 
   else 
     fHandle:=TWinControl(Sender).Parent.Handle; 
   while DragDetect(fHandle,ObjectPos) do 
   begin 
     GetCursorPos(MousePosMov); 
     Formulario.Left:=MousePosMov.X-X-3; 
     Formulario.Top:=MousePosMov.Y-Y-3; 
     Application.ProcessMessages;
   end; 
 end; 



procedure TFrmRecebBaixa.Atualiza_tela;
Var
  Varreg, varmarc, varvenc : integer;
Begin
  Varreg := ZQRecBai.RecNo;
  if vazio='T' then
  begin
    vazio:='F';
    Exit;
  end;  
  if ZQRecBai.active=false then
     ZQRecBai.open;

     
  ZQRecBai.First;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  YNEDesconto.Value:=0;
  XNEJuros.Value := 0;
  YNEJuros.Value := 0;
  varmarc := 0;
  varvenc := 0;
  Panel1.Visible:=true;
  ZQRecBai.DisableControls;
  PB.Visible:=true;
  PB.Position:=0;
  PB.Max:=ZQRecBai.RecordCount;
  Application.ProcessMessages;
  ZQRecBai.DisableControls;
  while not ZQRecBai.Eof do
  begin
    PB.Position:=ZQRecBai.RecNo;
    XNEAberto.Value := XNEAberto.Value + ZQRecBai.FieldByName('saldocorrig').Value;
    if ZQRecBai.FieldByName('marcar').Value = '0' Then
    Begin
      XNEMarcado.Value := XNEMarcado.Value + ZQRecBai.FieldByName('saldocorrig').Value;
      XNEJuros.Value := ZQRecBai.FieldByName('juros').value; //XNEJuros.Value + (ZQRecBaisaldocorrig.Value-ZQRecBaisaldo.Value); 17/05/2012
      XNEDesconto.Value:=ZQRecBai.FieldByName('descontos').Value;
      YNEJuros.Value:=YNEJuros.Value+XNEJuros.Value;
      YNEDesconto.Value:=YNEDesconto.Value+XNEDesconto.Value;
      inc(varmarc);
    end;
    if ZQRecBai.FieldByName('Dt_Vencimento').Value <=date Then Begin
      XNEVencido.Value := XNEVencido.Value + ZQRecBai.FieldByName('saldocorrig').Value;
      inc(varvenc);
    end;
    ZQRecBai.Next;
  end;
  ZQRecBai.EnableControls;
  PB.Position:=0;
  PB.Visible:=false;
  ZQRecBai.EnableControls;
  ZQRecBai.First;
  Application.ProcessMessages;
  Label31.Caption := inttostr(varmarc);
  Label33.Caption := inttostr(varvenc);
  XNERecebido.Value := (XNEMarcado.Value+YNEJuros.Value)-YNEDesconto.Value;
  XNEDesconto.Value:=YNEDesconto.Value;
  XNEJuros.Value:=YNEJuros.Value;
//mexi
{  if XNERecebido.Value < (XNEMarcado.Value+YNEJuros.Value) then
  begin
    if XNEResto.Value=0 then
       XNEDesconto.Value:=(XNEMarcado.Value+YNEJuros.Value)-XNERecebido.Value;
    if XNEDesconto.Value=0 then
       XNEResto.Value:= (XNEMarcado.Value+YNEJuros.Value)-XNERecebido.Value
    else
        XNEResto.Value:=0;
  end
  else
   //ate   }
   // XNEResto.Value:=0;
  ZQRecBai.RecNo := Varreg;
  if xrec>0 then
     ZQRecBai.Locate('idrecebimento',xrec,[]);
  xrec:=0;
  YNEJuros.Value:=0;
  YNEDesconto.Value:=0;
  Panel1.Visible:=false;
  Application.ProcessMessages;
End;

procedure TFrmRecebBaixa.DXBFecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRecebBaixa.FormShow(Sender: TObject);
begin
  vazio:='F';
  pini.Clear;
  pfim.Clear;
  Elote.Clear;
  Edit2.Clear;
  xrec:=0;

  //  DM_Tabelas.ZQDividido.open;
  dm_tabelas.ZQLoteamento.Open;
  DM_tabelas.ZQCobaRe.open;
  DM_Tabelas.ZQNumOrdem.open;
  DM_tabelas.contador_bx.open;
  DM_Tabelas.ZQMovBancaria.open;
  DM_tabelas.ZQCheque.open;
  DM_tabelas.ZQCaixa.open;

  DM_Tabelas.ZQCheque.Close;
  DM_Tabelas.ZQCheque.SQL.Clear;
  DM_Tabelas.ZQCheque.SQL.Add('select Banco,Dono,CPF_CNPJ,valor,emissao,deposito,alias,prorrogado,conta,idrecebimento,idpagamento,idcheque,idvinculo,numero,agencia,Substituicao,sq from cheque_rec limit 0, 1');
  DM_Tabelas.ZQCheque.open;


{  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_tabelas.ZQRecebimento.SQL.Add('Select * from Recebimento order by DT_Vencimento');
  DM_tabelas.ZQRecebimento.Open;}
//  if  DM_tabelas.ZQRecebimento.Active=false then
//       DM_tabelas.ZQRecebimento.Open;

  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and 1=0');
  DM_tabelas.ZQRecebimento.open;

  DM_tabelas.ZQReBxHi.close;
  DM_tabelas.ZQReBxHi.SQL.Clear;
  DM_tabelas.ZQReBxHi.SQL.Add('select idrecbxhist,refer,idrecib,valor,descon,juros,percent_usado,data,sq,valor_parcela from RecBxHist limit 0, 1');
  DM_tabelas.ZQReBxHi.open;


  DM_tabelas.ZQReceb_Baixa.close;
  DM_tabelas.ZQReceb_Baixa.SQL.Clear;
  DM_tabelas.ZQReceb_Baixa.SQL.Add('Select idreceb_baixa,Dt_rec,Vr_rec,TipDoc,Docum,Juros_Vr,Desc_Vr,RefBaixa,obsebx,Codcontabil,dataref,vencimento,substituicao,vr_ab,sq,valor_parcela from Receb_Baixa limit 0, 1');
  DM_tabelas.ZQReceb_Baixa.open;

  CDSParcelas.Close;
  CDSParcelas.CreateDataSet;
  CDSParcelas.Open;
  JDEntrada.DateText := datetostr(Date);
  JDBaixa.DateText := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
//  if empty(DM_tabelas.ZQRecebimentocliente.Text) Then Begin
//    Close;
//    Exit;
//  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.DateTEXT := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
  DM_tabelas.ZQCobaRe.post;

  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').Value;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;
 //aqui

  CBPeriodoEnt.ItemIndex := 0;
//  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;

  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
    Label5.Visible:=true;
    XNEParcelas.Visible:=true;
  end
  else
  begin
    XNEParcelas.Value:=1;
    Label5.Visible:=false;
    XNEParcelas.Visible:=false;
  end;

//  Lpassou.Caption:='N';
//  Task := Ttask.Create(procedure
//                       begin
//                            if ZQRecebBxTemp.Active=false then
//                            begin
//                              ZQRecebBxTemp.Close;
//                              ZQRecebBxTemp.SQL.Clear;
//                          //    ZQRecebBxTemp.SQL.Add('SELECT * FROM recbxhist as H left join recebimento as R ON R.idrecebimento=H.idrecib left join receb_baixa as B ON B.refbaixa=H.refer');
//                              ZQRecebBxTemp.SQL.Add('SELECT h.refer,h.idrecib,h.percent_usado,r.idrecebimento,b.dataref,b.vr_rec,b.refbaixa FROM recbxhist as H left join recebimento as R ON R.idrecebimento=H.idrecib left join receb_baixa as B ON B.refbaixa=H.refer');
//                              ZQRecebBxTemp.open;
//                              Lpassou.Caption:='S';
//                            end;
//
//                       end);
//  Task.Start;

//  if ZQRecebBxTemp.Active=false then
//  begin
//    ZQRecebBxTemp.Close;
//    ZQRecebBxTemp.SQL.Clear;
////    ZQRecebBxTemp.SQL.Add('SELECT * FROM recbxhist as H left join recebimento as R ON R.idrecebimento=H.idrecib left join receb_baixa as B ON B.refbaixa=H.refer');
//    ZQRecebBxTemp.SQL.Add('SELECT h.refer,h.idrecib,h.percent_usado,r.idrecebimento,b.dataref,b.vr_rec,b.refbaixa FROM recbxhist as H left join recebimento as R ON R.idrecebimento=H.idrecib left join receb_baixa as B ON B.refbaixa=H.refer');
//    ZQRecebBxTemp.open;
//  end;

  bxmanual.Text:=DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value;

  elote.SetFocus;
end;

procedure TFrmRecebBaixa.DBGBaixandoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);

begin
  if DBEidreg.Text = ZQRecBai.FieldByName('idrecebimento').Text Then
  Begin
    DBGBaixando.Canvas.Brush.Color := $006CFFFF;
    DBGBaixando.Canvas.Font.Color := $00A80000;
  end;
  DBGBaixando.DefaultDrawDataCell(Rect, DBGBaixando.columns[datacol].field, State);
  if Column.Field = ZQRecBai.FieldByName('marcar') then
  begin
    DBGBaixando.Canvas.FillRect(Rect);
    if ZQRecBai.FieldByName('marcar').Value = '0' then
      ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGBaixando.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
end;

procedure TFrmRecebBaixa.DBGBaixandoColEnter(Sender: TObject);
begin
  if DBGBaixando.SelectedField = ZQRecBai.FieldByName('marcar') then
    DBGBaixando.Options := DBGBaixando.Options - [dgEditing]
  else
    DBGBaixando.Options := DBGBaixando.Options + [dgEditing];
  dpl.Caption:='n';    
end;

procedure TFrmRecebBaixa.DBGBaixandoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if DBGBaixando.SelectedField.fieldname = 'marcar' Then
  Begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').Value = '0' Then
    begin
      ZQRecBai.FieldByName('marcar').Value := '1';
    end
    else
    begin
      ZQRecBai.FieldByName('marcar').Value := '0';
    end;
    ZQRecBai.post;
    if ZQRecBai.FieldByName('marcar').Value = '0' then
    begin
      DM_tabelas.ZQCobaRe.Insert;
      DTPIni.DateTEXT := datetostr(date);
      DTPFim.DateTEXT := datetostr(date);
      DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
      DM_tabelas.ZQCobaRe.post;
    end;
    Atualiza_tela;
  end;
end;

procedure TFrmRecebBaixa.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #27 Then
    close;
  if Key = #13 then Begin
    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;
end;

procedure TFrmRecebBaixa.XNERecebidoExit(Sender: TObject);
begin
  // 24/05/2012

  if (XNEMarcado.Value > (XNERecebido.Value+XNEJuros.value)) and (XNEDesconto.Value=0) Then
     XNEResto.Value := ((XNEMarcado.Value+XNEJuros.value)) - (XNERecebido.Value);
   //  XNEResto.Value :=xdesc;
  xdesc:=0;

end;

procedure TFrmRecebBaixa.DXBBaixaGravarClick(Sender: TObject);
Var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varsomarepasse, varsomarepasseautomatico : Double;
  contador,varreg : Integer;
  quadra,verif,varordem, varnumordem1, varnumordem2 : string;
begin
  if empty(CBTipobaixa.Text) then
  begin
    showmessage('Selecione o tipo de opera��o !');
    CBTipobaixa.SetFocus;
    exit;
  end;
  xrec:=0;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
    ZQReceb_Baixa.open;
    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(CDSParcelasDocum.Value))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value)));
    ZQReceb_Baixa.Filtered:=true;
    if ZQReceb_Baixa.RecordCount>0 then
       ZQReceb_Baixa.Delete;
    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.close;

    DXBBaixaGravar.Enabled := False;
    DXBFechar.Enabled := False;
    ZQRecBai.Filter := 'marcar=0';
    ZQRecBai.Filtered := True;
    LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
    varnumordem1 := '';
    varnumordem2 := '';

    // coloquei as proximas linha 25/06/2024 para mudar a referencia nas baixas parciais

        DM_tabelas.ZQCobaRe.Insert;
        DTPIni.Datetext := datetostr(date);
        DTPFim.Datetext := datetostr(date);
        DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
        DM_tabelas.ZQCobaRe.post;


    If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
      DM_tabelas.ZQCobaRe.Edit;
    DM_tabelas.ZQCobaRe.FieldByName('entrada').Value := strtodate(JDEntrada.DateText);
    DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').Value := CBTipobaixa.Text;
    DM_tabelas.ZQCobaRe.Post;

    CDSParcelas.First;
    varsomarepasse := 0;
    DM_tabelas.contador_bx.close;
    DM_tabelas.contador_bx.Open;
    DM_tabelas.contador_bx.Last;
    if DM_tabelas.contador_bx.RecordCount>0 then
       contador:=DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value
    else
      contador:=0;

    bar1.visible:=true;
    bar1.position:=0;
    bar1.max:=CDSParcelas.Recordcount;
    CDSParcelas.DisableControls;
    while not CDSParcelas.Eof do
    begin
      bar1.position:=CDSParcelas.Recno;
      inc(contador);
      if (CBTipobaixa.Text<>'SUBSTITUI��O')  then
      begin
        DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.DateText); //DM_tabelas.ZQCobareHoje.Value;
        DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
        if CDSParcelastip.Value<>'J' then
        begin
           DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
        end
        else
        begin
           DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasVrParc.Value;
           DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value := 0;
        end;
        DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;
      end
      else
      begin
        if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
        begin
          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.DateText); //DM_tabelas.ZQCobareHoje.Value;
          DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
          DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
          if CDSParcelastip.Value<>'J' then
          begin
             DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
          end
          else
          begin
             DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasVrParc.Value;
             DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value := 0;
          end;
          DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='S';
          DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQReceb_Baixa.Post;

        end
        else
        begin
          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQreceb_baixa.FieldByName('VR_AB').Value := CDSParcelasVrParc.Value;
    //      DM_tabelas.ZQReceb_BaixaDt_rec.Value := DM_tabelas.ZQCobareHoje.Value;
          DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
          DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
//          DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
  //        DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='S';
          DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQReceb_Baixa.Post;

        end;
      end;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
      if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').Value='S' Then
        varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;
      if DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').Value = 'S' Then
      Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          if empty(varnumordem1) Then
          Begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          DM_tabelas.ZQRecebimento.Insert;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
          DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
          DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
          DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('tip').Value   := CDSParcelastip.Value;
          DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

          DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
          DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;

          Dm_tabelas.ZQLoteamento.Filtered:=false;
          Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
          Dm_tabelas.ZQLoteamento.Filtered:=true;
          if Dm_tabelas.ZQLoteamento.RecordCount>0 then
             Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
           Dm_tabelas.ZQLoteamento.Filtered:=false;

          DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMObsbx.Text; //EMNovostit.Text;
          DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := ZQRecBai.FieldByName('recpag').Value;
          DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := strtoint(EcodAdversaNova.text);
          DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);
          DM_tabelas.ZQRecebimento.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
          DM_tabelas.ZQRecebimento.Post;
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O')  then
        begin
          if empty(varnumordem1) Then
          Begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          DM_tabelas.ZQRecebimento.Insert;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
          DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
          DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
          DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;

          DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

          DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
          DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;

          DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;
          Dm_tabelas.ZQLoteamento.Filtered:=false;
          Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
          Dm_tabelas.ZQLoteamento.Filtered:=true;
          if Dm_tabelas.ZQLoteamento.RecordCount>0 then
             Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
           Dm_tabelas.ZQLoteamento.Filtered:=false;

    //     if (CDSParcelasTipDoc.Value<>'CH') then
           DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;

          DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;;
          DM_tabelas.ZQRecebimento.FieldByName('Observ').Value   :=  EMObsbx.Text;//EMNovostit.Text;
          DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value   := ZQRecBai.FieldByName('recpag').Value;
          DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value  := strtoint(EcodAdversaNova.text);
          DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);

          DM_Tabelas.ZQRecebimento.FieldByName('Substituicao').Value:='S';
          DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
          DM_tabelas.ZQRecebimento.Post;

        end;
      end;
      if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').Value='S' Then
      Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O')then
        begin
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
            DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCaixa.Post;
        end
        else
        begin
          if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
          begin
            DM_tabelas.ZQCaixa.Insert;
            DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
            DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
            DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
            DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
            DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
            DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
            DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
            DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
            IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
              DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
            DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
            DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
            IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
              DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
            end;
            DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
            DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value := 'S';
            DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
            DM_tabelas.ZQCaixa.Post;

          end;
        end;
      end;
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then
      Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          DM_tabelas.ZQCheque.Insert;
          DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
          DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
          DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
          DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
          DM_tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasNcheque.Value;
          DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.DateText);
          DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
          DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
          DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
//          DM_Tabelas.ZQChequeidvinculo.Value := DM_Tabelas.ZQRecebimentoidrecebimento.Value;
          DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCheque.Post;
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O') then
        begin
          DM_tabelas.ZQCheque.Insert;
          DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
          DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
          DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
          DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
          DM_tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasNcheque.Value;
          DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.DateText);
          DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
          DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
          DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
//          DM_Tabelas.ZQChequeidvinculo.Value := DM_Tabelas.ZQRecebimentoidrecebimento.Value;
          DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value := 'S';
          DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCheque.Post;

        end;
      end;
      if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value = 'S' Then Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          DM_Tabelas.ZQMovBancaria.Insert;
          DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
          DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
          DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
          DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
          DM_Tabelas.ZQMovBancaria.Post;
          //      CalcSaldo('','',False);
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O') then
        begin
          DM_Tabelas.ZQMovBancaria.Insert;
          DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
          DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
          DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
          DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value := 'S';
          DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
          DM_Tabelas.ZQMovBancaria.Post;

          //      CalcSaldo('','',False);
        end;
      end;
      CDSParcelas.Next;
    end;
    bar1.position:=0;
    bar1.visible:=false;
    CDSParcelas.EnableControls;
    CDSParcelas.Last;

    DM_tabelas.contador_bx.insert;
    DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value:=contador;
    DM_tabelas.contador_bx.Post;
    DM_tabelas.contador_bx.close;

    if not empty(varnumordem1) Then Begin
      ZQDiv_usado.SQL.Clear;
      ZQDiv_usado.SQL.Add('select iddividido,participante,percentual,planodeconta,tipodocumento,ordem from dividido');
      ZQDiv_usado.SQL.Add(' where  ordem = '+quotedstr(ZQRecBai.FieldByName('numordem').Text));
      ZQDiv_usado.Open;
      varsomarepasseautomatico := 0;
      ZQDiv_usado.DisableControls;
      while not ZQDiv_usado.Eof do begin
        DM_Tabelas.ZQDividido.Insert;
        DM_Tabelas.ZQDividido.FieldByName('participante').Value := ZQDiv_usado.FieldByName('participante').Value;
        DM_Tabelas.ZQDividido.FieldByName('percentual').Value := ZQDiv_usado.FieldByName('percentual').Value;
        DM_Tabelas.ZQDividido.FieldByName('planodeconta').Value := ZQDiv_usado.FieldByName('planodeconta').Value;
        DM_Tabelas.ZQDividido.FieldByName('tipodocumento').Value := ZQDiv_usado.FieldByName('tipodocumento').Value;
        DM_Tabelas.ZQDividido.FieldByName('ordem').Value := strtoint(varnumordem1);
        DM_Tabelas.ZQDividido.Post;
        ZQDiv_usado.Next;
      end;
    end;
    ZQDiv_usado.EnableControls;
    ZQRecBai.First;
    varpago := 0;
    varjuros:=0;
    Memo1.Clear;
    baixatot := XNERecebido.Value;
    bar1.visible:=true;
    bar1.position:=0;
    bar1.max:=ZqRecBai.Recordcount;
    ZqRecBai.DisableControls;
    While not ZqRecBai.Eof do
    Begin
      bar1.position:=ZqRecBai.RecNo;
      DM_tabelas.ZQReBxHi.Insert;
      percentbaixa := ZQRecBai.FieldByName('saldocorrig').Value * 100 / XNEMarcado.Value;
      DM_Tabelas.ZQReBxHi.FieldByName('descon').Value := XNEDesconto.Value * percentbaixa/100;
      if ZQRecBai.FieldByName('saldocorrig').Value<=baixatot Then
        varpago := ZQRecBai.FieldByName('saldocorrig').Value
      else
        varpago := baixatot+DM_Tabelas.ZQReBxHi.FieldByName('descon').Value;
      varjuros := XNEJuros.Value* percentbaixa/100;

      if varjuros <=varpago Then
        DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := varjuros
      else
        DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := varpago;
      if XNEResto.value=0 then
         varpago := varpago + DM_Tabelas.ZQReBxHi.FieldByName('juros').Value;

      DM_Tabelas.ZQReBxHi.FieldByName('valor').Value := varpago-DM_Tabelas.ZQReBxHi.FieldByName('descon').Value;


      // mexi aqui 25/09/2024 para ver se nao sai duplicado no relatorio
    //  DM_Tabelas.ZQReBxHidata.Value := strtodate(JDEntrada.DateText); //JDBaixa.DateText;
      if CDSParcelasData_Quitacao.Value>0 then
         DM_Tabelas.ZQReBxHi.FieldByName('data').Value := CDSParcelasData_Quitacao.Value
      else
         DM_Tabelas.ZQReBxHi.FieldByName('data').Value := strtodate(JDEntrada.DateText);



      DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').Value := percentbaixa;
      DM_tabelas.ZQReBxHi.FieldByName('refer').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
      DM_tabelas.ZQReBxHi.FieldByName('idrecib').Value := ZQRecBai.FieldByName('idrecebimento').Value;
      DM_tabelas.ZQReBxHi.FieldByName('sq').Value := contador;
      varx := ExRound(varpago,2)-ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').Value,2);
      DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').Value :=varx;

      if ZQRecBai.FieldByName('tip').Value='J' then
      begin
        DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := varpago;
        DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').Value :=0;
      end;


      DM_Tabelas.ZQReBxHi.Post;

      ZQRecBai.Edit;
      if ZQRecBai.FieldByName('tip').Value<>'J' then
         varx := ExRound(varpago,2)-ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').Value,2);
//    varx := varx + DM_Tabelas.ZQReBxHidescon.Value;
//    ZQRecBaisaldo.Value := truncar(ZQRecBaisaldo.Value - varx,2);
      if ZQRecBai.FieldByName('tip').Value='J' then
      begin
        ZQRecBai.FieldByName('saldo').Value := ExRound(ZQRecBai.FieldByName('saldo').Value,2) - ExRound(varx,2);
        ZQRecBai.FieldByName('Valor').Value := ZQRecBai.FieldByName('saldo').Value;
      end
      else
      begin
        ZQRecBai.FieldByName('saldo').Value := ExRound(ZQRecBai.FieldByName('saldo').Value,2) - ExRound(varx,2);
      end;
      ZQRecBai.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;


      // 17/05/2012
      ZQRecBai.FieldByName('juros').Value:=0;
      ZQRecBai.FieldByName('descontos').Value:=0;
      ZQRecBai.post;

      // 22/09/2010
      DM_Tabelas.ZQRecebimento.Locate('idrecebimento',ZQRecBai.FieldByName('idrecebimento').Value,[]);
      quadra:= DM_Tabelas.ZQRecebimento.FieldByName('quadralote').Value;
      DM_TAbelas.ZQRecebimento.Edit;
      DM_Tabelas.ZQRecebimento.FieldByName('saldo').Value := ZQRecBai.FieldByName('saldo').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('Valor').Value := ZQRecBai.FieldByName('Valor').Value;
      if ZQRecBai.FieldByName('saldo').Value>0 then
         DM_Tabelas.ZQRecebimento.FieldByName('tip').Value:='P';

      if (ZQRecBai.FieldByName('saldo').Value=0) and (DM_Tabelas.ZQRecebimento.FieldByName('tip').Value='P') then
         DM_Tabelas.ZQRecebimento.FieldByName('tip').Value:='';

      DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').Value := ZQRecBai.FieldByName('RefBaixa').Value;
      DM_Tabelas.ZQRecebimento.Post;
      // at?aqui


      if (CBTipobaixa.Text='SUBSTITUI��O') then
      begin
//      ZQRecBaisubstituicao.Value:='S';
        ZQRecBai.edit;
        ZQRecBai.FieldByName('somar').Value:='N';
        ZQRecBai.post;
//        ZQRecBaisq.Value := contador;
      end;
      if pos('-',ZQRecBai.FieldByName('ordem').Value)=0 Then
        Varordem := Varordem + ZQRecBai.FieldByName('idrecebimento').Text+' - ';
      if pos(ZQRecBai.FieldByName('nomeadversa').Value,Memo1.Text)=0 Then
        Memo1.Lines.Add(ZQRecBai.FieldByName('nomeadversa').Value);
      ZqRecBai.Next;
    end;
    ZqRecBai.EnableControls;
    CDS_dividido.Close;
    CDS_dividido.CreateDataSet;
    ZQRatear.SQL.Clear;
    ZQRatear.SQL.Add('select idrecebimento, h.valor, percent_usado, b.vr_rec, participante, percentual,planodeconta, dv.tipodocumento, re.ordem, dv.ordem from recbxhist as H');
    ZQRatear.SQL.Add(' join recebimento as re ON H.idrecib=RE.idrecebimento');
    ZQRatear.SQL.Add(' join dividido as dv ON dv.ordem=numordem');
    ZQRatear.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
    ZQRatear.SQL.Add(' join tipodocumento as D ON B.tipdoc=D.tipodoc');
    ZQRatear.SQL.Add(' where somapaga = '+quotedstr('S')+' and refer = '+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text);
    ZQRatear.Open;
    varsomarepasseautomatico := 0;
    ZQRatear.DisableControls;
    while not ZQRatear.Eof do begin
      if CDS_dividido.Locate('participante;contabil;tipodoc',VarArrayOf([ZQRatear.FieldByName('participante').Value,ZQRatear.FieldByName('planodeconta').Value,ZQRatear.FieldByName('tipodocumento').Value]),[]) then
        CDS_dividido.Edit
      else
        CDS_dividido.Insert;
      CDS_divididoparticipante.Value := ZQRatear.FieldByName('participante').Value;
      CDS_divididovr_base.Value := CDS_divididovr_base.Value + ZQRatear.FieldByName('vr_base').Value;
      CDS_divididovr_calc.Value := CDS_divididovr_calc.Value + ZQRatear.FieldByName('vr_divi').Value;
      CDS_divididocontabil.Value := ZQRatear.FieldByName('planodeconta').Value;
      CDS_divididotipodoc.Value := ZQRatear.FieldByName('tipodocumento').Value;
      CDS_divididoobs.Value := CDS_divididoobs.Value+ZQRatear.FieldByName('idrecebimento').Text+' | '+ZQRatear.FieldByName('ordem').Value+' | '+ZQRatear.FieldByName('vr_base').Text+' x '+ZQRatear.FieldByName('percentual').Text+' = '+ZQRatear.FieldByName('vr_divi').Text+chr(13);
      varsomarepasseautomatico := varsomarepasseautomatico + ZQRatear.FieldByName('vr_divi').Value;
      ZQRatear.Next;
    end;
    ZQRatear.EnableControls;
    CDS_dividido.First;
    if CDS_dividido.RecordCount>0 Then Begin
      DM_Tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
    end;
    CDS_dividido.DisableControls;
    while not CDS_dividido.Eof do begin
      DM_tabelas.ZQRecebimento.Insert;
      DM_tabelas.ZQRecebimento.FieldByName('documento').Value := 'Rateio Bx - '+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
      DM_tabelas.ZQRecebimento.FieldByName('cliente').Value :=  CDS_divididoparticipante.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := strtodate(JDEntrada.DateText);
      DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := strtodate(JDEntrada.DateText);
      DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDS_divididocontabil.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMObsbx.Text+chr(13)+CDS_divididoobs.Value;//EMNovostit.Text
      DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := 'P';
      DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDS_divididotipodoc.Value;
      DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount);
      DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
      DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value :=  CDS_divididovr_base.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Valor').Value :=  CDS_divididovr_calc.Value;
      DM_tabelas.ZQRecebimento.FieldByName('saldo').Value :=  DM_tabelas.ZQRecebimento.FieldByName('Valor').Value;
      DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := DM_Tabelas.ZQConfiguracoes.FieldByName('paticipantedefault').Value;

      DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

      DM_tabelas.ZQRecebimento.Post;
      CDS_dividido.Next;
    end;
    CDS_dividido.EnableControls;
    varreg := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
    DM_Tabelas.ZQRecebimento.Close;
    DM_Tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where idrecebimento='+quotedstr(varreg.ToString)+' order by DT_Vencimento');
    DM_tabelas.ZQRecebimento.Open;
    DM_Tabelas.ZQRecebimento.Locate('idrecebimento',varreg,[]);

    FrmImpRecibo:=nil;
    if FrmImpRecibo=nil then
       FrmImpRecibo:=TFrmImpRecibo.Create(self);
    FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
    FrmImpRecibo.Nomecli.Text := Enomes.Text;
    FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

    ///
    ZQRecBai.Refresh;
    DBGBaixando.Refresh;
    DM_tabelas.ZQReceb_Baixa.Refresh;
    DM_Tabelas.ZQReBxHi.Refresh;
    CDSParcelas.Close;
    CDSParcelas.CreateDataSet;
    XNEJuros.Value:=0;
    XNERecebido.Value:=0;
    XNEDesconto.Value:=0;
    XNEResto.Value:=0;
    XNEParcelas.Value:=1;
    CBPeriodoEnt.Text:='HOJE';
    EMObsbx.Clear;
    EMNovostit.Clear;


    ZQRecBai.Filtered := False;
//    DM_tabelas.ZQRecebimento.Refresh;

    Panel1.Caption:='Aguarde! Ajustando Saldo da Venda...';
    Panel1.Visible:=true;
    application.ProcessMessages;
    ZQRecebimento.close;
    ZQRecebimento.sql.clear;
    ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    ZQRecebimento.sql.add(' from recebimento where quadralote='+quotedstr(quadra)+' group by quadralote order by quadralote');
    ZQRecebimento.Open;
    ZQRecebimento.First;
    bar1.Position := 0;
    bar1.Max := ZQRecebimento.RecordCount;
    bar1.visible:=true;
    application.ProcessMessages;
    ZQVenda.open;
    while not ZQRecebimento.Eof do
    begin
      bar1.Position := ZQRecebimento.RecNo;
      application.ProcessMessages;
      ZQEntrada.Close;
      ZQEntrada.SQL.Clear;
      ZQEntrada.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQEntrada.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQEntrada.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      ZQEntrada.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
      ZQEntrada.SQL.Add(', sum(valor) as entrada from recebimento where tip<>''J'' and quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').Value)+' and documento like  ''%-E-%''  order by ordem');
      ZQEntrada.open;

      ZQParcela.Close;
      ZQParcela.SQL.Clear;
      ZQParcela.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQParcela.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQParcela.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      ZQParcela.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
      ZQParcela.SQL.Add(', sum(valor) as parcela from recebimento where tip<>''J'' and quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').Value)+' and documento like  ''%-P-%''  order by ordem');
      ZQParcela.open;
      if (ZQEntrada.FieldByName('entrada').Value>0) or (ZQParcela.FieldByName('parcela').Value>0) then
      begin
        ZQVenda.SQL.Clear;
        ZQVenda.SQL.Add('Select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao from venda where idvenda='+quotedstr(ZQRecebimento.FieldByName('venda_idvenda').Text));
        ZQVenda.open;
        if ZQVenda.RecordCount>0 then
        begin
          ZQVenda.Edit;
          if (ZQParcela.FieldByName('tip').Value<>'J') and (ZQEntrada.FieldByName('tip').Value<>'J') then
             ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
          ZQVenda.Post;
        end;

      end;
      ZQRecebimento.Next;
    end;

    bar1.Position := 0;
    bar1.visible:=false;
    ZQVenda.close;
    ZQRecebimento.close;
    ZQEntrada.close;
    ZQParcela.close;
    Panel1.Caption:='AGUARDE! ATUALIZANDO A TELA.   ';
    Panel1.Visible:=false;
//    CBTipobaixa.SetFocus;
    Elote.SetFocus;
    DXBFechar.Enabled := true;
    application.ProcessMessages;
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(quadra)+' order by DT_Vencimento');
    DM_tabelas.ZQRecebimento.open;



  //  Close;
     // imprime o recibo
    //  FrmImpRecibo.showmodal;
  ///
  {  if perguntaSN('Gerar Recibo de Quita��o ?','S') Then Begin
      FrmRelRecibodeQuita.ZQForma.SQL.Clear;
      FrmRelRecibodeQuita.ZQForma.SQL.Add('Select * from receb_baixa where refbaixa='+DM_tabelas.ZQCobaRecotagem.Text);
      FrmRelRecibodeQuita.ZQForma.Open;
      FrmRelRecibodeQuita.RLLabel19.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
      FrmRelRecibodeQuita.RLLabel54.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
      FrmRelRecibodeQuita.RLMemo1.Lines := EMObsbx.Lines;
      FrmRelRecibodeQuita.RLMemo2.Lines := Memo1.Lines;
      FrmRelRecibodeQuita.RLMemo3.Lines := FrmRelRecibodeQuita.RLMemo2.Lines;
      FrmRelRecibodeQuita.RLLabel18.Caption := DM_tabelas.ZQCobaRecotagem.Text;
      FrmRelRecibodeQuita.RLLabel39.Caption := DM_tabelas.ZQCobaRecotagem.Text;
      ZQItensRecibo.Close;
      ZQItensRecibo.SQL[1] := 'where refer='+DM_tabelas.ZQCobaRecotagem.Text;
      ZQItensRecibo.Open;
      if FrmRelRecibodeQuita=nil then
        FrmRelRecibodeQuita:=TFrmRelRecibodeQuita.Create(Application);
      FrmRelRecibodeQuita.RLReport1.Preview;
    end; }
    ZQRecBai.Filtered := False;
//    DM_tabelas.ZQRecebimento.Refresh;
//    CBTipobaixa.SetFocus;
    Elote.SetFocus;
    DXBFechar.Enabled := true;
    mensagem('Opera��o Terminada!');
  end
  else
  begin

    verif:=CBTipobaixa.Text;
    CDSParcelas.First;
    CDSParcelas.DisableControls;
    while not CDSParcelas.Eof do
    begin
      ZQReceb_Baixa.open;
      ZQReceb_Baixa.Filtered:=false;
      ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(CDSParcelasDocum.Value))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value)));
      ZQReceb_Baixa.Filtered:=true;
      if ZQReceb_Baixa.RecordCount>0 then
         ZQReceb_Baixa.Delete;
      ZQReceb_Baixa.Filtered:=false;
      ZQReceb_Baixa.close;
      CDSParcelas.Next;
    end;
    CDSParcelas.First;
    CDSParcelas.EnableControls;

{    ZQReceb_Baixa.open;
    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.Filter:='docum='+quotedstr(alltrim(CDSParcelasDocum.Value))+' and sq='+quotedstr(alltrim(inttostr(CDSParcelassq.value)));
    ZQReceb_Baixa.Filtered:=true;
    if ZQReceb_Baixa.RecordCount>0 then
       ZQReceb_Baixa.Delete;
    ZQReceb_Baixa.Filtered:=false;
    ZQReceb_Baixa.close;}

    DXBBaixaGravar.Enabled := False;
    DXBFechar.Enabled := False;
    ZQRecBai.Filter := 'marcar=0';
    ZQRecBai.Filtered := True;
    LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
    varnumordem1 := '';
    varnumordem2 := '';
{    If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
      DM_tabelas.ZQCobaRe.Edit;
    DM_tabelas.ZQCobaReentrada.Value := CDSParcelasData_Quitacao.Value;
    DM_Tabelas.ZQCobaRetipobaixa.Value := CBTipobaixa.Text;
    DM_tabelas.ZQCobaRe.Post;}

    CDSParcelas.First;
    varsomarepasse := 0;
    DM_tabelas.contador_bx.close;
    DM_tabelas.contador_bx.Open;
    DM_tabelas.contador_bx.Last;
    if DM_tabelas.contador_bx.RecordCount>0 then
       contador:=DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value
    else
      contador:=0;

    bar1.visible:=true;
    bar1.position:=0;
    bar1.max:=CDSParcelas.Recordcount;
    CDSParcelas.DisableControls;
    while not CDSParcelas.Eof do
    begin
      bar1.position:=CDSParcelas.Recno;
      inc(contador);
      if (CBTipobaixa.Text<>'SUBSTITUI��O')  then
      begin

        If not (DM_tabelas.ZQCobaRe.State in [DsEdit,DsInsert]) Then
           DM_tabelas.ZQCobaRe.insert;
        DM_tabelas.ZQCobaRe.FieldByName('entrada').Value := CDSParcelasData_Quitacao.Value;
        DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').Value := CBTipobaixa.Text;
        DM_tabelas.ZQCobaRe.Post;

        CDSParcelas.edit;
        CDSParcelascontagem.Value:=DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        CDSParcelas.post;


        DM_tabelas.ZQReceb_Baixa.Insert;
        DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
        if CDSParcelastip.Value<>'J' then
           DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value)
        else
           DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=0;
        DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.DateText); //DM_tabelas.ZQCobareHoje.Value;
        if CDSParcelasData_Quitacao.Value>0 then
           DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := CDSParcelasData_Quitacao.Value
        else
           DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
        DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
        if CDSParcelastip.Value<>'J' then
           DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value //  XNEJuros.Value/CDSParcelas.RecordCount;
        else
           DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasVrParc.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
        DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
        DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
        DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
        DM_tabelas.ZQReceb_Baixa.Post;

      end
      else
      begin
        if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
        begin
          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQreceb_baixa.FieldByName('Vr_rec').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQReceb_Baixa.FieldByName('valor_parcela').Value :=(CDSParcelasVrParc.Value+CDSParcelasdesconto.Value-CDSParcelasjuros.value);
          DM_tabelas.ZQreceb_baixa.FieldByName('Dt_rec').Value := strtodate(JDBaixa.DateText); //DM_tabelas.ZQCobareHoje.Value;
          if CDSParcelasData_Quitacao.Value>0 then
             DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := CDSParcelasData_Quitacao.Value
          else
             DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
          DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Juros_Vr').Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
          DM_tabelas.ZQreceb_baixa.FieldByName('Desc_Vr').Value  := CDSParcelasdesconto.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQReceb_Baixa.Post;

        end
        else
        begin
          DM_tabelas.ZQReceb_Baixa.Insert;
          DM_tabelas.ZQreceb_baixa.FieldByName('VR_AB').Value := CDSParcelasVrParc.Value;
    //      DM_tabelas.ZQReceb_BaixaDt_rec.Value := DM_tabelas.ZQCobareHoje.Value;
          if CDSParcelasData_Quitacao.Value>0 then
             DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := CDSParcelasData_Quitacao.Value
          else
             DM_TAbelas.ZQreceb_baixa.FieldByName('dataref').Value := strtodate(JDEntrada.DateText);
          DM_TAbelas.ZQreceb_baixa.FieldByName('vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value  := CDSParcelasDocum.Value;
//          DM_tabelas.ZQReceb_BaixaJuros_Vr.Value := CDSParcelasjuros.Value; //  XNEJuros.Value/CDSParcelas.RecordCount;
  //        DM_tabelas.ZQReceb_BaixaDesc_Vr.Value  := CDSParcelasdesconto.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('RefBaixa').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Codcontabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('obsebx').Value := EMObsbx.Text;
          DM_tabelas.ZQreceb_baixa.FieldByName('Docum').Value:=CDSParcelasDocum.Value;
          DM_tabelas.ZQreceb_baixa.FieldByName('Substituicao').Value:='S';
          DM_tabelas.ZQreceb_baixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQReceb_Baixa.Post;

        end;
      end;
      DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
      if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').Value='S' Then
        varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;

      if DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').Value = 'S' Then
      Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          if empty(varnumordem1) Then
          Begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          DM_tabelas.ZQRecebimento.Insert;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
          DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
          DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
          DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;

          DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

          DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
          DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;

          Dm_tabelas.ZQLoteamento.Filtered:=false;
          Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
          Dm_tabelas.ZQLoteamento.Filtered:=true;
          if Dm_tabelas.ZQLoteamento.RecordCount>0 then
             Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
           Dm_tabelas.ZQLoteamento.Filtered:=false;

          DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMObsbx.Text; //EMNovostit.Text; // 24/05/2012 a pedido da Erika mar?
          DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := ZQRecBai.FieldByName('recpag').Value;
          DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := strtoint(EcodAdversaNova.text);
          DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);
          DM_tabelas.ZQRecebimento.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
          DM_tabelas.ZQRecebimento.Post;
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O')  then
        begin
          if empty(varnumordem1) Then
          Begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          quadra:=CDSParcelasquadralote.Value;
          DM_tabelas.ZQRecebimento.Insert;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQRecebimento.FieldByName('cliente').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := DM_tabelas.ZQCobaRe.FieldByName('hoje').Value;
          DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasVenci.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Valor').Value := CDSParcelasVrParc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value := XNERecebido.Value;
          DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQRecebimento.FieldByName('usuario').Value := DM_tabelas.ZQUsuario.FieldByName('idusuario').Value;
          DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := varnumordem1+'-'+CDSParcelasOrdem.Value;
          DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDSParcelasTipDoc.Value;

          DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

          DM_tabelas.ZQRecebimento.FieldByName('venda_idvenda').Value:=CDSParcelasvenda_idvenda.Value;
          DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value:=CDSParcelasquadralote.Value;

          DM_tabelas.ZQRecebimento.FieldByName('documento').Value:=CDSParcelasDocum.Value;
          Dm_tabelas.ZQLoteamento.Filtered:=false;
          Dm_tabelas.ZQLoteamento.Filter:='Apelido='+quotedstr(CDSParcelasnomedoempreendimento.Value);
          Dm_tabelas.ZQLoteamento.Filtered:=true;
          if Dm_tabelas.ZQLoteamento.RecordCount>0 then
             Dm_tabelas.ZQRecebimento.FieldByName('idloteamento').Value:=Dm_tabelas.ZQLoteamento.FieldByName('idloteamento').Value;
           Dm_tabelas.ZQLoteamento.Filtered:=false;

    //     if (CDSParcelasTipDoc.Value<>'CH') then
           DM_tabelas.ZQRecebimento.FieldByName('saldo').Value := CDSParcelasVrParc.Value;

          DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;;
          DM_tabelas.ZQRecebimento.FieldByName('Observ').Value :=  EMObsbx.Text;//EMNovostit.Text;
          DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := ZQRecBai.FieldByName('recpag').Value;
          DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := strtoint(EcodAdversaNova.text);
          DM_Tabelas.ZQRecebimento.FieldByName('numordem').Value := strtoint(varnumordem1);

          DM_Tabelas.ZQRecebimento.FieldByName('Substituicao').Value:='S';
          DM_tabelas.ZQRecebimento.FieldByName('sq').Value := contador;
          DM_tabelas.ZQRecebimento.Post;

        end;
      end;
      if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').Value='S' Then
      Begin
        if (CBTipobaixa.Text<>'SUBSTITUI��O')then
        begin
          quadra:=CDSParcelasquadralote.Value;
          DM_tabelas.ZQCaixa.Insert;
          DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
          DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
          DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
          DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
          IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
            DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
          end;
          DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
          DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCaixa.Post;
        end
        else
        begin
          if (CDSParcelasTipDoc.Value='DN') or (CDSParcelasTipDoc.Value='DP') then
          begin
            quadra:=CDSParcelasquadralote.Value;
            DM_tabelas.ZQCaixa.Insert;
            DM_tabelas.ZQCaixa.FieldByName('plano_contas_codigo').Value := CDSParcelascodContabil.Value;
            DM_tabelas.ZQCaixa.FieldByName('data_lan').Value := date;
            DM_tabelas.ZQCaixa.FieldByName('vr_lan').Value := CDSParcelasVrparc.Value;
            DM_tabelas.ZQCaixa.FieldByName('documento').Value := CDSParcelasDocum.Value;
            DM_tabelas.ZQCaixa.FieldByName('empresa').Value := 1;
            DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascodContabil.Value,[]);
            DM_tabelas.ZQCaixa.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
            DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
            IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
              DM_tabelas.ZQCaixa.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
            DM_tabelas.ZQCaixa.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
            DM_tabelas.ZQCaixa.FieldByName('obs').Value := 'Recebimento - '+' Ordem '+CDSParcelasOrdem.Value;
            IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then Begin
              DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value + ' Vencimento em: '+datetostr(CDSParcelasVenci.Value);
            end;
            DM_tabelas.ZQCaixa.FieldByName('obs').Value :=DM_tabelas.ZQCaixa.FieldByName('obs').Value +' | '+EMObsbx.Text;
            DM_tabelas.ZQCaixa.FieldByName('Substituicao').Value := 'S';
            DM_tabelas.ZQCaixa.FieldByName('sq').Value := contador;
            DM_tabelas.ZQCaixa.Post;

          end;
        end;
      end;
      IF DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' Then
      Begin
        quadra:=CDSParcelasquadralote.Value;
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          DM_tabelas.ZQCheque.Insert;
          DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
          DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
          DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
          DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
          DM_tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasNcheque.Value;
          DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.DateText);
          DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
          DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
          DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
//          DM_Tabelas.ZQChequeidvinculo.Value := DM_Tabelas.ZQRecebimentoidrecebimento.Value;
          DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCheque.Post;
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O') then
        begin
          DM_tabelas.ZQCheque.Insert;
          DM_tabelas.ZQCheque.FieldByName('Banco').Value := CdSParcelasBanco.Value;
          DM_tabelas.ZQCheque.FieldByName('agencia').Value := CDSParcelasagencia.Value;
          DM_tabelas.ZQCheque.FieldByName('Dono').Value := CDSParcelasdono.Value;
          DM_tabelas.ZQCheque.FieldByName('CPF_CNPJ').Value := CDSParcelasCPF_CNPJ.Value;
          DM_tabelas.ZQCheque.FieldByName('numero').Value := CDSParcelasNcheque.Value;
          DM_tabelas.ZQCheque.FieldByName('valor').Value := CDSParcelasVrparc.Value;
          DM_tabelas.ZQCheque.FieldByName('emissao').Value := strtodate(JDEntrada.DateText);
          DM_Tabelas.ZQCheque.FieldByName('conta').Value := CDSParcelasConta.Value;
          DM_tabelas.ZQCheque.FieldByName('deposito').Value := CDSParcelasVenci.Value;
          DM_Tabelas.ZQCheque.FieldByName('idrecebimento').Value := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value;
//          DM_Tabelas.ZQChequeidvinculo.Value := DM_Tabelas.ZQRecebimentoidrecebimento.Value;
          DM_Tabelas.ZQCheque.FieldByName('idvinculo').Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Value;
          DM_Tabelas.ZQCheque.FieldByName('Substituicao').Value := 'S';
          DM_tabelas.ZQCheque.FieldByName('sq').Value := contador;
          DM_tabelas.ZQCheque.Post;

        end;
      end;
      if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value = 'S' Then
      Begin
        quadra:=CDSParcelasquadralote.Value;
        if (CBTipobaixa.Text<>'SUBSTITUI��O') then
        begin
          DM_Tabelas.ZQMovBancaria.Insert;
          DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
          DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
          DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
          DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value:=CDSParcelassubstituicao.Value;
          DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
          DM_Tabelas.ZQMovBancaria.Post;
          //      CalcSaldo('','',False);
        end
        else if (CBTipobaixa.Text='SUBSTITUI��O') then
        begin
          DM_Tabelas.ZQMovBancaria.Insert;
          DM_Tabelas.ZQMovBancaria.FieldByName('conta_bancaria_cod_banco').Value := DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('plano_contas_codigo').Value := CDSParcelascontabanco.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('documento').Value := CDSParcelasDocum.Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('dt_lanc').Value := date;
          DM_tabelas.ZQMovBancaria.FieldByName('hora_lanc').Value := Time;
          DM_Tabelas.ZQMovBancaria.FieldByName('contracodigo').Value := strtoint(Ecodnome.Text);
          DM_TAbelas.ZQMovBancaria.FieldByName('dt_conciliado').Value := CDSParcelasVenci.Value;
          DM_TAbelas.ZQMovBancaria.FieldByName('vr_lanc').Value := CDSParcelasVrParc.Value;
          DM_Tabelas.ZQPlanodeContas.Locate('codigo',CDSParcelascontabanco.Value,[]);
          DM_TAbelas.ZQMovBancaria.FieldByName('credeb').Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').Value;
          DM_Tabelas.ZQMovBancaria.FieldByName('obs').Value := EMObsbx.Text;
          DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'R'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          IF ZQRecBai.FieldByName('recpag').Value = 'P' Then
            DM_tabelas.ZQMovBancaria.FieldByName('vinculo').Value := 'P'+DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
          DM_Tabelas.ZQMovBancaria.FieldByName('Substituicao').Value := 'S';
          DM_tabelas.ZQMovBancaria.FieldByName('sq').Value := contador;
          DM_Tabelas.ZQMovBancaria.Post;

          //      CalcSaldo('','',False);
        end;
      end;
      CDSParcelas.Next;
    end;
    bar1.position:=0;
    bar1.visible:=false;
    CDSParcelas.EnableControls;
    DM_tabelas.contador_bx.insert;
    DM_tabelas.contador_bx.FieldByName('idcontador_bx').Value:=contador;
    DM_tabelas.contador_bx.Post;
    DM_tabelas.contador_bx.close;

    if not empty(varnumordem1) Then Begin
      ZQDiv_usado.SQL.Clear;
      ZQDiv_usado.SQL.Add('select iddividido,participante,percentual,planodeconta,tipodocumento,ordem from dividido');
      ZQDiv_usado.SQL.Add(' where  ordem = '+quotedstr(ZQRecBai.FieldByName('numordem').Text));
      ZQDiv_usado.Open;
      varsomarepasseautomatico := 0;
      ZQDiv_usado.DisableControls;
      while not ZQDiv_usado.Eof do begin
        DM_Tabelas.ZQDividido.Insert;
        DM_Tabelas.ZQDividido.FieldByName('participante').Value := ZQDiv_usado.FieldByName('participante').Value;
        DM_Tabelas.ZQDividido.FieldByName('percentual').Value := ZQDiv_usado.FieldByName('percentual').Value;
        DM_Tabelas.ZQDividido.FieldByName('planodeconta').Value := ZQDiv_usado.FieldByName('planodeconta').Value;
        DM_Tabelas.ZQDividido.FieldByName('tipodocumento').Value := ZQDiv_usado.FieldByName('tipodocumento').Value;
        DM_Tabelas.ZQDividido.FieldByName('ordem').Value := strtoint(varnumordem1);
        DM_Tabelas.ZQDividido.Post;
        ZQDiv_usado.Next;
      end;
    end;
    ZQDiv_usado.EnableControls;
    ZQRecBai.First;
    varpago := 0;
    Memo1.Clear;
    baixatot := XNERecebido.Value;
    ZqRecBai.DisableControls;
    While not ZqRecBai.Eof do
    Begin
      CDSParcelas.Filtered:=false;
      CDSParcelas.Filter:='idrecebimento='+quotedstr(alltrim(ZQRecBai.FieldByName('idrecebimento').Text));
      CDSParcelas.Filtered:=true;
      CDSParcelas.RecordCount;

      DM_tabelas.ZQReBxHi.Insert;
//        percentbaixa := ZQRecBaisaldocorrig.Value * 100 / XNEMarcado.Value;


      DM_Tabelas.ZQReBxHi.FieldByName('descon').Value := ZQRecBai.FieldByName('descontos').Value;//XNEDesconto.Value * percentbaixa/100;      17/05/2012

      if ZQRecBai.FieldByName('saldocorrig').Value<=baixatot Then
         varpago := ZQRecBai.FieldByName('saldocorrig').Value
      else
         varpago := baixatot+DM_Tabelas.ZQReBxHi.FieldByName('descon').Value;

      varjuros := XNEJuros.Value* percentbaixa/100;
  //      if varjuros <=varpago Then
      if ZQRecBai.FieldByName('tip').Value<>'J' then
         DM_Tabelas.ZQReBxHi.FieldByName('juros').Value := ZQRecBai.FieldByName('juros').Value //varjuros   17/05/2012
      else
         DM_Tabelas.ZQReBxHi.FieldByName('juros').value:= ZQRecBai.FieldByName('saldo').Value;
    //    else
      //    DM_Tabelas.ZQReBxHijuros.Value := varpago;
      if ZQRecBai.FieldByName('tip').Value<>'J' then
         varpago := varpago + DM_Tabelas.ZQReBxHi.FieldByName('juros').Value;

      DM_Tabelas.ZQReBxHi.FieldByName('valor').Value := (ZQRecBai.FieldByName('saldo').Value+ZQRecBai.FieldByName('juros').Value)-ZQRecBai.FieldByName('descontos').Value;
      if CDSParcelasData_Quitacao.Value>0 then
         DM_Tabelas.ZQReBxHi.FieldByName('data').Value := CDSParcelasData_Quitacao.Value
      else
         DM_Tabelas.ZQReBxHi.FieldByName('data').Value := strtodate(JDEntrada.DateText); //JDBaixa.Date;
      DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').Value := 100;
      DM_tabelas.ZQReBxHi.FieldByName('refer').Value := CDSParcelascontagem.Value; //DM_tabelas.ZQCobaRecotagem.Value;
      DM_tabelas.ZQReBxHi.FieldByName('idrecib').Value := ZQRecBai.FieldByName('idrecebimento').Value;
      DM_tabelas.ZQReBxHi.FieldByName('sq').Value := contador;
//        DM_tabelas.ZQReBxHiValor_parcela.Value :=(CDSParcelasVrParc.Value+XNEResto.Value+XNEDesconto.Value-XNEJuros.value);
      varx := ExRound(varpago,2)-ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').Value,2);
      if ZQRecBai.FieldByName('tip').Value<>'J' then
         DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').Value :=ZQRecBai.FieldByName('saldo').Value
      else
         DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').Value :=0;
      DM_Tabelas.ZQReBxHi.Post;

      ZQRecBai.Edit;
      varx := ExRound(varpago,2)-ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').Value,2);
//     varx := varx + DM_Tabelas.ZQReBxHidescon.Value;
//    ZQRecBaisaldo.Value := truncar(ZQRecBaisaldo.Value - varx,2);
      ZQRecBai.FieldByName('saldo').Value := ExRound(ZQRecBai.FieldByName('saldo').Value,2) - ExRound(((DM_Tabelas.ZQReBxHi.FieldByName('valor').Value-ZQRecBai.FieldByName('juros').Value)+ZQRecBai.FieldByName('descontos').Value),2);
      ZQRecBai.FieldByName('RefBaixa').Value := CDSParcelascontagem.Value;//DM_tabelas.ZQCobaRecotagem.Value;

      // 17/05/2012
      ZQRecBai.FieldByName('juros').Value:=0;
      ZQRecBai.FieldByName('descontos').Value:=0;
      ZQRecBai.post;

      // 22/09/2010
      DM_Tabelas.ZQRecebimento.Locate('idrecebimento',ZQRecBai.FieldByName('idrecebimento').Value,[]);
      if DM_Tabelas.ZQRecebimento.recordcount>0 then
      begin
        quadra:= DM_Tabelas.ZQRecebimento.FieldByName('quadralote').Value;
        DM_TAbelas.ZQRecebimento.Edit;
        if ZQRecBai.FieldByName('saldo').Value>0 then
           DM_Tabelas.ZQRecebimento.FieldByName('tip').Value:='P';

        if (ZQRecBai.FieldByName('saldo').Value=0) and (DM_Tabelas.ZQRecebimento.FieldByName('tip').Value='P') then
           DM_Tabelas.ZQRecebimento.FieldByName('tip').Value:='';

        if (DM_Tabelas.ZQRecebimento.FieldByName('tip').Value='J') then
           DM_Tabelas.ZQRecebimento.FieldByName('Valor').Value := ZQRecBai.FieldByName('saldo').Value;

        DM_Tabelas.ZQRecebimento.FieldByName('saldo').Value := ZQRecBai.FieldByName('saldo').Value;
        DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').Value := ZQRecBai.FieldByName('RefBaixa').Value;
        DM_Tabelas.ZQRecebimento.Post;
      end;
      // at?aqui


      if (CBTipobaixa.Text='SUBSTITUI��O') then
      begin
//      ZQRecBaisubstituicao.Value:='S';
        ZQRecBai.edit;
        ZQRecBai.FieldByName('somar').Value:='N';
        ZQRecBai.post;
//        ZQRecBaisq.Value := contador;
      end;

      if pos('-',ZQRecBai.FieldByName('ordem').Value)=0 Then
        Varordem := Varordem + ZQRecBai.FieldByName('idrecebimento').Text+' - ';
      if pos(ZQRecBai.FieldByName('nomeadversa').Value,Memo1.Text)=0 Then
        Memo1.Lines.Add(ZQRecBai.FieldByName('nomeadversa').Value);
      ZqRecBai.Next;
    end;
    CDSParcelas.Filtered:=false;
    ZqRecBai.EnableControls;

    CDSParcelas.First;
    CDSParcelas.DisableControls;
    while not CDSParcelas.Eof do
    begin
      if not empty(CDSParcelascontagem.Text) then
      begin
        CDS_dividido.Close;
        CDS_dividido.CreateDataSet;
        ZQRatear.SQL.Clear;
        ZQRatear.SQL.Add('select idrecebimento, h.valor, percent_usado, b.vr_rec, participante, percentual,planodeconta, dv.tipodocumento, re.ordem, dv.ordem from recbxhist as H');
        ZQRatear.SQL.Add(' join recebimento as re ON H.idrecib=RE.idrecebimento');
        ZQRatear.SQL.Add(' join dividido as dv ON dv.ordem=numordem');
        ZQRatear.SQL.Add(' join receb_baixa as B ON B.refbaixa=H.refer');
        ZQRatear.SQL.Add(' join tipodocumento as D ON B.tipdoc=D.tipodoc');
        ZQRatear.SQL.Add(' where somapaga = '+quotedstr('S')+' and refer = '+CDSParcelascontagem.Text);
        ZQRatear.Open;
        varsomarepasseautomatico := 0;
        while not ZQRatear.Eof do
        begin
          if CDS_dividido.Locate('participante;contabil;tipodoc',VarArrayOf([ZQRatear.FieldByName('participante').Value,ZQRatear.FieldByName('planodeconta').Value,ZQRatear.FieldByName('tipodocumento').Value]),[]) then
            CDS_dividido.Edit
          else
            CDS_dividido.Insert;
          CDS_divididoparticipante.Value := ZQRatear.FieldByName('participante').Value;
          CDS_divididovr_base.Value := CDS_divididovr_base.Value + ZQRatear.FieldByName('vr_base').Value;
          CDS_divididovr_calc.Value := CDS_divididovr_calc.Value + ZQRatear.FieldByName('vr_divi').Value;
          CDS_divididocontabil.Value := ZQRatear.FieldByName('planodeconta').Value;
          CDS_divididotipodoc.Value := ZQRatear.FieldByName('tipodocumento').Value;
          CDS_divididoobs.Value := CDS_divididoobs.Value+ZQRatear.FieldByName('idrecebimento').Text+' | '+ZQRatear.FieldByName('ordem').Value+' | '+ZQRatear.FieldByName('vr_base').Text+' x '+ZQRatear.FieldByName('percentual').Text+' = '+ZQRatear.FieldByName('vr_divi').Text+chr(13);
          varsomarepasseautomatico := varsomarepasseautomatico + ZQRatear.FieldByName('vr_divi').Value;
          ZQRatear.Next;
        end;
        CDS_dividido.First;
        if CDS_dividido.RecordCount>0 Then
        Begin
          DM_Tabelas.ZQNumOrdem.Insert;
          DM_Tabelas.ZQNumOrdem.Post;
        end;
        while not CDS_dividido.Eof do
        begin
          DM_tabelas.ZQRecebimento.Insert;
          DM_tabelas.ZQRecebimento.FieldByName('documento').Value := 'Rateio Bx - '+CDSParcelascontagem.Text;
          DM_tabelas.ZQRecebimento.FieldByName('cliente').Value :=  CDS_divididoparticipante.Value;
          if CDSParcelasData_Quitacao.Value>0 then
          begin
             DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := CDSParcelasData_Quitacao.Value;
             DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := CDSParcelasData_Quitacao.Value;
          end
          else
          begin
             DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').Value := strtodate(JDEntrada.DateText);
             DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').Value := strtodate(JDEntrada.DateText);
          end;
          DM_tabelas.ZQRecebimento.FieldByName('contabil').Value := CDS_divididocontabil.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Observ').Value := EMObsbx.text+chr(13)+CDS_divididoobs.Value; //EMNovostit.Text+chr(13)+CDS_divididoobs.Value 24/12/2012; erika pediu
          DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value := 'P';
          DM_tabelas.ZQRecebimento.FieldByName('TipDoc').Value := CDS_divididotipodoc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('ordem').Value := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text+'-'+inttostr(CDS_dividido.RecNo)+'/'+inttostr(CDS_dividido.RecordCount);
          DM_tabelas.ZQRecebimento.FieldByName('refvinda').Value := CDSParcelascontagem.Value;
          DM_tabelas.ZQRecebimento.FieldByName('VrDoc').Value :=  CDS_divididovr_base.Value;
          DM_tabelas.ZQRecebimento.FieldByName('Valor').Value :=  CDS_divididovr_calc.Value;
          DM_tabelas.ZQRecebimento.FieldByName('saldo').Value :=  DM_tabelas.ZQRecebimento.FieldByName('Valor').Value;
          DM_Tabelas.ZQRecebimento.FieldByName('adversa').Value := DM_Tabelas.ZQConfiguracoes.FieldByName('paticipantedefault').Value;

          DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').Value := EAdversaNova.text;

          DM_tabelas.ZQRecebimento.Post;
          CDS_dividido.Next;
        end;
      end;
      CDSParcelas.next;
    end;
    CDSParcelas.EnableControls;
    ///


    FrmImpRecibo:=nil;
    if FrmImpRecibo=nil then
       FrmImpRecibo:=TFrmImpRecibo.Create(self);
    FrmImpRecibo.N_Baixa.Text := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
    FrmImpRecibo.Nomecli.Text := Enomes.Text;
    FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;


    ZQRecBai.Filtered := False;
    ZQRecBai.Refresh;
    DBGBaixando.Refresh;
    DM_tabelas.ZQReceb_Baixa.Refresh;
    DM_Tabelas.ZQReBxHi.Refresh;
    CDSParcelas.Close;
    CDSParcelas.CreateDataSet;
    XNEJuros.Value:=0;
    XNERecebido.Value:=0;
    XNEDesconto.Value:=0;
    XNEResto.Value:=0;
    XNEParcelas.Value:=1;
    CBPeriodoEnt.Text:='HOJE';
    EMObsbx.Clear;
    EMNovostit.Clear;



////    DM_tabelas.ZQRecebimento.Refresh;

    Panel1.Caption:='Aguarde! Ajustando Saldo da Venda...';
    Panel1.Visible:=true;
    application.ProcessMessages;
    ZQRecebimento.close;
    ZQRecebimento.sql.clear;
    ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    ZQRecebimento.sql.add(' from recebimento where quadralote='+quotedstr(quadra)+' group by quadralote order by quadralote');
    ZQRecebimento.Open;
    ZQRecebimento.First;
    bar1.Position := 0;
    bar1.Max := ZQRecebimento.RecordCount;
    bar1.visible:=true;
    application.ProcessMessages;
    ZQVenda.open;
    ZQRecebimento.DisableControls;
    while not ZQRecebimento.Eof do
    begin
      bar1.Position := ZQRecebimento.RecNo;
      application.ProcessMessages;
      ZQEntrada.Close;
      ZQEntrada.SQL.Clear;
      ZQEntrada.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQEntrada.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQEntrada.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      ZQEntrada.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
      ZQEntrada.SQL.Add(', sum(valor) as entrada from recebimento where quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').Value)+' and documento like  ''%-E-%''  order by ordem');
      ZQEntrada.open;

      ZQParcela.Close;
      ZQParcela.SQL.Clear;
      ZQParcela.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQParcela.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQParcela.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
      ZQParcela.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
      ZQParcela.SQL.Add(', sum(valor) as parcela from recebimento where quadralote='+quotedstr(ZQRecebimento.FieldByName('quadralote').Value)+' and documento like  ''%-P-%''  order by ordem');
      ZQParcela.open;
      if (ZQEntrada.FieldByName('entrada').Value>0) or (ZQParcela.FieldByName('parcela').Value>0) then
      begin
        ZQVenda.SQL.Clear;
        ZQVenda.SQL.Add('Select idvenda,datavenda,imovel,valorvenda,forma_reajuste,tabela_Price,Escriturado,marca,codigo_contrato_ref,Multa,Mora,Perc_comissao,vlr_comissao from venda where idvenda='+quotedstr(ZQRecebimento.FieldByName('venda_idvenda').Text));
        ZQVenda.open;
        if ZQVenda.RecordCount>0 then
        begin
          ZQVenda.Edit;
          ZQVenda.FieldByName('valorvenda').Value:=ZQEntrada.FieldByName('entrada').Value+ZQParcela.FieldByName('parcela').Value;
          ZQVenda.Post;
        end;

      end;
      ZQRecebimento.Next;
    end;
    ZQRecebimento.EnableControls;
    bar1.Position := 0;
    bar1.visible:=false;
    ZQVenda.close;
    ZQRecebimento.close;
    ZQEntrada.close;
    ZQParcela.close;
    Panel1.Caption:='AGUARDE! ATUALIZANDO A TELA.   ';
    Panel1.Visible:=false;
//    CBTipobaixa.SetFocus;
    Elote.SetFocus;
    DXBFechar.Enabled := true;
    application.ProcessMessages;



  //  Close;
     // imprime o recibo
    //  FrmImpRecibo.showmodal;
  ///
  {  if perguntaSN('Gerar Recibo de Quita��o ?','S') Then Begin
      FrmRelRecibodeQuita.ZQForma.SQL.Clear;
      FrmRelRecibodeQuita.ZQForma.SQL.Add('Select * from receb_baixa where refbaixa='+DM_tabelas.ZQCobaRecotagem.Text);
      FrmRelRecibodeQuita.ZQForma.Open;
      FrmRelRecibodeQuita.RLLabel19.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
      FrmRelRecibodeQuita.RLLabel54.Caption := Transform(XNERecebido.Value,'###,###,###,##0.00');
      FrmRelRecibodeQuita.RLMemo1.Lines := EMObsbx.Lines;
      FrmRelRecibodeQuita.RLMemo2.Lines := Memo1.Lines;
      FrmRelRecibodeQuita.RLMemo3.Lines := FrmRelRecibodeQuita.RLMemo2.Lines;
      FrmRelRecibodeQuita.RLLabel18.Caption := DM_tabelas.ZQCobaRecotagem.Text;
      FrmRelRecibodeQuita.RLLabel39.Caption := DM_tabelas.ZQCobaRecotagem.Text;
      ZQItensRecibo.Close;
      ZQItensRecibo.SQL[1] := 'where refer='+DM_tabelas.ZQCobaRecotagem.Text;
      ZQItensRecibo.Open;
      if FrmRelRecibodeQuita=nil then
        FrmRelRecibodeQuita:=TFrmRelRecibodeQuita.Create(Application);
      FrmRelRecibodeQuita.RLReport1.PreviewModal;
      FreeAndNil(FrmRelRecibodeQuita);
    end; }


    ZQRecBai.Filtered := False;
//    DM_tabelas.ZQRecebimento.Refresh;
    DM_Tabelas.ZQRecebimento.close;
    DM_Tabelas.ZQRecebimento.SQL.clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ');
    DM_tabelas.ZQRecebimento.SQL.Add(' from Recebimento where quadralote='+quotedstr(quadra)+' order by DT_Vencimento');
  //  showmessage(DM_tabelas.ZQRecebimento.SQL.Text);
    DM_tabelas.ZQRecebimento.Open;

//    CBTipobaixa.SetFocus;
    Elote.SetFocus;
    DXBFechar.Enabled := true;
    mensagem('Opera��o Terminada!');

  end;
  if FrmImpRecibo<>nil then
     FreeAndNil(FrmImpRecibo);
end;

procedure TFrmRecebBaixa.DBCBTIPDOCExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then
  Begin
    if Empty(DBCBTIPDOC.Text) Then
    Begin
      DBCBTIPDOC.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',DBCBTipDoc.Text,[]);
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').Value='S' Then
    Begin
      if CDSParcelasData_Quitacao.Value>0 then
      begin
        if (CDSParcelasVenci.Value<>CDSParcelasData_Quitacao.Value) and ((CDSParcelasTipDoc.Value<>'DP')) Then
        BEgin
          showmessage('Este tipo de documento s?aceita pagamento a vista!!!');
          EContabil.SetFocus;
          Exit;
        end;
      end
      else
      begin
        if (datetostr(CDSParcelasVenci.Value)<>JDEntrada.DateText) and ((CDSParcelasTipDoc.Value<>'DP')) Then
        BEgin
          showmessage('Este tipo de documento s?aceita pagamento a vista!!!');
          EContabil.SetFocus;
          Exit;
        end;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').Value='S' then
    Begin
      GBBanco.Visible := True;
      DBGBanco.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value='S' then Begin


       ZQCheque.close;
       ZQCheque.SQL.Clear;
       ZQCheque.SQL.Add('select ch.Banco,ch.Dono,ch.CPF_CNPJ,ch.valor,ch.emissao,ch.deposito,ch.alias,ch.prorrogado,ch.conta,ch.idrecebimento,ch.idpagamento,ch.idcheque,ch.idvinculo,ch.numero,ch.agencia,ch.Substituicao,ch.sq,');
       ZQCheque.SQL.Add('       rc.idrecebimento,rc.documento,rc.cliente,rc.usuario,rc.Dt_Entrada,rc.Dt_Vencimento,rc.Valor,rc.Observ,rc.VrDoc,rc.ordem,rc.TipDoc,rc.saldo,rc.marcar,');
       ZQCheque.SQL.Add('       rc.RefBaixa,rc.refvinda,rc.contabil,rc.empresa,rc.custodaparcela,rc.origem,rc.adversa,rc.recpag,rc.numordem,rc.idloteamento,rc.venda_idvenda,rc.quadralote,');
       ZQCheque.SQL.Add('       rc.numboleto,rc.Substituicao,rc.sq,rc.somar,rc.nomeadversa,rc.Reajustado,rc.Data_reajuste ');
       ZQCheque.SQL.Add(' from cheque_rec as ch join recebimento as rc on ch.idrecebimento=rc.idrecebimento');
       ZQCheque.SQL.Add(' where rc.idrecebimento= '+quotedstr(ZQRecBai.FieldByName('idrecebimento').text)+' order by numero');
       ZQCheque.Open;
       ZQCheque.First;
       if ZQCheque.RecordCount>0 then
       begin
         CDSParcelasagencia.Value := ZQCheque.FieldByName('agencia').Value;
         CDSParcelasConta.Value   := ZQCheque.FieldByName('conta').Value;
         CDSParcelasbanco.Value   := ZQCheque.FieldByName('Banco').Value;
         CDSParcelasdono.Value    := ZQCheque.FieldByName('Dono').Value;
         CDSParcelasncheque.Value := ZQCheque.FieldByName('numero').Value;
         CDSParcelasCpf_Cnpj.Value:= ZQCheque.FieldByName('CPF_CNPJ').Value; 
       end;
       ZQCheque.close;       
       GBCheque.Visible := True;
       DBEBanco.SetFocus;
       exit; 
    end;
  end;
end;

procedure TFrmRecebBaixa.DXBBaixaGravarEnter(Sender: TObject);
begin
  CDSParcelas.First;
  While not CDSParcelas.Eof do Begin
    if CDSParcelascodContabil.Value=0 Then Begin
      EContabil.SetFocus;
      Exit;
    end;
    if empty(CDSParcelasTipDoc.Value) Then Begin
      DBCBTIPDOC.SetFocus;
      Exit;
    end;
    CDSParcelas.Next;
  End;
end;

procedure TFrmRecebBaixa.DBCBTIPDOCEnter(Sender: TObject);
Var
  Varcheque,VarTipDoc, VarDocum, VarBanco, Varconta, VarAgencia, VarCPF_CNPJ, VarDono : string;
begin
  if CDSParcelas.Recno>1  Then
  begin
    CDSParcelas.Prior;
    VarTipDoc := CDSParcelasTipdoc.Value;
    VarDocum := CDSParcelasDocum.Value;
    VarBanco := CDSParcelasBanco.Value;
    VarAgencia := CDSParcelasagencia.Value;
    varconta := CDSParcelasconta.Value;
    VarCPF_CNPJ := CDSParcelascpf_cnpj.Value;
    VarDono := CDSParcelasDono.Value;
    VarCheque :=CDSParcelasNcheque.Value;
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
      CDSParcelasConta.Value := VarConta;
    if Empty(CDSParcelasNCheque.Value) then
      CDSParcelasNCheque.Value := VarCheque;

  end
  else
  Begin
    CDSParcelas.Edit;
    CDSParcelasDono.Value := Enomes.Text;
  End;
end;

procedure TFrmRecebBaixa.DBEBaixaVrRecExit(Sender: TObject);
Var
  varregis : Integer;
  varsoma, vardif, varjur, vardesc : Double;
begin
  if Empty(WDTBaixaEntra.Text) Then
    CDSParcelasVenci.Value := date;
  varsoma := 0;
  varjur := XNEJuros.Value;
  vardesc := XNEDesconto.Value;
  varregis := CDSParcelas.RecNo;
  CDSParcelas.First;
  while not (CDSParcelas.Eof) and (CDSParcelas.RecNo <= varregis) do
  Begin
    VarDoc := VarDoc + quotedstr(CDSParcelasDocum.Value)+',';
    varsoma := varsoma +CDSParcelasVrParc.Value;
    if CDSParcelasjuros.Value=0 then
    begin
      CDSParcelas.Edit;
      if CDSParcelasVrParc.Value>=varjur then
        CDSParcelasjuros.Value := varjur
      else
        CDSParcelasjuros.Value := CDSParcelasVrParc.Value;
      varjur := varjur - CDSParcelasjuros.Value;
      if CDSParcelasVrParc.Value>= vardesc Then
        CDSParcelasdesconto.Value := vardesc
      else
        CDSParcelasdesconto.Value := CDSParcelasVrParc.Value;
      vardesc := vardesc - CDSParcelasdesconto.Value;
      CDSParcelas.Post;
    end;
    CDSParcelas.next;
  end;
  VarDoc := copy(VarDoc,1,length(VarDoc)-1);
  CDSParcelas.RecNo := varregis;
  vardif := ExRound(XNERecebido.Value,2) - ExRound(varsoma,2);
  if CDSParcelas.RecNO< XNEParcelas.Value Then
  Begin
    CDSParcelas.RecNO := varregis+1;
    while not CDSParcelas.Eof do Begin
      CDSParcelas.Edit;
      CDSParcelasVrParc.Value := ExRound(vardif,2) / (ExRound(XNEParcelas.Value,2)-ExRound(varregis,2));
      vardesc := vardesc - CDSParcelasdesconto.Value;
      CDSParcelas.Next;
    end;
  end;
  CDSParcelas.RecNo := varregis;
  DXBBaixaGravar.Enabled := True;
  DXBBaixaGravar.SetFocus;
end;

procedure TFrmRecebBaixa.EContabilExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then
  Begin
    if empty(Econtabil.text) Then
    Begin
      showmessage('O Descri��o contabil nao pode ficar em branco....');
      Econtabil.SetFocus;
      exit;
    end;
    if not AchaPlanoDeConta(450, 30, 'S', EContabil.Text, '') then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
{      if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
        CDSParcelas.Edit;
      CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContascodigo.Value;
      CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContasdoccomum.Value;
  }
    CDSParcelas.First;
    while not CDSParcelas.Eof do
    begin
      if not (CDSParcelas.State in [DsInsert, DsEdit]) Then
        CDSParcelas.Edit;
      CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
      if not empty(DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value) then
         CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').Value;
      CDSParcelas.post;
      CDSParcelas.Next;
    end;
    CDSParcelas.First;

  end;
//    CDSParcelasTipDoc.Value := FrmAchaPlanodeConta.ZQTempPlanoContasdoccomum.Value;
end;

procedure TFrmRecebBaixa.XDBNumEdit1Exit(Sender: TObject);
begin
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
     CDSParcelasVrParc.Value := XNERecebido.Value*CDSParcelasPercent.Value/100;
  end;   
end;

procedure TFrmRecebBaixa.DBEBaixaDocumExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
{    if (not empty(DBEBaixaDocum.Text)) and (DM_tabelas.ZQRecebimento.Locate('documento',DBEBaixaDocum.Text,[])) Then Begin
      showmessage('Este documento ja foi lan?do anteriormente... corriga a numera��o.....');
      DBEBaixaDocum.SetFocus;
      exit;
    End;
    if DM_tabelas.ZQCheque.Locate('CH_Conta',DBEBaixaDocum.Text,[]) Then Begin
      showmessage('Este cheque ja foi lan?do anteriormente... corriga a numera��o.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (pos(quotedstr(CDSParcelasDocum.Value),VarDoc)>0) Then Begin
      showmessage('Este cheque ja foi lan?do nesta inclus?... corriga a numera��o.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (empty(CDSParcelasDocum.Value)) Then Begin
      showmessage('o campo N?do cheque deve ser preenchido...');
      DBEBaixaDocum.SetFocus;
      exit;
    end;}
  end;
end;

procedure TFrmRecebBaixa.EContabilEnter(Sender: TObject);
begin
  if empty(Econtabil.Text) Then Begin
    Econtabil.Text := '9';//DM_Tabelas.ZQConfiguracoesrecebe_planodecontas.Text;
  end;
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;
  DBText1.Caption := CDSParcelasOrdem.Value;
end;

procedure TFrmRecebBaixa.DBECpf_CnpjExit(Sender: TObject);
Var
  variavel : string;
  varx : integer;
begin
  If not DXBFechar.Focused Then Begin
    if not empty(CDSParcelascpf_cnpj.Value) then Begin
      variavel := '';
      for varx :=1 to length(CDSParcelascpf_cnpj.Value) do
        if pos(copy(CDSParcelascpf_cnpj.Value,varx,1),'0123456789')>0 then
          variavel := variavel + copy(CDSParcelascpf_cnpj.Value,varx,1);
      if not Verif_doc(variavel,True) then
        DBECpf_Cnpj.SetFocus
    end;
  end;
end;

procedure TFrmRecebBaixa.EnomesExit(Sender: TObject);
begin
  If not DXBFechar.Focused Then Begin
    if not AchaCLiente(450, 140,Enomes.Text) Then Begin
      Enomes.SetFocus;
      Exit;
    End;
    Ecodnome.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
    Enomes.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
  End;
end;

procedure TFrmRecebBaixa.Label1Click(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
end;

procedure TFrmRecebBaixa.RBChequeClick(Sender: TObject);
begin
  Label1.Visible := true;
  DTPIni.Visible := true;
  Label23.Visible := true;
  DTPFim.Visible := true;
  Enomes.Visible := False;
end;

procedure TFrmRecebBaixa.GBBancoExit(Sender: TObject);
begin
  GBBanco.Visible := False;
  if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value='S' then Begin
    CDSParcelasagencia.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').Value;
    CDSParcelasConta.Value := DM_TAbelas.ZQContaBancaria.FieldByName('n_conta').Value;
    CDSParcelasbanco.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').Value;
    CDSParcelasdono.Value := DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').Value;
    GBCheque.Visible := True;
    DBEBanco.SetFocus;
    exit;
  end;
end;

procedure TFrmRecebBaixa.GBChequeExit(Sender: TObject);
begin
  GBCheque.Visible := False;
end;

procedure TFrmRecebBaixa.EPlanocontabancoExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not AchaPlanoDeConta(450, 30, 'S', EPlanocontabanco.Text, '') then Begin
      EPlanocontabanco.SetFocus;
      exit;
    end;
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').Value;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').Value;
  end;
end;

procedure TFrmRecebBaixa.filtrabaixa;
var
  vartipo : string;
  posi:integer;
Begin
  Label37.Visible:=true;
  pini.Visible:=true;
  Label38.Visible:=true;
  pfim.Visible:=true;
  if ( not empty(alltrim(Elote.Text))) then
  begin
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
       if (not empty(Elote.text)) and  ( not empty(varrecpag)) and ( not empty(Ecodnome.Text)) then
       begin
         ZQRecBai.SQL.Add('update recebimento set marcar= CASE WHEN Dt_Vencimento<=now() then ''0'' else ''1'' end  where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+Ecodnome.Text+')');
         vazio:='F';
       end
       else
       begin
         showmessage('Dados Nao Encontrados!');
         vazio:='T';
         exit;
       end;  
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
       end;
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and (dt_vencimento between :dt1 and :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
      // ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
    end
    else if RBCheque.Checked then
    Begin
      Label37.Visible:=false;
      pini.Visible:=false;
      pini.Clear;
      Label38.Visible:=false;
      pfim.Visible:=false;
      pfim.Clear;
      vartipo := '';
      DM_TAbelas.ZQTipodoc.First;
      DM_TAbelas.ZQTipodoc.DisableControls;
      while not DM_Tabelas.ZQTipodoc.Eof do begin
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      DM_TAbelas.ZQTipodoc.EnableControls;
      vartipo := copy(vartipo,1,length(vartipo)-1);
      ZQRecBai.SQL.Add('update recebimento set marcar= CASE WHEN Dt_Vencimento<=now() then ''0'' else ''1'' end  where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null))and (dt_vencimento between :dt5 and :dt6)');
      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
//      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
    end;
    ZQRecBai.ExecSQL;
////////////////////////////////////






    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
       if (not empty(Elote.text)) and  ( not empty(varrecpag)) and ( not empty(Ecodnome.Text)) then
       begin
         ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
         ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
         ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
         ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+Ecodnome.Text+')');
         vazio:='F';
       end
       else
       begin
         showmessage('Dados Nao Encontrados!');
         vazio:='T';
         exit;
       end;
       if not empty(EcodAdversa.text) Then
       Begin
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
       end;
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and (dt_vencimento between :dt1 and :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
    end
    else if RBCheque.Checked then
    Begin
      Label37.Visible:=false;
      pini.Visible:=false;
      pini.Clear;
      Label38.Visible:=false;
      pfim.Visible:=false;
      pfim.Clear;
      vartipo := '';
      DM_TAbelas.ZQTipodoc.First;
      DM_TAbelas.ZQTipodoc.DisableControls;
      while not DM_Tabelas.ZQTipodoc.Eof do begin
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
        DM_Tabelas.ZQTipodoc.Next;
      end;
      DM_TAbelas.ZQTipodoc.EnableControls;
      vartipo := copy(vartipo,1,length(vartipo)-1);
      ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
      ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and ( tipdoc in ('+vartipo+')) and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (dt_vencimento between :dt5 and :dt6)');
      ZQRecBai.SQL.Add(' and adversa = '+quotedstr(EcodAdversa.Text));
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
      ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
    end;
    ZQRecBai.Open;
    XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').Value;
    Enomes.Text      := ZQRecBai.FieldByName('Nomecli').Value;
    Ecodnome.Text    := ZQRecBai.FieldByName('cliente').Text;
    LREg.Caption     := inttostr(ZQREcbai.RecordCount)+' Registros';
    if RBParticipante.Checked Then
{    ZQRecBai.DisableControls;
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
      if ZQRecBaiDt_Vencimento.Value<=date Then
        ZQRecBaimarcar.Value := '0'
      else
        ZQRecBaimarcar.Value := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
    ZQRecBai.enableControls;    }



    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value,[]);
    if RBPonteiro.Checked Then Begin
       ZQRecBai.Edit;
       ZQRecBai.FieldByName('marcar').Value := '0';
       ZQRecBai.Post;
    end;
    ZQRecBai.first;
  end
  else
  begin
    ZQRecBai.SQL.Clear;
    if (RBParticipante.Checked) or (RBPonteiro.Checked) Then
    begin
//       ZQRecBai.SQL.Add('select * from recebimento where ( quadralote='+quotedstr(alltrim(Equadra.text))+'-'+quotedstr(alltrim(Elote.Text))+') and (recpag='+quotedstr(varrecpag)+') and (saldo>0) and  (not saldo is null));
       ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
       ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
       ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
       ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
       if not empty(EcodAdversa.text) Then
       Begin
          ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
       end;
       if not empty(EDLoteamento.Text) Then
       begin
         posi:=pos('-',EDLoteamento.Text);
         dec(posi);
         ZQRecBai.SQL.Add(' and (idloteamento = '+quotedstr(copy(alltrim(EDLoteamento.Text),1,posi))+')');
       end;
       if pfim.Text<>'  /  /    ' then
       begin
          ZQRecBai.SQL.Add(' and (dt_vencimento between :dt1 and :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
       end;
       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
       if RBCheque.Checked then
       Begin
         Label37.Visible:=false;
         pini.Visible:=false;
         pini.Clear;
         Label38.Visible:=false;
         pfim.Visible:=false;
         pfim.Clear;

         vartipo := '';
         DM_TAbelas.ZQTipodoc.First;
         while not DM_Tabelas.ZQTipodoc.Eof do
         begin
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
           DM_Tabelas.ZQTipodoc.Next;
         end;
         vartipo := copy(vartipo,1,length(vartipo)-1);
         ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
         ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
         ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
         ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
         ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
         ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
         ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
       end;
    end;
    ZQRecBai.Open;
   // XEEmpresend.Text := ZQRecBainomeempreend.Value;
    LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
    if RBParticipante.Checked Then
    begin
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('update recebimento set marcar= case when (Dt_Vencimento<=now()) then ''0'' when (Dt_Vencimento>now()) then ''1'' end where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null))');
      if not empty(EcodAdversa.text) Then
      Begin
        ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
      end;
      if not empty(EDLoteamento.Text) Then
      begin
        posi:=pos('-',EDLoteamento.Text);
        dec(posi);
        ZQRecBai.SQL.Add(' and (idloteamento = '+quotedstr(copy(alltrim(EDLoteamento.Text),1,posi))+')');
      end;
      if pfim.Text<>'  /  /    ' then
      begin
         ZQRecBai.SQL.Add(' and (dt_vencimento between :dt1 and :dt2)');
         ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
         ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
      end;
//      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      if RBCheque.Checked then
      Begin
         Label37.Visible:=false;
         pini.Visible:=false;
         pini.Clear;
         Label38.Visible:=false;
         pfim.Visible:=false;
         pfim.Clear;

         vartipo := '';
         DM_TAbelas.ZQTipodoc.First;
         while not DM_Tabelas.ZQTipodoc.Eof do
         begin
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
           DM_Tabelas.ZQTipodoc.Next;
         end;
         vartipo := copy(vartipo,1,length(vartipo)-1);
         ZQRecBai.SQL.Clear;
         ZQRecBai.SQL.Add('update recebimento set marcar= case when (Dt_Vencimento<=now()) then ''0'' when (Dt_Vencimento>now()) then ''1'' end where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null))');
//         ZQRecBai.SQL.Add('select * from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
  {       ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
         ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
         ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);}
      end;
      ZQRecBai.ExecSQL;
      ZQRecBai.SQL.Clear;
      ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
      ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
      ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
      ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
      if not empty(EcodAdversa.text) Then
      Begin
          ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
      end;
      if not empty(EDLoteamento.Text) Then
      begin
         posi:=pos('-',EDLoteamento.Text);
         dec(posi);
         ZQRecBai.SQL.Add(' and (idloteamento = '+quotedstr(copy(alltrim(EDLoteamento.Text),1,posi))+')');
      end;
      if pfim.Text<>'  /  /    ' then
      begin
          ZQRecBai.SQL.Add(' and (dt_vencimento between :dt1 and :dt2)');
          ZQRecBai.ParamByName('dt1').AsDate:=strtodate(PIni.text);
          ZQRecBai.ParamByName('dt2').AsDate:=strtodate(PFim.text);
      end;
      ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
      if RBCheque.Checked then
      Begin
         Label37.Visible:=false;
         pini.Visible:=false;
         pini.Clear;
         Label38.Visible:=false;
         pfim.Visible:=false;
         pfim.Clear;

         vartipo := '';
         DM_TAbelas.ZQTipodoc.First;
         DM_TAbelas.ZQTipodoc.DisableControls;
         while not DM_Tabelas.ZQTipodoc.Eof do
         begin
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').Value = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').Value)+',';
           DM_Tabelas.ZQTipodoc.Next;
         end;
         DM_TAbelas.ZQTipodoc.EnableControls;
         vartipo := copy(vartipo,1,length(vartipo)-1);
         ZQRecBai.SQL.Clear;
         ZQRecBai.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
         ZQRecBai.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
         ZQRecBai.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,juros,descontos,Data_Quitacao,tip ');
         ZQRecBai.SQL.Add(' from recebimento where (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) ');
         ZQRecBai.SQL.Add(' and (adversa = '+quotedstr(EcodAdversa.Text)+')');
         ZQRecBai.SQL.Add(' order by Dt_Vencimento, documento');
         ZQRecBai.ParamByName('dt5').AsDate:=strtodate(DTPIni.Datetext);
         ZQRecBai.ParamByName('dt6').AsDate:=strtodate(DTPFim.Datetext);
      end;
      ZQRecBai.Open;
    end;



{      while not ZQRecBai.Eof do begin
        ZQRecBai.Edit;
      if ZQRecBaiDt_Vencimento.Value<=date Then
        ZQRecBaimarcar.Value := '0'
      else
        ZQRecBaimarcar.Value := '1';
        ZQRecBai.Post;
        ZQRecBai.Next;
      end;
    end;}
    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').Value,[]);
    if RBPonteiro.Checked Then
    Begin
      ZQRecBai.Edit;
      ZQRecBai.FieldByName('marcar').Value := '0';
      ZQRecBai.Post;
    end;
  end;
{  ZQRecBai.Open;
  XEEmpresend.Text := ZQRecBainomeempreend.Value;
  LREg.Caption := inttostr(ZQREcbai.RecordCount)+' Registros';
  if RBParticipante.Checked Then
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
//      if ZQRecBaiDt_Vencimento.Value<=date Then
        ZQRecBaimarcar.Value := '0';
  //    else
    //    ZQRecBaimarcar.Value := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
  ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimentoidRecebimento.Value,[]);
  if RBPonteiro.Checked Then Begin
    ZQRecBai.Edit;
    ZQRecBaimarcar.Value := '0';
    ZQRecBai.Post;
  end;}
end;


procedure TFrmRecebBaixa.GBContraExit(Sender: TObject);
begin
  if not RBPonteiro.Checked Then
//    Filtrabaixa;
end;

procedure TFrmRecebBaixa.RBParticipanteClick(Sender: TObject);
begin
  Label1.Visible := False;
  DTPIni.Visible := False;
  Label23.Visible := False;
  DTPFim.Visible := False;
  Enomes.Visible := True;
  Enomes.SetFocus;
end;

procedure TFrmRecebBaixa.RBPonteiroClick(Sender: TObject);
Var
  VArreg : integer;
begin
  if ZQRecBai.active=true then
  begin
    Label1.Visible := False;
    DTPIni.Visible := False;
    Label23.Visible := False;
    DTPFim.Visible := False;
    Enomes.Visible := True;
    varreg := ZQRecBai.RecNo;
    ZQRecBai.First;
    ZQRecBai.DisableControls;
    while not ZQRecBai.Eof do begin
      ZQRecBai.Edit;
      ZQRecBai.FieldByName('marcar').Value := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
    ZQRecBai.EnableControls;
    ZQRecBai.Edit;
    ZQRecBai.RecNo := varreg;
    ZQRecBai.FieldByName('marcar').Value := '0';
    ZQRecBai.Post;
    XNEMarcado.Value := ZQRecBai.FieldByName('saldo').Value;
    Atualiza_tela;
  end;
end;

procedure TFrmRecebBaixa.XDBNumEdit1Enter(Sender: TObject);
begin
  if CDSParcelaspercent.Value = 0 Then
    CBPeriodoEnt.SetFocus;

end;

procedure TFrmRecebBaixa.JDEntradaExit(Sender: TObject);
begin
  if not (DXBFechar.Focused or GBContra.Focused or CBSaldoCorrigido.Focused or RBPonteiro.Focused or RBCheque.Focused or RBParticipante.Focused ) Then
  Begin
    if ZQRecBai.Active=true then
    begin
      if ZQRecBai.RecordCount>0 then
         Atualiza_tela;
    end;
    JDBaixa.DateText:=JDEntrada.DateText;
//    DBGBaixando.SetFocus;
  end;
end;

procedure TFrmRecebBaixa.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   Elote.clear;
   //if Task.Status = TTaskStatus.Completed then
//   begin
     dm_tabelas.ZQLoteamento.close;
     JDEntrada.Enabled := False;
     JDBaixa.Enabled   := False;
     ZQRecBai.Close;
     ZQRecebBxTemp.Close;
     ZQrecebimento.close;
  {  ZQrecebimento.SQL.Clear;
    ZQrecebimento.SQL.Add('update recebimento set saldo=0 where saldo<=5');
    ZQrecebimento.ExecSQL;
    ZQrecebimento.Close;}
     dm_tabelas.ZQLoteamento.close;
     DM_tabelas.ZQCobaRe.close;
     DM_Tabelas.ZQNumOrdem.close;
     DM_tabelas.ZQReBxHi.close;
     DM_Tabelas.ZQDividido.close;
     DM_tabelas.contador_bx.close;
     DM_Tabelas.ZQMovBancaria.close;
     DM_tabelas.ZQCheque.close;
     DM_tabelas.ZQCaixa.close;

//   end
//   else
//   begin
//     showmessage('Aguarde a Leitura dos Dados. Fechar� Automaticamente em Seguida.');
//     sleep(12000);
//     if Task.Status = TTaskStatus.Completed then
//     begin
//       dm_tabelas.ZQLoteamento.close;
//       JDEntrada.Enabled := False;
//       JDBaixa.Enabled   := False;
//       ZQRecBai.Close;
//       ZQrecebimento.close;
//    {  ZQrecebimento.SQL.Clear;
//      ZQrecebimento.SQL.Add('update recebimento set saldo=0 where saldo<=5');
//      ZQrecebimento.ExecSQL;
//      ZQrecebimento.Close;}
//       ZQRecebBxTemp.close;
//       dm_tabelas.ZQLoteamento.close;
//       DM_tabelas.ZQCobaRe.close;
//       DM_Tabelas.ZQNumOrdem.close;
//       DM_tabelas.ZQReBxHi.close;
//       DM_Tabelas.ZQDividido.close;
//       DM_tabelas.contador_bx.close;
//       DM_Tabelas.ZQMovBancaria.close;
//       DM_tabelas.ZQCheque.close;
//       DM_tabelas.ZQCaixa.close;
//     end
//     else
//     begin
//       showmessage('Aguarde a Leitura dos Dados. Fechar� Automaticamente em Seguida.');
//       sleep(15000);
//       if Task.Status = TTaskStatus.Completed then
//       begin
//         dm_tabelas.ZQLoteamento.close;
//         JDEntrada.Enabled := False;
//         JDBaixa.Enabled   := False;
//         ZQRecBai.Close;
//         ZQrecebimento.close;
//      {  ZQrecebimento.SQL.Clear;
//        ZQrecebimento.SQL.Add('update recebimento set saldo=0 where saldo<=5');
//        ZQrecebimento.ExecSQL;
//        ZQrecebimento.Close;}
//         ZQRecebBxTemp.close;
//         dm_tabelas.ZQLoteamento.close;
//         DM_tabelas.ZQCobaRe.close;
//         DM_Tabelas.ZQNumOrdem.close;
//         DM_tabelas.ZQReBxHi.close;
//         DM_Tabelas.ZQDividido.close;
//         DM_tabelas.contador_bx.close;
//         DM_Tabelas.ZQMovBancaria.close;
//         DM_tabelas.ZQCheque.close;
//         DM_tabelas.ZQCaixa.close;
//       end;
//     end;
//   end;
   FrmRecebBaixa:=nil;
   Action:=CaFree;


end;


procedure TFrmRecebBaixa.FormCreate(Sender: TObject);
begin
  EnsureRuntimeFields(Self);
  // definindo o componete a ser arrastado
  GroupBox2.OnMouseDown := ControlMouseDown;
  GroupBox2.OnMouseMove := ControlMouseMove;
  GroupBox2.OnMouseUp   := ControlMouseUp;

  GroupBox3.OnMouseDown := ControlMouseDown;
  GroupBox3.OnMouseMove := ControlMouseMove;
  GroupBox3.OnMouseUp   := ControlMouseUp;

end;

procedure TFrmRecebBaixa.ZQRecBaiCalcFields(DataSet: TDataSet);
var
  vardI, vardT : Tdate;
  varsal : Double;
begin
    if (ZQRecBai.FieldByName('saldo').Value>0) and (ZQRecBai.FieldByName('recpag').Value='R') and (CBSaldoCorrigido.Checked)  Then Begin
    varsal := ZQRecBai.FieldByName('Valor').value;

    ZQRecebBxTemp.Close;
    ZQRecebBxTemp.SQL.Clear;
    ZQRecebBxTemp.SQL.Add('SELECT H.idrecbxhist,H.refer,H.idrecib,H.valor,H.descon,H.juros,H.percent_usado,H.data,H.sq,H.valor_parcela,');
    ZQRecebBxTemp.SQL.Add('       R.documento,R.cliente,R.dt_entrada,R.dt_vencimento,R.Valor,R.VrDoc,R.contabil,R.usuario,R.ordem,R.Tipdoc,R.nomeadversa,R.venda_idvenda,R.quadralote,R.idloteamento,R.saldo,R.refvinda,R.Observ,R.recpag,');
    ZQRecebBxTemp.SQL.Add('       R.juros,R.descontos,R.Data_Quitacao,');
    ZQRecebBxTemp.SQL.Add('       R.marcar,R.adversa,R.numordem,R.substituicao,R.sq,R.Parcelas_fixas,R.origem,R.adversa,R.recpag,R.numordem,R.idrecebimento,R.empresa,R.custodaparcela,R.origem,R.adversa,R.recpag,R.numordem,R.RefBaixa,');
    ZQRecebBxTemp.SQL.Add('       B.idreceb_baixa,B.Dt_rec,B.Vr_rec,B.TipDoc,B.Docum,B.Juros_Vr,B.Desc_Vr,B.RefBaixa,B.obsebx,B.Codcontabil,B.dataref,B.vencimento,B.substituicao,B.vr_ab,B.sq,B.valor_parcela ');
    ZQRecebBxTemp.SQL.Add(' FROM recbxhist as H join recebimento as R ON R.idrecebimento=H.idrecib join receb_baixa as B ON B.refbaixa=H.refer');
    ZQRecebBxTemp.SQL.Add(' where H.idrecib='+quotedstr(ZQRecBai.FieldByName('idrecebimento').Text));
    ZQRecebBxTemp.open;
    ZQRecebBxTemp.First;
//    ZQRecebBxTemp.Refresh;
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').Value;
    vardT := strtodate(FrmRecebBaixa.JDEntrada.DateText);
    ZQRecebBxTemp.DisableControls;
    while (not ZQRecebBxTemp.Eof) and (vardT>=ZQRecebBxTemp.FieldByName('dataref').Value) do begin
      varsal := Calcjuros(vardI,ZQRecebBxTemp.FieldByName('dataref').Value,varsal,XNJuros.Value);
      varsal := varsal-(ZQRecebBxTemp.FieldByName('Vr_rec').Value*ZQRecebBxTemp.FieldByName('percent_usado').Value/100);
      vardI := ZQRecebBxTemp.FieldByName('dataref').Value;
      ZQRecebBxTemp.Next;
    end;
    ZQRecebBxTemp.EnableControls;
    ZQRecBai.FieldByName('saldocorrig').Value := Calcjuros(vardI,vardT,varsal,XNJuros.Value);
  end
  else if ZQRecBai.FieldByName('saldo').Value=0 Then
    ZQRecBai.FieldByName('saldocorrig').Value := 0
  else
    ZQRecBai.FieldByName('saldocorrig').Value := ZQRecBai.FieldByName('saldo').Value;
end;

procedure TFrmRecebBaixa.CDSParcelasCalcFields(DataSet: TDataSet);
begin
//  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
    CDsParcelasPercent.Value := CDsParcelasVrParc.Value*100/XNERecebido.Value;
  end
  else
  begin
    CDsParcelasPercent.Value :=100;
  end;

end;

procedure TFrmRecebBaixa.ZQRatearCalcFields(DataSet: TDataSet);
begin
    ZQRatear.FieldByName('vr_base').Value := ZQRatear.FieldByName('vr_rec').Value * ZQRatear.FieldByName('percent_usado').Value / 100;
    ZQRatear.FieldByName('vr_divi').Value := ZQRatear.FieldByName('vr_base').Value * ZQRatear.FieldByName('percentual').Value /100;

end;

procedure TFrmRecebBaixa.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').Value = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmRecebBaixa.CDS_divididoCalcFields(DataSet: TDataSet);
begin
  CDS_divididodescrcontabil2.Value := CDS_divididocontabil.Text+' - '+CDS_divididodescrcontabil.Value;
end;

procedure TFrmRecebBaixa.CBPeriodoEntExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma : double;
  varvenci : Tdatetime;
  vartip : string;
begin

  if CBPeriodoEnt.ItemIndex <0 Then Begin
    CBPeriodoEnt.SetFocus;
    exit;
  end;
  varvezes:=1;
  ZQRecBai.Filtered := false;
  ZQRecBai.Filter := 'marcar=0';
  ZQRecBai.Filtered := True;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
    if ZQRecBai.FieldByName('Data_Quitacao').Value>0 then
       varvenci :=  ZQRecBai.FieldByName('Data_Quitacao').Value
    else
       varvenci :=  strtodate(JDEntrada.DateText);
    CDSParcelas.Close;
    CDSParcelas.CreateDataSet;
    CDSParcelas.Open;
    CDSParcelas.EmptyDataSet;
    CDSParcelas.IndexFieldNames := 'ordem';
    varfinal := strtoint(floattostr(XNEParcelas.Value));
    if StrIsNumero(CBPeriodoEnt.Text) Then Begin
      vardias := strtoint(CBPeriodoEnt.Text);
      vartip := 'N';
    end
    else Begin
      if CBPeriodoEnt.ItemIndex<0 Then Begin
        CBPeriodoEnt.SetFocus;
        Exit;
      end;
      vartip := 'S';
      case CBPeriodoEnt.ItemIndex of
        0 : vardias := 0;
        1 : vardias := 1;
        2 : vardias := 7;
        3 : vardias := 15;
        4 : vardias := 1;
        5 : vardias := 2;
        6 : vardias := 3;
        7 : vardias := 4;
        8 : vardias := 6;
        9 : vardias := 12;
        10: vardias := 24;
      end;
      if CBPeriodoEnt.ItemIndex<=2 Then
        vartip := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value,[]);
    varsoma := 0;
    for varvezes := 1 to varfinal do
    Begin
      CDSParcelas.Insert;
      CDSParcelasOrdem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
      CDSParcelasjuros.Value :=  ExRound(XNEJuros.Value,2)/ExRound(varfinal,2);
      CDSParcelasVrParc.Value := ExRound(XNERecebido.Value,2)/ExRound(varfinal,2);

      if ZQRecBai.FieldByName('recpag').Value='R' Then
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').Value
      else
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').Value;
        CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').Value;
      if ZQRecBai.FieldByName('Data_Quitacao').Value>0 then
      begin
        if varvezes = 1 Then
           CDSParcelasVenci.Value := ZQRecBai.FieldByName('Data_Quitacao').Value
        else Begin
          if vartip = 'S' then
            CDSParcelasVenci.Value := IncMonth(ZQRecBai.FieldByName('Data_Quitacao').Value,(varvezes*vardias)-(vardias))
          else
            CDSParcelasVenci.Value := varvenci+vardias;
          varvenci := CDSParcelasVenci.Value;
        end;
      end
      else
      begin
        if varvezes = 1 Then
           CDSParcelasVenci.Value := strtodate(JDEntrada.DateText)
        else Begin
        if vartip = 'S' then
            CDSParcelasVenci.Value := IncMonth(strtodate(JDEntrada.DateText),(varvezes*vardias)-(vardias))
          else
            CDSParcelasVenci.Value := varvenci+vardias;
          varvenci := CDSParcelasVenci.Value;
        end;
      end;
      CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').Value;
      CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').Value;
      CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').Value;
      CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
      CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').Value;
      CDSParcelassq.Value := ZQRecBai.FieldByName('sq').Value;
      CDSParcelasData_Quitacao.Value:=ZQRecBai.FieldByName('Data_Quitacao').Value;
      CDSParcelastip.Value:=ZQRecBai.FieldByName('tip').Value;
      CDSParcelas.Post;
      varsoma := varsoma + CDSParcelasVrParc.Value;
     end;
    CDSParcelas.IndexFieldNames := 'ordem';
    CDSParcelas.First;
    CDSParcelas.Edit;
    CDSParcelasVrParc.Value := CDSParcelasVrParc.Value+((XNERecebido.Value)-varsoma);
    CDSParcelas.Post;
    ZQRecBai.Filtered := false;
  end
  else
  begin
    if ZQRecBai.FieldByName('Data_Quitacao').Value>0 then
       varvenci :=  ZQRecBai.FieldByName('Data_Quitacao').Value
    else
       varvenci :=  strtodate(JDEntrada.DateText);

    CDSParcelas.Close;
    CDSParcelas.CreateDataSet;
    CDSParcelas.Open;
    CDSParcelas.EmptyDataSet;    
    CDSParcelas.IndexFieldNames := 'ordem';
    varfinal := strtoint(floattostr(XNEParcelas.Value));
    if StrIsNumero(CBPeriodoEnt.Text) Then Begin
      vardias := strtoint(CBPeriodoEnt.Text);
      vartip := 'N';
    end
    else Begin
      if CBPeriodoEnt.ItemIndex<0 Then Begin
        CBPeriodoEnt.SetFocus;
        Exit;
      end;
      vartip := 'S';
      case CBPeriodoEnt.ItemIndex of
        0 : vardias := 0;
        1 : vardias := 1;
        2 : vardias := 7;
        3 : vardias := 15;
        4 : vardias := 1;
        5 : vardias := 2;
        6 : vardias := 3;
        7 : vardias := 4;
        8 : vardias := 6;
        9 : vardias := 12;
        10: vardias := 24;
      end;
      if CBPeriodoEnt.ItemIndex<=2 Then
        vartip := 'N';
    end;
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').Value,[]);
    varsoma := 0;
    ZQRecBai.First;
    varfinal:=1;
    //for varvezes := 1 to varfinal do Begin
    ZQRecBai.DisableControls;
    while not ZQRecBai.Eof do
    begin
      CDSParcelas.Insert;
      CDSParcelasidrecebimento.Value:= ZQRecBai.FieldByName('idrecebimento').Value;
      CDSParcelasOrdem.Value := '001/001';
      CDSParcelasjuros.Value :=  ExRound(ZQRecBai.FieldByName('juros').Value,2);
      CDSParcelasdesconto.Value :=  ExRound(ZQRecBai.FieldByName('descontos').Value,2);
      CDSParcelasVrParc.Value := (ExRound(ZQRecBai.FieldByName('saldo').Value,2)+ExRound(ZQRecBai.FieldByName('juros').Value,2))-ExRound(ZQRecBai.FieldByName('descontos').Value,2);
      if ZQRecBai.FieldByName('recpag').Value='R' Then
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').Value
      else
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').Value;
        CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').Value;
      if ZQRecBai.FieldByName('Data_Quitacao').Value>0 then
      begin
        if varvezes = 1 Then
           CDSParcelasVenci.Value := ZQRecBai.FieldByName('Data_Quitacao').Value
        else Begin
          if vartip = 'S' then
            CDSParcelasVenci.Value := IncMonth(ZQRecBai.FieldByName('Data_Quitacao').Value,(varfinal*vardias)-(vardias))
          else
            CDSParcelasVenci.Value := varvenci+vardias;
          varvenci := CDSParcelasVenci.Value;
        end;
      end
      else
      begin
        if varvezes = 1 Then
           CDSParcelasVenci.Value := strtodate(JDEntrada.DateText)
        else Begin
        if vartip = 'S' then
            CDSParcelasVenci.Value := IncMonth(strtodate(JDEntrada.DateText),(varvezes*vardias)-(vardias))
          else
            CDSParcelasVenci.Value := varvenci+vardias;
          varvenci := CDSParcelasVenci.Value;
        end;
      end;
      CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').Value;
      CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').Value;
      CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').Value;
      CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
      CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').Value;
      CDSParcelassq.Value := ZQRecBai.FieldByName('sq').Value;
      CDSParcelasData_Quitacao.Value:=ZQRecBai.FieldByName('Data_Quitacao').Value;
      CDSParcelastip.Value:=ZQRecBai.FieldByName('tip').Value;
      CDSParcelas.Post;

      varsoma := varsoma + CDSParcelasVrParc.Value;
      ZQRecBai.Next;
    end;
    ZQRecBai.EnableControls;
    CDSParcelas.IndexFieldNames := 'ordem';
    CDSParcelas.First;
{    CDSParcelas.Edit;
    CDSParcelasVrParc.Value := CDSParcelasVrParc.Value+((XNERecebido.Value)-varsoma);
    CDSParcelas.Post;}
    ZQRecBai.Filtered := false;

  end;
end;

procedure TFrmRecebBaixa.EadversaExit(Sender: TObject);
var
posi:integer;
begin
  if not DXBFechar.Focused Then
  Begin
    if not empty(Eadversa.Text) Then
    Begin
//      if Lpassou.Caption<>'S' then
//      begin
//       showmessage('Aguarde o Carregamento dos Dados e Tente Novamente');
//       Eadversa.SetFocus;
//       exit;
//      end;

      if AchaCLiente(400, 300, Eadversa.Text) Then
      Begin
        Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
        EcodAdversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
        EcodAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
        if not empty(EDLoteamento.Text) then
        begin
          posi:=pos('-',EDLoteamento.Text);
          dec(posi);
          DM_tabelas.ZQRecebimento.Close;
          DM_tabelas.ZQRecebimento.SQL.Clear;
          DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
          if pfim.Text='  /  /    ' then
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+' and adversa='+quotedstr(EcodAdversa.Text)+' and saldo>0');
            DM_tabelas.ZQRecebimento.open;
          end
          else
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+' and adversa='+quotedstr(EcodAdversa.Text)+' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2 and saldo>0');
            DM_tabelas.ZQRecebimento.ParamByName('dt1').Value:=strtodate(pini.text);
            DM_tabelas.ZQRecebimento.ParamByName('dt2').Value:=strtodate(pfim.text);
            DM_tabelas.ZQRecebimento.open;
          end;
        end
        else
        begin
          DM_tabelas.ZQRecebimento.Close;
          DM_tabelas.ZQRecebimento.SQL.Clear;
          DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
          DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
          if pfim.Text='  /  /    ' then
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and adversa='+quotedstr(EcodAdversa.Text)+' and saldo>0');
            DM_tabelas.ZQRecebimento.open;
          end
          else
          begin
            DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and adversa='+quotedstr(EcodAdversa.Text)+' and Dt_Vencimento>= :dt1 and Dt_Vencimento<= :dt2 and saldo>0');
            DM_tabelas.ZQRecebimento.ParamByName('dt1').Value:=strtodate(pini.text);
            DM_tabelas.ZQRecebimento.ParamByName('dt2').Value:=strtodate(pfim.text);
            DM_tabelas.ZQRecebimento.open;
          end;
        end;
        DM_tabelas.ZQRecebimento.Recordcount;
//        JDEntrada.Date := Date;
        JDBaixa.DateText   := JDEntrada.DateText;
        DBText1.Caption := '';
        XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;

        DM_tabelas.ZQCobaRe.Insert;
        DTPIni.Datetext := datetostr(date);
        DTPFim.Datetext := datetostr(date);
        DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
        DM_tabelas.ZQCobaRe.post;

        GBCheque.Visible := False;
        DXBBaixaGravar.Enabled := false;
        EMObsbx.Clear;
        EMNovostit.Clear;
        EContabil.Text := '';
        XNEParcelas.Value := 1;
        XNEAberto.Value := 0;
        XNEVencido.Value := 0;
        XNEMarcado.Value := 0;
        XNERecebido.Value := 0;
        XNEDesconto.Value := 0;
        XNEJuros.Value := 0;
        DBCBTipDoc.Clear;
        DM_tabelas.ZQTipodoc.First;
        while not DM_tabelas.ZQTipodoc.Eof do
        Begin
          DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
          DM_tabelas.ZQTipodoc.Next;
        end;
        Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
        Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
        Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
        EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
        EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').Value;
        varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;
      end
      else
        JDEntrada.SetFocus;
      filtrabaixa;  
    end
    else
      EcodAdversa.Text := '';

  end;
end;

procedure TFrmRecebBaixa.EAdversaNovaExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
    if not empty(EAdversaNova.Text) Then Begin
      if AchaCLiente(400, 300, EAdversaNova.Text) Then Begin
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').Value;
        EcodAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
      end
      else
        EAdversaNova.SetFocus;
    end
    else
      EAdversaNova.Text := '';
  end;
end;

procedure TFrmRecebBaixa.CBTipobaixaExit(Sender: TObject);
begin
  if CBTipobaixa.Text='NORMAL' then
  begin
    Label15.Caption:='Dt. Quita��o.......';
  end
  else
  begin
    Label15.Caption:='Dt. 1?agto........';
  end;
  //JDEntrada.SetFocus;
  DBGBaixando.SetFocus;
end;

procedure TFrmRecebBaixa.DXBPesquisarClick(Sender: TObject);
begin
  FrmPesqRecebimento_bx.Top := FrmRecebBaixa.Top+100;
  FrmPesqRecebimento_bx.Left := FrmRecebBaixa.Left;
  FrmPesqRecebimento_bx.showmodal;


//  JDEntrada.Date := Date;
  JDBaixa.DateText   := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
  DM_tabelas.ZQCobaRe.post;

  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').Value;  
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


  CBPeriodoEnt.ItemIndex := 0;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True; 
  Atualiza_tela;


  DBGBaixando.SetFocus;

end;

procedure TFrmRecebBaixa.EloteExit(Sender: TObject);
var
  vartip : string;

begin
 if not empty(Elote.Text) then
 begin
//   if Lpassou.Caption<>'S' then
//   begin
//     showmessage('Aguarde o Carregamento dos Dados e Tente Novamente');
//     Elote.SetFocus;
//     exit;
//   end;
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and quadralote='+quotedstr(Elote.text)+' and saldo>0');
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.Recordcount;
//  JDEntrada.DateText := datetostr(Date);
  JDBaixa.DateText   := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;

  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
  DM_tabelas.ZQCobaRe.post;

  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  DM_tabelas.ZQTipodoc.DisableControls;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').Value;
  varrecpag            := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;

  Atualiza_tela;

  DBGBaixando.SetFocus;
 end;

end;

procedure TFrmRecebBaixa.btndesClick(Sender: TObject);
begin
  ZQRecBai.First;
  while not ZQRecBai.Eof do
  begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').Value = '0' Then
      ZQRecBai.FieldByName('marcar').Value := '1'
    else
      ZQRecBai.FieldByName('marcar').Value := '0';
    ZQRecBai.post;
    ZQRecBai.Next;
  end;
  Atualiza_tela;

end;

procedure TFrmRecebBaixa.bxmanualExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.Edit;
  DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value:=bxmanual.Text;
  DM_Tabelas.ZQConfiguracoes.Post;
  DM_Tabelas.ZQConfiguracoes.Refresh;
end;

procedure TFrmRecebBaixa.EDLoteamentoDropDown(Sender: TObject);
begin
//   if Lpassou.Caption<>'S' then
//   begin
//     showmessage('Aguarde o Carregamento dos Dados e Tente Novamente');
//     EDLoteamento.SetFocus;
//     exit;
//   end;

  EDLoteamento.Clear;
  DM_tabelas.ZQLoteamento.First;
  DM_tabelas.ZQLoteamento.DisableControls;
  while not  DM_tabelas.ZQLoteamento.Eof do
  begin
    EDLoteamento.Items.Add( DM_tabelas.ZQLoteamento.FieldByName('idloteamento').AsString+'-'+DM_tabelas.ZQLoteamento.FieldByName('apelido').AsString);
    DM_tabelas.ZQLoteamento.Next;
  end;
  DM_tabelas.ZQLoteamento.EnableControls;

end;

procedure TFrmRecebBaixa.Edit2Exit(Sender: TObject);
begin
  if not empty(Edit2.text) then
  begin
//   if Lpassou.Caption<>'S' then
//   begin
//     showmessage('Aguarde o Carregamento dos Dados e Tente Novamente');
//     Edit2.SetFocus;
//     exit;
//   end;

    ZQCheque.close;
    ZQCheque.SQL.Clear;
    ZQCheque.SQL.Add('select ch.Banco,ch.Dono,ch.CPF_CNPJ,ch.valor,ch.emissao,ch.deposito,ch.alias,ch.prorrogado,ch.conta,ch.idrecebimento,ch.idpagamento,ch.idcheque,ch.idvinculo,ch.numero,ch.agencia,ch.Substituicao,ch.sq,');
    ZQCheque.SQL.Add('       rc.idrecebimento,rc.documento,rc.cliente,rc.usuario,rc.Dt_Entrada,rc.Dt_Vencimento,rc.Valor,rc.Observ,rc.VrDoc,rc.ordem,rc.TipDoc,rc.saldo,rc.marcar,');
    ZQCheque.SQL.Add('       rc.RefBaixa,rc.refvinda,rc.contabil,rc.empresa,rc.custodaparcela,rc.origem,rc.adversa,rc.recpag,rc.numordem,rc.idloteamento,rc.venda_idvenda,rc.quadralote,');
    ZQCheque.SQL.Add('       rc.numboleto,rc.Substituicao,rc.sq,rc.somar,rc.nomeadversa,rc.Reajustado,rc.Data_reajuste ');
    ZQCheque.SQL.Add(' from cheque_rec as ch join recebimento as rc on ch.idrecebimento=rc.idrecebimento');
    ZQCheque.SQL.Add(' where ch.numero like '+quotedstr('%'+edit2.Text+'%')+' order by numero');
    ZQCheque.Open;
    ZQCheque.First;
    if (ZQCheque.Active) and (ZQCheque.RecordCount>0) Then
    begin
      while not ZQCheque.Eof do
      begin
        DM_tabelas.ZQRecebimento.Close;
        DM_tabelas.ZQRecebimento.SQL.Clear;
        DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
        DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
        DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where idrecebimento ='+quotedstr(ZQCheque.FieldByName('idrecebimento').text)+' and recpag=''R'' and saldo>0');
        DM_tabelas.ZQRecebimento.open;
        if DM_tabelas.ZQRecebimento.RecordCount>0 then
        begin
          xrec:=DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').value;
          Elote.Text:=DM_tabelas.ZQRecebimento.FieldByName('quadralote').Value;
          elote.setfocus;
          ZQCheque.close;
          exit;
        end;
        ZQCheque.Next;
      end;
    end
    else
    begin
      ZQCheque.close;
      showmessage('N�mero de Cheque n�o Encontrado.');
      dxbfechar.setfocus;
    end;
  end;
  showmessage('Cheque n�o Encontrado nas Parcelas em Aberto.');
  DM_tabelas.ZQRecebimento.Close;
  DM_tabelas.ZQRecebimento.SQL.Clear;
  DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
  DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
  DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and 1=0');
  DM_tabelas.ZQRecebimento.open;
  dxbfechar.setfocus;


end;

procedure TFrmRecebBaixa.JDEntradaEnter(Sender: TObject);
begin
  FrmRecebBaixa.KeyPreview:=true;
end;


// procedures para arrastar componentes. No caso aqui groupbox2
procedure TFrmRecebBaixa.ControlMouseDown(
  Sender: TObject;
  Button: TMouseButton;
  Shift: TShiftState;
  X, Y: Integer);
begin
    inReposition:=True;
    SetCapture(TWinControl(Sender).Handle);
    GetCursorPos(oldPos);
end; (*ControlMouseDown*)


procedure TFrmRecebBaixa.ControlMouseMove(
  Sender: TObject;
  Shift: TShiftState;
  X, Y: Integer);
const
  minWidth = 20;
  minHeight = 20;
var
  newPos: TPoint;
  frmPoint : TPoint;
begin
  if inReposition then
  begin
    with TWinControl(Sender) do
    begin
      GetCursorPos(newPos);

      if ssShift in Shift then
      begin //resize
        Screen.Cursor := crSizeNWSE;
        frmPoint := ScreenToClient(Mouse.CursorPos);
        if frmPoint.X > minWidth then 
          Width := frmPoint.X;
        if frmPoint.Y > minHeight then 
          Height := frmPoint.Y;
      end
      else //move
      begin
        Screen.Cursor := crSize;
        Left := Left - oldPos.X + newPos.X;
        Top := Top - oldPos.Y + newPos.Y;
        oldPos := newPos;
      end;
    end;
  end;
end; (*ControlMouseMove*)


procedure TFrmRecebBaixa.ControlMouseUp(
  Sender: TObject;
  Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if inReposition then
  begin
    Screen.Cursor := crDefault;
    ReleaseCapture;
    inReposition := False;
  end;
end; (*ControlMouseUp*)


procedure TFrmRecebBaixa.XNEDescontoEnter(Sender: TObject);
begin
  xdesc:=XNEDesconto.Value;
end;

procedure TFrmRecebBaixa.XNEDescontoExit(Sender: TObject);
begin
  // 01/02/2021
  if (XNEResto.Value)=(XNEDesconto.Value) Then
  begin
      xdesc:=XNEResto.Value;
      XNEResto.Value:=0;
  end;

end;

procedure TFrmRecebBaixa.DBGBaixandoDblClick(Sender: TObject);
begin
  dpl.Caption:='s';
  fim.Caption:='n';
  FrmRecebBaixa.KeyPreview:=false;
  xnrecebido.Value:=XNERecebido.Value;
  ZQRecBai.Edit;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').value='DATA QUITA��O GERAL' then
  begin
     GroupBox2.Visible:=true;
     DBGrid1.SetFocus;
     DBGrid1.SelectedIndex:=0; // setfocus na 1?coluna
     DBGrid1.Options := DBGrid1.Options + [dgEditing];
  end
  else
  begin
     GroupBox3.Visible:=true;
     DBGrid3.SetFocus;
     DBGrid3.SelectedIndex:=0; // setfocus na 1?coluna
     DBGrid3.Options := DBGrid3.Options + [dgEditing];
  end;


end;

procedure TFrmRecebBaixa.DBGBaixandoExit(Sender: TObject);
begin
  if dpl.Caption<>'s' then
     JDEntrada.SetFocus;

end;

procedure TFrmRecebBaixa.DBGrid1ColEnter(Sender: TObject);
begin
  if (DBGrid1.SelectedField = ZQRecBai.FieldByName('juros')) or (DBGrid1.SelectedField = ZQRecBai.FieldByName('descontos')) then
    DBGrid1.Options := DBGrid1.Options + [dgEditing]
  else
    DBGrid1.Options := DBGrid1.Options - [dgEditing];

end;

procedure TFrmRecebBaixa.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGrid1.DefaultDrawDataCell(Rect, DBGrid1.columns[datacol].field, State);
  if Column.Field = ZQRecBai.FieldByName('marcar') then
  begin
    DBGrid1.Canvas.FillRect(Rect);
    if ZQRecBai.FieldByName('marcar').Value = '0' then
      ImageList1.Draw(DBGrid1.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGrid1.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
  //Remove barra Horizontal}
 ShowScrollBar(DBGrid1.Handle,SB_HORZ,False);

{Remove barra Vertical}
 ShowScrollBar(DBGrid1.Handle,SB_VERT,False);



end;

procedure TFrmRecebBaixa.DBGrid1Exit(Sender: TObject);
begin
  if fim.Caption='n' then
  begin
    if ZQRecBai.State in [DsEdit] then
       ZQRecBai.post;
    XNERecebido.Value:=xnrecebido.Value;
    xnrecebido.Value:=0;
    XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').Value)-ZQRecBai.FieldByName('descontos').Value;
    GroupBox2.Visible:=false;
    dpl.Caption:='n';
    DBGBaixando.SetFocus;
    fim.Caption:='s';
  end;

end;

procedure TFrmRecebBaixa.DBGrid1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If ( Chr(Key) = #13) Then
  begin
    If ( DBGrid1.SelectedIndex + 1 <> DBGrid1.FieldCount ) Then
    begin
       DBGRid1.SelectedIndex := DBGRid1.SelectedIndex + 1;
    end
    else
    begin
      if ZQRecBai.State in [DsEdit] then
         ZQRecBai.post;
      XNERecebido.Value:=xnrecebido.Value;
      xnrecebido.Value:=0;
      XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').Value)-ZQRecBai.FieldByName('descontos').Value;
      GroupBox2.Visible:=false;
      dpl.Caption:='n';
      DBGBaixando.SetFocus;
      fim.Caption:='s';
    end;
  end;

end;

procedure TFrmRecebBaixa.DBGrid3ColEnter(Sender: TObject);
begin
  if (DBGrid3.SelectedField = ZQRecBai.FieldByName('juros')) or (DBGrid3.SelectedField = ZQRecBai.FieldByName('descontos')) or (DBGrid3.SelectedField = ZQRecBai.FieldByName('Data_Quitacao')) then
    DBGrid3.Options := DBGrid3.Options + [dgEditing]
  else
    DBGrid3.Options := DBGrid3.Options - [dgEditing];
end;

procedure TFrmRecebBaixa.DBGrid3DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  DBGrid3.DefaultDrawDataCell(Rect, DBGrid3.columns[datacol].field, State);
  if Column.Field = ZQRecBai.FieldByName('marcar') then
  begin
    DBGrid3.Canvas.FillRect(Rect);
    if ZQRecBai.FieldByName('marcar').Value = '0' then
      ImageList1.Draw(DBGrid3.Canvas, Rect.Left + 10, Rect.Top + 1, 0)
    else
      ImageList1.Draw(DBGrid3.Canvas, Rect.Left + 10, Rect.Top + 1, 1);
  end;
  //Remove barra Horizontal}
 ShowScrollBar(DBGrid3.Handle,SB_HORZ,False);

{Remove barra Vertical}
 ShowScrollBar(DBGrid3.Handle,SB_VERT,False);
end;

procedure TFrmRecebBaixa.DBGrid3Exit(Sender: TObject);
begin
  if fim.Caption='n' then
  begin
    if ZQRecBai.State in [DsEdit] then
       ZQRecBai.post;
    XNERecebido.Value:=xnrecebido.Value;
    xnrecebido.Value:=0;
    XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').Value)-ZQRecBai.FieldByName('descontos').Value;
    GroupBox3.Visible:=false;
    dpl.Caption:='n';
    DBGBaixando.SetFocus;
    fim.Caption:='s';
  end;
end;

procedure TFrmRecebBaixa.DBGrid3KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If ( Chr(Key) = #13) Then
  begin
    If ( DBGrid3.SelectedIndex + 1 <> DBGrid3.FieldCount ) Then
    begin
       DBGRid3.SelectedIndex := DBGRid3.SelectedIndex + 1;
    end
    else
    begin
      if ZQRecBai.State in [DsEdit] then
         ZQRecBai.post;
      XNERecebido.Value:=xnrecebido.Value;
      xnrecebido.Value:=0;
      XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').Value)-ZQRecBai.FieldByName('descontos').Value;
      GroupBox3.Visible:=false;
      dpl.Caption:='n';
      DBGBaixando.SetFocus;
      fim.Caption:='s';
    end;
  end;

end;

procedure TFrmRecebBaixa.DBGBaixandoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If ( Chr(Key) = #13) Then
  begin
    if dpl.Caption<>'s' then
      JDEntrada.SetFocus;
  end;
end;

procedure TFrmRecebBaixa.Label9MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  //mover formulario sem bordas
  MovimentaObject(sender,button,shift,x,y,FrmRecebBaixa);

end;

procedure TFrmRecebBaixa.pfimExit(Sender: TObject);
var
posi : integer;
begin
 if pfim.Text<>'  /  /    ' then
 begin
//   if Lpassou.Caption<>'S' then
//   begin
//     showmessage('Aguarde o Carregamento dos Dados e Tente Novamente');
//     pfim.SetFocus;
//     exit;
//   end;

  DM_tabelas.ZQRecebimento.Filtered:=false;

  if not empty(EDLoteamento.Text) then
  begin
    posi:=pos('-',EDLoteamento.Text);
    dec(posi);
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
    DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and  idloteamento='+quotedstr(copy(EDLoteamento.Text,1,posi))+' and (dt_vencimento between :dt1 and :dt2) and saldo>0');
    DM_tabelas.ZQRecebimento.ParamByName('dt1').AsDate:=strtodate(PIni.text);
    DM_tabelas.ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(PFim.text);
  end
  else
  begin
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_Tabelas.ZQRecebimento.SQL.Add('Select  idrecebimento,documento,cliente,usuario,Dt_Entrada,Dt_Vencimento,Valor,Observ,VrDoc,ordem,TipDoc,saldo,marcar,RefBaixa,refvinda,contabil,empresa,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        custodaparcela,origem,adversa,recpag,numordem,idloteamento,venda_idvenda,quadralote,numboleto,Substituicao,sq,nomeadversa,Reajustado,Data_reajuste,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        somar,Proximo_Reajuste,Parcelas_fixas,observ_estorno,tip,juros,descontos,Data_Quitacao,sld_antes_reajuste,Percentual_reajuste,juridico,data_juridico,');
    DM_Tabelas.ZQRecebimento.SQL.Add('        dt_nao_pagou_no_mes,descricao_juridico,multa,mora ' );
    DM_tabelas.ZQRecebimento.SQL.Add(' from recebimento where recpag=''R'' and (dt_vencimento between :dt1 and :dt2) and saldo>0');
    DM_tabelas.ZQRecebimento.ParamByName('dt1').AsDate:=strtodate(PIni.text);
    DM_tabelas.ZQRecebimento.ParamByName('dt2').AsDate:=strtodate(PFim.text);


  end;
  DM_tabelas.ZQRecebimento.open;
  DM_tabelas.ZQRecebimento.Recordcount;
//  JDEntrada.DateText := datetostr(Date);
  JDBaixa.DateText   := JDEntrada.DateText;
  DBText1.Caption := '';
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').Value;

  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').Value := date;
  DM_tabelas.ZQCobaRe.post;

  GBCheque.Visible := False;
  DXBBaixaGravar.Enabled := false;
  EMObsbx.Clear;
  EMNovostit.Clear;
  EContabil.Text := '';
  XNEParcelas.Value := 1;
  XNEAberto.Value := 0;
  XNEVencido.Value := 0;
  XNEMarcado.Value := 0;
  XNERecebido.Value := 0;
  XNEDesconto.Value := 0;
  XNEJuros.Value := 0;
  DBCBTipDoc.Clear;
  DM_tabelas.ZQTipodoc.First;
  while not DM_tabelas.ZQTipodoc.Eof do Begin
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').Value);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').Value;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').Value;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').Value;
  varrecpag            := DM_Tabelas.ZQRecebimento.FieldByName('recpag').Value;


  CBPeriodoEnt.ItemIndex := 4;
  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;

 // Atualiza_tela;
  JDEntrada.SetFocus;
//  DBGBaixando.SetFocus;
 end;
end;


procedure TFrmRecebBaixa.AfterConstruction;
begin
  inherited AfterConstruction;
  EnsureRuntimeFields(Self);
end;

initialization
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQRecBai', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaisaldocorrig', 'saldocorrig', TFloatField, fkCalculated, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaidocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaicliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaisaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaimarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBairefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaicontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaicustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBairecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBainumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBainomeadversa', 'nomeadversa', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'adversa', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaivenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBainumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBainomeempreend', 'nomeempreend', TWideStringField, fkLookup, 100, 0, False, '', '', '', '', 0, 'idloteamento', 'DM_Tabelas.ZQLoteamento', 'idloteamento', 'apelido', True);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaisq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaisomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiProximo_Reajuste', 'Proximo_Reajuste', TWideStringField, fkData, 7, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiParcelas_fixas', 'Parcelas_fixas', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaijuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaidescontos', 'descontos', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '########0.00', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiData_Quitacao', 'Data_Quitacao', TDateField, fkData, 0, 0, False, '', '', '', '!99/99/9999;1;_', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaiNomecli', 'Nomecli', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'nome_parte', True);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecBai', 'ZQRecBaitip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQRecebBxTemp', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTemprefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTemppercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTempidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTempdataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTempvr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRecebBxTemp', 'ZQRecebBxTemprefbaixa', 'refbaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQRatear', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearvalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearpercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearvr_rec', 'vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearpercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRateartipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearvr_base', 'vr_base', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearvr_divi', 'vr_divi', TFloatField, fkCalculated, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQRatear', 'ZQRatearordem_1', 'ordem_1', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQItensRecibo', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboidrecbxhist', 'idrecbxhist', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciborefer', 'refer', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboidrecib', 'idrecib', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibovalor', 'valor', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibodescon', 'descon', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibojuros', 'juros', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibopercent_usado', 'percent_usado', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibodata', 'data', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciborefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciborecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensRecibonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQItensRecibo', 'ZQItensReciboSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQDiv_usado', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadoiddividido', 'iddividido', TLargeintField, fkData, 0, 0, True, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadoparticipante', 'participante', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadopercentual', 'percentual', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadoplanodeconta', 'planodeconta', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadotipodocumento', 'tipodocumento', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQDiv_usado', 'ZQDiv_usadoordem', 'ordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQReceb_Baixa', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixaidreceb_baixa', 'idreceb_baixa', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDt_rec', 'Dt_rec', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaVr_rec', 'Vr_rec', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDocum', 'Docum', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaJuros_Vr', 'Juros_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaDesc_Vr', 'Desc_Vr', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixaobsebx', 'obsebx', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaCodcontabil', 'Codcontabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixadataref', 'dataref', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixavencimento', 'vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_BaixaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixavr_ab', 'vr_ab', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQReceb_Baixa', 'ZQReceb_Baixasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQrecebimento', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentodocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentocliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentousuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentosaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentomarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentorefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentocontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentocustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentorecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentonumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentovenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentonumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentoSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentosq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQrecebimento', 'ZQrecebimentosomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQCheque', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidcheque', 'idcheque', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequebanco', 'banco', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeagencia', 'agencia', TWideStringField, fkData, 6, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequedono', 'dono', TWideStringField, fkData, 50, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeCPF_CNPJ', 'CPF_CNPJ', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequevalor', 'valor', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeemissao', 'emissao', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequedeposito', 'deposito', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequealias', 'alias', TWideStringField, fkData, 12, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeprorrogado', 'prorrogado', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeconta', 'conta', TWideStringField, fkData, 15, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidrecebimento', 'idrecebimento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidpagamento', 'idpagamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequenumero', 'numero', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidvinculo', 'idvinculo', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequesq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidrecebimento_1', 'idrecebimento_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequedocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequecliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeusuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeValor_1', 'Valor_1', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequesaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequemarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequerefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequecontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequecustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequerecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequenumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequevenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequequadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequenumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeSubstituicao_1', 'Substituicao_1', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequesq_1', 'sq_1', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequesomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequenomeadversa', 'nomeadversa', TWideStringField, fkData, 140, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeReajustado', 'Reajustado', TWideStringField, fkData, 3, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQCheque', 'ZQChequeData_reajuste', 'Data_reajuste', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQVenda', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQVenda', 'ZQVendaidvenda', 'idvenda', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQVenda', 'ZQVendadatavenda', 'datavenda', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQVenda', 'ZQVendaimovel', 'imovel', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQVenda', 'ZQVendavalorvenda', 'valorvenda', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQVenda', 'ZQVendaforma_reajuste', 'forma_reajuste', TWideStringField, fkData, 45, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQEntrada', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradadocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradasq', 'sq', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradasomar', 'somar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradaentrada', 'entrada', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQEntrada', 'ZQEntradatip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeDataSet(TFrmRecebBaixa, 'ZQParcela', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaidrecebimento', 'idrecebimento', TLargeintField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParceladocumento', 'documento', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelacliente', 'cliente', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelausuario', 'usuario', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaDt_Entrada', 'Dt_Entrada', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaDt_Vencimento', 'Dt_Vencimento', TDateField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaValor', 'Valor', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaObserv', 'Observ', TWideMemoField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaVrDoc', 'VrDoc', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaordem', 'ordem', TWideStringField, fkData, 14, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaTipDoc', 'TipDoc', TWideStringField, fkData, 2, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelasaldo', 'saldo', TFloatField, fkData, 0, 0, False, '', '###,###,###,##0.00', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelamarcar', 'marcar', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaRefBaixa', 'RefBaixa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelarefvinda', 'refvinda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelacontabil', 'contabil', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaempresa', 'empresa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelacustodaparcela', 'custodaparcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaorigem', 'origem', TWideStringField, fkData, 10, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaadversa', 'adversa', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelarecpag', 'recpag', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelanumordem', 'numordem', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaidloteamento', 'idloteamento', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelavenda_idvenda', 'venda_idvenda', TIntegerField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaquadralote', 'quadralote', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelanumboleto', 'numboleto', TWideStringField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaSubstituicao', 'Substituicao', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelanomecliente', 'nomecliente', TWideStringField, fkLookup, 50, 0, False, '', '', '', '', 0, 'cliente', 'DM_Tabelas.ZqParticipante', 'idpaticipante', 'Nome_Firma', True);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelaparcela', 'parcela', TFloatField, fkData, 0, 0, False, '', '', '', '', 0, '', '', '', '', False);
  RegisterRuntimeField(TFrmRecebBaixa, 'ZQParcela', 'ZQParcelatip', 'tip', TWideStringField, fkData, 1, 0, False, '', '', '', '', 0, '', '', '', '', False);

end.


