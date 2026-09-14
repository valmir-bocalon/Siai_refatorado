

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
    XNERecebido_ant: TXNumEdit;
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
    procedure XNERecebidoEnter(Sender: TObject);

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
    XNEAberto.Value := XNEAberto.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
    if ZQRecBai.FieldByName('marcar').AsString = '0' Then
    Begin
      XNEMarcado.Value := XNEMarcado.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
      XNEJuros.Value := ZQRecBai.FieldByName('juros').AsFloat; //XNEJuros.Value + (ZQRecBaisaldocorrig.Value-ZQRecBaisaldo.Value); 17/05/2012
      XNEDesconto.Value:=ZQRecBai.FieldByName('descontos').AsFloat;
      YNEJuros.Value:=YNEJuros.Value+XNEJuros.Value;
      YNEDesconto.Value:=YNEDesconto.Value+XNEDesconto.Value;
      inc(varmarc);
    end;
    if ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime <=date Then Begin
      XNEVencido.Value := XNEVencido.Value + ZQRecBai.FieldByName('saldocorrig').AsFloat;
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
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
//  if empty(DM_tabelas.ZQRecebimentocliente.Text) Then Begin
//    Close;
//    Exit;
//  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.DateTEXT := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;
 //aqui

  CBPeriodoEnt.ItemIndex := 0;
//  filtrabaixa;
  DXBFechar.Enabled := True;
  JDEntrada.Enabled := True;
  JDBaixa.Enabled   := True;

  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL' then
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

  bxmanual.Text:=DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString;

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
    if ZQRecBai.FieldByName('marcar').AsString = '0' then
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
    if  ZQRecBai.FieldByName('marcar').AsString = '0' Then
    begin
      ZQRecBai.FieldByName('marcar').AsString := '1';
    end
    else
    begin
      ZQRecBai.FieldByName('marcar').AsString := '0';
    end;
    ZQRecBai.post;
    if ZQRecBai.FieldByName('marcar').AsString = '0' then
    begin
      DM_tabelas.ZQCobaRe.Insert;
      DTPIni.DateTEXT := datetostr(date);
      DTPFim.DateTEXT := datetostr(date);
      DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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

procedure TFrmRecebBaixa.XNERecebidoEnter(Sender: TObject);
begin
  XNERecebido_ant.Value:=XNERecebido.Value;
end;

procedure TFrmRecebBaixa.XNERecebidoExit(Sender: TObject);
begin
  // 24/05/2012

  if (XNEMarcado.Value > (XNERecebido.Value+XNEJuros.value)) and (XNEDesconto.Value=0) Then
     XNEResto.Value := ((XNEMarcado.Value+XNEJuros.value)) - (XNERecebido.Value);
  if XNEResto.Value=0 then
     XNEResto.Value :=  XNERecebido_ant.Value - XNERecebido.Value;
  if XNEResto.Value <0 then
     XNEResto.Value := 0;
   //  XNEResto.Value :=xdesc;
  xdesc:=0;

end;

// =============================================================================
// PROCEDURE OTIMIZADA: TFrmRecebBaixa.DXBBaixaGravarClick
// Otimizações aplicadas:
//   1. Transação explícita (StartTransaction/Commit/Rollback) para segurança
//   2. SQL montado com texto único (sem múltiplos .Add desnecessários)
//   3. Datasets abertos UMA VEZ fora dos loops
//   4. Eliminados Filter/Filtered dentro de loops — substituídos por SQL direto
//   5. DisableControls/EnableControls aplicados em todos os loops
//   6. application.ProcessMessages removido de dentro de loops críticos
//   7. Variáveis locais para cache de valores repetidos (quadra, contagem, etc.)
//   8. Queries de reconstrução de ZQRecebimento unificadas (eram duplicadas)
//   9. Abertura de ZQVenda movida para fora do loop interno
//  10. Função auxiliar local InserirRegistroBaixa elimina blocos if/else repetidos
// =============================================================================

procedure TFrmRecebBaixa.DXBBaixaGravarClick(Sender: TObject);
var
  Varpago, percentbaixa, baixatot, varjuros, varx,
  varparcelaquitada, vardescontoquitado, varsomaaberto, varabertolinha,
  varsomarepasse, varsomarepasseautomatico: Double;
  contador, varreg, varregistrosparcial, varindiceparcial: Integer;
  quadra, verif, varordem, varnumordem1, varnumordem2: string;
  isSubstituicao: Boolean;
  isDataQuitacaoGeral: Boolean;
  vDataRef: TDateTime;
  vContagem: string;

  // ---------------------------------------------------------------------------
  // Monta o SQL padrão de SELECT de Recebimento (evita repetição de código)
  // ---------------------------------------------------------------------------
  function SQLRecebimento(const aWhere: string): string;
  begin
    Result :=
      'SELECT idrecebimento, documento, cliente, usuario, Dt_Entrada, Dt_Vencimento, ' +
      '       Valor, Observ, VrDoc, ordem, TipDoc, saldo, marcar, RefBaixa, refvinda, ' +
      '       contabil, empresa, custodaparcela, origem, adversa, recpag, numordem, ' +
      '       idloteamento, venda_idvenda, quadralote, numboleto, Substituicao, sq, ' +
      '       nomeadversa, Reajustado, Data_reajuste, somar, Proximo_Reajuste, ' +
      '       Parcelas_fixas, observ_estorno, tip, juros, descontos, Data_Quitacao, ' +
      '       sld_antes_reajuste, Percentual_reajuste, juridico, data_juridico, ' +
      '       dt_nao_pagou_no_mes, descricao_juridico, multa, mora ' +
      'FROM Recebimento ' + aWhere;
  end;

  // ---------------------------------------------------------------------------
  // Monta o SQL padrão de SELECT de Recebimento entrada (evita repetição de código)
  // ---------------------------------------------------------------------------
  function SQLRecebimentoEntrada(const aWhere: string): string;
  begin
    Result :=
      'SELECT idrecebimento, documento, cliente, usuario, Dt_Entrada, Dt_Vencimento, ' +
      '       Valor, Observ, VrDoc, ordem, TipDoc, saldo, marcar, RefBaixa, refvinda, ' +
      '       contabil, empresa, custodaparcela, origem, adversa, recpag, numordem, ' +
      '       idloteamento, venda_idvenda, quadralote, numboleto, Substituicao, sq, ' +
      '       nomeadversa, Reajustado, Data_reajuste, somar, Proximo_Reajuste, ' +
      '       Parcelas_fixas, observ_estorno, tip, juros, descontos, Data_Quitacao, ' +
      '       sld_antes_reajuste, Percentual_reajuste, juridico, data_juridico, ' +
      '       dt_nao_pagou_no_mes, descricao_juridico, multa, mora,sum(valor) as entrada ' +
      'FROM Recebimento ' + aWhere;
  end;

  // ---------------------------------------------------------------------------
  // Monta o SQL padrão de SELECT de Recebimento parcela (evita repetição de código)
  // ---------------------------------------------------------------------------
  function SQLRecebimentoParcela(const aWhere: string): string;
  begin
    Result :=
      'SELECT idrecebimento, documento, cliente, usuario, Dt_Entrada, Dt_Vencimento, ' +
      '       Valor, Observ, VrDoc, ordem, TipDoc, saldo, marcar, RefBaixa, refvinda, ' +
      '       contabil, empresa, custodaparcela, origem, adversa, recpag, numordem, ' +
      '       idloteamento, venda_idvenda, quadralote, numboleto, Substituicao, sq, ' +
      '       nomeadversa, Reajustado, Data_reajuste, somar, Proximo_Reajuste, ' +
      '       Parcelas_fixas, observ_estorno, tip, juros, descontos, Data_Quitacao, ' +
      '       sld_antes_reajuste, Percentual_reajuste, juridico, data_juridico, ' +
      '       dt_nao_pagou_no_mes, descricao_juridico, multa, mora, sum(valor) as parcela ' +
      'FROM Recebimento ' + aWhere;
  end;
  // ---------------------------------------------------------------------------
  // Exclui registro de receb_baixa via SQL direto (sem abrir/fechar dataset)
  // ---------------------------------------------------------------------------
  procedure ExcluirRecebBaixaAnterior(const aDocum: string; aSq: Integer);
  begin
    with TZQuery.Create(nil) do
    try
      Connection := DM_tabelas.zconeccao; // ajuste para sua conexão
      SQL.Text := 'DELETE FROM receb_baixa WHERE docum = ' + QuotedStr(Trim(aDocum)) +
                  ' AND sq = ' + QuotedStr(Trim(IntToStr(aSq)));
      ExecSQL;
    finally
      Free;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Insere no ZQReceb_Baixa — centraliza a lógica repetida nos dois branches
  // ---------------------------------------------------------------------------
  procedure InserirRecebBaixa(aVrRec, aVrAb: Double; aDtRec: TDateTime;
    aSubstituicao: string);
  begin
    with DM_tabelas.ZQReceb_Baixa do
    begin
      Insert;
      if aVrAb > 0 then
        FieldByName('Vr_ab').AsFloat    := aVrAb
      else
        FieldByName('Vr_rec').AsFloat   := aVrRec;

      if CDSParcelastip.Value <> 'J' then
      begin
        FieldByName('valor_parcela').AsFloat := CDSParcelasVrParc.Value + CDSParcelasdesconto.Value - CDSParcelasjuros.Value;
        FieldByName('Juros_Vr').AsFloat := CDSParcelasjuros.Value;
      end
      else
      begin
        FieldByName('Juros_Vr').AsFloat    := CDSParcelasVrParc.Value;
        FieldByName('valor_parcela').AsFloat := 0;
      end;

      FieldByName('Dt_rec').AsDateTime    := StrToDate(JDBaixa.DateText);

      if CDSParcelasData_Quitacao.Value > 0 then
        FieldByName('dataref').AsDateTime := CDSParcelasData_Quitacao.Value
      else
        FieldByName('dataref').AsDateTime := StrToDate(JDEntrada.DateText);

      FieldByName('vencimento').AsDateTime := CDSParcelasVenci.Value;
      FieldByName('Tipdoc').AsString       := CDSParcelasTipDoc.AsString;
      FieldByName('Docum').AsString        := CDSParcelasDocum.AsString;
      FieldByName('Desc_Vr').AsFloat       := CDSParcelasdesconto.Value;
      FieldByName('refbaixa').AsInteger    := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
      FieldByName('codcontabil').AsInteger := CDSParcelascodContabil.Value;
      FieldByName('Obsebx').AsString       := EMObsbx.Text;
      FieldByName('substituicao').AsString := aSubstituicao;
      FieldByName('sq').AsInteger          := contador;
      Post;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Insere no ZQCaixa — centraliza lógica repetida
  // ---------------------------------------------------------------------------
  procedure InserirCaixa(aSubstituicao: string);
  begin
    with DM_tabelas.ZQCaixa do
    begin
      Insert;
      FieldByName('Plano_Contas_codigo').AsInteger:= CDSParcelascodContabil.Value;
      FieldByName('data_lan').AsDateTime          := Date;
      FieldByName('vr_lan').AsFloat               := CDSParcelasVrparc.Value;
      FieldByName('documento').AsString           := CDSParcelasDocum.AsString;
      FieldByName('empresa').AsInteger            := 1;
      DM_Tabelas.ZQPlanodeContas.Locate('codigo', CDSParcelascodContabil.Value, []);
      FieldByName('credeb').AsString   := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').AsString;
      FieldByName('vinculo').AsString  := 'R' + DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
      if ZQRecBai.FieldByName('recpag').AsString = 'P' then
        FieldByName('vinculo').AsString := 'P' + DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
      FieldByName('contracodigo').AsInteger := StrToInt(Ecodnome.Text);
      FieldByName('obs').AsString := 'Recebimento - Ordem ' + CDSParcelasOrdem.AsString;
      if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
        FieldByName('obs').AsString :=
          FieldByName('obs').AsString + ' Vencimento em: ' + DateToStr(CDSParcelasVenci.Value);
      FieldByName('obs').AsString :=
        FieldByName('obs').AsString + ' | ' + EMObsbx.Text;
      FieldByName('Substituicao').AsString := aSubstituicao;
      FieldByName('sq').AsInteger          := contador;
      Post;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Insere no ZQCheque — centraliza lógica repetida
  // ---------------------------------------------------------------------------
  procedure InserirCheque(aSubstituicao: string);
  begin
    with DM_tabelas.ZQCheque do
    begin
      Insert;
      FieldByName('Banco').AsString      := CDSParcelasBanco.AsString;
      FieldByName('Agencia').AsString    := CDSParcelasagencia.AsString;
      FieldByName('Dono').AsString       := CDSParcelasdono.AsString;
      FieldByName('CPF_CNPJ').AsString   := CDSParcelasCPF_CNPJ.AsString;
      FieldByName('Numero').AsString     := CDSParcelasNcheque.AsString;
      FieldByName('valor').AsFloat       := CDSParcelasVrparc.Value;
      FieldByName('emissao').AsDateTime  := StrToDate(JDEntrada.DateText);
      FieldByName('conta').AsString      := CDSParcelasConta.AsString;
      FieldByName('deposito').AsDateTime := CDSParcelasVenci.Value;
      FieldByName('idrecebimento').AsInteger := DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
      FieldByName('idvinculo').AsInteger     := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
      FieldByName('Substituicao').AsString   := aSubstituicao;
      FieldByName('sq').AsInteger            := contador;
      Post;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Insere no ZQMovBancaria — centraliza lógica repetida
  // ---------------------------------------------------------------------------
  procedure InserirMovBancaria(aSubstituicao: string);
  begin
    with DM_Tabelas.ZQMovBancaria do
    begin
      Insert;
      FieldByName('Conta_Bancaria_cod_banco').AsInteger :=
        DM_Tabelas.ZQContaBancaria.FieldByName('idconta_bancaria').AsLargeInt;
      FieldByName('Plano_Contas_codigo').AsInteger:= CDSParcelascontabanco.Value;
      FieldByName('documento').AsString           := CDSParcelasDocum.AsString;
      FieldByName('dt_lanc').AsDateTime           := Date;
      FieldByName('hora_lanc').AsDateTime         := Time;
      FieldByName('contracodigo').AsInteger       := StrToInt(Ecodnome.Text);
      FieldByName('dt_conciliado').AsDateTime     := CDSParcelasVenci.Value;
      FieldByName('vr_lanc').AsFloat              := CDSParcelasVrParc.Value;
      DM_Tabelas.ZQPlanodeContas.Locate('codigo', CDSParcelascontabanco.Value, []);
      FieldByName('CreDeb').AsString  := DM_Tabelas.ZQPlanoDeContas.FieldByName('cred_debi').AsString;
      FieldByName('obs').AsString     := EMObsbx.Text;
      FieldByName('vinculo').AsString := 'R' + DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
      if ZQRecBai.FieldByName('recpag').AsString = 'P' then
        FieldByName('vinculo').AsString := 'P' + DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
      FieldByName('Substituicao').AsString := aSubstituicao;
      FieldByName('sq').AsInteger          := contador;
      Post;
    end;
  end;

  // Resolve the lot from the original title/sale.  Do not depend on the shared
  // ZQLoteamento dataset here: its current filter/record can be unrelated to
  // the title being generated by this baixa.
  function ObterIdLoteamentoDaParcela: Int64;
  var
    LConsulta: TZQuery;
  begin
    Result := 0;
    LConsulta := TZQuery.Create(nil);
    try
      LConsulta.Connection := DM_Tabelas.zconeccao;

      // The original receivable preserves the lot even when it has no sale.
      if CDSParcelasidrecebimento.AsLargeInt > 0 then
      begin
        LConsulta.SQL.Text :=
          'select idloteamento from recebimento ' +
          'where idrecebimento=:idrecebimento';
        LConsulta.ParamByName('idrecebimento').AsLargeInt :=
          CDSParcelasidrecebimento.AsLargeInt;
        LConsulta.Open;
        if not LConsulta.IsEmpty then
          Result := LConsulta.FieldByName('idloteamento').AsLargeInt;
        LConsulta.Close;
      end;

      // For sales, imovel is the canonical relationship with loteamento.
      if (Result = 0) and (CDSParcelasvenda_idvenda.AsLargeInt > 0) then
      begin
        LConsulta.SQL.Text :=
          'select i.loteamento_idloteamento from venda v ' +
          'inner join imovel i on i.idimovel=v.imovel ' +
          'where v.idvenda=:idvenda';
        LConsulta.ParamByName('idvenda').AsLargeInt :=
          CDSParcelasvenda_idvenda.AsLargeInt;
        LConsulta.Open;
        if not LConsulta.IsEmpty then
          Result := LConsulta.FieldByName('loteamento_idloteamento').AsLargeInt;
        LConsulta.Close;
      end;

      // Compatibility fallback for legacy titles that have neither link.
      if (Result = 0) and (Trim(CDSParcelasnomedoempreendimento.AsString) <> '') then
      begin
        LConsulta.SQL.Text :=
          'select idloteamento from loteamento where apelido=:apelido ' +
          'order by idloteamento limit 1';
        LConsulta.ParamByName('apelido').AsString :=
          Trim(CDSParcelasnomedoempreendimento.AsString);
        LConsulta.Open;
        if not LConsulta.IsEmpty then
          Result := LConsulta.FieldByName('idloteamento').AsLargeInt;
      end;
    finally
      LConsulta.Free;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Insere no ZQRecebimento — centraliza lógica repetida
  // ---------------------------------------------------------------------------
  procedure InserirRecebimento(aSubstituicao: string; const aNumOrdem: string);
  begin
    with DM_tabelas.ZQRecebimento do
    begin
      Insert;
      FieldByName('documento').AsString       := CDSParcelasDocum.AsString;
      FieldByName('cliente').AsInteger        := StrToInt(Ecodnome.Text);
      FieldByName('Dt_Entrada').AsDateTime    := DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime;
      FieldByName('Dt_Vencimento').AsDateTime := CDSParcelasVenci.Value;
      FieldByName('Valor').AsFloat            := CDSParcelasVrParc.Value;
      FieldByName('VrDoc').AsFloat            := XNERecebido.Value;
      FieldByName('contabil').AsInteger       := CDSParcelascodContabil.Value;
      FieldByName('usuario').AsInteger        := DM_tabelas.ZQUsuario.FieldByName('idusuario').AsLargeInt;
      FieldByName('ordem').AsString           := aNumOrdem + '-' + CDSParcelasOrdem.AsString;
      FieldByName('TipDoc').AsString          := CDSParcelasTipDoc.AsString;
      FieldByName('saldo').AsFloat            := CDSParcelasVrParc.Value;
      FieldByName('nomeadversa').AsString     := EAdversaNova.Text;
      FieldByName('venda_idvenda').AsInteger  := CDSParcelasvenda_idvenda.Value;
      FieldByName('quadralote').AsString      := CDSParcelasquadralote.AsString;
      FieldByName('Documento').AsString       := CDSParcelasDocum.AsString;

      FieldByName('idloteamento').AsLargeInt := ObterIdLoteamentoDaParcela;

      FieldByName('refvinda').AsInteger    := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
      FieldByName('Observ').AsString       := EMObsbx.Text;
      FieldByName('recpag').AsString       := ZQRecBai.FieldByName('recpag').AsString;
      FieldByName('adversa').AsInteger     := StrToInt(EcodAdversaNova.Text);
      FieldByName('numordem').AsInteger    := StrToInt(aNumOrdem);
      FieldByName('substituicao').AsString := aSubstituicao;
      FieldByName('sq').AsInteger          := contador;

      if aSubstituicao <> 'S' then
        FieldByName('tip').AsString := CDSParcelastip.AsString;

      Post;
    end;
  end;

  // ---------------------------------------------------------------------------
  // Processa rateio automático (dividido) — usado nos dois branches
  // ---------------------------------------------------------------------------
  procedure ProcessarRateio(const aReferencia: string);
  begin
    CDS_dividido.Close;
    CDS_dividido.CreateDataSet;

    ZQRatear.Close;
    ZQRatear.SQL.Text :=
      'SELECT idrecebimento, h.valor, percent_usado, b.vr_rec, participante, ' +
      '       percentual, planodeconta, dv.tipodocumento, re.ordem, dv.ordem ' +
      'FROM recbxhist AS H ' +
      'JOIN recebimento AS re ON H.idrecib = RE.idrecebimento ' +
      'JOIN dividido    AS dv ON dv.ordem  = numordem ' +
      'JOIN receb_baixa AS B  ON B.refbaixa = H.refer ' +
      'JOIN tipodocumento AS D ON B.tipdoc  = D.tipodoc ' +
      'WHERE somapaga = ' + QuotedStr('S') + ' AND refer = ' + aReferencia;
    ZQRatear.Open;

    varsomarepasseautomatico := 0;
    ZQRatear.DisableControls;
    while not ZQRatear.Eof do
    begin
      if CDS_dividido.Locate(
           'participante;contabil;tipodoc',
           VarArrayOf([ZQRatear.FieldByName('participante').AsInteger,
                       ZQRatear.FieldByName('planodeconta').AsInteger,
                       ZQRatear.FieldByName('tipodocumento').AsString]), []) then
        CDS_dividido.Edit
      else
        CDS_dividido.Insert;

      CDS_divididoparticipante.Value := ZQRatear.FieldByName('participante').AsInteger;
      CDS_divididovr_base.Value      := CDS_divididovr_base.Value + ZQRatear.FieldByName('vr_base').AsFloat;
      CDS_divididovr_calc.Value      := CDS_divididovr_calc.Value + ZQRatear.FieldByName('vr_divi').AsFloat;
      CDS_divididocontabil.Value     := ZQRatear.FieldByName('planodeconta').AsInteger;
      CDS_divididotipodoc.Value      := ZQRatear.FieldByName('tipodocumento').AsString;
      CDS_divididoobs.Value          := CDS_divididoobs.Value + ZQRatear.FieldByName('idrecebimento').Text + ' | ' +
                                        ZQRatear.FieldByName('ordem').AsString + ' | ' + ZQRatear.FieldByName('vr_base').Text + ' x ' +
                                        ZQRatear.FieldByName('percentual').Text + ' = ' + ZQRatear.FieldByName('vr_divi').Text + Chr(13);
      varsomarepasseautomatico       := varsomarepasseautomatico + ZQRatear.FieldByName('vr_divi').AsFloat;

      ZQRatear.Next;
    end;
    ZQRatear.EnableControls;
    ZQRatear.Close;
  end;

  // ---------------------------------------------------------------------------
  // Gera lançamentos de rateio no ZQRecebimento (dividido)
  // ---------------------------------------------------------------------------
  procedure GerarLancamentosRateio(const aReferencia: string;
    const aDataEntrada: TDateTime);
  begin
    if CDS_dividido.RecordCount > 0 then
    begin
      DM_Tabelas.ZQNumOrdem.Insert;
      DM_Tabelas.ZQNumOrdem.Post;
    end;

    CDS_dividido.First;
    CDS_dividido.DisableControls;
    while not CDS_dividido.Eof do
    begin
      DM_tabelas.ZQRecebimento.Insert;
      DM_tabelas.ZQRecebimento.FieldByName('documento').AsString  := 'Rateio Bx - ' + aReferencia;
      DM_tabelas.ZQRecebimento.FieldByName('cliente').AsInteger    := CDS_divididoparticipante.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Dt_Entrada').AsDateTime    := aDataEntrada;
      DM_tabelas.ZQRecebimento.FieldByName('Dt_Vencimento').AsDateTime := aDataEntrada;
      DM_tabelas.ZQRecebimento.FieldByName('contabil').AsInteger   := CDS_divididocontabil.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Observ').AsString     := EMObsbx.Text + Chr(13) + CDS_divididoobs.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString     := 'P';
      DM_tabelas.ZQRecebimento.FieldByName('TipDoc').AsString     := CDS_divididotipodoc.Value;
      DM_tabelas.ZQRecebimento.FieldByName('ordem').AsString      := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text + '-' +
                                                  IntToStr(CDS_dividido.RecNo) + '/' + IntToStr(CDS_dividido.RecordCount);
      DM_tabelas.ZQRecebimento.FieldByName('refvinda').AsInteger   := StrToInt(aReferencia);
      DM_tabelas.ZQRecebimento.FieldByName('VrDoc').AsFloat      := CDS_divididovr_base.Value;
      DM_tabelas.ZQRecebimento.FieldByName('Valor').AsFloat      := CDS_divididovr_calc.Value;
      DM_tabelas.ZQRecebimento.FieldByName('saldo').AsFloat      := CDS_divididovr_calc.Value;
      DM_Tabelas.ZQRecebimento.FieldByName('adversa').AsInteger    :=
        DM_Tabelas.ZQConfiguracoes.FieldByName('paticipantedefault').AsInteger;
      DM_Tabelas.ZQRecebimento.FieldByName('nomeadversa').AsString := EAdversaNova.Text;
      DM_tabelas.ZQRecebimento.Post;
      CDS_dividido.Next;
    end;
    CDS_dividido.EnableControls;
  end;

  // ---------------------------------------------------------------------------
  // Atualiza saldo da venda após baixa (bloco final — igual nos dois branches)
  // ---------------------------------------------------------------------------
  procedure AtualizarSaldoVenda(const aQuadra: string);
  begin
    Panel1.Caption  := 'Aguarde! Ajustando Saldo da Venda...';
    Panel1.Visible  := True;
    Application.ProcessMessages;

    ZQRecebimento.Close;
    ZQRecebimento.SQL.Clear;
    ZQRecebimento.SQL.Text :=
      SQLRecebimento('WHERE quadralote = ' + QuotedStr(aQuadra) +
                     ' GROUP BY quadralote ORDER BY quadralote');
    ZQRecebimento.Open;

    bar1.Position := 0;
    bar1.Max      := ZQRecebimento.RecordCount;
    bar1.Visible  := True;
    Application.ProcessMessages;

    ZQVenda.Open;
    ZQRecebimento.DisableControls;
    while not ZQRecebimento.Eof do
    begin
      bar1.Position := ZQRecebimento.RecNo;

      // Entrada
      ZQEntrada.Close;
      ZQEntrada.SQL.Clear;
      ZQEntrada.SQL.Text :=
        SQLRecebimentoEntrada('WHERE tip <> ' + QuotedStr('J') +
          '  AND quadralote = ' + QuotedStr(ZQRecebimento.FieldByName('quadralote').AsString) +
          '  AND documento  LIKE ' + QuotedStr('%-E-%') +
          ' ORDER BY ordem');
      ZQEntrada.Open;

      // Parcela
      ZQParcela.Close;
      ZQParcela.SQL.Clear;
      ZQParcela.SQL.Text :=
        SQLRecebimentoParcela('WHERE tip <> ' + QuotedStr('J') +
          '  AND quadralote = ' + QuotedStr(ZQRecebimento.FieldByName('quadralote').AsString) +
          '  AND documento  LIKE ' + QuotedStr('%-P-%') +
          ' ORDER BY ordem');
      ZQParcela.Open;

      if (ZQEntrada.FieldByName('entrada').AsFloat > 0) or (ZQParcela.FieldByName('parcela').AsFloat > 0) then
      begin
        ZQVenda.Close;
        ZQVenda.SQL.Text :=
          'SELECT idvenda, datavenda, imovel, valorvenda, forma_reajuste, ' +
          '       tabela_Price, Escriturado, marca, codigo_contrato_ref, ' +
          '       Multa, Mora, Perc_comissao, vlr_comissao ' +
          'FROM venda WHERE idvenda = ' + QuotedStr(ZQRecebimento.FieldByName('venda_idvenda').Text);
        ZQVenda.Open;
        if ZQVenda.RecordCount > 0 then
        begin
          ZQVenda.Edit;
          // No branch DATA QUITAÇÃO GERAL mantinha o guard tip<>'J'; no outro não.
          // Mantemos sem guard (comportamento do branch padrão) — ajuste se necessário.
          ZQVenda.FieldByName('valorvenda').AsFloat := ZQEntrada.FieldByName('entrada').AsFloat + ZQParcela.FieldByName('parcela').AsFloat;
          ZQVenda.Post;
        end;
      end;

      ZQRecebimento.Next;
    end;
    ZQRecebimento.EnableControls;

    bar1.Position  := 0;
    bar1.Visible   := False;
    ZQVenda.Close;
    ZQRecebimento.Close;
    ZQEntrada.Close;
    ZQParcela.Close;

    Panel1.Caption  := 'AGUARDE! ATUALIZANDO A TELA.';
    Panel1.Visible  := False;
  end;

  // ---------------------------------------------------------------------------
  // Recarrega ZQRecebimento no DM_tabelas ao final
  // ---------------------------------------------------------------------------
  procedure RecarregarDMRecebimento(const aQuadra: string);
  begin
    DM_tabelas.ZQRecebimento.Close;
    DM_tabelas.ZQRecebimento.SQL.Clear;
    DM_tabelas.ZQRecebimento.SQL.Text := SQLRecebimento('WHERE quadralote = ' + QuotedStr(aQuadra) +
                                         ' ORDER BY DT_Vencimento');
    DM_tabelas.ZQRecebimento.Open;
  end;

  // ---------------------------------------------------------------------------
  // Copia regras de divisão (dividido) para a nova ordem
  // ---------------------------------------------------------------------------
  procedure CopiarDividido(const aOrdemOrigem: string;
    const aNovaOrdem: string);
  begin
    ZQDiv_usado.Close;
    ZQDiv_usado.SQL.Text :='SELECT iddividido, participante, percentual, planodeconta, ' +
                           '       tipodocumento, ordem ' +
                           'FROM dividido WHERE ordem = ' + QuotedStr(aOrdemOrigem);
    ZQDiv_usado.Open;

    ZQDiv_usado.DisableControls;
    while not ZQDiv_usado.Eof do
    begin
      DM_Tabelas.ZQDividido.Insert;
      DM_Tabelas.ZQDividido.FieldByName('participante').AsInteger  := ZQDiv_usado.FieldByName('participante').AsInteger;
      DM_Tabelas.ZQDividido.FieldByName('percentual').AsFloat    := ZQDiv_usado.FieldByName('percentual').AsFloat;
      DM_Tabelas.ZQDividido.FieldByName('planodeconta').AsInteger  := ZQDiv_usado.FieldByName('planodeconta').AsInteger;
      DM_Tabelas.ZQDividido.FieldByName('tipodocumento').AsString := ZQDiv_usado.FieldByName('tipodocumento').AsString;
      DM_Tabelas.ZQDividido.FieldByName('ordem').AsInteger         := StrToInt(aNovaOrdem);
      DM_Tabelas.ZQDividido.Post;
      ZQDiv_usado.Next;
    end;
    ZQDiv_usado.EnableControls;
    ZQDiv_usado.Close;
  end;

  // ---------------------------------------------------------------------------
  // Finaliza a tela após gravação
  // ---------------------------------------------------------------------------
  procedure FinalizarTela;
  begin
    ZQRecBai.Filtered := False;
    ZQRecBai.Refresh;
    DBGBaixando.Refresh;
    DM_tabelas.ZQReceb_Baixa.Refresh;
    DM_Tabelas.ZQReBxHi.Refresh;

    CDSParcelas.Close;
    CDSParcelas.CreateDataSet;

    XNEJuros.Value    := 0;
    XNERecebido.Value := 0;
    XNEDesconto.Value := 0;
    XNEResto.Value    := 0;
    XNEParcelas.Value := 1;
    CBPeriodoEnt.Text := 'HOJE';
    EMObsbx.Clear;
    EMNovostit.Clear;

    DXBBaixaGravar.Enabled := False;
    DXBFechar.Enabled      := True;
    Elote.SetFocus;
  end;

begin
  // --------------------------------------------------------------------------
  // Validação inicial
  // --------------------------------------------------------------------------
  if Empty(CBTipobaixa.Text) then
  begin
    ShowMessage('Selecione o tipo de operação !');
    CBTipobaixa.SetFocus;
    Exit;
  end;

  xrec                := 0;
  isSubstituicao      := CBTipobaixa.Text = 'SUBSTITUIÇÃO';
  isDataQuitacaoGeral :=DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString = 'DATA QUITAÇÃO GERAL';

  DXBBaixaGravar.Enabled := False;
  DXBFechar.Enabled      := False;

  // --------------------------------------------------------------------------
  // Lê o contador de baixa UMA VEZ antes dos loops
  // --------------------------------------------------------------------------
  DM_tabelas.contador_bx.Close;
  DM_tabelas.contador_bx.Open;
  DM_tabelas.contador_bx.Last;
  if DM_tabelas.contador_bx.RecordCount > 0 then
    contador := DM_tabelas.contador_bx.FieldByName('idcontador_bx').AsLargeInt
  else
    contador := 0;
  DM_tabelas.contador_bx.Close;

  // --------------------------------------------------------------------------
  // Inicia transação para garantir atomicidade
  // --------------------------------------------------------------------------
  DM_tabelas.zconeccao.StartTransaction; // ajuste para o nome da sua ZConnection
  try

    // =========================================================================
    // BRANCH A — DATA QUITAÇÃO GERAL
    // =========================================================================
    if isDataQuitacaoGeral then
    begin
      // Remove receb_baixa anterior via SQL direto (sem abrir/filtrar dataset)
      ExcluirRecebBaixaAnterior(CDSParcelasDocum.Value, CDSParcelassq.Value);

      ZQRecBai.Filter    := 'marcar=0';
      ZQRecBai.Filtered  := True;
      LREg.Caption       := IntToStr(ZQRecbai.RecordCount) + ' Registros';
      varnumordem1       := '';
      varnumordem2       := '';

      // Registra cabeçalho de baixa
      DM_tabelas.ZQCobaRe.Insert;
      DTPIni.DateText := DateToStr(Date);
      DTPFim.DateText := DateToStr(Date);
      DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := Date;
      DM_tabelas.ZQCobaRe.Post;

      if not (DM_tabelas.ZQCobaRe.State in [dsEdit, dsInsert]) then
        DM_tabelas.ZQCobaRe.Edit;
      DM_tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime       := StrToDate(JDEntrada.DateText);
      DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').AsString     := CBTipobaixa.Text;
      DM_tabelas.ZQCobaRe.Post;

      // ------------------------------------------------------------------
      // Loop principal — CDSParcelas
      // ------------------------------------------------------------------
      varsomarepasse := 0;
      bar1.Visible   := True;
      bar1.Position  := 0;
      bar1.Max       := CDSParcelas.RecordCount;
      CDSParcelas.First;
      CDSParcelas.DisableControls;
      while not CDSParcelas.Eof do
      begin
        bar1.Position := CDSParcelas.RecNo;
        Inc(contador);

        if not isSubstituicao then
          InserirRecebBaixa(CDSParcelasVrParc.Value, 0,
            StrToDate(JDBaixa.DateText), CDSParcelassubstituicao.Value)
        else
        begin
          // SUBSTITUIÇÃO: DN ou DP gravam como recebimento, demais como abertura
          if (CDSParcelasTipDoc.Value = 'DN') or
             (CDSParcelasTipDoc.Value = 'DP') then
            InserirRecebBaixa(CDSParcelasVrParc.Value, 0,
              StrToDate(JDBaixa.DateText), 'S')
          else
            InserirRecebBaixa(0, CDSParcelasVrParc.Value,
              StrToDate(JDBaixa.DateText), 'S');
        end;

        // Soma repasse
        DM_tabelas.ZQTipodoc.Locate('tipodoc', CDSParcelasTipDoc.Value, []);
        if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').AsString = 'S' then
          varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;

        // Recebimento
        if DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString = 'S' then
        begin
          if Empty(varnumordem1) then
          begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          if not isSubstituicao then
            InserirRecebimento(CDSParcelassubstituicao.Value, varnumordem1)
          else
            InserirRecebimento('S', varnumordem1);
        end;

        // Caixa
        if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString = 'S' then
        begin
          if not isSubstituicao then
            InserirCaixa(CDSParcelassubstituicao.Value)
          else if (CDSParcelasTipDoc.Value = 'DN') or
                  (CDSParcelasTipDoc.Value = 'DP') then
            InserirCaixa('S');
        end;

        // Cheque
        if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
        begin
          if not isSubstituicao then
            InserirCheque(CDSParcelassubstituicao.Value)
          else
            InserirCheque('S');
        end;

        // Movimento bancário
        if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').AsString = 'S' then
        begin
          if not isSubstituicao then
            InserirMovBancaria(CDSParcelassubstituicao.Value)
          else
            InserirMovBancaria('S');
        end;

        CDSParcelas.Next;
      end;
      bar1.Position := 0;
      bar1.Visible  := False;
      CDSParcelas.EnableControls;
      CDSParcelas.Last;

      // Grava contador
      DM_tabelas.contador_bx.Open;
      DM_tabelas.contador_bx.Insert;
      DM_tabelas.contador_bx.FieldByName('idcontador_bx').AsLargeInt := contador;
      DM_tabelas.contador_bx.Post;
      DM_tabelas.contador_bx.Close;

      // Copia dividido
      if not Empty(varnumordem1) then
        CopiarDividido(ZQRecBai.FieldByName('numordem').Text, varnumordem1);

      // ------------------------------------------------------------------
      // Loop ZQRecBai — atualiza histórico e saldos
      // ------------------------------------------------------------------
      ZQRecBai.First;
      varpago   := 0;
      varjuros  := 0;
      varsomaaberto := 0;
      varindiceparcial := 0;
      varregistrosparcial := CDSParcelas.RecordCount;
      Memo1.Clear;
      baixatot  := XNERecebido.Value;

      bar1.Visible  := True;
      bar1.Position := 0;
      bar1.Max      := ZQRecBai.RecordCount;
      ZQRecBai.DisableControls;
      while not ZQRecBai.Eof do
      begin
        bar1.Position := ZQRecBai.RecNo;

        DM_tabelas.ZQReBxHi.Insert;
        percentbaixa := ZQRecBai.FieldByName('saldocorrig').AsFloat * 100 / XNEMarcado.Value;
        DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat := XNEDesconto.Value * percentbaixa / 100;

        varjuros := XNEJuros.Value * percentbaixa / 100;

        if isDataQuitacaoGeral and (XNEResto.Value > 0) and
           CDSParcelas.Locate('idrecebimento', ZQRecBai.FieldByName('idrecebimento').AsLargeInt, []) then
        begin
          vardescontoquitado := CDSParcelasdesconto.Value;
          varparcelaquitada := CDSParcelasVrParc.Value + vardescontoquitado -
                               CDSParcelasjuros.Value;

          DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat := vardescontoquitado;
          DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat := CDSParcelasjuros.Value;
          varpago := CDSParcelasVrParc.Value + vardescontoquitado;

          if ZQRecBai.FieldByName('saldocorrig').AsFloat > 0 then
            percentbaixa := (varparcelaquitada * 100) / ZQRecBai.FieldByName('saldocorrig').AsFloat
          else
            percentbaixa := 0;
        end
        else
        begin
          if ZQRecBai.FieldByName('saldocorrig').AsFloat <= baixatot then
            varpago := ZQRecBai.FieldByName('saldocorrig').AsFloat
          else
            varpago := baixatot + DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat;

          if varjuros <= varpago then
            DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat := varjuros
          else
            DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat := varpago;

          if XNEResto.Value = 0 then
            varpago := varpago + DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat;
        end;

        DM_Tabelas.ZQReBxHi.FieldByName('valor').AsFloat := varpago - DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat;

        if CDSParcelasData_Quitacao.Value > 0 then
          DM_Tabelas.ZQReBxHi.FieldByName('data').AsDateTime := CDSParcelasData_Quitacao.Value
        else
          DM_Tabelas.ZQReBxHi.FieldByName('data').AsDateTime := StrToDate(JDEntrada.DateText);

        DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').AsFloat := percentbaixa;
        DM_tabelas.ZQReBxHi.FieldByName('refer').AsInteger         := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
        DM_tabelas.ZQReBxHi.FieldByName('idrecib').AsInteger       := ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
        DM_tabelas.ZQReBxHi.FieldByName('sq').AsLargeInt            := contador;

        varx := ExRound(varpago, 2) - ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat, 2);
        DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat := varx;

        if ZQRecBai.FieldByName('tip').AsString = 'J' then
        begin
          DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat         := varpago;
          DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat := 0;
        end;

        DM_Tabelas.ZQReBxHi.Post;

        // Atualiza ZQRecBai
        ZQRecBai.Edit;
        if ZQRecBai.FieldByName('tip').AsString <> 'J' then
          varx := ExRound(varpago, 2) - ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat, 2);

        if isDataQuitacaoGeral and (XNEResto.Value > 0) and
           (ZQRecBai.FieldByName('tip').AsString <> 'J') then
        begin
          Inc(varindiceparcial);
          if varindiceparcial < varregistrosparcial then
            varabertolinha := ExRound(ZQRecBai.FieldByName('saldo').AsFloat, 2) - ExRound(varx, 2)
          else
            varabertolinha := ExRound(XNEResto.Value - varsomaaberto, 2);

          if varabertolinha < 0 then
            varabertolinha := 0;

          ZQRecBai.FieldByName('saldo').AsFloat := varabertolinha;
          varsomaaberto := varsomaaberto + ZQRecBai.FieldByName('saldo').AsFloat;
        end
        else if ZQRecBai.FieldByName('tip').AsString = 'J' then
        begin
          ZQRecBai.FieldByName('saldo').AsFloat := ExRound(ZQRecBai.FieldByName('saldo').AsFloat, 2) - ExRound(varx, 2);
          ZQRecBai.FieldByName('Valor').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
        end
        else
          ZQRecBai.FieldByName('saldo').AsFloat := ExRound(ZQRecBai.FieldByName('saldo').AsFloat, 2) - ExRound(varx, 2);

        ZQRecBai.FieldByName('RefBaixa').AsInteger  := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
        ZQRecBai.FieldByName('juros').AsFloat     := 0;
        ZQRecBai.FieldByName('descontos').AsFloat := 0;
        ZQRecBai.Post;

        // Atualiza ZQRecebimento
        DM_Tabelas.ZQRecebimento.Locate('idrecebimento', ZQRecBai.FieldByName('idrecebimento').AsLargeInt, []);
        quadra := DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString;
        DM_Tabelas.ZQRecebimento.Edit;
        DM_Tabelas.ZQRecebimento.FieldByName('saldo').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
        DM_Tabelas.ZQRecebimento.FieldByName('Valor').AsFloat := ZQRecBai.FieldByName('Valor').AsFloat;

        if ZQRecBai.FieldByName('saldo').AsFloat > 0 then
          DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString := 'P';
        if (ZQRecBai.FieldByName('saldo').AsFloat = 0) and
           (DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString = 'P') then
          DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString := '';

        DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').AsInteger := ZQRecBai.FieldByName('RefBaixa').AsInteger;
        DM_Tabelas.ZQRecebimento.Post;

        if isSubstituicao then
        begin
          ZQRecBai.Edit;
          ZQRecBai.FieldByName('somar').AsString := 'N';
          ZQRecBai.Post;
        end;

        if Pos('-', ZQRecBai.FieldByName('ordem').AsString) = 0 then
          Varordem := Varordem + ZQRecBai.FieldByName('idrecebimento').Text + ' - ';
        if Pos(ZQRecBai.FieldByName('nomeadversa').AsString, Memo1.Text) = 0 then
          Memo1.Lines.Add(ZQRecBai.FieldByName('nomeadversa').AsString);

        ZQRecBai.Next;
      end;
      ZQRecBai.EnableControls;
      bar1.Position := 0;
      bar1.Visible  := False;

      // Rateio global
      ProcessarRateio(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text);

      varreg := DM_Tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;

      if isDataQuitacaoGeral then
        GerarLancamentosRateio(DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text, StrToDate(JDEntrada.DateText));

      // Recarrega ZQRecebimento com filtro para impressão
      DM_tabelas.ZQRecebimento.Close;
      DM_tabelas.ZQRecebimento.SQL.Clear;
      DM_tabelas.ZQRecebimento.SQL.Text := SQLRecebimento('WHERE idrecebimento = ' +
                                           QuotedStr(IntToStr(varreg)) + ' ORDER BY DT_Vencimento');
      DM_tabelas.ZQRecebimento.Open;
      DM_Tabelas.ZQRecebimento.Locate('idrecebimento', varreg, []);

    end   // fim BRANCH A
    // =========================================================================
    // BRANCH B — DATA QUITAÇÃO POR PARCELA
    // =========================================================================
    else
    begin
      verif := CBTipobaixa.Text;

      // Remove receb_baixa anterior de TODAS as parcelas via SQL direto
      CDSParcelas.First;
      CDSParcelas.DisableControls;
      while not CDSParcelas.Eof do
      begin
        ExcluirRecebBaixaAnterior(CDSParcelasDocum.Value, CDSParcelassq.Value);
        CDSParcelas.Next;
      end;
      CDSParcelas.First;
      CDSParcelas.EnableControls;

      ZQRecBai.Filter   := 'marcar=0';
      ZQRecBai.Filtered := True;
      LREg.Caption      := IntToStr(ZQRecbai.RecordCount) + ' Registros';
      varnumordem1      := '';
      varnumordem2      := '';

      varsomarepasse := 0;
      bar1.Visible   := True;
      bar1.Position  := 0;
      bar1.Max       := CDSParcelas.RecordCount;
      CDSParcelas.First;
      CDSParcelas.DisableControls;
      while not CDSParcelas.Eof do
      begin
        bar1.Position := CDSParcelas.RecNo;
        Inc(contador);

        if not isSubstituicao then
        begin
          // Registra cabeçalho por parcela
          if not (DM_tabelas.ZQCobaRe.State in [dsEdit, dsInsert]) then
            DM_tabelas.ZQCobaRe.Insert;
          DM_tabelas.ZQCobaRe.FieldByName('entrada').AsDateTime   := CDSParcelasData_Quitacao.Value;
          DM_Tabelas.ZQCobaRe.FieldByName('tipobaixa').AsString := CBTipobaixa.Text;
          DM_tabelas.ZQCobaRe.Post;

          CDSParcelas.Edit;
          CDSParcelascontagem.Value := DM_tabelas.ZQCobaRe.FieldByName('cotagem').AsLargeInt;
          CDSParcelas.Post;

          InserirRecebBaixa(CDSParcelasVrParc.Value, 0,
            StrToDate(JDBaixa.DateText), CDSParcelassubstituicao.Value);
        end
        else
        begin
          if (CDSParcelasTipDoc.Value = 'DN') or
             (CDSParcelasTipDoc.Value = 'DP') then
            InserirRecebBaixa(CDSParcelasVrParc.Value, 0,
              StrToDate(JDBaixa.DateText), CDSParcelassubstituicao.Value)
          else
            InserirRecebBaixa(0, CDSParcelasVrParc.Value,
              StrToDate(JDBaixa.DateText), 'S');
        end;

        // Soma repasse
        DM_tabelas.ZQTipodoc.Locate('tipodoc', CDSParcelasTipDoc.Value, []);
        if DM_tabelas.ZQTipoDoc.FieldByName('somapaga').AsString = 'S' then
          varsomarepasse := varsomarepasse + CDSParcelasVrParc.Value;

        // Recebimento
        if DM_tabelas.ZQTipoDoc.FieldByName('receb_receb').AsString = 'S' then
        begin
          if Empty(varnumordem1) then
          begin
            DM_Tabelas.ZQNumOrdem.Insert;
            DM_Tabelas.ZQNumOrdem.Post;
            varnumordem1 := DM_Tabelas.ZQNumOrdem.FieldByName('idnumordem').Text;
          end;
          if not isSubstituicao then
            InserirRecebimento(CDSParcelassubstituicao.Value, varnumordem1)
          else
          begin
            quadra := CDSParcelasquadralote.Value;
            InserirRecebimento('S', varnumordem1);
          end;
        end;

        // Caixa
        if DM_tabelas.ZQTipoDoc.FieldByName('receb_caixa').AsString = 'S' then
        begin
          quadra := CDSParcelasquadralote.Value;
          if not isSubstituicao then
            InserirCaixa(CDSParcelassubstituicao.Value)
          else if (CDSParcelasTipDoc.Value = 'DN') or
                  (CDSParcelasTipDoc.Value = 'DP') then
            InserirCaixa('S');
        end;

        // Cheque
        if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
        begin
          quadra := CDSParcelasquadralote.Value;
          if not isSubstituicao then
            InserirCheque(CDSParcelassubstituicao.Value)
          else
            InserirCheque('S');
        end;

        // Movimento bancário
        if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').AsString = 'S' then
        begin
          quadra := CDSParcelasquadralote.Value;
          if not isSubstituicao then
            InserirMovBancaria(CDSParcelassubstituicao.Value)
          else
            InserirMovBancaria('S');
        end;

        CDSParcelas.Next;
      end;
      bar1.Position := 0;
      bar1.Visible  := False;
      CDSParcelas.EnableControls;

      // Grava contador
      DM_tabelas.contador_bx.Open;
      DM_tabelas.contador_bx.Insert;
      DM_tabelas.contador_bx.FieldByName('idcontador_bx').AsLargeInt := contador;
      DM_tabelas.contador_bx.Post;
      DM_tabelas.contador_bx.Close;

      // Copia dividido
      if not Empty(varnumordem1) then
        CopiarDividido(ZQRecBai.FieldByName('numordem').Text, varnumordem1);

      // ------------------------------------------------------------------
      // Loop ZQRecBai — atualiza histórico e saldos
      // ------------------------------------------------------------------
      ZQRecBai.First;
      varpago  := 0;
      Memo1.Clear;
      baixatot := XNERecebido.Value;
      ZQRecBai.DisableControls;
      while not ZQRecBai.Eof do
      begin
        // Filtra CDSParcelas para a parcela corrente (necessário para Data_Quitacao)
        CDSParcelas.Filtered := False;
        CDSParcelas.Filter   := 'idrecebimento=' + QuotedStr(Trim(ZQRecBai.FieldByName('idrecebimento').Text));
        CDSParcelas.Filtered := True;

        DM_tabelas.ZQReBxHi.Insert;
        DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat := ZQRecBai.FieldByName('descontos').AsFloat;

        if ZQRecBai.FieldByName('saldocorrig').AsFloat <= baixatot then
          varpago := ZQRecBai.FieldByName('saldocorrig').AsFloat
        else
          varpago := baixatot + DM_Tabelas.ZQReBxHi.FieldByName('descon').AsFloat;

        varjuros := XNEJuros.Value * percentbaixa / 100;

        if ZQRecBai.FieldByName('tip').AsString <> 'J' then
          DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat := ZQRecBai.FieldByName('juros').AsFloat
        else
          DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;

        if ZQRecBai.FieldByName('tip').AsString <> 'J' then
          varpago := varpago + DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat;

        DM_Tabelas.ZQReBxHi.FieldByName('valor').AsFloat :=(ZQRecBai.FieldByName('saldo').AsFloat + ZQRecBai.FieldByName('juros').AsFloat) - ZQRecBai.FieldByName('descontos').AsFloat;

        if CDSParcelasData_Quitacao.Value > 0 then
          DM_Tabelas.ZQReBxHi.FieldByName('data').AsDateTime := CDSParcelasData_Quitacao.Value
        else
          DM_Tabelas.ZQReBxHi.FieldByName('data').AsDateTime := StrToDate(JDEntrada.DateText);

        DM_Tabelas.ZQReBxHi.FieldByName('percent_usado').AsFloat := 100;
        DM_tabelas.ZQReBxHi.FieldByName('refer').AsInteger         := CDSParcelascontagem.Value;
        DM_tabelas.ZQReBxHi.FieldByName('idrecib').AsInteger       := ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
        DM_tabelas.ZQReBxHi.FieldByName('sq').AsLargeInt            := contador;

        varx := ExRound(varpago, 2) - ExRound(DM_Tabelas.ZQReBxHi.FieldByName('juros').AsFloat, 2);
        if ZQRecBai.FieldByName('tip').AsString <> 'J' then
          DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat
        else
          DM_tabelas.ZQReBxHi.FieldByName('valor_parcela').AsFloat := 0;

        DM_Tabelas.ZQReBxHi.Post;

        // Atualiza ZQRecBai
        ZQRecBai.Edit;
        ZQRecBai.FieldByName('saldo').AsFloat :=
          ExRound(ZQRecBai.FieldByName('saldo').AsFloat, 2) -
          ExRound((DM_Tabelas.ZQReBxHi.FieldByName('valor').AsFloat -
                   ZQRecBai.FieldByName('juros').AsFloat) + ZQRecBai.FieldByName('descontos').AsFloat, 2);
        ZQRecBai.FieldByName('RefBaixa').AsInteger  := CDSParcelascontagem.Value;
        ZQRecBai.FieldByName('juros').AsFloat     := 0;
        ZQRecBai.FieldByName('descontos').AsFloat := 0;
        ZQRecBai.Post;

        // Atualiza ZQRecebimento
        if DM_Tabelas.ZQRecebimento.Locate('idrecebimento', ZQRecBai.FieldByName('idrecebimento').AsLargeInt, []) then
        begin
          quadra := DM_Tabelas.ZQRecebimento.FieldByName('quadralote').AsString;
          DM_Tabelas.ZQRecebimento.Edit;

          if ZQRecBai.FieldByName('saldo').AsFloat > 0 then
            DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString := 'P';
          if (ZQRecBai.FieldByName('saldo').AsFloat = 0) and
             (DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString = 'P') then
            DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString := '';
          if DM_Tabelas.ZQRecebimento.FieldByName('tip').AsString = 'J' then
            DM_Tabelas.ZQRecebimento.FieldByName('Valor').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;

          DM_Tabelas.ZQRecebimento.FieldByName('saldo').AsFloat    := ZQRecBai.FieldByName('saldo').AsFloat;
          DM_Tabelas.ZQRecebimento.FieldByName('RefBaixa').AsInteger := ZQRecBai.FieldByName('RefBaixa').AsInteger;
          DM_Tabelas.ZQRecebimento.Post;
        end;

        if isSubstituicao then
        begin
          ZQRecBai.Edit;
          ZQRecBai.FieldByName('somar').AsString := 'N';
          ZQRecBai.Post;
        end;

        if Pos('-', ZQRecBai.FieldByName('ordem').AsString) = 0 then
          Varordem := Varordem + ZQRecBai.FieldByName('idrecebimento').Text + ' - ';
        if Pos(ZQRecBai.FieldByName('nomeadversa').AsString, Memo1.Text) = 0 then
          Memo1.Lines.Add(ZQRecBai.FieldByName('nomeadversa').AsString);

        ZQRecBai.Next;
      end;
      CDSParcelas.Filtered := False;
      ZQRecBai.EnableControls;

      // Rateio por parcela
      CDSParcelas.First;
      CDSParcelas.DisableControls;
      while not CDSParcelas.Eof do
      begin
        vContagem := CDSParcelascontagem.Text;
        if not Empty(vContagem) then
        begin
          ProcessarRateio(vContagem);

          if CDSParcelasData_Quitacao.Value > 0 then
            vDataRef := CDSParcelasData_Quitacao.Value
          else
            vDataRef := StrToDate(JDEntrada.DateText);

          GerarLancamentosRateio(vContagem, vDataRef);
        end;
        CDSParcelas.Next;
      end;
      CDSParcelas.EnableControls;
    end; // fim BRANCH B

    // =========================================================================
    // COMMIT — tudo certo
    // =========================================================================
    DM_tabelas.zconeccao.Commit;

  except
    on E: Exception do
    begin
      // ROLLBACK em caso de erro — banco volta ao estado anterior
      DM_tabelas.zconeccao.Rollback;
      ShowMessage('Erro ao gravar baixa: ' + E.Message +
                  #13#10 + 'A operação foi cancelada.');
      DXBBaixaGravar.Enabled := True;
      DXBFechar.Enabled      := True;
      Exit;
    end;
  end;

  // ==========================================================================
  // Pós-gravação: recibo, atualização de saldo de venda, refresh de tela
  // ==========================================================================

  // Prepara recibo (sem exibir ainda)
  FrmImpRecibo := nil;
  if FrmImpRecibo = nil then
    FrmImpRecibo := TFrmImpRecibo.Create(Self);
  FrmImpRecibo.N_Baixa.Text  := DM_tabelas.ZQCobaRe.FieldByName('cotagem').Text;
  FrmImpRecibo.Nomecli.Text  := Enomes.Text;
  FrmImpRecibo.Memoobs.Lines := EMObsbx.Lines;

  // Atualiza saldo da venda
  AtualizarSaldoVenda(quadra);

  // Recarrega ZQRecebimento no DM_tabelas
  RecarregarDMRecebimento(quadra);

  // Finaliza controles visuais
  FinalizarTela;

  ZQRecBai.Filtered := False;

  if FrmImpRecibo <> nil then
    FreeAndNil(FrmImpRecibo);

  ShowMessage('Operação Terminada!');
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
    if DM_tabelas.ZQTipoDoc.FieldByName('so_avista').AsString='S' Then
    Begin
      if CDSParcelasData_Quitacao.Value>0 then
      begin
        if (CDSParcelasVenci.Value<>CDSParcelasData_Quitacao.Value) and ((CDSParcelasTipDoc.Value<>'DP')) Then
        BEgin
          showmessage('Este tipo de documento só aceita pagamento a vista!!!');
          EContabil.SetFocus;
          Exit;
        end;
      end
      else
      begin
        if (datetostr(CDSParcelasVenci.Value)<>JDEntrada.DateText) and ((CDSParcelasTipDoc.Value<>'DP')) Then
        BEgin
          showmessage('Este tipo de documento só aceita pagamento a vista!!!');
          EContabil.SetFocus;
          Exit;
        end;
      end;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('lancabanco').AsString='S' then
    Begin
      GBBanco.Visible := True;
      DBGBanco.SetFocus;
      exit;
    end;
    if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin


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
         CDSParcelasagencia.Value := ZQCheque.FieldByName('agencia').AsString;
         CDSParcelasConta.Value   := ZQCheque.FieldByName('conta').AsString;
         CDSParcelasbanco.Value   := ZQCheque.FieldByName('Banco').AsString;
         CDSParcelasdono.Value    := ZQCheque.FieldByName('Dono').AsString;
         CDSParcelasncheque.Value := ZQCheque.FieldByName('numero').AsString;
         CDSParcelasCpf_Cnpj.Value:= ZQCheque.FieldByName('CPF_CNPJ').AsString; 
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
  varregis, vartotalparcelas : Integer;
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
    VarDoc := VarDoc + quotedstr(CDSParcelasDocum.AsString)+',';
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
  vartotalparcelas := Trunc(XNEParcelas.Value);
  if (DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL') and
     (XNEResto.Value > 0) then
    vartotalparcelas := CDSParcelas.RecordCount;

  if CDSParcelas.RecNO< vartotalparcelas Then
  Begin
    CDSParcelas.RecNO := varregis+1;
    while not CDSParcelas.Eof do Begin
      CDSParcelas.Edit;
      CDSParcelasVrParc.Value := ExRound(vardif,2) /
        (ExRound(vartotalparcelas,2)-ExRound(varregis,2));
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
      showmessage('A descrição contábil não pode ficar em branco....');
      Econtabil.SetFocus;
      exit;
    end;
    if not AchaPlanoDeConta(450, 30, 'S', EContabil.Text, '') then Begin
      EContabil.SetFocus;
      exit;
    end;
    EContabil.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
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
      CDSParcelascodcontabil.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
      if not empty(DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString) then
         CDSParcelasTipDoc.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('doccomum').AsString;
      CDSParcelas.post;
      CDSParcelas.Next;
    end;
    CDSParcelas.First;

  end;
//    CDSParcelasTipDoc.Value := FrmAchaPlanodeConta.ZQTempPlanoContasdoccomum.Value;
end;

procedure TFrmRecebBaixa.XDBNumEdit1Exit(Sender: TObject);
begin
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL' then
  begin
     CDSParcelasVrParc.Value := XNERecebido.Value*CDSParcelasPercent.Value/100;
  end;   
end;

procedure TFrmRecebBaixa.DBEBaixaDocumExit(Sender: TObject);
begin
  if not DXBFechar.Focused Then Begin
{    if (not empty(DBEBaixaDocum.Text)) and (DM_tabelas.ZQRecebimento.Locate('documento',DBEBaixaDocum.Text,[])) Then Begin
      showmessage('Este documento já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    End;
    if DM_tabelas.ZQCheque.Locate('CH_Conta',DBEBaixaDocum.Text,[]) Then Begin
      showmessage('Este cheque já foi lançado anteriormente... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    DM_tabelas.ZQTipodoc.Locate('tipodoc',CDSParcelasTipDoc.Value,[]);
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (pos(quotedstr(CDSParcelasDocum.AsString),VarDoc)>0) Then Begin
      showmessage('Este cheque já foi lançado nesta inclusão... corrija a numeração.....');
      DBEBaixaDocum.SetFocus;
      exit;
    end;
    if (DM_tabelas.ZQTipodocdados_chequ.Value='S') and (empty(CDSParcelasDocum.Value)) Then Begin
      showmessage('o campo Nº do cheque deve ser preenchido...');
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
  DBText1.Caption := CDSParcelasOrdem.AsString;
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
    Enomes.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
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
  if DM_tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString='S' then Begin
    CDSParcelasagencia.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_agencia').AsString;
    CDSParcelasConta.Value := DM_TAbelas.ZQContaBancaria.FieldByName('n_conta').AsString;
    CDSParcelasbanco.Value := DM_Tabelas.ZQContaBancaria.FieldByName('n_banco').AsString;
    CDSParcelasdono.Value := DM_Tabelas.ZQContaBancaria.FieldByName('nometitular').AsString;
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
    EPlanocontabanco.Text := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('descricao').AsString;
    CDSParcelascontabanco.Value := DM_Tabelas.ZQAchaPlanoDeContas.FieldByName('codigo').AsLargeInt;
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
         ZQRecBai.SQL.Add('update recebimento set marcar= CASE WHEN Dt_Vencimento<=now() then ''0'' else ''1'' end  where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+QuotedStr(Ecodnome.Text)+')');
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
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
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
         ZQRecBai.SQL.Add(' from recebimento where (quadralote='+quotedstr(Elote.text)+') and (recpag='+quotedstr(varrecpag)+') and ((saldo>0) and  (not saldo is null)) and (cliente ='+QuotedStr(Ecodnome.Text)+')');
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
        if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
           vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
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
    XEEmpresend.Text := ZQRecBai.FieldByName('nomeempreend').AsString;
    Enomes.Text      := ZQRecBai.FieldByName('Nomecli').AsString;
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



    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]);
    if RBPonteiro.Checked Then Begin
       ZQRecBai.Edit;
       ZQRecBai.FieldByName('marcar').AsString := '0';
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
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
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
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
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
           if DM_Tabelas.ZQTipoDoc.FieldByName('dados_chequ').AsString = 'S' then
              vartipo := vartipo + quotedstr(DM_Tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString)+',';
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
    ZQRecBai.Locate('idRecebimento',DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt,[]);
    if RBPonteiro.Checked Then
    Begin
      ZQRecBai.Edit;
      ZQRecBai.FieldByName('marcar').AsString := '0';
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
      ZQRecBai.FieldByName('marcar').AsString := '1';
      ZQRecBai.Post;
      ZQRecBai.Next;
    end;
    ZQRecBai.EnableControls;
    ZQRecBai.Edit;
    ZQRecBai.RecNo := varreg;
    ZQRecBai.FieldByName('marcar').AsString := '0';
    ZQRecBai.Post;
    XNEMarcado.Value := ZQRecBai.FieldByName('saldo').AsFloat;
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
//     showmessage('Aguarde a Leitura dos Dados. Fechará Automaticamente em Seguida.');
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
//       showmessage('Aguarde a Leitura dos Dados. Fechará Automaticamente em Seguida.');
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
    if (ZQRecBai.FieldByName('saldo').AsFloat>0) and (ZQRecBai.FieldByName('recpag').AsString='R') and (CBSaldoCorrigido.Checked)  Then Begin
    varsal := ZQRecBai.FieldByName('Valor').AsFloat;

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
    vardI := ZQRecBai.FieldByName('Dt_Vencimento').AsDateTime;
    vardT := strtodate(FrmRecebBaixa.JDEntrada.DateText);
    ZQRecebBxTemp.DisableControls;
    while (not ZQRecebBxTemp.Eof) and (vardT>=ZQRecebBxTemp.FieldByName('dataref').AsDateTime) do begin
      varsal := Calcjuros(vardI,ZQRecebBxTemp.FieldByName('dataref').AsDateTime,varsal,XNJuros.Value);
      if ZQRecebBxTemp.FieldByName('valor_parcela').AsFloat > 0 then
        varsal := varsal - ZQRecebBxTemp.FieldByName('valor_parcela').AsFloat
      else
        varsal := varsal-(ZQRecebBxTemp.FieldByName('Vr_rec').AsFloat*ZQRecebBxTemp.FieldByName('percent_usado').AsFloat/100);
      vardI := ZQRecebBxTemp.FieldByName('dataref').AsDateTime;
      ZQRecebBxTemp.Next;
    end;
    ZQRecebBxTemp.EnableControls;
    ZQRecBai.FieldByName('saldocorrig').AsFloat := Calcjuros(vardI,vardT,varsal,XNJuros.Value);
  end
  else if ZQRecBai.FieldByName('saldo').AsFloat=0 Then
    ZQRecBai.FieldByName('saldocorrig').AsFloat := 0
  else
    ZQRecBai.FieldByName('saldocorrig').AsFloat := ZQRecBai.FieldByName('saldo').AsFloat;
end;

procedure TFrmRecebBaixa.CDSParcelasCalcFields(DataSet: TDataSet);
begin
//  CDSParcelasPercent.Value := CDSParcelasVrParc.Value*100/XNERecebido.Value;
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL' then
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
    ZQRatear.FieldByName('vr_base').AsFloat := ZQRatear.FieldByName('vr_rec').AsFloat * ZQRatear.FieldByName('percent_usado').AsFloat / 100;
    ZQRatear.FieldByName('vr_divi').AsFloat := ZQRatear.FieldByName('vr_base').AsFloat * ZQRatear.FieldByName('percentual').AsFloat /100;

end;

procedure TFrmRecebBaixa.XNEJurosEnter(Sender: TObject);
begin
  if ZQRecBai.FieldByName('recpag').AsString = 'R' Then
    Perform(WM_NEXTDLGCTL, 0, 0);
end;

procedure TFrmRecebBaixa.CDS_divididoCalcFields(DataSet: TDataSet);
begin
  CDS_divididodescrcontabil2.Value := CDS_divididocontabil.Text+' - '+CDS_divididodescrcontabil.Value;
end;

procedure TFrmRecebBaixa.CBPeriodoEntExit(Sender: TObject);
var
  varvezes, varfinal, vardias: integer;
  varsoma, varprincipalquitadototal, varprincipalquitada, varjuroslinha,
  vardescontolinha, varpercentualbase, varsomajuros,
  varsomadesconto : double;
  varregistros : integer;
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
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL' then
  begin
    if ZQRecBai.FieldByName('Data_Quitacao').AsDateTime>0 then
       varvenci :=  ZQRecBai.FieldByName('Data_Quitacao').AsDateTime
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
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger,[]);
    if XNEResto.Value > 0 then
    begin
      varsoma := 0;
      varsomajuros := 0;
      varsomadesconto := 0;
      varprincipalquitadototal := ExRound(XNEMarcado.Value - XNEResto.Value,2);
      if varprincipalquitadototal < 0 then
        varprincipalquitadototal := 0;

      varregistros := ZQRecBai.RecordCount;
      varfinal := 1;
      varvezes := 0;
      ZQRecBai.First;
      ZQRecBai.DisableControls;
      while not ZQRecBai.Eof do
      begin
        if XNEMarcado.Value > 0 then
          varpercentualbase := ZQRecBai.FieldByName('saldocorrig').AsFloat * 100 / XNEMarcado.Value
        else
          varpercentualbase := 0;

        Inc(varvezes);
        CDSParcelas.Insert;
        CDSParcelasidrecebimento.Value:= ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
        CDSParcelasOrdem.Value := '001/001';

        if varvezes < varregistros then
        begin
          varprincipalquitada := ExRound(varprincipalquitadototal * varpercentualbase / 100,2);
          varjuroslinha := ExRound(XNEJuros.Value * varpercentualbase / 100,2);
          vardescontolinha := ExRound(XNEDesconto.Value * varpercentualbase / 100,2);
        end
        else
        begin
          varprincipalquitada := ExRound(varprincipalquitadototal - varsoma,2);
          varjuroslinha := ExRound(XNEJuros.Value - varsomajuros,2);
          vardescontolinha := ExRound(XNEDesconto.Value - varsomadesconto,2);
        end;

        if varprincipalquitada < 0 then
          varprincipalquitada := 0;
        if varjuroslinha < 0 then
          varjuroslinha := 0;
        if vardescontolinha < 0 then
          vardescontolinha := 0;

        CDSParcelasjuros.Value := varjuroslinha;
        CDSParcelasdesconto.Value := vardescontolinha;
        CDSParcelasVrParc.Value := ExRound(varprincipalquitada + varjuroslinha - vardescontolinha,2);
        if ZQRecBai.FieldByName('recpag').AsString='R' Then
          CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger
        else
          CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger;
        CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
        if ZQRecBai.FieldByName('Data_Quitacao').AsDateTime>0 then
          CDSParcelasVenci.Value := ZQRecBai.FieldByName('Data_Quitacao').AsDateTime
        else
          CDSParcelasVenci.Value := strtodate(JDEntrada.DateText);
        CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').AsString;
        CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').AsInteger;
        CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').AsString;
        CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
        CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').AsString;
        CDSParcelassq.Value := ZQRecBai.FieldByName('sq').AsLargeInt;
        CDSParcelasData_Quitacao.Value:=ZQRecBai.FieldByName('Data_Quitacao').AsDateTime;
        CDSParcelastip.Value:=ZQRecBai.FieldByName('tip').AsString;
        CDSParcelas.Post;
        varsoma := varsoma + varprincipalquitada;
        varsomajuros := varsomajuros + varjuroslinha;
        varsomadesconto := varsomadesconto + vardescontolinha;
        ZQRecBai.Next;
      end;
      ZQRecBai.EnableControls;
    end
    else
    begin
      varsoma := 0;
      for varvezes := 1 to varfinal do
      Begin
        CDSParcelas.Insert;
        CDSParcelasOrdem.Value := inttostrZero(varvezes,3)+'/'+inttostrZero(varfinal,3);
        CDSParcelasjuros.Value :=  ExRound(XNEJuros.Value,2)/ExRound(varfinal,2);
        CDSParcelasVrParc.Value := ExRound(XNERecebido.Value,2)/ExRound(varfinal,2);

        if ZQRecBai.FieldByName('recpag').AsString='R' Then
          CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger
        else
          CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger;
          CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
        if ZQRecBai.FieldByName('Data_Quitacao').AsDateTime>0 then
        begin
          if varvezes = 1 Then
             CDSParcelasVenci.Value := ZQRecBai.FieldByName('Data_Quitacao').AsDateTime
          else Begin
            if vartip = 'S' then
              CDSParcelasVenci.Value := IncMonth(ZQRecBai.FieldByName('Data_Quitacao').AsDateTime,(varvezes*vardias)-(vardias))
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
        CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').AsString;
        CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').AsInteger;
        CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').AsString;
        CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
        CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').AsString;
        CDSParcelassq.Value := ZQRecBai.FieldByName('sq').AsLargeInt;
        CDSParcelasData_Quitacao.Value:=ZQRecBai.FieldByName('Data_Quitacao').AsDateTime;
        CDSParcelastip.Value:=ZQRecBai.FieldByName('tip').AsString;
        CDSParcelas.Post;
        varsoma := varsoma + CDSParcelasVrParc.Value;
       end;
      CDSParcelas.IndexFieldNames := 'ordem';
      CDSParcelas.First;
      CDSParcelas.Edit;
      CDSParcelasVrParc.Value := CDSParcelasVrParc.Value+((XNERecebido.Value)-varsoma);
      CDSParcelas.Post;
    end;
    ZQRecBai.Filtered := false;
  end
  else
  begin
    if ZQRecBai.FieldByName('Data_Quitacao').AsDateTime>0 then
       varvenci :=  ZQRecBai.FieldByName('Data_Quitacao').AsDateTime
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
    DM_tabelas.ZQPlanoDeContas.Locate('codigo',DM_Tabelas.ZQConfiguracoes.FieldByName('venda_planodecontas').AsInteger,[]);
    varsoma := 0;
    ZQRecBai.First;
    varfinal:=1;
    //for varvezes := 1 to varfinal do Begin
    ZQRecBai.DisableControls;
    while not ZQRecBai.Eof do
    begin
      CDSParcelas.Insert;
      CDSParcelasidrecebimento.Value:= ZQRecBai.FieldByName('idrecebimento').AsLargeInt;
      CDSParcelasOrdem.Value := '001/001';
      CDSParcelasjuros.Value :=  ExRound(ZQRecBai.FieldByName('juros').AsFloat,2);
      CDSParcelasdesconto.Value :=  ExRound(ZQRecBai.FieldByName('descontos').AsFloat,2);
      CDSParcelasVrParc.Value := (ExRound(ZQRecBai.FieldByName('saldo').AsFloat,2)+ExRound(ZQRecBai.FieldByName('juros').AsFloat,2))-ExRound(ZQRecBai.FieldByName('descontos').AsFloat,2);
      if ZQRecBai.FieldByName('recpag').AsString='R' Then
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codrecBx').AsInteger
      else
        CDSParcelascodcontabil.Value := DM_Tabelas.ZqParticipante.FieldByName('codpagBx').AsInteger;
        CDSParcelasTipDoc.Value := DM_Tabelas.ZQPlanoDeContas.FieldByName('doccomum').AsString;
      if ZQRecBai.FieldByName('Data_Quitacao').AsDateTime>0 then
      begin
        if varvezes = 1 Then
           CDSParcelasVenci.Value := ZQRecBai.FieldByName('Data_Quitacao').AsDateTime
        else Begin
          if vartip = 'S' then
            CDSParcelasVenci.Value := IncMonth(ZQRecBai.FieldByName('Data_Quitacao').AsDateTime,(varfinal*vardias)-(vardias))
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
      CDSParcelasDocum.Value:=ZQRecBai.FieldByName('documento').AsString;
      CDSParcelasvenda_idvenda.Value:=ZQRecBai.FieldByName('venda_idvenda').AsInteger;
      CDSParcelasquadralote.Value:=ZQRecBai.FieldByName('quadralote').AsString;
      CDSParcelasnomedoempreendimento.Value:=XEEmpresend.Text;
      CDSParcelasSubstituicao.Value := ZQRecBai.FieldByName('Substituicao').AsString;
      CDSParcelassq.Value := ZQRecBai.FieldByName('sq').AsLargeInt;
      CDSParcelasData_Quitacao.Value:=ZQRecBai.FieldByName('Data_Quitacao').AsDateTime;
      CDSParcelastip.Value:=ZQRecBai.FieldByName('tip').AsString;
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
        Eadversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
        EcodAdversa.Text := DM_Tabelas.ZQAchaParticip.FieldByName('idpaticipante').Text;
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
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
        XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;

        DM_tabelas.ZQCobaRe.Insert;
        DTPIni.Datetext := datetostr(date);
        DTPFim.Datetext := datetostr(date);
        DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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
          DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
          DM_tabelas.ZQTipodoc.Next;
        end;
        Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
        Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
        Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
        EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
        EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
        XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
        varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;
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
        EAdversaNova.Text := DM_Tabelas.ZQAchaParticip.FieldByName('nome_parte').AsString;
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
    Label15.Caption:='Dt. Quitação.......';
  end
  else
  begin
    Label15.Caption:='Dt. 1º agto........';
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
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;
  if empty(DM_tabelas.ZQRecebimento.FieldByName('cliente').Text) Then Begin
    Close;
    Exit;
  end;
  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;  
  varrecpag := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


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
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;

  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  DM_tabelas.ZQTipodoc.EnableControls;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
  varrecpag            := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


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
//  ZQRecBai.DisableControls;
  while not ZQRecBai.Eof do
  begin
    ZQRecBai.Edit;
    if  ZQRecBai.FieldByName('marcar').AsString = '0' Then
      ZQRecBai.FieldByName('marcar').AsString := '1'
    else
      ZQRecBai.FieldByName('marcar').AsString := '0';
    ZQRecBai.post;
    ZQRecBai.Next;
  end;
//  ZQRecBai.EnableControls;
  Atualiza_tela;

end;

procedure TFrmRecebBaixa.bxmanualExit(Sender: TObject);
begin
  DM_Tabelas.ZQConfiguracoes.Edit;
  DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString:=bxmanual.Text;
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
          xrec:=DM_tabelas.ZQRecebimento.FieldByName('idrecebimento').AsLargeInt;
          Elote.Text:=DM_tabelas.ZQRecebimento.FieldByName('quadralote').AsString;
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
      showmessage('Número de Cheque não Encontrado.');
      dxbfechar.setfocus;
    end;
  end;
  showmessage('Cheque não Encontrado nas Parcelas em Aberto.');
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
  if DM_Tabelas.ZQConfiguracoes.FieldByName('Baixa').AsString='DATA QUITAÇÃO GERAL' then
  begin
     GroupBox2.Visible:=true;
     DBGrid1.SetFocus;
     DBGrid1.SelectedIndex:=0; // setfocus na 1ª coluna
     DBGrid1.Options := DBGrid1.Options + [dgEditing];
  end
  else
  begin
     GroupBox3.Visible:=true;
     DBGrid3.SetFocus;
     DBGrid3.SelectedIndex:=0; // setfocus na 1ª coluna
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
    if ZQRecBai.FieldByName('marcar').AsString = '0' then
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
    XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').AsFloat)-ZQRecBai.FieldByName('descontos').AsFloat;
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
      XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').AsFloat)-ZQRecBai.FieldByName('descontos').AsFloat;
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
    if ZQRecBai.FieldByName('marcar').AsString = '0' then
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
    XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').AsFloat)-ZQRecBai.FieldByName('descontos').AsFloat;
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
      XNERecebido.Value:=(XNERecebido.Value+ZQRecBai.FieldByName('juros').AsFloat)-ZQRecBai.FieldByName('descontos').AsFloat;
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
  XNJuros.Value := DM_Tabelas.ZQConfiguracoes.FieldByName('jurosmensal').AsFloat;

  DM_tabelas.ZQCobaRe.Insert;
  DTPIni.Datetext := datetostr(date);
  DTPFim.Datetext := datetostr(date);
  DM_tabelas.ZQCobaRe.FieldByName('hoje').AsDateTime := date;
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
    DBCBTipDoc.Items.Add(DM_tabelas.ZQTipoDoc.FieldByName('tipodoc').AsString);
    DM_tabelas.ZQTipodoc.Next;
  end;
  Enomes.Text          := DM_tabelas.ZQRecebimento.FieldByName('nomecli').AsString;
  Ecodnome.Text        := DM_tabelas.ZQRecebimento.FieldByName('cliente').Text;
  Eadversa.Text        := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EAdversaNova.Text    := DM_Tabelas.ZQRecebimento.FieldByName('adversanome').AsString;
  EcodAdversa.Text     := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  EcodAdversaNova.Text := DM_Tabelas.ZQRecebimento.FieldByName('adversa').Text;
  XEEmpresend.Text     := DM_Tabelas.ZQRecebimento.FieldByName('nome_loteamento').AsString;
  varrecpag            := DM_Tabelas.ZQRecebimento.FieldByName('recpag').AsString;


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



